"""
Reconstrói a base da escp771 a partir de várias extrações ("fotos") do relatório de programações em aberto.

Uso:
    python reconstruir_escp771.py maio.xlsx [junho.xlsx ...] pasta_de_saida/

Cada aba de cada arquivo é tratada como uma extração. Todas as extrações de todos os arquivos são
empilhadas e ordenadas pelo horário do título; por isso, ao incluir um novo mês, basta passar todos
os arquivos juntos: registros que atravessam a virada do mês são unidos pela mesma chave.

Fatos sobre o relatório que guiam as regras (observados em maio/2026):
  * Chave do registro = (COD PROGR, SUB-ITEM): é única dentro de cada extração.
  * Campos de cadastro (máquina, quantidade, condutor, terminais, comprimento...) não mudam entre extrações.
  * QUANT CORTADA é ACUMULADA; já MC REALIZADO, início/fim e RE mostram só a ÚLTIMA sessão de corte.
    Uma ordem cortada em 2 ou 3 vezes aparece no relatório final com a quantidade total e o horário
    só da última sessão. As sessões anteriores só existem nas extrações antigas -> são recuperadas aqui.
  * QUANT PROD é sempre igual a QUANT CORTADA (redundante).
  * Algumas abas gravam as datas como número de série do Excel (ex.: 46140 = 28/04/2026): convertidas.
  * Colunas sem cabeçalho contêm anotações manuais (ex.: número de máquina, '*', 'PRIORIZAR'): preservadas.
  * O relatório "CORTE AUTOMOTIVA EM ABERTO" é outro universo (programações 459xxx, com CLIENTE).
"""
import re, sys, datetime as dt
from pathlib import Path
import numpy as np
import pandas as pd
from openpyxl import load_workbook

LACUNA_H = 24          # intervalo entre extrações acima do qual há risco de sessões não observadas
DUR_MAX_H = 24         # sessão acima disso (já descontadas as paradas de calendário) recebe flag
# Paradas de calendário: a fábrica não opera aos domingos (a base horária da MC 68 não tem nenhuma hora de domingo)
# nem nos feriados nacionais abaixo. Ajuste a lista se a fábrica trabalhar em algum deles ou parar em outros dias.
FERIADOS = {"2026-01-01", "2026-04-03", "2026-04-21", "2026-05-01", "2026-09-07", "2026-10-12", "2026-11-02",
            "2026-11-15", "2026-11-20", "2026-12-25"}
K = ["cod_prog", "sub_item"]

MAPA = {  # cabeçalho do relatório -> nome canônico
    "MÁQUINA": "maquina_prog", "MAQUINA": "maquina_prog", "COD PROGR": "cod_prog", "DATA PROGR": "data_prog",
    "DATA CRIAÇÃO": "data_criacao", "HORA CRIAÇÃO": "hora_criacao", "CLIENTE": "cliente", "KIT": "kit", "ITEM": "item",
    "SUB-ITEM": "sub_item", "QUANTIDADE": "quantidade", "QUANT PROD": "quant_prod", "CORTE": "corte", "TEMPO": "tempo_prog_min",
    "CONDUTOR": "condutor", "TERMINAL A": "terminal_a", "MA LADO A": "ma_lado_a", "TERMINAL B": "terminal_b",
    "MA LADO B": "ma_lado_b", "COMPR": "comprimento", "COR CABO": "cor_cabo", "ESTOQUE MPB CONDUTOR": "estoque_condutor",
    "ESTOQUE MPB TERMINAL A": "estoque_terminal_a", "ESTOQUE MPB TERMINAL B": "estoque_terminal_b",
    "MC REALIZADO": "mc_realizado", "DATA INI CORTE": "data_ini", "HORA INI CORTE": "hora_ini", "RE INI CORTE": "re_ini",
    "DATA FIM CORTE": "data_fim", "HORA FIM CORTE": "hora_fim", "RE FIM CORTE": "re_fim", "QUANT CORTADA": "quant_cortada"}
TEXTO = ["kit", "item", "sub_item", "corte", "condutor", "terminal_a", "ma_lado_a", "terminal_b", "ma_lado_b", "cor_cabo",
         "mc_realizado", "re_ini", "re_fim", "cliente"]
NUM = ["maquina_prog", "cod_prog", "quantidade", "quant_prod", "tempo_prog_min", "comprimento", "estoque_condutor",
       "estoque_terminal_a", "estoque_terminal_b", "quant_cortada"]
DATAS = ["data_prog", "data_criacao", "data_ini", "data_fim"]
HORAS = ["hora_criacao", "hora_ini", "hora_fim"]
CADASTRO = ["maquina_prog", "data_prog", "data_criacao", "hora_criacao", "cliente", "kit", "item", "quantidade", "tempo_prog_min",
            "condutor", "terminal_a", "ma_lado_a", "terminal_b", "ma_lado_b", "comprimento", "cor_cabo",
            "estoque_condutor", "estoque_terminal_a", "estoque_terminal_b"]


def txt(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None
    if isinstance(v, (float, np.floating)) and float(v).is_integer():
        v = int(v)
    s = str(v).strip()
    if s.startswith('="') and s.endswith('"'):
        s = s[2:-1].strip()
    return s or None


def data_excel(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return pd.NaT
    if isinstance(v, (int, float, np.integer, np.floating)):          # número de série do Excel
        return pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(v))
    if isinstance(v, str):
        return pd.to_datetime(v, dayfirst=True, errors="coerce")
    return pd.Timestamp(v)


def min_parada_calendario(a, b):
    """Minutos do intervalo [a, b] que caem em domingo ou feriado (fábrica parada)."""
    if pd.isna(a) or pd.isna(b) or b <= a:
        return 0.0
    t, d = 0.0, a.normalize()
    while d <= b:
        if d.dayofweek == 6 or d.strftime("%Y-%m-%d") in FERIADOS:
            ini, fim = max(a, d), min(b, d + pd.Timedelta(days=1))
            if fim > ini: t += (fim - ini).total_seconds() / 60
        d += pd.Timedelta(days=1)
    return t


def hora(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None
    if isinstance(v, dt.time):
        return v.strftime("%H:%M:%S")
    if isinstance(v, dt.datetime):
        return v.strftime("%H:%M:%S")
    if isinstance(v, (float, np.floating)) and 0 <= v < 1:             # fração do dia
        s = int(round(v * 86400)); return f"{s // 3600:02d}:{s % 3600 // 60:02d}:{s % 60:02d}"
    return txt(v)


# ------------------------------------------------------------------ 1) leitura de todas as abas
def ler(arquivos):
    extr, linhas, anot = [], [], []
    for arq in arquivos:
        wb = load_workbook(arq, read_only=True, data_only=True)
        for ws in wb.worksheets:
            rows = list(ws.iter_rows(values_only=True))
            hi = next((i for i, r in enumerate(rows[:20]) if any(str(v).strip().upper() in ("MÁQUINA", "MAQUINA") for v in r if v is not None)), None)
            if hi is None:
                extr.append(dict(arquivo=Path(arq).name, aba=ws.title, titulo=None, relatorio=None, extracao=pd.NaT, linhas=0, obs="sem cabeçalho reconhecível; ignorada"))
                continue
            titulo = next((" ".join(str(v) for v in r if v is not None) for r in rows[:hi] if any("BRASCABOS" in str(v).upper() for v in r if v is not None)), None)
            m = re.search(r"(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}:\d{2}:\d{2})", titulo or "")
            relatorio = "automotiva" if titulo and "AUTOMOTIVA" in titulo.upper() else "corte"
            hdr = [(str(v).strip() or None) if v is not None else None for v in rows[hi]]   # célula vazia ou "" = coluna de anotação
            idx = {MAPA[h]: j for j, h in enumerate(hdr) if h in MAPA}
            sem_nome = [j for j, h in enumerate(hdr) if h is None]
            n = 0
            for r in rows[hi + 1:]:
                if r is None or all(v is None for v in r):
                    continue
                d = {c: (r[j] if j < len(r) else None) for c, j in idx.items()}
                d["arquivo"], d["aba"] = Path(arq).name, ws.title
                linhas.append(d); n += 1
                for j in sem_nome:
                    if j < len(r) and r[j] is not None:
                        anot.append(dict(arquivo=Path(arq).name, aba=ws.title, cod_prog=d.get("cod_prog"), sub_item=d.get("sub_item"), coluna=j + 1, valor=txt(r[j])))
            extr.append(dict(arquivo=Path(arq).name, aba=ws.title, titulo=titulo, relatorio=relatorio,
                             extracao=pd.to_datetime(m.group(1) + " " + m.group(2), format="%d/%m/%Y %H:%M:%S") if m else pd.NaT,
                             linhas=n, colunas=len([h for h in hdr if h]), colunas_ausentes=", ".join(sorted(set(MAPA.values()) - set(idx) - {"maquina_prog"})),
                             colunas_sem_nome=len(sem_nome), obs=""))
    return pd.DataFrame(extr), pd.DataFrame(linhas), pd.DataFrame(anot)


def normalizar(L):
    for c in TEXTO:
        if c in L: L[c] = L[c].map(txt)
    for c in NUM:
        if c in L: L[c] = pd.to_numeric(L[c].map(txt), errors="coerce")
    for c in DATAS:
        if c in L: L[c] = L[c].map(data_excel)
    for c in HORAS:
        if c in L: L[c] = L[c].map(hora)
    for c in set(MAPA.values()) - set(L.columns):
        L[c] = np.nan
    L["kit"] = L["kit"].str.upper()
    L["ini"] = pd.to_datetime(L["data_ini"].dt.strftime("%Y-%m-%d") + " " + L["hora_ini"].fillna(""), errors="coerce")
    L["fim"] = pd.to_datetime(L["data_fim"].dt.strftime("%Y-%m-%d") + " " + L["hora_fim"].fillna(""), errors="coerce")
    return L


# ------------------------------------------------------------------ 1b) extrações sem a coluna SUB-ITEM
CHAVE_CADASTRO = ["cod_prog", "kit", "quantidade", "condutor", "terminal_a", "terminal_b", "comprimento", "cor_cabo"]
ESTADO = ["mc_realizado", "ini", "fim", "quant_cortada"]


def recuperar_sub_item(L, E):
    """Para abas exportadas sem SUB-ITEM: identifica cada linha pelos campos de cadastro, usando as extrações
    vizinhas (anterior e seguinte do mesmo relatório). Grupos com cadastro idêntico são desempatados pelo estado
    de corte; se as linhas do grupo têm o mesmo estado, a atribuição não muda nenhum valor. O que não puder ser
    resolvido sem ambiguidade é descartado e registrado."""
    L = L.copy(); L["sub_item_recuperado"] = False
    alvo = E[(E.linhas > 0) & E.colunas_ausentes.fillna("").str.split(", ").apply(lambda c: "sub_item" in c)]
    stats = []
    sig = lambda df, cols: df[cols].astype(object).apply(lambda c: c.map(lambda v: "" if pd.isna(v) else str(v))).agg("|".join, axis=1)
    for r in alvo.itertuples():
        viz = E[(E.relatorio == r.relatorio) & (E.ordem != r.ordem) & ~E.colunas_ausentes.fillna("").str.contains(r"\bsub_item\b")]
        ant, pos = viz[viz.ordem < r.ordem].tail(1), viz[viz.ordem > r.ordem].head(1)
        vz = pd.concat([ant, pos])
        P = L.merge(vz[["arquivo", "aba"]], on=["arquivo", "aba"])            # linhas das vizinhas
        P = P[P.sub_item.notna()]
        m = (L.arquivo == r.arquivo) & (L.aba == r.aba)
        X = L[m].copy()
        X["_c"] = sig(X, CHAVE_CADASTRO); P["_c"] = sig(P, CHAVE_CADASTRO)
        cand = P.groupby("_c").sub_item.apply(lambda s: sorted(set(s)))
        X["_e"] = sig(X, ESTADO); P["_e"] = sig(P, ESTADO)
        resolvido = {}
        for c, g in X.groupby("_c"):
            subs = cand.get(c, [])
            if len(subs) == 1 and len(g) == 1:
                resolvido[g.index[0]] = subs[0]; continue
            if len(subs) != len(g):                                          # sem correspondência clara
                continue
            livres, pend = list(subs), []
            for i, row in g.iterrows():                                      # desempate pelo estado de corte
                ok = P[(P._c == c) & (P._e == row._e) & P.sub_item.isin(livres)].sub_item.unique()
                if len(ok) == 1:
                    resolvido[i] = ok[0]; livres.remove(ok[0])
                else:
                    pend.append(i)
            if pend and g.loc[pend, "_e"].nunique() == 1:                   # estados idênticos: atribuição indiferente
                for i, sbi in zip(pend, sorted(livres)):
                    resolvido[i] = sbi
        idx = list(resolvido)
        L.loc[idx, "sub_item"] = [resolvido[i] for i in idx]
        L.loc[idx, "sub_item_recuperado"] = True
        stats.append(dict(arquivo=r.arquivo, aba=r.aba, linhas=int(m.sum()), recuperadas=len(idx),
                          vizinhas=" e ".join(vz.aba)))
    return L, pd.DataFrame(stats)


# ------------------------------------------------------------------ 2) ordem das extrações
def ordenar_extracoes(E, L):
    E = E.copy()
    E["extracao_estimada"] = False
    for i, r in E[E.extracao.isna() & (E.linhas > 0)].iterrows():   # sem horário no título: limite inferior = último corte registrado
        g = L[(L.arquivo == r.arquivo) & (L.aba == r.aba)]
        E.at[i, "extracao"] = max(g["fim"].max(), g["ini"].max())
        E.at[i, "extracao_estimada"] = True
        E.at[i, "obs"] = "título sem data/hora: horário estimado pelo último corte registrado (limite inferior)"
    E = E.sort_values(["extracao", "arquivo", "aba"]).reset_index(drop=True)
    E["ordem"] = np.arange(len(E))
    for rel in E.relatorio.dropna().unique():
        m = E.relatorio == rel
        E.loc[m, "horas_desde_anterior"] = E.loc[m, "extracao"].diff().dt.total_seconds() / 3600
    E["lacuna_antes"] = E["horas_desde_anterior"] > LACUNA_H
    return E


# ------------------------------------------------------------------ 3) reconstrução
def reconstruir(L, E):
    L = L.merge(E[["arquivo", "aba", "extracao", "ordem", "relatorio"]], on=["arquivo", "aba"])
    L = L.sort_values("ordem")
    ultima = E.groupby("relatorio").ordem.max().to_dict()
    prox = {}
    for rel, g in E.groupby("relatorio"):                       # extração seguinte (para medir lacunas)
        o = g.sort_values("ordem")
        prox.update(dict(zip(o.ordem, o.extracao.shift(-1))))
    primeira_extr = E.groupby("relatorio").extracao.min().to_dict()
    ordens_rel = {rel: g.ordem.values for rel, g in E.groupby("relatorio")}
    ordens, sessoes = [], []
    for (cp, si), g in L.groupby(K, sort=False):
        rel = g.relatorio.iloc[-1]
        o = {"cod_prog": cp, "sub_item": si, "relatorio": rel}
        for c in CADASTRO:                                      # cadastro: último valor não vazio
            s = g[c].dropna(); o[c] = s.iloc[-1] if len(s) else np.nan
        o["ma_lado_a_mudou"] = g.ma_lado_a.nunique() > 1
        o["ma_lado_b_mudou"] = g.ma_lado_b.nunique() > 1
        o["primeira_vista"], o["ultima_vista"] = g.extracao.min(), g.extracao.max()
        o["n_extracoes"] = g.ordem.nunique()
        o["arquivos"] = " | ".join(sorted(g.arquivo.unique()))
        no_intervalo = ((ordens_rel[rel] >= g.ordem.min()) & (ordens_rel[rel] <= g.ordem.max())).sum()
        o["ausente_em_extracoes_intermediarias"] = int(no_intervalo - o["n_extracoes"])   # sumiu e reapareceu
        o["presente_na_ultima"] = g.ordem.max() == ultima[rel]
        nx = prox.get(g.ordem.max())
        o["horas_ate_extracao_seguinte"] = (nx - g.extracao.max()).total_seconds() / 3600 if pd.notna(nx) else np.nan
        o["quant_cortada_final"] = g.quant_cortada.iloc[-1]
        o["corte_final"] = g.corte.iloc[-1]
        # sessões = valores distintos de início observados
        obs = g[g.ini.notna()]
        ses = []
        for ini, s in obs.groupby("ini", sort=True):
            cf = s[s.fim.notna()]
            antes = g[g.ordem < s.ordem.min()]
            ses.append(dict(
                cod_prog=cp, sub_item=si, relatorio=rel, ini=ini,
                fim=cf.fim.iloc[-1] if len(cf) else pd.NaT,
                mc_realizado=s.mc_realizado.dropna().iloc[-1] if s.mc_realizado.notna().any() else None,
                re_ini=s.re_ini.dropna().iloc[-1] if s.re_ini.notna().any() else None,
                re_fim=cf.re_fim.dropna().iloc[-1] if len(cf) and cf.re_fim.notna().any() else None,
                quant_acum_fim=cf.quant_cortada.max() if len(cf) else np.nan,
                quant_acum_vista=s.quant_cortada.max(),
                quant_acum_antes=antes.quant_cortada.max() if len(antes) else np.nan,
                primeira_vista=s.extracao.min(), ultima_vista=s.extracao.max(), n_extracoes=s.ordem.nunique(),
                horas_desde_obs_anterior=(s.extracao.min() - antes.extracao.max()).total_seconds() / 3600 if len(antes) else np.nan,
                antes_da_primeira_extracao=ini < primeira_extr[rel]))
        prev = np.nan
        for k, x in enumerate(ses):
            x["n_sessao"] = k + 1
            base = prev if k > 0 else x["quant_acum_antes"]
            q = x["quant_acum_fim"] if pd.notna(x["quant_acum_fim"]) else x["quant_acum_vista"]
            x["incremento"] = q - base if pd.notna(base) else q
            x["incremento_confiavel"] = bool(pd.notna(base) and pd.notna(x["quant_acum_fim"]))
            prev = q if pd.notna(q) else prev
            x["dur_min"] = (x["fim"] - x["ini"]).total_seconds() / 60 if pd.notna(x["fim"]) else np.nan
            x["min_parada_calendario"] = min_parada_calendario(x["ini"], x["fim"])
            x["dur_util_min"] = x["dur_min"] - x["min_parada_calendario"] if pd.notna(x["dur_min"]) else np.nan
            x["atravessa_parada"] = x["min_parada_calendario"] > 0
            x["mes_sessao"] = x["ini"].strftime("%Y-%m")
            x["sessao_sem_fim"] = pd.isna(x["fim"])
            x["dur_mais_24h"] = bool(x["dur_util_min"] > DUR_MAX_H * 60) if pd.notna(x["dur_util_min"]) else False
            x["possivel_sessao_oculta"] = bool(pd.notna(x["horas_desde_obs_anterior"]) and x["horas_desde_obs_anterior"] > LACUNA_H and k > 0)
            x["pecas_por_min"] = x["incremento"] / x["dur_util_min"] if x["incremento_confiavel"] and pd.notna(x["dur_util_min"]) and x["dur_util_min"] > 0 else np.nan
        sessoes += ses
        o["n_sessoes"] = len(ses)
        o["mc_realizado_final"] = ses[-1]["mc_realizado"] if ses else None
        o["mcs_realizado"] = " / ".join(dict.fromkeys(x["mc_realizado"] for x in ses if x["mc_realizado"])) if ses else None
        o["primeiro_ini"] = ses[0]["ini"] if ses else pd.NaT
        o["ultimo_fim"] = max((x["fim"] for x in ses if pd.notna(x["fim"])), default=pd.NaT)
        o["min_em_sessoes"] = np.nansum([x["dur_min"] for x in ses]) if ses else np.nan
        o["min_uteis_em_sessoes"] = np.nansum([x["dur_util_min"] for x in ses]) if ses else np.nan
        o["ini_relatorio_final"] = g.ini.iloc[-1]                     # o que uma foto única mostraria
        o["fim_relatorio_final"] = g.fim.iloc[-1]
        cortou = (g.quant_cortada.fillna(0) > 0).any() or len(ses) > 0
        if o["quantidade"] == 0:
            o["status"] = "qtd_zero"                          # linha com quantidade programada 0: o sistema já marca "Completo"
        else:
            o["status"] = ("completo" if o["corte_final"] == "Completo" else "parcial") if cortou else "nao_cortado"
        o["completo_sem_apontamento"] = bool(o["corte_final"] == "Completo" and not ses and o["quantidade"] > 0)
        o["excesso_corte"] = bool(pd.notna(o["quant_cortada_final"]) and o["quant_cortada_final"] > o["quantidade"])
        o["sumiu_sem_corte"] = bool(not o["presente_na_ultima"] and not cortou and o["quantidade"] > 0)
        o["sumiu_apos_lacuna"] = bool(o["sumiu_sem_corte"] and pd.notna(o["horas_ate_extracao_seguinte"]) and o["horas_ate_extracao_seguinte"] > LACUNA_H)
        o["sessao_sem_fim_na_ultima"] = bool(ses and pd.isna(ses[-1]["fim"]))
        o["alguma_sessao_mais_24h"] = any(x["dur_mais_24h"] for x in ses)
        o["mes_ultimo_corte"] = o["ultimo_fim"].strftime("%Y-%m") if pd.notna(o["ultimo_fim"]) else None
        ordens.append(o)
    return pd.DataFrame(ordens), pd.DataFrame(sessoes)


DICIONARIO = [
    ("ordens", "cod_prog + sub_item", "Chave do registro (única em cada extração)."),
    ("ordens", "relatorio", "corte = DETALHE PROGRAMAÇÕES EM ABERTO CORTE; automotiva = CORTE AUTOMOTIVA EM ABERTO (outro universo de programações)."),
    ("ordens", "campos de cadastro", "maquina_prog, quantidade, condutor, terminais, MA, comprimento, cor, tempo, estoques, kit, item, datas: último valor não vazio (não variam entre extrações, exceto MA em poucos casos)."),
    ("ordens", "status", "completo / parcial (cortado, mas CORTE = Falta) / nao_cortado (nunca teve quantidade nem sessão) / qtd_zero (quantidade programada 0; o sistema marca Completo automaticamente — excluir das análises)."),
    ("ordens", "n_sessoes, mcs_realizado", "Quantas sessões de corte foram observadas e em quais máquinas (na ordem)."),
    ("ordens", "primeiro_ini, ultimo_fim, min_em_sessoes, min_uteis_em_sessoes", "Início da 1ª sessão, fim da última e soma das durações (bruta e útil) das sessões. Use estes no lugar do ini/fim do relatório final."),
    ("ordens", "ini_relatorio_final, fim_relatorio_final", "O que uma única extração final mostraria (só a última sessão). Mantidos para comparação."),
    ("ordens", "presente_na_ultima", "Ainda em aberto na última extração do período (continua no mês seguinte)."),
    ("ordens", "sumiu_sem_corte, sumiu_apos_lacuna", "Saiu do relatório sem nunca ter corte registrado; 'após lacuna' = a extração seguinte veio mais de 24 h depois (pode ter sido cortado e encerrado sem ser visto)."),
    ("ordens", "arquivos, ausente_em_extracoes_intermediarias", "Arquivos mensais em que a ordem aparece; nº de extrações, entre a primeira e a última em que foi vista, nas quais ela não apareceu (ex.: extração de 19/06 13:21 trouxe programações já encerradas)."),
 ("ordens", "completo_sem_apontamento", "CORTE = Completo com quantidade, mas sem máquina/horário de corte."),
    ("ordens", "excesso_corte", "Quantidade cortada maior que a programada."),
    ("sessoes", "ini, fim, mc_realizado, re_ini, re_fim", "Uma linha por sessão de corte (cada início distinto observado)."),
    ("sessoes", "quant_acum_fim, incremento", "Quantidade acumulada ao fim da sessão e peças cortadas na sessão (diferença para a anterior)."),
    ("sessoes", "incremento_confiavel", "True quando a quantidade antes da sessão foi observada e a sessão foi vista terminada."),
    ("sessoes", "possivel_sessao_oculta", "Entre a observação anterior da ordem e esta sessão passaram mais de 24 h sem extração: pode ter havido outra sessão não vista."),
    ("sessoes", "dur_min, min_parada_calendario, dur_util_min", "Duração bruta, minutos em domingo/feriado (fábrica parada) e duração útil. Sessões iniciadas no sábado à noite costumam terminar na segunda 00h-01h: a duração bruta passa de 24 h."),
 ("sessoes", "pecas_por_min", "incremento ÷ dur_util_min (só quando o incremento é confiável)."),
 ("sessoes", "sessao_sem_fim", "Sessão vista começando, mas nunca terminada nas extrações disponíveis."),
    ("sessoes", "antes_da_primeira_extracao", "Sessão iniciada antes da primeira extração do período (provavelmente pertence ao mês anterior)."),
    ("extracoes", "extracao_estimada", "Aba sem horário no título: horário estimado pelo último corte registrado nela (limite inferior)."),
    ("anotacoes", "valor", "Conteúdo de colunas sem cabeçalho (anotações manuais: nº de máquina, '*', 'PRIORIZAR')."),
]


def salvar(saida, E, O, S, A, Q):
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    saida = Path(saida); saida.mkdir(parents=True, exist_ok=True)
    O.to_csv(saida / "escp771_ordens_reconstruidas.csv", index=False, encoding="utf-8")
    S.to_csv(saida / "escp771_sessoes_corte.csv", index=False, encoding="utf-8")
    with pd.ExcelWriter(saida / "escp771_reconstruida.xlsx", engine="openpyxl", datetime_format="dd/mm/yyyy hh:mm:ss", date_format="dd/mm/yyyy") as w:
        Q.to_excel(w, sheet_name="resumo", index=False)
        E.to_excel(w, sheet_name="extracoes", index=False)
        O[O.relatorio == "corte"].to_excel(w, sheet_name="ordens", index=False)
        S[S.relatorio == "corte"].to_excel(w, sheet_name="sessoes", index=False)
        O[O.relatorio == "automotiva"].to_excel(w, sheet_name="automotiva", index=False)
        A.to_excel(w, sheet_name="anotacoes", index=False)
        pd.DataFrame(DICIONARIO, columns=["aba", "coluna(s)", "descrição"]).to_excel(w, sheet_name="dicionario", index=False)
        for ws in w.book.worksheets:
            for c in ws[1]:
                c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="1F3864")
                c.alignment = Alignment(wrap_text=True, vertical="center")
            ws.freeze_panes = "A2"
            for j, col in enumerate(ws.iter_cols(min_row=1, max_row=min(ws.max_row, 200)), 1):
                ws.column_dimensions[get_column_letter(j)].width = min(max(10, max(len(str(c.value)) if c.value is not None else 0 for c in col) + 2), 70)
        for nome in ("resumo", "dicionario"):
            for row in w.sheets[nome].iter_rows(min_row=2):
                for c in row: c.alignment = Alignment(wrap_text=True, vertical="top")


def resumo(E, O, S):
    oc, sc = O[O.relatorio == "corte"], S[S.relatorio == "corte"]
    multi = oc[oc.n_sessoes >= 2]
    lac = E[E.lacuna_antes & (E.relatorio == "corte")]
    linhas = [
        ("Extrações lidas", f"{len(E)} abas ({(E.relatorio == 'corte').sum()} do relatório de corte, {(E.relatorio == 'automotiva').sum()} automotivo); "
                            f"{int(E.extracao_estimada.sum())} com horário estimado"),
        ("Período das extrações", f"{E.extracao.min():%d/%m/%Y %H:%M} a {E.extracao.max():%d/%m/%Y %H:%M}"),
        ("Lacunas > 24 h entre extrações", "; ".join(f"{r.extracao - pd.Timedelta(hours=r.horas_desde_anterior):%d/%m %H:%M} → {r.extracao:%d/%m %H:%M} ({r.horas_desde_anterior:.0f} h)" for r in lac.itertuples()) or "nenhuma"),
        ("Linhas lidas (todas as extrações)", f"{int(E.linhas.sum()):,}".replace(",", ".")),
        ("Ordens distintas (relatório de corte)", f"{len(oc):,}".replace(",", ".")),
        ("  completas / parciais / não cortadas / quantidade zero", f"{(oc.status == 'completo').sum()} / {(oc.status == 'parcial').sum()} / {(oc.status == 'nao_cortado').sum()} / {(oc.status == 'qtd_zero').sum()}"),
        ("  ainda em aberto na última extração", int(oc.presente_na_ultima.sum())),
        ("  sumiram sem corte registrado", f"{int(oc.sumiu_sem_corte.sum())} (das quais {int(oc.sumiu_apos_lacuna.sum())} logo antes de uma lacuna > 24 h)"),
        ("Sessões de corte reconstruídas", f"{len(sc):,}".replace(",", ".")),
        ("  ordens com 2 ou mais sessões", f"{len(multi)} ({len(multi) / max((oc.n_sessoes > 0).sum(), 1):.1%} das ordens cortadas)"),
        ("  sessões sem fim observado", int(sc.sessao_sem_fim.sum())),
        ("  sessões que atravessam domingo ou feriado", f"{int(sc.atravessa_parada.sum())} (duração útil desconta o dia parado)"),
        ("  sessões > 24 h mesmo descontando domingo/feriado", int(sc.dur_mais_24h.sum())),
        ("Sessões por mês de início", " | ".join(f"{k}: {v}" for k, v in sc.mes_sessao.value_counts().sort_index().items())),
        ("  sessões com possível sessão oculta antes", int(sc.possivel_sessao_oculta.sum())),
        ("  sessões iniciadas antes da 1ª extração", int(sc.antes_da_primeira_extracao.sum())),
        ("Ordens 'Completo' sem apontamento de corte", int(oc.completo_sem_apontamento.sum())),
        ("Ordens com corte acima do programado", int(oc.excesso_corte.sum())),
        ("Ordens em que o aplicador (MA) mudou", int((oc.ma_lado_a_mudou | oc.ma_lado_b_mudou).sum())),
        ("Ordens cortadas em mais de uma máquina", int(oc.mcs_realizado.fillna("").str.contains("/").sum())),
        ("Ordens que somem e reaparecem entre extrações", int((oc.ausente_em_extracoes_intermediarias > 0).sum())),
        ("Ordens por arquivo de origem", " | ".join(f"{k}: {v}" for k, v in oc.arquivos.value_counts().items())),
    ]
    return pd.DataFrame(linhas, columns=["item", "resultado"])


if __name__ == "__main__":
    *arquivos, saida = sys.argv[1:]
    E, L, A = ler(arquivos)
    L = normalizar(L)
    E = ordenar_extracoes(E, L)
    L, REC = recuperar_sub_item(L, E)
    for r in REC.itertuples():
        i = E.index[(E.arquivo == r.arquivo) & (E.aba == r.aba)][0]
        E.at[i, "obs"] = (E.at[i, "obs"] + "; " if E.at[i, "obs"] else "") + \
            f"sem coluna SUB-ITEM: {r.recuperadas} de {r.linhas} linhas identificadas pelo cadastro (vizinhas: {r.vizinhas}); demais descartadas"
        print(f"aviso: {r.aba} ({r.arquivo}) sem SUB-ITEM -> {r.recuperadas} de {r.linhas} linhas recuperadas")
    sem_chave = L["cod_prog"].isna() | L["sub_item"].isna()      # ex.: célula solta digitada abaixo da tabela
    if sem_chave.any():
        print(f"aviso: {int(sem_chave.sum())} linha(s) sem COD PROGR/SUB-ITEM descartada(s):",
              L.loc[sem_chave, ["arquivo", "aba"]].value_counts().to_dict())
    L = L[~sem_chave].copy()
    if len(A):
        A["sub_item"] = A["sub_item"].map(txt)
        A = A.merge(E[["arquivo", "aba", "extracao"]], on=["arquivo", "aba"]).sort_values(["extracao", "cod_prog", "sub_item"])
    O, S = reconstruir(L, E)
    Q = resumo(E, O, S)
    salvar(saida, E, O, S, A, Q)
    print(Q.to_string(index=False))
