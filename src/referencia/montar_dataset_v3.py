"""
Dataset de uma máquina de corte — versão 3 (1 linha = 1 SESSÃO de corte na máquina escolhida), montado sobre a escp771 reconstruída (maio a outubro).

Diferenças em relação à versão 1 (que usava uma única extração da escp771, de 02/10):
  * a espinha são as sessões reconstruídas (reconstruir_escp771.py): cada início de corte distinto vira uma linha,
    com a máquina, o operador, o fim e as PEÇAS DAQUELA SESSÃO (incremento). Ordens cortadas em várias vezes deixam
    de ter duração subestimada e ritmo inflado;
  * a duração é a ÚTIL (desconta domingos e feriados em que a fábrica está parada);
  * o tempo programado da sessão é proporcional às peças da sessão (TEMPO da ordem × peças da sessão ÷ quantidade);
  * cada fonte auxiliar só é usada dentro do período que ela cobre. Fora dele, as colunas ficam VAZIAS (nunca zero)
    e uma coluna "cobre_*" diz se a sessão está coberta.

Uso:
  python montar_dataset_v3.py MAQUINA sessoes.csv ordens.csv programacao.xlsx "paradas1.xlsx,paradas2.xlsx" "oee1.xlsx,oee2.xlsx" base_horaria.csv revisao_manual.csv pasta_saida/
  - MAQUINA: número da máquina (ex.: 68, 63), comparado com MC REALIZADO de cada sessão.
  - paradas e OEE: um ou mais arquivos separados por vírgula (ex.: 1º semestre e 2º semestre). As abas usadas são
    PARADAS e "OEE MCs"/"Colar - dados mês"; OEE Parque e MANUTENÇÃO POR MÁQUINA são ignoradas. Eventos repetidos
    entre arquivos ou dentro do mesmo arquivo (mesma estação, início, evento e duração) entram uma vez só.
  - base_horaria.csv: usada só se tiver a máquina; senão as colunas h_* ficam vazias.
"""
import re, sys, unicodedata
from pathlib import Path
import numpy as np
import pandas as pd

if __name__ == "__main__":
    MC, ARQ_SES, ARQ_ORD, ARQ_PROG, ARQS_PAR, ARQS_OEE, ARQ_HORARIA, ARQ_REVISAO, SAIDA = sys.argv[1:10]
SAIDA = Path(SAIDA); SAIDA.mkdir(parents=True, exist_ok=True)
MC = str(int(MC))
ARQS_PAR = [a.strip() for a in ARQS_PAR.split(",") if a.strip()]
ARQS_OEE = [a.strip() for a in ARQS_OEE.split(",") if a.strip()]
JANELA_CONTEXTO_H = 6          # horas antes do início usadas como contexto de paradas
LIMITE_SESSAO_MIN = 1440       # sessão com duração ÚTIL acima de 24 h = suspeita
DUR_MIN_PLAUSIVEL = 3          # sessão com menos de 3 min = suspeita
RITMO_MAX_PLAUSIVEL = 100      # > 100 peças/min (mediana da MC 68 ~43) = suspeita
TOL_OEE_DIAS = 4               # OEE "do dia anterior" só vale se o dia de referência estiver a no máximo 4 dias
log = {}


# ---------------------------------------------------------------- utilidades
def norm(s):
    return unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().strip().lower()


def num_mc(serie):
    return pd.to_numeric(serie.astype(str).str.extract(r"^\s*MC\s*(\d+)")[0], errors="coerce")


def cod(x):
    if pd.isna(x):
        return "SEM"
    if isinstance(x, (int, float, np.integer, np.floating)) and float(x).is_integer():
        return str(int(x))
    s = str(x).strip()
    return s[:-2] if s.endswith(".0") and s[:-2].isdigit() else s


FAMILIAS = ["programada", "setup", "abastecimento", "ajuste_manutencao",
            "processo_qualidade", "organizacao_pessoal", "sem_categoria", "outros"]


def familia(ev):
    e = norm(ev)
    if any(k in e for k in ["lanche", "almoc", "treinamento", "reuniao"]): return "programada"
    if "setup" in e: return "setup"
    if "abastecimento" in e: return "abastecimento"
    if any(k in e for k in ["ajuste", "manutencao"]): return "ajuste_manutencao"
    if any(k in e for k in ["plano de controle", "bobina", "produto novo", "embramada"]): return "processo_qualidade"
    if any(k in e for k in ["organizacao", "saida do posto", "ambulatorio"]): return "organizacao_pessoal"
    if any(k in e for k in ["desconhecida", "nao categorizada"]): return "sem_categoria"
    return "outros"


# ---------------------------------------------------------------- 1) espinha: sessões da MC 68
txtc = {c: str for c in ["sub_item", "mc_realizado", "re_ini", "re_fim", "kit", "item", "condutor", "terminal_a",
                         "terminal_b", "ma_lado_a", "ma_lado_b", "cor_cabo", "mcs_realizado", "mc_realizado_final"]}
S = pd.read_csv(ARQ_SES, dtype=txtc, low_memory=False, parse_dates=["ini", "fim"])
O = pd.read_csv(ARQ_ORD, dtype=txtc, low_memory=False, parse_dates=["data_prog", "data_criacao"])
S["mc_realizado"] = S["mc_realizado"].map(cod)
s = S[(S.relatorio == "corte") & (S.mc_realizado == MC)].copy()
log[f"sessões de corte da MC {MC} na base reconstruída"] = len(s)
sem_fim = s["fim"].isna()
log["sessões removidas por não terem fim observado"] = int(sem_fim.sum())
s = s[~sem_fim]
cad = ["cod_prog", "sub_item", "maquina_prog", "kit", "item", "quantidade", "tempo_prog_min", "condutor", "terminal_a",
       "terminal_b", "ma_lado_a", "ma_lado_b", "comprimento", "cor_cabo", "n_sessoes", "mcs_realizado", "status",
       "data_prog", "data_criacao"]
b = s.merge(O[O.relatorio == "corte"][cad], on=["cod_prog", "sub_item"], how="left", validate="many_to_one")
assert len(b) == len(s)
b = b.sort_values("ini").reset_index(drop=True)
b["id_sessao"] = np.arange(1, len(b) + 1)

# peças da sessão: observadas (incremento confiável) ou inferidas quando a ordem tem uma única sessão
b["qtd_sessao"] = b["incremento"]
b["qtd_sessao_fonte"] = np.where(b["incremento_confiavel"], "observada",
                                 np.where(b["n_sessoes"] == 1, "inferida_sessao_unica", "incerta"))
b["quantidade_ordem"] = b["quantidade"]
b["tempo_prog_ordem_min"] = b["tempo_prog_min"]
frac = (b["qtd_sessao"] / b["quantidade_ordem"]).where(b["quantidade_ordem"] > 0)
b["tempo_prog_sessao_min"] = b["tempo_prog_ordem_min"] * frac                     # tempo programado proporcional às peças da sessão
b["sessao_completa_ordem"] = (b["quant_acum_fim"] >= b["quantidade_ordem"]).astype(int)
b["dur_bruta_min"] = b["dur_min"]
b["dur_min"] = b["dur_util_min"]                                                  # alvo principal: duração útil
b["pecas_por_min"] = (b["qtd_sessao"] / b["dur_min"]).where(b["dur_min"] > 0)
b["razao_real_prog"] = (b["dur_min"] / b["tempo_prog_sessao_min"]).where(b["tempo_prog_sessao_min"] > 0)
b["flag_sessao_maior_24h"] = b["dur_min"] > LIMITE_SESSAO_MIN
b["flag_suspeito_auto"] = (b["flag_sessao_maior_24h"] | (b["dur_min"] < DUR_MIN_PLAUSIVEL)
                           | (b["pecas_por_min"] > RITMO_MAX_PLAUSIVEL))
log["sessões suspeitas pela regra automática"] = int(b["flag_suspeito_auto"].sum())

# ---------------------------------------------------------------- 2) contexto sequencial na MC 68 (só passado)
for c in ["condutor", "terminal_a", "terminal_b", "cor_cabo", "ma_lado_a", "ma_lado_b", "sub_item", "kit"]:
    b[c] = b[c].map(cod)
b["cod_prog"] = b["cod_prog"].map(cod)
prev = b.shift(1)
b["ocioso_antes_min"] = (b["ini"] - prev["fim"]).dt.total_seconds() / 60
b["flag_sobrepoe_anterior"] = b["ocioso_antes_min"] < 0
b["ocioso_antes_min"] = b["ocioso_antes_min"].clip(lower=0)
b["lacuna_antes_6h"] = b["ocioso_antes_min"] > 360            # a sessão anterior pode não ser a imediatamente anterior
b["dur_anterior_min"] = prev["dur_min"]
b["qtd_anterior"] = prev["qtd_sessao"]
b["compr_anterior"] = prev["comprimento"]
for nome, col in [("troca_condutor", "condutor"), ("troca_terminal_a", "terminal_a"), ("troca_terminal_b", "terminal_b"),
                  ("troca_cor", "cor_cabo"), ("troca_ma_a", "ma_lado_a"), ("troca_ma_b", "ma_lado_b")]:
    b[nome] = np.where(prev[col].isna(), np.nan, (b[col] != prev[col]).astype(float))
b["mesmo_sub_item"] = np.where(prev["sub_item"].isna(), np.nan, (b["sub_item"] == prev["sub_item"]).astype(float))
b["mesma_programacao"] = np.where(prev["cod_prog"].isna(), np.nan, (b["cod_prog"] == prev["cod_prog"]).astype(float))
b["ordem_no_dia"] = b.groupby(b["ini"].dt.date).cumcount() + 1
b["n_sessoes_mesma_prog"] = b.groupby("cod_prog")["id_sessao"].transform("size")
b["n_subitens_kit"] = b.groupby("kit")["sub_item"].transform("nunique")

# ---- revisão manual dos suspeitos (chave = início + programação + sub-item)
b["flag_registro_suspeito"] = b["flag_suspeito_auto"]
b["flag_revisada_manual"] = False
try:
    rev = pd.read_csv(ARQ_REVISAO, dtype={"ini": str, "cod_prog": str, "sub_item": str})
    decis = dict(zip(rev["ini"] + "|" + rev["cod_prog"] + "|" + rev["sub_item"], rev["flag_manual"].astype(bool)))
    k = b["ini"].dt.strftime("%Y-%m-%d %H:%M:%S") + "|" + b["cod_prog"] + "|" + b["sub_item"]
    achou = k.isin(decis)
    b.loc[achou, "flag_registro_suspeito"] = k[achou].map(decis).astype(bool)
    b.loc[achou, "flag_revisada_manual"] = True
    log["decisões manuais (revisão de 07/10) encontradas na base nova"] = f"{int(achou.sum())} de {len(rev)}"
except FileNotFoundError:
    log["decisões manuais"] = "arquivo de revisão não encontrado"
log["sessões com registro suspeito (flag final)"] = int(b["flag_registro_suspeito"].sum())

ok = ~b["flag_registro_suspeito"]
g = b.assign(_d=b["dur_min"].where(ok), _r=b["razao_real_prog"].where(ok), _p=b["pecas_por_min"].where(ok)).groupby("sub_item")
b["n_ocorrencias_previas_sub"] = g.cumcount()
for nome, col in [("dur_media_hist_sub", "_d"), ("razao_media_hist_sub", "_r"), ("ritmo_medio_hist_sub", "_p")]:
    b[nome] = g[col].transform(lambda x: x.shift(1).expanding().mean())

# ---------------------------------------------------------------- 3) programação do PCP
prog = pd.read_excel(ARQ_PROG, sheet_name="Programacao")
for c in ["ITEM", "KIT", "SUB-ITEM"]:
    prog[c] = prog[c].astype(str).str.replace("'", "", regex=False).str.strip()
chave = (prog.sort_values("Data").groupby(["PROGRAMACAO", "SUB-ITEM"], as_index=False)
             .agg(setup_pcp=("SETUP", "first"), seq_pcp=("SEQ", "first"), qtde_pcp=("QTDE", "first"),
                  item_pai_pcp=("ITEM", "first"), data_pcp=("Data", "first")))
chave["PROGRAMACAO"] = chave["PROGRAMACAO"].map(cod)
n0 = len(b)
b = b.merge(chave, left_on=["cod_prog", "sub_item"], right_on=["PROGRAMACAO", "SUB-ITEM"], how="left").drop(columns=["PROGRAMACAO", "SUB-ITEM"])
assert len(b) == n0
b["casou_programacao"] = b["setup_pcp"].notna()
b["maquina_diferente_programada"] = (pd.to_numeric(b["maquina_prog"], errors="coerce") != int(MC)).astype(int)
log["período da programação do PCP"] = f"{prog.Data.min():%d/%m/%Y} a {prog.Data.max():%d/%m/%Y}"

# ---------------------------------------------------------------- 4) paradas (sobreposição de intervalos), só no período coberto
par = pd.concat([pd.read_excel(a, sheet_name="PARADAS") for a in ARQS_PAR], ignore_index=True)


def para_dt(x):
    if hasattr(x, "year"): return pd.Timestamp(x)
    return pd.to_datetime(x, format="%d/%m/%Y %H:%M", errors="coerce")


def para_min(x):
    if hasattr(x, "total_seconds"): return x.total_seconds() / 60
    try:
        h, m, s_ = str(x).split(":"); return int(h) * 60 + int(m) + float(s_) / 60
    except Exception:
        return np.nan


par["di"] = par["Data inicial"].map(para_dt)
par["dur"] = par["Duração"].map(para_min)
par["fim_c"] = par["di"] + pd.to_timedelta(par["dur"], unit="m")
par["mc"] = num_mc(par["Estação"])
n_par = len(par)
par = par.drop_duplicates(subset=["Estação", "di", "Evento", "dur"])          # repetidos entre arquivos ou no mesmo arquivo
log["paradas: eventos lidos / repetidos removidos"] = f"{n_par} / {n_par - len(par)}"
p = par[(par["mc"] == int(MC)) & par["di"].notna() & par["dur"].notna()].copy()
PAR_INI, PAR_FIM = par["di"].min().normalize(), par["fim_c"].max()
log["período da base de paradas"] = f"{PAR_INI:%d/%m/%Y} a {PAR_FIM:%d/%m/%Y %H:%M}"
p["evento"] = p["Evento"].map(lambda x: " ".join(str(x).split()))
p["familia"] = p["evento"].map(familia)
fam_idx = p["familia"].map({f: i for i, f in enumerate(FAMILIAS)}).values
ps = p["di"].values.astype("datetime64[s]").astype("int64")
pe = p["fim_c"].values.astype("datetime64[s]").astype("int64")
nao_prog = fam_idx != FAMILIAS.index("programada")
b["cobre_paradas"] = (b["ini"] - pd.Timedelta(hours=JANELA_CONTEXTO_H) >= PAR_INI) & (b["fim"] <= PAR_FIM)
mat = np.full((len(b), len(FAMILIAS)), np.nan); n_ev = np.full(len(b), np.nan)
ctx_min = np.full(len(b), np.nan); ctx_n = np.full(len(b), np.nan); detalhe = []
for i, (ini, fim, cob) in enumerate(zip(b["ini"], b["fim"], b["cobre_paradas"])):
    if not cob:
        continue
    a, z = int(ini.timestamp()), int(fim.timestamp())
    ov = np.clip(np.minimum(pe, z) - np.maximum(ps, a), 0, None); m = ov > 0
    mat[i] = np.bincount(fam_idx[m], weights=ov[m] / 60, minlength=len(FAMILIAS)); n_ev[i] = m.sum()
    for ev, mins in pd.Series(ov[m] / 60).groupby(p["evento"].values[m]).sum().items():
        detalhe.append((int(b.at[i, "id_sessao"]), ev, round(mins, 2)))
    a0 = a - JANELA_CONTEXTO_H * 3600
    ov0 = np.clip(np.minimum(pe, a) - np.maximum(ps, a0), 0, None) * nao_prog
    ctx_min[i] = ov0.sum() / 60; ctx_n[i] = (ov0 > 0).sum()
for j, f in enumerate(FAMILIAS):
    b[f"par_{f}_min"] = mat[:, j]
b["par_total_min"] = np.nansum(mat, axis=1); b.loc[~b["cobre_paradas"], "par_total_min"] = np.nan
b["par_nao_programada_min"] = b["par_total_min"] - b["par_programada_min"]
b["par_n_eventos"] = n_ev
b["pct_nao_programada_na_sessao"] = b["par_nao_programada_min"] / b["dur_min"]
b["dur_sem_programadas_min"] = b["dur_min"] - b["par_programada_min"]
b["par_ctx_nao_prog_min"] = ctx_min; b["par_ctx_n_eventos"] = ctx_n
paradas_detalhe = pd.DataFrame(detalhe, columns=["id_sessao", "evento", "minutos_na_sessao"])
log["sessões cobertas pela base de paradas"] = int(b["cobre_paradas"].sum())

# ---------------------------------------------------------------- 5) OEE e produção do dia anterior (com tolerância)
oee_x = [pd.read_excel(a, sheet_name=["OEE MCs", "Colar - dados mês"]) for a in ARQS_OEE]
o = pd.concat([x["OEE MCs"] for x in oee_x], ignore_index=True).drop_duplicates(subset=["Data", "Estação"], keep="last")
o["mc"] = num_mc(o["Estação"])
for c in ["Disponibilidade", "Desempenho", "Qualidade", "OEE Total"]:
    o[c] = pd.to_numeric(o[c].astype(str).str.replace(",", ".", regex=False), errors="coerce")
o = o[o["mc"] == int(MC)].sort_values("Data").reset_index(drop=True)
for c, n in [("Disponibilidade", "disp"), ("Desempenho", "desemp"), ("OEE Total", "oee")]:
    o[f"{n}_med7d"] = o[c].rolling(7, min_periods=3).mean()
o = o.rename(columns={"Disponibilidade": "disp_dia_ant", "Desempenho": "desemp_dia_ant", "Qualidade": "qual_dia_ant", "OEE Total": "oee_dia_ant"})
cols_o = ["Data", "disp_dia_ant", "desemp_dia_ant", "qual_dia_ant", "oee_dia_ant", "disp_med7d", "desemp_med7d", "oee_med7d"]
log["período da base de OEE"] = f"{o.Data.min():%d/%m/%Y} a {o.Data.max():%d/%m/%Y}"
c_ = pd.concat([x["Colar - dados mês"] for x in oee_x], ignore_index=True).drop_duplicates()
c_["mc"] = num_mc(c_["Estação"])
c_ = c_[(c_["mc"] == int(MC)) & c_["Produção Programada"].notna()]
dia = c_.groupby("Data").agg(prod_real=("Produção Realizada", "sum"), prod_prog=("Produção Programada", "sum")).reset_index()
dia["ratio_prod_prog_dia_ant"] = dia["prod_real"] / dia["prod_prog"].where(dia["prod_prog"] > 0)
dia["ratio_prod_prog_med7d"] = dia["ratio_prod_prog_dia_ant"].rolling(7, min_periods=3).mean()
dia = dia[["Data", "ratio_prod_prog_dia_ant", "ratio_prod_prog_med7d"]]
b["_dia"] = b["ini"].dt.normalize().astype("datetime64[ns]")
b = b.sort_values("_dia", kind="stable")
o_ref = o[cols_o].copy(); o_ref["Data"] = o_ref["Data"].astype("datetime64[ns]"); o_ref["dia_ref"] = o_ref["Data"]
dia["Data"] = dia["Data"].astype("datetime64[ns]")
tol = pd.Timedelta(days=TOL_OEE_DIAS)
b = pd.merge_asof(b, o_ref, left_on="_dia", right_on="Data", direction="backward", allow_exact_matches=False, tolerance=tol).drop(columns="Data")
b["dias_desde_oee_ref"] = (b["_dia"] - b["dia_ref"]).dt.days
b = b.drop(columns="dia_ref")
b = pd.merge_asof(b, dia, left_on="_dia", right_on="Data", direction="backward", allow_exact_matches=False, tolerance=tol).drop(columns="Data")
b = b.sort_values("id_sessao").reset_index(drop=True)
b["cobre_oee"] = b["oee_dia_ant"].notna()

# ---------------------------------------------------------------- 5b) base horária (carimbo = início do intervalo)
HORA_CTX = ["h_pecas_ult_hora", "h_pecas_ctx3_med", "h_pecas_ctx6_med", "h_pecas_ctx24_med", "h_oee_ctx3", "h_oee_ctx6",
            "h_disp_ctx3", "h_disp_ctx6", "h_desemp_ctx3", "h_desemp_ctx6", "h_horas_zero_ctx6", "h_horas_dado_ctx6"]
HORA_POS = ["h_pecas_na_sessao", "h_min_cobertos", "h_cobertura_pct", "h_ritmo_pecas_min_sessao",
            "h_oee_na_sessao", "h_disp_na_sessao", "h_desemp_na_sessao", "h_razao_cortada_vs_horaria"]
hh = pd.read_csv(ARQ_HORARIA)
hh.columns = ["estacao", "data", "turno", "hora", "oee", "disp", "desemp", "qual", "pecas"]
hh = hh[num_mc(hh["estacao"]) == int(MC)].copy()
if hh.empty:                                                  # não há base horária desta máquina
    hh = pd.DataFrame({"ts": pd.to_datetime(["2000-01-01"]), "pecas": [np.nan], "oee": [np.nan], "disp": [np.nan], "desemp": [np.nan], "data": ["01/01/2000"], "hora": ["00:00"]})
    log["base horária"] = f"sem dados da MC {MC}: colunas h_* vazias"
hh["ts"] = pd.to_datetime(hh["data"] + " " + hh["hora"], format="%d/%m/%Y %H:%M")
hh = hh.sort_values(["ts", "pecas"], ascending=[True, False]).drop_duplicates("ts", keep="first")
# horas com produção mas disponibilidade 0: os índices não foram calculados (ex.: MC 68 de 02/01 a 28/03/2026).
# As peças são válidas; OEE, disponibilidade e desempenho dessas horas viram "sem dado" (não zero).
sem_indice = (hh["pecas"] > 0) & (hh["disp"] == 0)
hh.loc[sem_indice, ["oee", "disp", "desemp"]] = np.nan
if sem_indice.any():
    log["base horária: horas com peças mas sem índices de OEE (tratadas como sem dado)"] = int(sem_indice.sum())
grade = pd.date_range(hh["ts"].min().normalize(), hh["ts"].max().normalize() + pd.Timedelta(hours=23), freq="h")
H = hh.set_index("ts")[["pecas", "oee", "disp", "desemp"]].reindex(grade)
if "base horária" not in log:
    log["período da base horária"] = f"{grade[0]:%d/%m/%Y} a {grade[-1]:%d/%m/%Y}"
hpec, hoee, hdis, hdes = (H[c].values for c in ["pecas", "oee", "disp", "desemp"])
pos_h = {t: i for i, t in enumerate(grade)}
nm = lambda a: float(np.nanmean(a)) if np.isfinite(a).any() else np.nan
ctx = {c: np.full(len(b), np.nan) for c in HORA_CTX}; pos = {c: np.full(len(b), np.nan) for c in HORA_POS}
for i, (ini, fim, q, dur) in enumerate(zip(b["ini"], b["fim"], b["qtd_sessao"], b["dur_bruta_min"])):
    i0 = pos_h.get(ini.floor("h"))
    if i0 is not None:
        for n in (3, 6, 24):
            lo = max(i0 - n, 0)
            if i0 - lo >= 1:
                ctx[f"h_pecas_ctx{n}_med"][i] = nm(hpec[lo:i0])
                if n in (3, 6):
                    ctx[f"h_oee_ctx{n}"][i] = nm(hoee[lo:i0]); ctx[f"h_disp_ctx{n}"][i] = nm(hdis[lo:i0]); ctx[f"h_desemp_ctx{n}"][i] = nm(hdes[lo:i0])
        if i0 >= 1: ctx["h_pecas_ult_hora"][i] = hpec[i0 - 1]
        lo = max(i0 - 6, 0); w = hpec[lo:i0]
        ctx["h_horas_dado_ctx6"][i] = np.isfinite(w).sum(); ctx["h_horas_zero_ctx6"][i] = (w == 0).sum()
    if not np.isfinite(dur) or dur > LIMITE_SESSAO_MIN * 3:
        continue
    tot = cob = ow = dw = sw = 0.0
    for t in pd.date_range(ini.floor("h"), fim.floor("h"), freq="h"):
        j = pos_h.get(t)
        if j is None or not np.isfinite(hpec[j]): continue
        ov = (min(fim, t + pd.Timedelta(hours=1)) - max(ini, t)).total_seconds() / 60
        if ov <= 0: continue
        tot += hpec[j] * ov / 60; cob += ov; ow += hoee[j] * ov; dw += hdis[j] * ov; sw += hdes[j] * ov
    if cob > 0:
        pos["h_pecas_na_sessao"][i] = tot; pos["h_min_cobertos"][i] = cob; pos["h_cobertura_pct"][i] = cob / dur
        pos["h_ritmo_pecas_min_sessao"][i] = tot / cob; pos["h_oee_na_sessao"][i] = ow / cob
        pos["h_disp_na_sessao"][i] = dw / cob; pos["h_desemp_na_sessao"][i] = sw / cob
        pos["h_razao_cortada_vs_horaria"][i] = q / tot if tot > 0 else np.nan
for c, v in {**ctx, **pos}.items(): b[c] = v
b["cobre_horaria"] = b["h_pecas_ctx6_med"].notna()

# ---------------------------------------------------------------- 6) colunas finais, cobertura e partição temporal
b["hora_ini"] = b["ini"].dt.hour; b["dia_semana"] = b["ini"].dt.dayofweek; b["mes"] = b["ini"].dt.month
b["mes_sessao"] = b["ini"].dt.strftime("%Y-%m")
b["cobertura_fontes"] = np.select(
    [b.cobre_horaria & b.cobre_paradas & b.cobre_oee, b.cobre_paradas & b.cobre_oee, b.cobre_paradas | b.cobre_oee],
    ["completa (PCP + paradas + OEE + horária)", "PCP + paradas + OEE", "parcial"], default="só escp771 + PCP")
corte = int(len(b) * 0.75)
b["particao_sugerida"] = np.where(b["id_sessao"] <= corte, "treino", "teste")
b["operador_ini"] = b["re_ini"]; b["operador_fim"] = b["re_fim"]

ordem = (["id_sessao", "particao_sugerida", "mes_sessao", "cobertura_fontes", "cobre_paradas", "cobre_oee", "cobre_horaria",
          "ini", "fim", "cod_prog", "kit", "sub_item", "item", "item_pai_pcp", "casou_programacao",
          "n_sessao", "n_sessoes", "mcs_realizado", "qtd_sessao", "qtd_sessao_fonte", "quantidade_ordem", "sessao_completa_ordem",
          "comprimento", "tempo_prog_ordem_min", "tempo_prog_sessao_min",
          "condutor", "terminal_a", "terminal_b", "ma_lado_a", "ma_lado_b", "cor_cabo", "operador_ini", "operador_fim",
          "hora_ini", "dia_semana", "mes", "maquina_diferente_programada", "setup_pcp", "seq_pcp", "qtde_pcp",
          "ordem_no_dia", "n_sessoes_mesma_prog", "n_subitens_kit", "ocioso_antes_min", "lacuna_antes_6h", "flag_sobrepoe_anterior",
          "dur_anterior_min", "qtd_anterior", "compr_anterior", "troca_condutor", "troca_terminal_a", "troca_terminal_b",
          "troca_cor", "troca_ma_a", "troca_ma_b", "mesmo_sub_item", "mesma_programacao",
          "n_ocorrencias_previas_sub", "dur_media_hist_sub", "razao_media_hist_sub", "ritmo_medio_hist_sub",
          "par_ctx_nao_prog_min", "par_ctx_n_eventos",
          "disp_dia_ant", "desemp_dia_ant", "qual_dia_ant", "oee_dia_ant", "disp_med7d", "desemp_med7d", "oee_med7d",
          "ratio_prod_prog_dia_ant", "ratio_prod_prog_med7d", "dias_desde_oee_ref"] + HORA_CTX + [
          "dur_min", "dur_bruta_min", "min_parada_calendario", "atravessa_parada", "razao_real_prog", "pecas_por_min",
          "dur_sem_programadas_min",
          "incremento_confiavel", "possivel_sessao_oculta", "flag_sessao_maior_24h", "flag_suspeito_auto",
          "flag_registro_suspeito", "flag_revisada_manual"]
         + [f"par_{f}_min" for f in FAMILIAS] + ["par_total_min", "par_nao_programada_min", "par_n_eventos",
                                                  "pct_nao_programada_na_sessao"] + HORA_POS)
final = b[ordem]

DICIONARIO = [
 ("id_sessao", "chave", "Número sequencial da sessão de corte na máquina, em ordem cronológica de início."),
 ("particao_sugerida", "divisão", "treino = 75% mais antigo, teste = 25% mais recente (por tempo)."),
 ("mes_sessao, cobertura_fontes, cobre_paradas, cobre_oee, cobre_horaria", "cobertura", "Quais bases auxiliares cobrem a sessão. Fora da cobertura as colunas daquela base ficam VAZIAS (não zero). Paradas e OEE: jan-08/10; horária: ago-set (só MC 68); PCP: todo o período."),
 ("ini, fim", "chave", "Início e fim da sessão (recuperados das várias extrações da escp771)."),
 ("cod_prog, kit, sub_item, item, item_pai_pcp", "identificação", "Programação, kit, sub-item, item e item pai (PCP)."),
 ("n_sessao, n_sessoes, mcs_realizado", "identificação", "Número da sessão dentro da ordem, total de sessões da ordem e máquinas em que a ordem foi cortada."),
 ("qtd_sessao, qtd_sessao_fonte", "feature/alvo", "Peças cortadas nesta sessão. 'observada' = diferença entre acumulados vistos; 'inferida_sessao_unica' = ordem com uma só sessão, peças = acumulado; 'incerta' = demais."),
 ("quantidade_ordem, sessao_completa_ordem", "feature", "Quantidade programada da ordem e se esta sessão completou a ordem."),
 ("tempo_prog_ordem_min, tempo_prog_sessao_min", "feature", "Tempo programado (TEMPO) da ordem e proporcional às peças da sessão."),
 ("condutor ... cor_cabo, operador_ini, operador_fim", "feature", "Componentes, aplicadores e RE do operador no início/fim. 'SEM' = lado sem terminal."),
 ("setup_pcp, seq_pcp, qtde_pcp", "feature", "Setup, sequência e quantidade do PCP."),
 ("ocioso_antes_min ... mesma_programacao", "contexto (passado)", "Relação com a sessão anterior NA MESMA MÁQUINA. lacuna_antes_6h = mais de 6 h desde a anterior (contexto menos confiável, ex.: fins de semana e o buraco de setembro)."),
 ("*_hist_sub", "contexto (passado)", "Histórico do mesmo sub-item só com sessões ANTERIORES e não suspeitas."),
 ("par_ctx_*", "contexto (passado)", "Paradas não programadas nas 6 h antes do início (só com cobre_paradas)."),
 ("*_dia_ant, *_med7d, ratio_prod_prog_*", "contexto (passado)", "OEE e produção do último dia ANTERIOR (no máximo 4 dias antes). Desempenho/OEE da MC 68 passam de 100%."),
 ("h_*_ctx*", "contexto (passado)", "Base horária: horas completas antes do início (ago-set)."),
 ("dur_min", "ALVO", "Duração ÚTIL da sessão (desconta domingos e feriados)."),
 ("dur_bruta_min, min_parada_calendario, atravessa_parada", "análise", "Duração bruta, minutos em domingo/feriado e se a sessão atravessou a parada."),
 ("razao_real_prog", "ALVO", "dur_min / tempo_prog_sessao_min."),
 ("pecas_por_min", "ALVO", "qtd_sessao / dur_min."),
 ("incremento_confiavel, possivel_sessao_oculta", "qualidade", "Peças da sessão observadas diretamente; possível sessão não vista antes desta (lacuna > 24 h entre extrações)."),
 ("flag_suspeito_auto, flag_registro_suspeito, flag_revisada_manual", "qualidade", "Regra automática (> 24 h úteis, < 3 min ou > 100 peças/min) e flag final após a revisão manual de 07/10."),
 ("par_<familia>_min, par_total_min ...", "ANÁLISE (pós-sessão)", "Minutos de parada DENTRO da sessão por família. Nunca usar como feature para prever a duração."),
 ("h_*_na_sessao, h_razao_cortada_vs_horaria", "ANÁLISE (pós-sessão)", "Peças e OEE da base horária dentro da sessão."),
]

if __name__ == "__main__":
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    final.to_csv(SAIDA / f"dataset_mc{MC}_v3.csv", index=False, encoding="utf-8")
    pd.to_pickle((final, log, paradas_detalhe), SAIDA / f"_dataset_mc{MC}_v3.pkl")
    with pd.ExcelWriter(SAIDA / f"dataset_mc{MC}_v3.xlsx", engine="openpyxl", datetime_format="dd/mm/yyyy hh:mm") as w:
        final.to_excel(w, sheet_name="dataset", index=False)
        pd.DataFrame(DICIONARIO, columns=["coluna(s)", "papel", "descrição"]).to_excel(w, sheet_name="dicionario", index=False)
        pd.DataFrame(list(log.items()), columns=["verificação", "resultado"]).to_excel(w, sheet_name="qualidade", index=False)
        paradas_detalhe.to_excel(w, sheet_name="paradas_por_evento", index=False)
        for ws in w.book.worksheets:
            for c in ws[1]:
                c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="1F3864")
                c.alignment = Alignment(wrap_text=True, vertical="center")
            ws.freeze_panes = "B2" if ws.title == "dataset" else "A2"
            for j, col in enumerate(ws.iter_cols(min_row=1, max_row=60), 1):
                ws.column_dimensions[get_column_letter(j)].width = min(max(10, max(len(str(c.value)) if c.value is not None else 0 for c in col) + 2), 60)
        w.sheets["dicionario"].column_dimensions["C"].width = 110
        for row in w.sheets["dicionario"].iter_rows(min_row=2):
            for c in row: c.alignment = Alignment(wrap_text=True, vertical="top")
    print(final.shape); [print(f"  {k}: {v}") for k, v in log.items()]
