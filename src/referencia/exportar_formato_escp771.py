"""
Grava a base reconstruída no MESMO formato das extrações da escp771:
título em A1, cabeçalho na linha 2 com filtro, Calibri 11, datas como data curta (formato 14),
códigos, horas, máquina realizada e RE como texto, quantidades como número.

Uso: python exportar_formato_escp771.py ordens.csv sessoes.csv saida.xlsx "JUNHO/2026" "01/06/2026" "01/07/2026" [junho.xlsx]
  O último argumento (opcional) filtra as ordens que aparecem naquele arquivo mensal; o histórico delas
  (inclusive sessões de meses anteriores) vem completo da reconstrução conjunta.

Abas geradas:
  escp771   -> 1 linha por ordem, exatamente como o relatório mostraria a ordem na última vez que foi vista
               (máquina, início, fim e RE da ÚLTIMA sessão; quantidade cortada acumulada).
  Sessoes   -> 1 linha por sessão de corte, no mesmo layout: cada linha é a "foto" da ordem ao fim daquela sessão
               (quantidade cortada acumulada até ali; CORTE = Completo quando acumulado >= quantidade).
  Automotiva-> relatório "CORTE AUTOMOTIVA EM ABERTO" no layout original dele (com CLIENTE, sem MA).
"""
import sys
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

CAB_CORTE = ["MÁQUINA", "COD PROGR", "DATA PROGR", "DATA CRIAÇÃO", "HORA CRIAÇÃO", "KIT", "ITEM", "SUB-ITEM", "QUANTIDADE",
             "QUANT PROD", "CORTE", "TEMPO", "CONDUTOR", "TERMINAL A", "MA LADO A", "TERMINAL B", "MA LADO B", "COMPR",
             "COR CABO", "ESTOQUE MPB CONDUTOR", "ESTOQUE MPB TERMINAL A", "ESTOQUE MPB TERMINAL B", "MC REALIZADO",
             "DATA INI CORTE", "HORA INI CORTE", "RE INI CORTE", "DATA FIM CORTE", "HORA FIM CORTE", "RE FIM CORTE", "QUANT CORTADA"]
CAB_AUTO = ["MÁQUINA", "COD PROGR", "DATA PROGR", "DATA CRIAÇÃO", "HORA CRIAÇÃO", "CLIENTE", "KIT", "ITEM", "SUB-ITEM",
            "QUANTIDADE", "QUANT PROD", "CORTE", "TEMPO", "CONDUTOR", "TERMINAL A", "TERMINAL B", "COMPR", "COR CABO",
            "ESTOQUE MPB CONDUTOR", "ESTOQUE MPB TERMINAL A", "ESTOQUE MPB TERMINAL B", "MC REALIZADO", "DATA INI CORTE",
            "HORA INI CORTE", "RE INI CORTE", "DATA FIM CORTE", "HORA FIM CORTE", "RE FIM CORTE", "QUANT CORTADA"]
DATAS = {"DATA PROGR", "DATA CRIAÇÃO", "DATA INI CORTE", "DATA FIM CORTE"}
NUMEROS = {"MÁQUINA", "COD PROGR", "QUANTIDADE", "QUANT PROD", "TEMPO", "COMPR", "ESTOQUE MPB CONDUTOR",
           "ESTOQUE MPB TERMINAL A", "ESTOQUE MPB TERMINAL B", "QUANT CORTADA"}
TEXTO_COLS = ["kit", "item", "sub_item", "condutor", "terminal_a", "terminal_b", "ma_lado_a", "ma_lado_b", "cor_cabo",
              "hora_criacao", "mc_realizado", "mc_realizado_final", "re_ini", "re_fim", "cliente", "corte_final"]


def vazio(v):
    return v is None or (isinstance(v, float) and np.isnan(v)) or v is pd.NaT or (isinstance(v, str) and v.strip() in ("", "nan", "None", "NaT"))


def cel(col, v):
    if vazio(v):
        return None
    if col in DATAS:
        return pd.Timestamp(v).to_pydatetime().replace(hour=0, minute=0, second=0, microsecond=0)
    if col in NUMEROS:
        f = float(v)
        return int(f) if f.is_integer() else f
    s = str(v).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s


def hora(ts):
    return None if vazio(ts) else pd.Timestamp(ts).strftime("%H:%M:%S")


def linha_cadastro(o, cab):
    base = {"MÁQUINA": o.maquina_prog, "COD PROGR": o.cod_prog, "DATA PROGR": o.data_prog, "DATA CRIAÇÃO": o.data_criacao,
            "HORA CRIAÇÃO": o.hora_criacao, "CLIENTE": o.cliente, "KIT": o.kit, "ITEM": o.item, "SUB-ITEM": o.sub_item,
            "QUANTIDADE": o.quantidade, "TEMPO": o.tempo_prog_min, "CONDUTOR": o.condutor, "TERMINAL A": o.terminal_a,
            "MA LADO A": o.ma_lado_a, "TERMINAL B": o.terminal_b, "MA LADO B": o.ma_lado_b, "COMPR": o.comprimento,
            "COR CABO": o.cor_cabo, "ESTOQUE MPB CONDUTOR": o.estoque_condutor, "ESTOQUE MPB TERMINAL A": o.estoque_terminal_a,
            "ESTOQUE MPB TERMINAL B": o.estoque_terminal_b}
    return {c: base.get(c) for c in cab}


def corte_campos(s):
    """Campos de corte a partir de uma linha de sessão (ou None)."""
    if s is None:
        return {"MC REALIZADO": None, "DATA INI CORTE": None, "HORA INI CORTE": None, "RE INI CORTE": None,
                "DATA FIM CORTE": None, "HORA FIM CORTE": None, "RE FIM CORTE": None}
    return {"MC REALIZADO": s.mc_realizado, "DATA INI CORTE": s.ini, "HORA INI CORTE": hora(s.ini), "RE INI CORTE": s.re_ini,
            "DATA FIM CORTE": s.fim, "HORA FIM CORTE": hora(s.fim), "RE FIM CORTE": s.re_fim}


def aba(wb, nome, titulo, cab, linhas):
    ws = wb.create_sheet(nome)
    ws.append([titulo])
    ws.append(cab)
    for r in linhas:
        ws.append([cel(c, r.get(c)) for c in cab])
    for j, c in enumerate(cab, 1):
        if c in DATAS:                                        # mesmo formato e largura das extrações
            for (cell,) in ws.iter_rows(min_row=2, min_col=j, max_col=j):
                cell.number_format = "mm-dd-yy"
            ws.column_dimensions[get_column_letter(j)].width = 10.7109375
    ws.auto_filter.ref = f"A2:{get_column_letter(len(cab))}{ws.max_row}"
    return ws


if __name__ == "__main__":
    f_ord, f_ses, saida, rotulo, d0, d1 = sys.argv[1:7]
    filtro = sys.argv[7] if len(sys.argv) > 7 else None
    dt = {c: str for c in TEXTO_COLS}
    O = pd.read_csv(f_ord, dtype=dt, low_memory=False, parse_dates=["data_prog", "data_criacao"])
    S = pd.read_csv(f_ses, dtype=dt, low_memory=False, parse_dates=["ini", "fim"])
    if filtro:
        alvos = {f.strip() for f in filtro.split(",")}                       # aceita vários arquivos separados por vírgula
        O = O[O.arquivos.fillna("").str.split(" | ", regex=False).apply(lambda a: bool(alvos & set(a)))]
        S = S.merge(O[["cod_prog", "sub_item"]], on=["cod_prog", "sub_item"])
    S = S.sort_values(["cod_prog", "sub_item", "n_sessao"])
    ultima = {k: g.iloc[-1] for k, g in S.groupby(["cod_prog", "sub_item"], sort=False)}
    O = O.sort_values(["cod_prog", "kit", "sub_item"])

    wb = Workbook(); wb.remove(wb.active)
    # 1) uma linha por ordem (relatório de corte)
    linhas = []
    for o in O[O.relatorio == "corte"].itertuples(index=False):
        r = linha_cadastro(o, CAB_CORTE)
        r.update(corte_campos(ultima.get((o.cod_prog, o.sub_item))))
        r.update({"QUANT PROD": o.quant_cortada_final, "QUANT CORTADA": o.quant_cortada_final, "CORTE": o.corte_final})
        linhas.append(r)
    aba(wb, "escp771", f"BRASCABOS - DETALHE PROGRAMAÇÕES EM ABERTO CORTE - BASE RECONSTRUÍDA {rotulo} - {len(linhas)} REGISTROS DE {d0} A {d1}", CAB_CORTE, linhas)
    # 2) uma linha por sessão
    cad = {(o.cod_prog, o.sub_item): o for o in O.itertuples(index=False)}
    linhas = []
    for s in S[S.relatorio == "corte"].itertuples(index=False):
        o = cad[(s.cod_prog, s.sub_item)]
        q = s.quant_acum_fim if not vazio(s.quant_acum_fim) else s.quant_acum_vista
        r = linha_cadastro(o, CAB_CORTE)
        r.update(corte_campos(s))
        r.update({"QUANT PROD": q, "QUANT CORTADA": q,
                  "CORTE": "Completo" if (not vazio(q) and not vazio(o.quantidade) and q >= o.quantidade) else "Falta"})
        linhas.append(r)
    aba(wb, "Sessoes", f"BRASCABOS - DETALHE PROGRAMAÇÕES EM ABERTO CORTE - SESSÕES DE CORTE RECONSTRUÍDAS {rotulo} - {len(linhas)} SESSÕES", CAB_CORTE, linhas)
    # 3) relatório automotivo, no layout dele
    A = O[O.relatorio == "automotiva"]
    if len(A):
        linhas = []
        for o in A.itertuples(index=False):
            r = linha_cadastro(o, CAB_AUTO)
            r.update(corte_campos(ultima.get((o.cod_prog, o.sub_item))))
            r.update({"QUANT PROD": o.quant_cortada_final, "QUANT CORTADA": o.quant_cortada_final, "CORTE": o.corte_final})
            linhas.append(r)
        aba(wb, "Automotiva", f"BRASCABOS - DETALHE PROGRAMAÇÕES CORTE AUTOMOTIVA EM ABERTO - BASE RECONSTRUÍDA {rotulo} - {len(linhas)} REGISTROS", CAB_AUTO, linhas)
    for ws in wb.worksheets:                                   # Calibri 11 normal, como no original
        ws.sheet_format.defaultRowHeight = 15
    wb.save(saida)
    print("ok", {ws.title: ws.max_row - 2 for ws in wb.worksheets})
