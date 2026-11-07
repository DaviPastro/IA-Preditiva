#criando a preparação dos dados para o modelo de machine learning, dessa forma não é necessário repetir o mesmo código
# em todos os notebooks.

import pandas as pd

X_COLS = [
    "comprimento", "qtd_sessao", "tempo_prog_sessao_min", "quantidade_ordem",
    "troca_condutor", "troca_terminal_a", "troca_terminal_b", "troca_cor", "troca_ma_a", "troca_ma_b",
    "mesmo_sub_item", "mesma_programacao", "ocioso_antes_min", "lacuna_antes_6h",
    "hora_ini", "dia_semana", "ordem_no_dia", "n_subitens_kit",
    "n_ocorrencias_previas_sub", "dur_media_hist_sub", "razao_media_hist_sub",
    "setup_pcp", "maquina_diferente_programada", "mc_68",
]

def carregar_base(caminho="../data/processed/dataset_modelagem.parquet"):
    """Lê o dataset de modelagem em ordem cronológica e cria a variável da máquina."""
    base = pd.read_parquet(caminho).sort_values("ini").reset_index(drop=True)
    base["mc_68"] = (base["mc"] == "68").astype(int)
    return base

def separar_x_y(base, alvo="y_abastecimento"):
    """Devolve as entradas (X) e o alvo (y) escolhido."""
    return base[X_COLS].astype(float), base[alvo]