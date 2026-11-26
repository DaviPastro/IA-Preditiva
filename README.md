# IA Clássica – Máquinas de Corte

Projeto de aprendizado de máquina aplicado ao setor de corte e crimpagem de circuitos elétricos (chicotes elétricos), desenvolvido como trabalho prático da disciplina de Inteligência Artificial I do curso de Engenharia de Computação da Fundação Hermínio Ometto (FHO).

**Autores:** Davi Furtado Pastro e Gabriel Vinicius Lima

## Objetivo

Prever, **antes do início de cada sessão de corte**, a ocorrência de três tipos de parada em duas máquinas automáticas de corte e crimpagem (MC 68 e MC 63), utilizando apenas dados tabulares do MES e da programação da produção e provenientes do ERP:

- **Abastecimento**: a sessão terá parada por reabastecimento de condutor ou terminal?
- **Setup**: a sessão terá parada de setup?
- **Parada longa**: a sessão terá alguma parada não programada com 15 minutos ou mais?

## Dados

| Fonte | Conteúdo | Período |
|---|---|---|
| Relatório de ordens em aberto (ERP) | Ordens, características do produto e apontamentos de corte; sessões reconstruídas a partir de extrações periódicas | mai.–out./2026 |
| Eventos de parada (MES) | 1.225.472 eventos de 93 estações | jan.–out./2026 |
| Programação de corte (PCP) | Sequência planejada, máquina de origem e setup previsto | 2025–2026 |
| OEE e produção hora a hora (MES) | Indicadores por hora (MC 68 e MC 63) | jan.–set./2026 |

Conjunto final de modelagem: **4.055 sessões de corte** (22/04 a 07/10/2026).

## Metodologia

Segue as etapas do CRISP-DM:

1. **Exploração e limpeza** das paradas: conversão de formatos, remoção de duplicadas e durações nulas, padronização dos motivos.
2. **Famílias de paradas**: 29 motivos agrupados em 8 famílias (programada, setup, abastecimento, ajuste/manutenção, processo/qualidade, organização do posto, externo e sem categoria).
3. **Rótulos**: cruzamento por intervalo de tempo entre sessões de corte e paradas.
4. **Variáveis de entrada**: 24 variáveis conhecidas antes do início da sessão (produto, sequência, calendário, histórico, planejamento e máquina), sem vazamento de informação.
5. **Modelos**: Regressão Logística, Árvore de Decisão, Random Forest e Gradient Boosting, comparados a um *baseline*.
6. **Validação temporal**: 80% mais antigos para desenvolvimento (busca de hiperparâmetros com `TimeSeriesSplit` de 5 janelas) e 20% mais recentes como teste final, avaliado uma única vez.
7. **Análises complementares**: importância por permutação, calibração das probabilidades e escolha de limiar para a classe rara.

## Principais resultados (conjunto de teste)

| Alvo | Melhor modelo | AUC | Observação |
|---|---|---|---|
| Abastecimento | Random Forest / Gradient Boosting | 0,881 | Explicado principalmente pelo volume de material da ordem |
| Setup | Gradient Boosting | 0,867 | Associado a diferenças de apontamento entre turnos, não às trocas de material |
| Parada longa | Gradient Boosting | 0,752 | Classe rara (~13%); AP de 0,312 contra 0,131 do *baseline* |

## Estrutura do repositório

```
├── notebooks/
│   ├── 00_teste_ambiente.ipynb             # verificação do ambiente
│   ├── 01_exploracao_paradas.ipynb         # exploração e limpeza da base de paradas
│   ├── 02_alvos.ipynb                      # famílias, cruzamento sessões × paradas, dataset de modelagem
│   ├── 03_modelo_abastecimento.ipynb       # primeiro modelo (regressão logística)
│   ├── 04_comparacao_modelos.ipynb         # 4 algoritmos; validação temporal × k-fold
│   ├── 05_calibração_modelos.ipynb         # ajuste de hiperparâmetros, teste final, importância e calibração
│   └── 06_teste_setup_paradalonga.ipynb    # alvos de setup e parada longa
├── src/
│   ├── preparacao.py                       # carregamento da base e definição das variáveis
│   └── referencia/                         # scripts de reconstrução da escp771 e montagem do dataset
├── reports/figures/                        # figuras do artigo
├── data/                                   # dados (ver seção "Dados e confidencialidade")
├── requirements.txt
└── requirements-lock.txt                   # versões exatas das bibliotecas
```

## Como executar

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

Abra a pasta no VS Code, selecione o interpretador `.venv` e execute os notebooks em ordem numérica.

## Tecnologias

Python · pandas · NumPy · scikit-learn · Matplotlib · Jupyter · Git

## Dados e confidencialidade

Os dados são de um processo produtivo real e de uso restrito. A identificação de operadores foi removida das bases derivadas, em atendimento à LGPD. Este repositório é de acesso restrito e não deve ser redistribuído.

## Próximos passos

- Classificar as paradas sem causa registrada (~19% das horas paradas) em famílias, com nível de confiança
- Estimar as peças esperadas por turno como meta realista
- Padronizar o apontamento de paradas entre turnos e reavaliar o modelo de setup
- Integrar os modelos ao MES e a um painel no Power BI