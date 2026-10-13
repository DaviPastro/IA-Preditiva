# IA Clássica – Máquinas de Corte

Projeto de aprendizado de máquina aplicado ao setor de corte e aplicação de circuitos elétricos,
desenvolvido no trabalho e como projeto acadêmico de Engenharia de Computação.

A partir dos dados do MES (programação de corte(PCP), relatório escp771 (datasul),
OEE/produção hora a hora e eventos de parada(sistema MES)), o projeto busca:

- **Prever, antes de a ordem começar**, a chance de ocorrer parada para
  abastecimento, parada longa (> 15 min) e setup;
- **Estimar a probabilidade de cada tipo de parada** por ordem;
- **Classificar paradas "sem categoria"** com um percentual de confiança;
- **Estimar as peças esperadas por turno**, como meta realista.

Além disso, o projeto analisa a aderência ao tempo padrão, o custo das trocas
(setup), o Pareto de perdas, o programado vs. realizado e o consumo de material,
para apoiar ações de melhoria focada.

**Máquinas analisadas:** MC 68 e MC 63 (dataset inicial)
**Tecnologias:** Python, pandas, scikit-learn, Jupyter, Power BI
**Metodologia:** CRISP-DM, validação temporal e comparação com baseline

> Os dados da empresa não são versionados neste repositório.