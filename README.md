# Analisador de Dados de Excel

Este projeto é um script em Python que permite a leitura, análise e visualização de dados a partir de planilhas Excel. Ele calcula estatísticas chave, como média, soma e desvio padrão, e também filtra e separa dados com base em critérios específicos ('Ativo' e 'Inativo'). Além disso, gera gráficos visuais dos resultados e salva resumos em arquivos Excel e CSV.

## Funcionalidades

- Leitura de arquivos Excel.
- Cálculo de estatísticas como:
  - Média, soma, mediana, moda e desvio padrão.
  - Contagem de itens 'Ativo'.
- Geração de gráficos de:
  - Barra (valores ativos).
  - Pizza (distribuição de status).
- Exportação dos dados processados para um novo arquivo Excel.
- Salvamento do resumo de estatísticas em formato CSV.

## Requisitos

- Python 3.x
- Bibliotecas:
  - `pandas`
  - `matplotlib`
  - `openpyxl` (para leitura e escrita de arquivos Excel)
