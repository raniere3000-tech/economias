# Painel de economias de esgoto 2026

Painel em HTML que consolida as bases mensais de faturamento (exportadas em .xlsx) e mostra quantas economias passaram a pagar esgoto em 2026, mês a mês, comparando com dezembro de 2025.

## Como usar

1. Abra o painel no Chrome ou no Edge.
2. Clique em **Carregar arquivos** (um .xlsx com todas as abas, ou vários arquivos) ou em **Carregar pasta**.
3. Aguarde a leitura. A base fica guardada no navegador e abre sozinha nas próximas vezes.
4. Para atualizar, carregue os arquivos novos: eles substituem a base anterior.

## Listas de comparação

Planilhas com "Virar" no nome (por exemplo, "Julho - Virar para Água e Esgoto.xlsx") são tratadas como listas das matrículas que deveriam passar a faturar esgoto no mês do nome. Coloque-as na mesma pasta da base. Elas não entram nos totais: o painel só compara, no cartão "Comparativo com as listas de virar para Água e Esgoto", em que mês cada matrícula começou a faturar esgoto.

## Privacidade

Os arquivos são lidos só no seu computador. Nenhum dado é enviado para a internet nem para este repositório. Nunca suba arquivos .xlsx ou .csv da base para cá.

## Colunas esperadas em cada aba

NUM_LIGACAO, NOM_CLIENTE, CIDADE, SUB_CATEGORIA, SIT_LIG, SIT_CONTRATO, Referência, ULTIMO_CONTRATO, COD_GRUPO, TIPO_FATURAMENTO, Faturamento_bruto_direta_agua, Faturamento_bruto_direta_esgoto, Qtd_eco_fat_agua, Qtd_eco_fat_esgoto, Eco_total.

Entram todos os contratos da matrícula, inclusive os anteriores a uma troca de titularidade. Cada matrícula conta uma vez por mês: se houver mais de uma linha dela no mesmo mês, vale a do último contrato (ULTIMO_CONTRATO = SIM). Abas sem a coluna NUM_LIGACAO são ignoradas.
