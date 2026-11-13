# Painel de economias de esgoto 2026

Painel em HTML que consolida as bases mensais de faturamento (exportadas em .xlsx) e mostra quantas economias passaram a pagar esgoto em 2026, mês a mês, comparando com dezembro de 2025.

## Como usar

1. Abra o painel no Chrome ou no Edge.
2. Clique em **Carregar arquivos** (um .xlsx com todas as abas, ou vários arquivos) ou em **Carregar pasta**.
3. Aguarde a leitura. A base fica guardada no navegador e abre sozinha nas próximas vezes.
4. Para atualizar, carregue os arquivos novos: eles substituem a base anterior.

## Estrutura da pasta

Conecte a pasta que tem as subpastas **Esgoto** e **Água**. A base de esgoto (e as listas "Virar") fica em Esgoto; a base de água, em Água. Os arquivos da subpasta Água não entram na base de esgoto: viram a base do painel Água, que mostra as economias a mais de água em 2026 (novas entradas e aumentos de economias), com as mesmas regras e telas do painel de esgoto, usando Qtd_eco_fat_agua e Faturamento_bruto_direta_agua.

A subpasta **Cadastro** recebe o cadastro de dez/25. Ele define o grupo de cada matrícula que não faturava em dez/25 (novas na base, recuperadas inativas/desativadas, cortadas ou ativas sem faturar e, no esgoto, as que passaram a faturar esgoto) e alimenta o cartão de conferências. As regras completas ficam na aba **Dados** do painel.

A mesma subpasta pode ter também o **cadastro atual** (ex.: `Base Cadastro - 11.11.2026.xlsx`, com a coluna ULTIMO_CONTRATO). O arquivo com "dez" ou "2025" no nome é o de dez/25; o outro é o atual. Com o cadastro atual carregado, aparece o botão **Só último contrato ativo**: ligado, só contam as matrículas cujo último contrato (ULTIMO_CONTRATO = SIM) está com SIT_CONTRATO = ATIVO no cadastro atual; as outras e as que não estão nele saem de toda a contagem. No painel Esgoto há também o botão **Só AGUA E ESGOTO**: ligado, só contam as matrículas cujo último contrato está com TIPO_FATURAMENTO = AGUA E ESGOTO no cadastro atual (o saldo das entradas que continuam com esgoto). Os dois filtros podem ser usados juntos.

A aba **Arquivos do faturamento** foi retirada na v1.35: o painel usa só a base do estudo (os números do BI). Se existir uma subpasta Faturamento, ela é ignorada.

```
Nova pasta/
  Esgoto/        base de esgoto e listas "Virar"
  Água/          base de água
  Cadastro/      cadastro de dez/25 (Base_Cadastro_Dezembro-2025.xlsx) e cadastro atual (Base Cadastro - 11.11.2026.xlsx)
```

## Listas de comparação

Planilhas com "Virar" no nome (por exemplo, "Julho - Virar para Água e Esgoto.xlsx") são tratadas como listas das matrículas que deveriam passar a faturar esgoto no mês do nome. Coloque-as na mesma pasta da base. Elas não entram nos totais: o painel só compara, no cartão "Comparativo com as listas de virar para Água e Esgoto", em que mês cada matrícula começou a faturar esgoto.

## Privacidade

Os arquivos são lidos só no seu computador. Nenhum dado é enviado para a internet nem para este repositório. Nunca suba arquivos .xlsx ou .csv da base para cá.

## Colunas esperadas em cada aba

NUM_LIGACAO, NOM_CLIENTE, CIDADE, SUB_CATEGORIA, SIT_LIG, SIT_CONTRATO, Referência, ULTIMO_CONTRATO, COD_GRUPO, TIPO_FATURAMENTO, Faturamento_bruto_direta_agua, Faturamento_bruto_direta_esgoto, Qtd_eco_fat_agua, Qtd_eco_fat_esgoto, Eco_total.

Entram todos os contratos da matrícula, inclusive os anteriores a uma troca de titularidade. Cada matrícula conta uma vez por mês: se houver mais de uma linha dela no mesmo mês, vale a do último contrato (ULTIMO_CONTRATO = SIM). Abas sem a coluna NUM_LIGACAO são ignoradas.
