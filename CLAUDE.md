# Painel de entradas de economias de esgoto 2026

Contexto para continuar o desenvolvimento no Claude Code. Este arquivo resume tudo o que foi decidido até agora.

## Regras de trabalho

- **Pergunte antes de executar.** Quando houver dúvida sobre regra de negócio, layout ou formato, faça as perguntas necessárias antes de mudar o código. O dono do painel prefere responder perguntas a receber algo diferente do que pediu.
- **Valide os números.** Confira toda mudança de cálculo com `tests/conferir_totais.py`, que recalcula os números em Python de forma independente.
- **Mantenha um arquivo só.** O painel é um único `index.html`, sem build e sem dependências além das Google Fonts.
- **Escreva em português do Brasil.** Os textos da interface e os comentários do código ficam em português.
- **Aumente a versão a cada atualização.** A constante `VERSAO_SITE` no `index.html` aparece no cabeçalho (`<span id="versao">`). Suba a versão em toda mudança publicada, para identificar a versão em uso e poder voltar à anterior pelo histórico do Git.

## Objetivo

Medir quantas **economias** passaram a faturar esgoto em 2026, mês a mês, comparando com dezembro de 2025, e quanto isso representa no faturamento de esgoto. O foco são as economias, não as matrículas. As saídas aparecem só para conhecimento e **nunca são descontadas** do total.

## Origem

A base vinha do sistema em .xlsx. Como tinha mais linhas do que o Excel suporta, foi exportada em partes: 3 abas, de dezembro de 2025 a fevereiro de 2026, de março a maio de 2026 e de junho a setembro de 2026. As referências vão até setembro de 2026.

- **Primeira tentativa:** consolidar no próprio Excel, com PROCX, EMPILHARV e SOMASES. Travou por causa do volume.
- **Restrições:** o usuário não quer Power Query, e o sistema só exporta .xlsx (CSV não é opção).

O painel nasceu para substituir essa consolidação. Ele é publicado no GitHub Pages, num repositório novo em outra conta do GitHub.

## Estrutura do repositório

```
index.html                  painel completo (HTML + CSS + JS + worker inline)
README.md                   instruções de uso
.gitignore                  bloqueia *.xlsx, *.xls, *.csv e ~$* (dados nunca vão para o repositório)
CLAUDE.md                   este arquivo
tests/gerar_base_teste.py   gera base sintética no formato da exportação
tests/conferir_totais.py    recalcula os números do Resumo em Python para conferência
```

Publicação: GitHub Pages, branch `main`, pasta raiz. Para atualizar, basta substituir o `index.html`.

## Dados de entrada

Cada aba precisa ter, com o cabeçalho na linha em que aparece `NUM_LIGACAO` (a leitura procura nas primeiras 30 linhas), as colunas abaixo. A comparação dos nomes ignora acentos e maiúsculas.

`NUM_LIGACAO, NOM_CLIENTE, CIDADE, SUB_CATEGORIA, SIT_LIG, SIT_CONTRATO, Referência, ULTIMO_CONTRATO, COD_GRUPO, TIPO_FATURAMENTO, Faturamento_bruto_direta_agua, Faturamento_bruto_direta_esgoto, Qtd_eco_fat_agua, Qtd_eco_fat_esgoto, Eco_total`

- **Linhas consideradas:** só entram as linhas com `ULTIMO_CONTRATO = SIM`. As demais aparecem no rodapé como "contratos anteriores desconsiderados".
- **Referência:** pode vir como data real do Excel (número de série) ou como texto `dd/mm/aaaa`. É normalizada para `AAAA-MM`.
- **Valores repetidos:** se a mesma matrícula aparecer duas vezes na mesma referência, vale a última linha lida.
- **Abas ignoradas:** abas sem `NUM_LIGACAO` são ignoradas e listadas no rodapé.
- **Volume esperado:** cerca de 3 milhões de linhas no total. No teste, 600 mil linhas foram lidas em cerca de 7 segundos.

## Arquitetura (index.html)

### Leitura (worker inline, `<script type="text/plain" id="worker-src">`)

- **Leitor de .xlsx próprio:** não usa SheetJS. O worker lê o diretório central do zip (com suporte a zip64), descompacta cada parte com `DecompressionStream('deflate-raw')` e processa o XML em fluxo, linha a linha (`</row>`). Isso evita carregar a planilha inteira na memória.
- **Textos compartilhados:** os `sharedStrings` são lidos primeiro. As abas são localizadas via `workbook.xml` e os `.rels`.
- **Acumulação:** a classe `Store` acumula as linhas em typed arrays que crescem conforme a leitura. O método `build()` gera matrizes planas com o índice `m * nR + r` (m = matrícula, r = referência ordenada):
  - `present`, `grp`, `tipo` (índice no dicionário), `fa`, `fe`, `qa`, `qe`, `et`.
  - Por matrícula: `mats`, `nomes`, `cid`, `sub`, `sl`, `sc` (índices de dicionário), com os dados cadastrais da referência mais recente.
- **Resultado:** volta para a página com `postMessage`, transferindo os buffers.

### Armazenamento no navegador

O banco IndexedDB `painel-esgoto-2026`, store `dados`, guarda duas chaves:

- **`atual`:** a base processada, que abre direto na próxima visita.
- **`pasta`:** o handle da File System Access API.

O botão "📁 Conectar pasta" usa `showDirectoryPicker`, que só existe no Chrome e no Edge. Depois disso, "↻ Atualizar" relê a pasta, inclusive as subpastas, e ignora os arquivos `~$`. Em outros navegadores, o painel usa `<input webkitdirectory>`, que não lembra a pasta. "Apagar dados do navegador" remove as duas chaves.

### Derivados calculados ao carregar (`derive()`)

| Derivado | O que é |
|---|---|
| `hasE[j]` | Faturando esgoto no mês: 1 quando `Qtd_eco_fat_esgoto > 0` (desde a v1.12). O `TIPO_FATURAMENTO` da exportação é o tipo **atual** da matrícula, não o de cada referência, por isso não entra em nenhuma conta. Ausente na referência = não faturando. |
| `div[j]` | Divergência, só em mês faturando esgoto: `1` = `Qtd_eco_fat_esgoto ≠ Eco_total`; `2` = economias de esgoto sem valor em `Faturamento_bruto_direta_esgoto`. |
| `entra[m]` | Referência em que a matrícula passa a faturar esgoto **pela primeira vez** em 2026, só para quem não faturava em dez/25; `-1` se não houver. |
| `nova[m]` | 1 para matrícula nova: aparece na base pela primeira vez em 2026 já faturando esgoto (`entra[m]` igual à primeira referência em que a matrícula aparece). |
| `ligMes[m]` | Mês em que a matrícula passa a contar como ligação que passou a faturar esgoto: o primeiro mês em que soma economias novas (troca ou aumento); `-1` se não somar. Usado no Resumo e nos filtros do Consolidado. |
| `ecoNov[m]` | Economias novas da matrícula em 2026 (troca + aumentos). |
| `anyOut[m]` | 1 se deixou de faturar esgoto em algum mês. |
| `yStat[m]` | Situação no ano (coluna "Esgoto em 2026"): 0 sem esgoto faturado, 1 já faturava, 2 nova faturando em 2026, 3 deixou de faturar, 4 já faturava e passou a faturar mais economias em 2026. 2 e 4 são exatamente as ligações com `ligMes > 0`. |

## Regras de negócio (definidas com o usuário)

Calculadas em `fluxos(k)`, onde `k` é a máscara dos filtros. A comparação é sempre com o mês anterior.

- **Faturando esgoto (regra da v1.12, definida pelo usuário):** a matrícula fatura esgoto num mês quando a `Qtd_eco_fat_esgoto` é maior que 0, qualquer que seja o tipo. O `TIPO_FATURAMENTO` da base é o atual: a matrícula pode estar como AGUA hoje e ter faturado esgoto antes. Veio da conferência de jul/26 em Miracema (planilha do usuário com 744 matrículas e 774 economias, painel com +732), em que 4 matrículas com tipo AGUA faturavam esgoto e matrículas AGUA E ESGOTO com 0 economias em dez/25 eram tratadas como já faturando.
- **Troca (entrada):** a matrícula não faturava esgoto no mês anterior (ou não existia) e passou a faturar. Conta a `Qtd_eco_fat_esgoto` do mês, **só na primeira vez** em 2026 (`r === entra[m]`). Matrícula nova que já aparece faturando esgoto entra aqui.
- **Aumento:** a matrícula faturava esgoto no mês anterior e no atual, e a `Qtd_eco_fat_esgoto` subiu. Conta a diferença; por exemplo, de 10 para 12 soma 2. **Soma ao total** de economias novas.
- **Contagem de economias (decidido na v1.3):** vale a `Qtd_eco_fat_esgoto`, a quantidade de economias que fatura esgoto de fato. Na v1.2 foi testado contar o `Eco_total` em toda ligação com AGUA E ESGOTO; o usuário achou o número muito abaixo e descartou.
- **Economias novas no mês:** troca + aumento. O total do ano é a soma dos meses, sem desconto de saídas.
- **Origem (como o Resumo mostra as economias novas):** o usuário não quer ver a separação "X por troca e Y por aumento". As economias novas aparecem divididas pela origem da ligação: **ligações que já estavam na base** (passaram de AGUA para AGUA E ESGOTO) e **matrículas novas já com AGUA E ESGOTO** (`nova[m]`). O aumento continua somando, no grupo da ligação. Matrícula que não aparece em dez/25 e surge em 2026 já com esgoto conta como nova.
- **Ligações que passaram a faturar esgoto (regra da v1.6):** são as mesmas matrículas que somam economias novas, por troca ou por aumento. Cada uma conta uma vez no ano, no primeiro mês em que soma economias, no grupo da sua origem. Ligação que já tinha esgoto em dez/25 e aumentou economias conta (caso da decisão judicial de jul/26). Troca com `Qtd_eco_fat_esgoto` 0 não soma economias e por isso também não conta ligação naquele mês. Antes da v1.6 só contava a troca de tipo, e o usuário viu +1.419 economias com só +311 ligações.
- **Saídas (só para conhecimento):** a matrícula deixou de faturar esgoto (economias de esgoto foram a 0); conta as economias do mês anterior. A redução de economias em quem continua com esgoto também entra como saída informativa.
- **Reentrada (só para conhecimento, não soma):** a ligação voltou a ter esgoto depois de sair. Isso vale para quem já tinha entrado em 2026 e também para quem já pagava em dez/25.
- **Faturamento de esgoto das economias novas:** é calculado por coorte, em cada mês em que a ligação continua faturando esgoto.
  - Ligação que entrou por troca: conta o `Faturamento_bruto_direta_esgoto` inteiro, de `entra[m]` em diante.
  - Ligação que já pagava em dez/25 e aumentou economias: conta só a parte proporcional, `fe * (qe_atual − qe_dez) / qe_atual`, quando `qe_atual > qe_dez`.
- **Faturando esgoto na base:** "matrículas faturando esgoto na base" e "economias faturando esgoto na base" (dez/25 na linha do tempo) e "Economias com esgoto faturadas" no gráfico de faturamento usam a mesma regra (`o.ligTot`, `o.ecoTot`). Na v1.11 exigiam também o tipo AGUA E ESGOTO; na v1.12 o tipo saiu.
- **Faturamento existente:** faturamento de esgoto total do mês menos o faturamento das economias novas.
- **Divergências:** aparecem numa linha discreta no Consolidado (desde a v1.7; no Resumo o usuário achou grande demais), contadas por ligação no último mês, com o link "Ver só as divergentes" que aplica o filtro.

## Vocabulário da interface (definido na v1.8)

- **"Novas" é sempre "novas faturando":** as economias e matrículas não são necessariamente novas na base, e sim novas faturando esgoto. Use "economias novas faturando esgoto" e "matrículas novas faturando esgoto", em indicadores, gráficos, linha do tempo e tabelas.
- **Não misture economias com matrículas no mesmo texto.** A divisão de origem segue o padrão do gráfico de faturamento ("Economias que já faturavam" / "Economias novas faturando em 2026"), desde a v1.10: **"Economias que já existiam na base"** e **"Economias novas faturando em 2026"** (o mesmo com matrículas) em legendas, gráficos, tabelas e dicas; na linha do tempo, "X já existiam na base" e "Y economias novas faturando". As linhas de soma se chamam "Total de economias no mês" e "Total de matrículas no mês", para não repetir o nome da parte.
- **Faturamento:** "economias que já faturavam" e "economias novas faturando em 2026".
- **Saídas:** "economias que deixaram de faturar" (só para conhecimento) e "economias que voltaram a faturar" (não somadas).
- **"Ligação" não aparece no Resumo;** a unidade de contagem é a matrícula.

## Layout atual

O estilo copia o **painel de faturamento** do usuário (repositório painel-faturamento-unificado).

- **Fontes e cores:**
  - Fontes IBM Plex Sans e IBM Plex Sans Condensed.
  - Paleta: marinho `#102840`, azul `#176b9c`, azul-claro `#64748b`, fundo `#f3f6f9`, linhas `#dbe4ec` e tint `#e7f0f6`.
  - **Positivo é marinho e azul, negativo é laranja** (`#E8710A`, texto `#C2560C`).
- **Cabeçalho:**
  - Gradiente `#0b2447 → #14375e → #1b4a76`, com o selo "Águas do Rio · Esgoto" em `#5ee7c8`.
  - Botões translúcidos com ícone e status à direita.
  - O progresso da leitura aparece dentro do próprio cabeçalho; não existe janela sobreposta.
- **Toolbar:**
  - Botão segmentado com as abas Resumo e Consolidado.
  - Filtros de Cidade e Sub-categoria em listas de seleção (Todas, Nenhuma e um selo com a quantidade marcada), no mesmo padrão do outro painel.

### Resumo

1. **Faixa de 8 indicadores:**
   - economias novas faturando esgoto em 2026, em destaque;
   - matrículas novas faturando esgoto em 2026, em destaque (títulos dos indicadores quebram em até 2 linhas);
   - média por mês;
   - melhor mês;
   - faturamento de esgoto das novas no último mês, com a participação no total;
   - faturamento das novas somado no ano;
   - faturamento de esgoto total no último mês, com a variação sobre dez/25;
   - economias que deixaram de faturar, com borda laranja.
2. **Cartão "Linha do tempo de 2026"**, recriado a partir do print do painel antigo, na paleta clara, com um seletor **Economias / Matrículas / Faturamento** (`st.lt`) que mostra uma informação por vez:
   - Economias: dez/25 com as economias faturando esgoto na base; cada mês com as economias novas faturando, quantas já existiam na base e quantas são economias novas faturando, e as que deixaram de faturar (laranja);
   - Matrículas: o mesmo com matrículas, sem as que deixaram de faturar;
   - Faturamento: dez/25 com o faturamento de esgoto do mês; cada mês com o faturamento das economias novas faturando (com a participação no total) e o faturamento de esgoto do mês;
   - última coluna "Total 2026", mais larga e com fundo tint: os mesmos itens somados no ano;
   - os blocos ficam alinhados entre as colunas com `grid-template-rows: subgrid` (3 linhas por coluna);
   - não mostra troca e aumento.
3. **Cartão "Economias novas faturando esgoto, por mês":**
   - colunas empilhadas pela origem: economias que já existiam na base em marinho e economias novas faturando em 2026 em azul;
   - **o valor exato fica sempre em cima da coluna**, e as partes aparecem dentro dela quando a altura é de pelo menos 17px;
   - embaixo, uma **tabela compacta alinhada às colunas** com as duas origens, economias novas no mês, acumulado, ligações que passaram a faturar esgoto, ligações acumuladas, saídas e reentradas.
   - **Interativo:** ao passar o mouse (ou tocar) numa coluna, ela fica destacada e uma dica mostra os dados do mês: as duas origens, economias novas, acumulado, ligações que passaram a faturar esgoto (já na base, matrículas novas, acumuladas), economias acrescentadas em ligações que já tinham esgoto e quantas ligações foram, saídas e reentradas. O mesmo vale para o gráfico de faturamento (existente, novas, total, participação e economias com esgoto). A dica é montada por `ligaDica(el)` a partir de `el.dica = { n, titulo, linhas }`, definido em `renderResumo()`.
   - O alinhamento funciona porque o SVG tem `viewBox` de 1000 de largura e área de rótulos de 180, enquanto a tabela usa `table-layout: fixed` com a primeira coluna em 18%. As demais colunas têm larguras iguais nos dois.
4. **Cartão "Faturamento de esgoto por mês":** colunas empilhadas entre o que já existia (`#BCD0E1`) e as economias novas (marinho), de dez/25 a set/26, com o total em reais em cima. A tabela alinhada traz existente, novas, total, participação das novas e economias com esgoto faturadas.
5. **Tabelas Por cidade e Por sub-categoria:**
   - meses, Total 2026 e as que deixaram de faturar (só para conhecimento);
   - na tabela por cidade, o seletor **"Cidades na tabela"** (`st.cidTab`, mesmo componente dos filtros) escolhe as linhas mostradas; o Total soma só as cidades escolhidas e os indicadores não mudam;
   - um seletor Economias / Faturamento (R$), que fica no cartão de cidade e vale para as duas;
   - ordenação ao clicar nos títulos.
6. **Rodapé** com a origem da base, as linhas desconsideradas e a regra de medição.

### Consolidado

Uma linha por matrícula. As colunas cadastrais e "Esgoto em 2026" ficam fixas, seguidas das 8 colunas por referência: `ULTIMO_CONTRATO, COD_GRUPO, TIPO_FATURAMENTO, Fat. água, Fat. esgoto, Eco. água, Eco. esgoto, Eco. total`.

- **Busca:** por matrícula ou nome.
- **Filtro de movimento:**
  - passou a faturar esgoto em 2026: as mesmas ligações do Resumo (`ligMes > 0`);
  - passou a faturar esgoto num mês: `ligMes` igual ao mês escolhido, igual à linha do tempo;
  - deixou de faturar em algum mês;
  - deixou de pagar num mês específico (troca de tipo);
  - divergentes;
  - aumentou, reduziu ou sem mudança.
- **Ordenação:** por matrícula, maior aumento ou maior redução.
- **Colunas visíveis:** escolhidas por chips.
- **Paginação:** 100 por página.
- **Destaques:** a troca de tipo fica destacada no mês (tint para entrada, laranja-claro para saída), e as economias divergentes ficam em amarelo, com o motivo no `title`.
- **Aviso de divergências** acima da contagem, quando houver.
- **Exportação CSV:** separador `;`, BOM, vírgula decimal, mais as colunas `MES_PASSOU_A_FATURAR_ESGOTO` (`ligMes`), `ECONOMIAS_NOVAS_2026` (`ecoNov`), `MESES_TROCA_PARA_ESGOTO`, `MESES_DEIXOU_DE_PAGAR` e `MESES_COM_DIVERGENCIA`. Com o filtro "Passou a faturar esgoto em 2026", a soma de `ECONOMIAS_NOVAS_2026` é igual às economias novas do Resumo.

## Preferências do usuário

- Quer ver **valores exatos** sobre as colunas dos gráficos, sem abreviação como "mil" nos rótulos. Só o eixo usa abreviação.
- Quer a **tabela compacta alinhada embaixo do gráfico** e também a **linha do tempo** em cartões, que voltou na v1.2.
- O que o Resumo precisa responder logo de cara: **entradas de economias por mês e total do ano** e **faturamento de esgoto das novas economias e do total**.
- Não pediu ranking de cidades, comparação com meta nem destaques automáticos. Pergunte antes de incluir.

## Pendências e ideias em aberto

- **Testes automatizados:** ainda não existem testes nem CI de interface. A validação foi feita com Playwright e Chromium sobre a base sintética, comparando com `tests/conferir_totais.py`.

## Como validar uma mudança

```bash
python tests/gerar_base_teste.py 3000 base_teste.xlsx     # gera a base sintética (requer openpyxl)
python tests/conferir_totais.py base_teste.xlsx           # números de referência
# abra o index.html no Chrome/Edge, carregue base_teste.xlsx e compare com a saída acima
```

Valores de referência com a base sintética padrão (3000 matrículas, semente 7). Desde a v1.2 o gerador inclui ligações com AGUA E ESGOTO e economias de esgoto zeradas em mar/26 e abr/26 (1 a cada 37 matrículas):

Base em dez/25: **154** matrículas e **245** economias faturando esgoto. **2.430** economias novas faturando (2.183 já existiam na base + 247 de matrículas novas), **1.516** matrículas novas faturando (1.374 já existiam na base + 142 novas), faturamento das novas em set/26 de **R$ 247.441**, no ano de **R$ 1.402.791**, economias que deixaram de faturar **536** e que voltaram a faturar **25**.

Nunca suba arquivos `.xlsx` ou `.csv` para o repositório: o `.gitignore` já bloqueia, e os dados reais têm nomes de clientes.
