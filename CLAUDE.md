# Painel de entradas de economias de esgoto 2026

Contexto para continuar o desenvolvimento no Claude Code. Este arquivo resume tudo o que foi decidido até agora.

## Regras de trabalho

- **Pergunte antes de executar.** Quando houver dúvida sobre regra de negócio, layout ou formato, faça as perguntas necessárias antes de mudar o código. O dono do painel prefere responder perguntas a receber algo diferente do que pediu.
- **Valide os números.** Confira toda mudança de cálculo com `tests/conferir_totais.py`, que recalcula os números em Python de forma independente.
- **Mantenha um arquivo só.** O painel é um único `index.html`, sem build e sem dependências além das Google Fonts.
- **Escreva em português do Brasil.** Os textos da interface e os comentários do código ficam em português.
- **Aumente a versão a cada atualização.** A constante `VERSAO_SITE` no `index.html` aparece no cabeçalho (`<span id="versao">`). Suba a versão em toda mudança publicada, para identificar a versão em uso e poder voltar à anterior pelo histórico do Git.

## Objetivo

Desde a v1.23 o painel tem dois painéis com as mesmas regras: **Esgoto** e **Água** (economias a mais de água: novas entradas e aumentos de economias, mês a mês, contra dez/25). O texto abaixo fala de esgoto; para água vale o mesmo trocando as colunas (ver "Painel Água").

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
tests/conferir_totais.py    recalcula os números do Resumo em Python para conferência (e o comparativo com listas Virar)
tests/gerar_lista_teste.py  gera uma lista "Virar para Água e Esgoto" a partir da base sintética
tests/conferir_faturamento.py  totais por mês dos arquivos do painel de faturamento (aba Arquivos do faturamento)
```

Publicação: GitHub Pages, branch `main`, pasta raiz. Para atualizar, basta substituir o `index.html`.

## Dados de entrada

Cada aba precisa ter, com o cabeçalho na linha em que aparece `NUM_LIGACAO` (a leitura procura nas primeiras 30 linhas), as colunas abaixo. A comparação dos nomes ignora acentos e maiúsculas.

`NUM_LIGACAO, NOM_CLIENTE, CIDADE, SUB_CATEGORIA, SIT_LIG, SIT_CONTRATO, Referência, ULTIMO_CONTRATO, COD_GRUPO, TIPO_FATURAMENTO, Faturamento_bruto_direta_agua, Faturamento_bruto_direta_esgoto, Qtd_eco_fat_agua, Qtd_eco_fat_esgoto, Eco_total`

- **Linhas consideradas (v1.18):** entram **todos os contratos** da matrícula, sem filtro de `ULTIMO_CONTRATO`. Com troca de titularidade, os meses anteriores vêm só com o contrato antigo (NÃO) e precisam contar; até a v1.17 o filtro SIM apagava esses meses, e a matrícula parecia nova na base no mês do contrato novo. **Cada matrícula conta uma vez por mês:** com mais de uma linha da mesma matrícula na mesma referência, vale a do último contrato (SIM); entre contratos iguais, a de mais `Qtd_eco_fat_esgoto`; no empate, a última lida (`build()`). O cadastro (nome, cidade…) é o da referência mais recente, preferindo o último contrato. O rodapé mostra as linhas de contratos anteriores e as repetidas descartadas. A coluna "Último contrato" do Consolidado mostra SIM/NÃO do contrato usado no mês (`D.ult`).
- **Referência:** pode vir como data real do Excel (número de série) ou como texto `dd/mm/aaaa`. É normalizada para `AAAA-MM`.
- **Abas ignoradas:** abas sem `NUM_LIGACAO` são ignoradas e listadas no rodapé.
- **Listas de comparação (v1.15):** arquivo com "virar" no nome (ex.: "Julho - Virar para Água e Esgoto.xlsx") é uma lista de matrículas que deveriam passar a faturar esgoto no mês do nome. O worker guarda só as matrículas (com nome, cidade, sub-categoria e economias da própria lista) em `d.listas`, agrupadas por mês, e **nunca soma essas linhas à base**. O mês vem do nome do arquivo (janeiro a dezembro, sem acento) e, desde a v1.16, é procurado entre as referências de 2026 da base pelo mês, sem depender do ano. Na v1.15 o ano vinha do nome ou da Referência da lista: sem Referência ou com outro ano, o mês previsto não era achado e todas as matrículas que começaram a faturar caíam em "antes" (o usuário viu 706 de 770 em julho). Na v1.17 a página também tira o mês do nome do arquivo (`mesDoNome()` em `montaListas()`), porque uma base guardada no IndexedDB por versão anterior não tem o campo `mes` das listas (o usuário abriu o painel com a base lida pela v1.15 e viu o aviso de mês não identificado). Sem mês no nome, vale a Referência da lista; se nem assim o mês estiver na base, o cartão mostra um aviso e essas matrículas ficam na situação 6 ("mês previsto não identificado"), nunca em "antes". Também não filtram `ULTIMO_CONTRATO` (v1.18) e não exigem a coluna Referência.
- **Estrutura da pasta (v1.22):** a pasta conectada (hoje "Nova pasta", dentro de "Painel Análise Esgoto" no OneDrive) tem as subpastas **Esgoto** (base de esgoto e listas Virar) e **Água** (base de água). Arquivo dentro de uma subpasta chamada "Água" (sem acento ou maiúscula também vale) **não entra na base de esgoto**: o caminho vem de `f.caminho` (montado em `coletaXlsx`) ou de `webkitRelativePath`. A lista desses arquivos fica em `st.aguaArquivos` e é guardada em `D.meta.agua`.
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
| `pri[m]` | Primeiro mês em que a matrícula aparece faturando esgoto na base (dez/25 incluído); `-1` se nunca. Coluna `PRIMEIRO_MES_FATURANDO_ESGOTO` nos analíticos e no CSV do Consolidado. |
| `cmp` | Comparativo com as listas (`montaListas()`): por lista, o mês previsto `rL` e a situação de cada matrícula pelo `pri[m]`: 0 começou no mês previsto, 1 antes, 2 depois, 3 já faturava em dez/25, 4 ainda não faturou, 5 não está na base, 6 começou a faturar mas o mês previsto não foi identificado; `sitMat` guarda a situação por matrícula da base. |
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

## Painel Água (v1.23)

- **Regras (confirmadas pelo usuário):** as mesmas do esgoto, com `Qtd_eco_fat_agua` no lugar de `Qtd_eco_fat_esgoto` e `Faturamento_bruto_direta_agua` no lugar de `Faturamento_bruto_direta_esgoto`. Faturando água = `Qtd_eco_fat_agua > 0`, sem olhar o `TIPO_FATURAMENTO` (decidido com o usuário depois da v1.23: mesmo padrão de contagem por economias faturadas do esgoto, sem restringir ao tipo AGUA); entradas e aumentos mês a mês com total no fim; nova entrada = quem não faturava água em dez/25 (matrículas novas ou que estavam cortadas em dez/25). Mesmas colunas e meses da base de esgoto.
- **Regra do teto (v1.24, só no painel Água, `SVS.agua.regra = 'teto'`):** definida pelo usuário, que achou alto o "já existiam na base" da água (muitos cortes e religações). Cada matrícula tem um teto de economias: o faturado em dez/25 para quem já faturava, 0 para quem não faturava ou não estava na base. Em cada mês, só soma o que passar do teto, e o teto sobe (3 → 5 soma 2; volta a 3 e a 5, não soma; vai a 6, soma 1). Cortada e religada com as mesmas economias não soma. **Três grupos:** 1 novas na base (não estavam na base em dez/25; `D.nova` na água = ausente em dez/25 e faturou em 2026), 2 recuperadas da base de dez/25 (estavam em dez/25 sem faturar; contam uma vez) e 3 aumentos (o que passa do teto depois, inclusive em novas e recuperadas). O esgoto segue a regra 'mes' (entrada só na primeira vez, aumento mês a mês), a pedido do usuário; perguntar se quer o teto também no esgoto.
- **Resultado por matrícula e mês (v1.24):** `derive()` grava `D.ent` (economias novas contadas), `D.cat` (grupo: 1, 2, 3) e `D.mv` (movimento: 1 passou a faturar, 2 passou a faturar mais economias — somam —, 3 voltou a faturar, 4 deixou de faturar, 5 menos economias, 6 voltou a economias que já tinha). `fluxos()`, o Resumo (`GR`: grupos e cores do painel) e o analítico de movimentos só leem esses arrays; o analítico ganhou a coluna `GRUPO`. Na base de casos do teste (`casos_agua.py`, cinco matrículas): 12 economias = 4 novas na base + 2 recuperadas + 6 aumentos, em 4 matrículas.
- **Duas bases (`DS.esgoto`, `DS.agua`):** cada uma lida num worker próprio (`lerBase(files, servico)`), a de água com os arquivos da subpasta Água. `D` aponta para a do painel escolhido (`ativa()`). `prepara(d, servico)` define `d.Q`/`d.F` (economias e faturamento do serviço: `qe`/`fe` no esgoto, `qa`/`fa` na água) e `d.qk`; todo o cálculo (`derive`, `fluxos`, Consolidado, analíticos, comparativo) usa `D.Q`/`D.F`, nunca `D.qe`/`D.fe` direto. No worker, `build(servico)` desempata linhas repetidas pelas economias do serviço. Listas Virar só no esgoto.
- **Armazenamento:** a base de água fica na chave `agua` do IndexedDB. Ela é relida (ou apagada, se a subpasta Água estiver vazia) só quando a leitura vem de uma pasta; "Carregar arquivo" mexe só na base de esgoto.
- **Textos:** escritos para esgoto. No painel Água, `T()` troca "esgoto/Esgoto/ESGOTO" por "água/Água/AGUA" (palavra inteira) e `traduzTela()` aplica isso aos textos e aos atributos `title`/`aria-label` de `#view-resumo` e `#view-cons`, guardando o original (volta ao trocar de painel). O corpo da tabela do Consolidado não passa por ela, para não mexer em dados como "AGUA E ESGOTO"; ali a situação e as dicas usam `T()` direto. Nomes de colunas calculadas nos downloads usam `HC()`, que troca ESGOTO por AGUA também dentro do nome (`PRIMEIRO_MES_FATURANDO_AGUA`); colunas da base (`Faturamento_bruto_direta_esgoto`…) não mudam. Arquivos: `analitico_agua_…`, `consolidado_agua.csv`. Ao escrever texto novo, evite construções que ficam erradas com "água" (ex.: "do esgoto").
- **Sem base de água:** o painel Água mostra `#view-agua` com os arquivos achados na subpasta (ou a orientação de onde colocar).
- **Teste:** a base de água de teste é a base de esgoto com as colunas de água e esgoto trocadas; o painel Água tem de dar os mesmos números do esgoto da base original (2.430 economias, 1.516 matrículas, set/26 2.164 economias, R$ 271.628 e R$ 247.441). `python tests/conferir_totais.py --agua base_agua.xlsx` confere a água.

## Aba "Arquivos do faturamento" (v1.25)

- **Pedido do usuário:** uma aba nova, com os painéis Esgoto e Água, que faz as mesmas contas sobre os arquivos do painel de faturamento (repositório analise-faturamento-agua-esgoto, mesma conta do GitHub) e compara com a base do estudo. Escolhas: "mesmo painel + comparação" e "as mesmas regras de hoje" (água 'teto', esgoto 'mes').
- **Pasta:** subpasta **Faturamento** dentro da pasta conectada, ao lado de Esgoto e Água, com os mesmos arquivos do painel de faturamento: a fatura e um consumo por mês, com o mês no nome (ex.: "Consumo 12-2025.csv"). Aceita .csv (`;`, UTF-8 ou Latin-1) e .xlsx. Cronograma, RF e "avulso" são ignorados (pelas colunas ou pelo nome). A mesma cópia (nome + tamanho) em duas pastas conta uma vez.
- **Leitura (worker, `processFat()`, `FatStore`):** fatura = colunas Rubrica, Valor Parcela e Data de Vencimento; vale só VALOR DE AGUA e VALOR DE ESGOTO, no mês da Referencia de Leitura; faturamento = soma do Valor Parcela. Consumo = Leitura Atual, Consumo Medido e Consumo Faturado; mês do nome do arquivo (MM-AAAA ou mes-AA), da pasta ou de Mes/Ano Lancamento. Economias da ligação no mês = Residencial + Comercial + Industrial + Publica, quando o Consumo Faturado > 0 e a fatura tem a rubrica do serviço (`qa`/`qe`), **uma vez por ligação**. O painel de faturamento soma por linha de fatura: isso fica em `qaL`/`qeL`, só para a comparação. Cadastro da fatura mais recente (Nome da Localidade = cidade, Categoria = sub-categoria). `ULTIMO_CONTRATO` = SIM e tipo vazio.
- **Tela:** toggle "Base do estudo | Arquivos do faturamento" (`st.fonte`), antes de Esgoto | Água (`st.painel` continua sendo o serviço). `DS.fat_esgoto` e `DS.fat_agua` (`preparaFat()`) vêm da mesma leitura, guardada na chave `fat` do IndexedDB (relida ou apagada só em leitura de pasta). `chave()` escolhe a base; downloads com `ARQ()` (`consolidado_faturamento_agua.csv`…). Rodapé próprio (nomes de rubrica em `data-literal`, que `traduzTela()` não traduz).
- **Cartão "Comparação com a base do estudo"** (`renderCmpFat()`, só nessa aba): por mês, economias faturando, faturamento, economias novas faturando e matrículas faturando, estudo × faturamento, com a diferença, e a linha cinza das economias somadas por linha de fatura (a conta do painel de faturamento). Filtro de cidade vale pelo nome nas duas bases; o de sub-categoria não entra.
- **Teste:** a pasta sintética do painel de faturamento (`dados_sinteticos.gera_pasta(..., n_ligacoes=600, n_grupos=6, semente=11, mes_atual=(9,2026), formato_fatura='csv', n_meses=10, com_dre=True)`) dá dez/25 500 matrículas e 799 economias, R$ água 84.219 e esgoto 67.375; água (teto) +73 economias (28 novas na base + 45 recuperadas) e +41 matrículas; esgoto +73 e +41. `python tests/conferir_faturamento.py <pasta Faturamento>` mostra os totais por mês (por ligação e por linha) e `--base saida.xlsx` gera a base no formato do estudo para o `conferir_totais.py`.

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
- **Ícone da aba (v1.21):** SVG embutido como data URI no `<head>` (sem arquivo separado): quadrado arredondado com o gradiente do cabeçalho e uma gota d'água no verde-água `#5ee7c8`, com uma onda marinho. A gota ocupa quase todo o quadrado para ficar legível em 16 px.
- **Toolbar:**
  - Painéis **Esgoto | Água** (v1.22, `st.painel`, `showPainel()`): o painel muda o título, o selo e o `<title>` ("Entradas de economias de água 2026"). Os filtros recomeçam em "Todas" ao trocar de painel.
  - Subabas Resumo e Consolidado nos dois painéis (somem no painel sem base).
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
   - **Analítico (v1.13):** no topo do cartão, o seletor de mês ("Todos os meses" ou um mês) e o botão "⬇ Baixar analítico" (`exportAnalitico()`) geram um CSV (`;`, BOM, vírgula decimal) com uma linha por matrícula e mês com movimento de esgoto, ordenado por mês e matrícula, respeitando os filtros de cidade e sub-categoria. Colunas: `REFERENCIA, NUM_LIGACAO, NOM_CLIENTE, CIDADE, SUB_CATEGORIA, SIT_LIG, SIT_CONTRATO, TIPO_FATURAMENTO_ATUAL, ORIGEM, MOVIMENTO, SOMA_NO_TOTAL, ECO_ESGOTO_MES_ANTERIOR, ECO_ESGOTO_MES, ECONOMIAS_CONTADAS, CONTA_COMO_MATRICULA, ECO_TOTAL_MES, FAT_ESGOTO_MES`. Movimentos: passou a faturar esgoto, passou a faturar mais economias (os dois somam), voltou a faturar, deixou de faturar e passou a faturar menos economias (não somam). Com `SOMA_NO_TOTAL = SIM`, a soma de `ECONOMIAS_CONTADAS` é igual às economias novas faturando do Resumo, e `CONTA_COMO_MATRICULA = SIM` dá as matrículas novas faturando. Na base sintética: todos os meses 2.430 economias e 1.516 matrículas; jul/26 298 e 189.
   - **Analítico de todas as economias (v1.14):** o seletor "Analítico" escolhe entre "Movimentos do mês" (o arquivo acima) e "Todas as economias faturando" (`exportTodas()`): uma linha por matrícula e mês com economias de esgoto faturadas ou valor de esgoto, de dez/25 em diante (ao escolher esse tipo, o mês começa no último, porque todos os meses podem passar do limite de 1.048.576 linhas do Excel). Colunas: `REFERENCIA`, cadastro, `TIPO_FATURAMENTO_ATUAL, SITUACAO_NO_MES` (Faturando na base, Já faturava, Nova faturando em 2026, Já faturava e fatura mais economias, Voltou a faturar, Valor de esgoto sem economias de esgoto), `ORIGEM, MES_PASSOU_A_FATURAR_ESGOTO, ECO_ESGOTO_MES, ECO_NOVAS_FATURANDO, ECO_TOTAL_MES, FAT_ESGOTO_MES, FAT_ESGOTO_NOVAS`. Por mês, a soma de `ECO_ESGOTO_MES` é a linha "Economias com esgoto faturadas", `FAT_ESGOTO_MES` o faturamento de esgoto total e `FAT_ESGOTO_NOVAS` o das economias novas faturando. Na base sintética, set/26: 2.164 economias, R$ 271.628 e R$ 247.441.
   - **Colunas novas (v1.15):** os dois analíticos ganharam `PRIMEIRO_MES_FATURANDO_ESGOTO` e `LISTAS_VIRAR` (ex.: "jul/26: Começou no mês previsto").
   - **Primeiro mês faturando em todo download (v1.20, pedido do usuário):** `PRIMEIRO_MES_FATURANDO_ESGOTO` fica sempre logo depois dos dados da matrícula: nos analíticos depois de `SIT_CONTRATO`, no CSV do Consolidado depois de `ESGOTO_2026`, no comparativo depois de `SITUACAO`. Todo download novo deve trazer essa coluna.
4. **Cartão "Comparativo com as listas de virar para Água e Esgoto"** (v1.15, só aparece quando há listas): uma linha por lista com as matrículas na lista, que começaram no mês previsto, antes, depois, já faturavam em dez/25, ainda não faturaram e não estão na base, mais as economias de esgoto da própria lista e as faturadas na base no mês previsto. Os números são botões que abrem o Consolidado com o filtro "Lista virar". Respeita os filtros de cidade e sub-categoria (fora da base, pela cidade da lista). **Seletor Matrículas / Economias (v1.20, `st.lsv`):** na visão Economias as mesmas situações somam `ecoLista()`: de quem começou a faturar, a `Qtd_eco_fat_esgoto` do primeiro mês faturando; de quem ainda não faturou, o `Eco_total` do último mês em que aparece; fora da base, as economias da própria lista (`Qtd_eco_fat_esgoto`, ou `Eco_total`); e há a coluna "De esgoto faturadas na base no mês previsto". O CSV traz essa conta em `ECONOMIAS_COMPARATIVO`. O botão "⬇ Baixar comparativo" (`exportListas()`) gera um CSV com uma linha por matrícula das listas, inclusive as que não estão na base. Na base sintética da v1.18 com as listas de `gerar_lista_teste.py`: jul/26 225 matrículas (113 no mês, 40 antes, 16 depois, 15 já faturavam, 26 ainda não, 15 fora da base) e 351 economias (175, 63, 29, 23, 46, 15; 252 de esgoto faturadas em jul); set/26 215 matrículas (106, 53, 0, 15, 26, 15) e 328 economias (157, 91, 0, 22, 43, 15; 250).
5. **Cartão "Faturamento de esgoto por mês":** colunas empilhadas entre o que já existia (`#BCD0E1`) e as economias novas (marinho), de dez/25 a set/26, com o total em reais em cima. A tabela alinhada traz existente, novas, total, participação das novas e economias com esgoto faturadas.
6. **Tabelas Por cidade e Por sub-categoria:**
   - meses, Total 2026 e as que deixaram de faturar (só para conhecimento);
   - na tabela por cidade, o seletor **"Cidades na tabela"** (`st.cidTab`, mesmo componente dos filtros) escolhe as linhas mostradas; o Total soma só as cidades escolhidas e os indicadores não mudam;
   - um seletor Economias / Faturamento (R$), que fica no cartão de cidade e vale para as duas;
   - ordenação ao clicar nos títulos.
7. **Rodapé** com a origem da base, as linhas desconsideradas, as listas de comparação (fora dos totais) e a regra de medição.

### Consolidado

Uma linha por matrícula. As colunas cadastrais, "Esgoto em 2026" e "1º mês faturando esgoto" (`pri[m]`, v1.19, pedido do usuário; "–" se nunca faturou) ficam fixas, seguidas das 8 colunas por referência: `ULTIMO_CONTRATO, COD_GRUPO, TIPO_FATURAMENTO, Fat. água, Fat. esgoto, Eco. água, Eco. esgoto, Eco. total`.

- **Busca:** por matrícula ou nome.
- **Filtro de movimento:**
  - passou a faturar esgoto em 2026: as mesmas ligações do Resumo (`ligMes > 0`);
  - passou a faturar esgoto num mês: `ligMes` igual ao mês escolhido, igual à linha do tempo;
  - deixou de faturar em algum mês;
  - deixou de pagar num mês específico (troca de tipo);
  - divergentes;
  - aumentou, reduziu ou sem mudança.
- **Filtro "Lista virar"** (aparece quando há listas): todas da lista na base ou uma situação do comparativo.
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

- **Base real de água:** ainda não chegou (o usuário vai mandar para análise quando tiver). Quando chegar, conferir os números com ele e a dúvida abaixo.

- **Testes automatizados:** ainda não existem testes nem CI de interface. A validação foi feita com Playwright e Chromium sobre a base sintética, comparando com `tests/conferir_totais.py`.

## Como validar uma mudança

```bash
python tests/gerar_base_teste.py 3000 base_teste.xlsx     # gera a base sintética (requer openpyxl)
python tests/conferir_totais.py base_teste.xlsx           # números de referência
python tests/conferir_totais.py --agua base_agua.xlsx     # painel Água (base de água de teste = base de esgoto com as colunas trocadas)
python tests/gerar_lista_teste.py base_teste.xlsx 2026-07 "Julho - Virar para Agua e Esgoto.xlsx"
python tests/conferir_totais.py base_teste.xlsx "Julho - Virar para Agua e Esgoto.xlsx"   # comparativo
# no teste com Playwright, use nomes de arquivo sem acento: o setInputFiles daqui descarta arquivos com acento no nome
# abra o index.html no Chrome/Edge, carregue base_teste.xlsx e compare com a saída acima
```

Valores de referência com a base sintética padrão (3000 matrículas, semente 7). Desde a v1.2 o gerador inclui ligações com AGUA E ESGOTO e economias de esgoto zeradas em mar/26 e abr/26 (1 a cada 37 matrículas); desde a v1.18, trocas de titularidade (1 a cada 41 matrículas só com o contrato anterior, NÃO, de dez/25 a fev/26), com 572 linhas repetidas no mesmo mês descartadas. Com o filtro antigo (só SIM) essa base daria 150 matrículas na base e 157 matrículas novas na base, em vez de 154 e 142:

Base em dez/25: **154** matrículas e **245** economias faturando esgoto. **2.430** economias novas faturando (2.183 já existiam na base + 247 de matrículas novas), **1.516** matrículas novas faturando (1.374 já existiam na base + 142 novas), faturamento das novas em set/26 de **R$ 247.441**, no ano de **R$ 1.402.791**, economias que deixaram de faturar **536** e que voltaram a faturar **25**.

Nunca suba arquivos `.xlsx` ou `.csv` para o repositório: o `.gitignore` já bloqueia, e os dados reais têm nomes de clientes.
