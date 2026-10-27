# Release Notes — advpl-excel

## [Unreleased]

### Data e hora (DateTime)

- A classe `DateTime` deixou de usar as funções `DEC_*` (ponto fixo). O `DEC_DIV` imprimia `Err 001/002 Attempt to perform an illogical operation division by zero.` no console do AppServer em alguns builds (24.3.1.x), mesmo com o resultado correto. Os cálculos passaram a usar o numérico nativo (double): o número serial gravado no xml é gerado por divisão longa com inteiros (17 casas, exato) em `StrNumber` e, nos demais casos, em dois blocos de 8 casas, contornando o limite de 8 casas do `cValToChar`/`Str`.
- Correção: os segundos eram arredondados separadamente e a hora podia sair `hh:mm:60`. Agora o total de segundos é arredondado antes de separar hora, minuto e segundo; ao chegar em 24:00:00, avança para o dia seguinte.
- Correção: em alguns builds do AppServer, `STOD("19000101")-2+n` resulta em número (a subtração cai antes de 1900), e a data lida ficava numérica em `NumToDateTime`, `GetDate` e na leitura de células de data. A expressão passou a ser `STOD("19000101")+(n-2)`.
- Correção: no modo `"B"` (banco), o número serial gravado com `DateTime` perdia o alinhamento das casas decimais a partir da 9ª (faltavam zeros à esquerda), e a leitura de célula de data subtraía número do campo texto `VLR`.
- Correção: os operadores de `DateTime` com hora em texto (`oData + "01:00:00"`) tratavam a hora como data, e a data (`oData - dData`) dependia do formato de data configurado (`CToD("30/12/1899")`). O `New` passou a aceitar o número serial também como texto sem aviso de tipo (`nData` como `Variant`).

### Células mescladas: borda e desempenho

- Correção ([#24](https://github.com/saulogm/advpl-excel/issues/24)): a borda (e o preenchimento) ficava só na primeira célula de um intervalo mesclado. Agora as demais células recebem o estilo da primeira:
  - Modo `"A"`: são gravadas vazias com o estilo da primeira quando o posicionamento passa por elas, inclusive em mesclagens de várias linhas, entre linhas com `SetRowH` e na troca de planilha/`Save`. Continua sendo possível gravar à direita e abaixo da mesclagem; o estilo pode ser aplicado antes ou depois do `mergeCells`.
  - Modo `"M"`: a cópia do estilo não funcionava (condição invertida). Agora é feita célula a célula e, se as células criadas ficarem fora de ordem, o `Save` reorganiza o xml.
- Desempenho: o `mergeCells` deixou de alterar o xml da planilha a cada chamada (XPath com `[last()]`, custo quadrático); as referências são guardadas e gravadas de uma vez no `Save`. Com 10 mil mesclagens no modo `"A"`, o `mergeCells` caiu de 4,6 s para 0,5 s e o `Save` de 12,8 s para 0,2 s; no modo `"M"`, a gravação das mesclagens caiu de 6,2 s para 0,7 s, e a cópia do estilo para as células mescladas (que antes não ocorria) acrescenta cerca de 4,3 s, o custo de criar 20 mil células.
- Correção: o atributo `count` de `<mergeCells>` ficava sempre `"1"`.
- `Save2Html` lê as mesclagens pendentes junto com as do arquivo aberto.

### Altura de linhas no modo arquivo

- Correção: no modo `"A"` (arquivo), `SetRowH(nHeight, nLinha, nLinha2)` ignorava o intervalo e alterava a altura de todas as linhas seguintes. Agora aplica a altura só às linhas `nLinha` a `nLinha2` (inclusive no `SetBulkLine`); linhas do intervalo sem células são gravadas vazias com a altura. Sem linha informada, continua valendo para as próximas linhas.
- Como a gravação é sequencial, informar uma linha já gravada gera erro: chame `SetRowH` antes de gravar o intervalo.

### Dashboard em tema escuro e cores nos gráficos

- Novo método `Chart:SetCores(cFundo, cTexto, cGrade)` (`src/custom.excel.chart.tlpp`): fundo do gráfico (`c:chartSpace/c:spPr`, área de plotagem transparente), cor dos textos (`c:txPr`, também aplicada no título) e das linhas de grade. Sem chamar o método, o XML gerado não muda.
- Html (`Save2Html`): gráficos respeitam `SetCores` (fundo, moldura, título, legenda, eixos, grade e contorno das fatias).
- Correção (html): o gráfico ficava uma linha e uma coluna menor que no Excel; a âncora final do `Insert` passa a ser inclusiva, como no `xdr:to` do xlsx.
- Exemplo: aba "Dashboard" em `u_tstXlsx` (`tstDashboard`, `exemplo/custom.excel.teste.tlpp`) com cards de KPI, combinação colunas + linha, rosca e tabelas com formatação condicional, tudo por fórmulas.

### Fórmula em português e L1C1

- Novo método `oExcel:Formula(cFormula, lPtBr = .T., lL1C1 = .T.)` (classe `custom.excel.Formula`, `src/custom.excel.formula.tlpp`): fórmula escrita como no Excel em português (nomes das funções, `VERDADEIRO`/`FALSO`, erros, `;` e `,` decimal) e com referências L1C1 (`L2C3`, `L[-1]C`, linhas/colunas inteiras). A conversão para inglês/A1 é feita no `SetValue`/`Cell`, relativa à célula posicionada, permitindo reutilizar o mesmo objeto em um laço.
- `SetValue` e `Cell` aceitam o objeto no parâmetro `cFormula`.
- Correção: no `Open` (Windows com 7-Zip) a pasta de extração que sobrou de uma execução anterior com o mesmo nome temporário é apagada e o 7-Zip recebe `-y`; antes a leitura travava esperando a confirmação de substituição.
- Testes `u_tstFormula` e desempenho `u_tstFormulaPerf` (`exemplo/custom.excel.calc.teste.tlpp`) e exemplo na segunda planilha de `u_tstXlsx` (`exemplo/custom.excel.teste.tlpp`).

### Cálculo automático de fórmulas

- `SetValue(, cFormula)` sem valor: a classe calcula a fórmula no `Save`/`Save2Html` e grava o valor em cache (`<v>`), nos modos A (arquivo), M (memória) e B (banco). `GetValue` e o html passam a mostrar o resultado.
- Novo motor de fórmulas em AdvPL (`src/custom.excel.calc*.tlpp`, namespace `custom.excel.calc`): parser com precedência do Excel, referências entre planilhas, nomes definidos, referências estruturadas de tabela, matrizes, LET/LAMBDA e 539 funções do Excel (504 calculadas; as dependentes do ambiente, como CUBE*/RTD/WEBSERVICE, ficam para o Excel).
- Fórmulas não calculáveis ficam sem valor e o workbook recebe `calcPr fullCalcOnLoad="1"`.
- Novos métodos: `SetCalcFormula`, `AddFuncFormula`, `SetCalcLocal`.
- Html (`Save2Html`): as fórmulas vão para a página (`data-f`) e são recalculadas no navegador quando um input muda, só as dependentes, inclusive entre planilhas, nomes definidos e `Tab[Coluna]`. Motor JS próprio com as funções mais usadas e as demais via formula.js (CDN, `oHtml:SetFormulaJs(cUrl)` para ambiente sem internet). API `xlsxHtml.recalcula/calcula/valorFormula/formata/addFuncao` e evento `xlsxhtml:recalculo`.
- Html: filtro e agrupamento recalculam `SUBTOTAL`/`AGGREGATE` só com as linhas visíveis (regras do Excel), os gráficos passam a ler as referências da série e se redesenham após input, recálculo ou filtro (sem as linhas ocultas) e a formatação condicional é reavaliada. Exemplo `u_tstHtml` com total calculado, subtotal visível e resumo ligado às vendas.
- Html: `data-v` grava o número com até 15 dígitos significativos (antes 8 decimais); no modo B a fórmula deixa de vir com escape XML.
- Correções: subtotais MAX/MIN/MULT/STDEV/VAR do `Alias2Tab`; `LocRef` aceita o prefixo `Plan!`; no modo B, células só com fórmula não são mais descartadas no `Save`.
- Documentação: `skills/advpl-excel-xlsx/references/formulas.md`.

### Migração para TLPP — `src/custom.tools.excel.tlpp`

Criado novo fonte TLPP com migração completa da biblioteca de geração de planilhas Excel XLSX.

#### Novo arquivo: `src/custom.tools.excel.tlpp`

- **Namespace** `custom.tools.excel` aplicado a todas as classes
- **9 classes renomeadas** (prefixos `YExcel_`/`YExcel` removidos):

  | Classe legada | Nova classe |
  |---|---|
  | `YExcel` | `Xlsx` |
  | `YExcel_Style` | `Style` |
  | `YExcel_StyleRules` | `StyleRules` |
  | `YExcel_RegraLinha` | `RegraLinha` |
  | `YExcelTag` | `Tag` |
  | `YExcel_Table` | `Table` |
  | `YExcel_DateTime` | `DateTime` |
  | `YExcelVar` | `Var` |
  | `YExcelfunction` | `Formula` |

- **Tipagem TLPP completa**:
  - `Public Data` com `As <Tipo>` em todos os membros de dados
  - Parâmetros tipados em todas as assinaturas de métodos e funções estáticas
  - Declarações `Local`/`Private`/`Static` tipadas com sintaxe correta (`Local var := valor As Tipo`)
  - Visibilidade `Public Method` nas declarações dentro da classe; implementações sem modificador
- **Sobrecarga de operadores** na classe `DateTime`:
  - `Operator Add`, `Operator Sub`, `Operator Mult`, `Operator Div` — operando direito aceita Numeric (dias), Date (serial Excel) ou Character (HH:mm[:ss])
  - `Operator Compare` — retorna -1, 0 ou 1
  - `Operator ToString` — retorna `DD/MM/YYYY HH:MM:SS`
- **`Static Function CopyFile(cOrigem, cDestino)`** — substitui `__COPYFILE`; usa `FOpen`, `FREAD`, `FClose` e `FWFileWriter`; retorna Logical
- Encoding **CP-1252**, fim de linha **CRLF**

#### Arquivo modificado: `src/YEXCEL.prw`

- Reduzido a *thin wrappers*: cada classe legada mantém apenas `New()` delegando para `custom.tools.excel.*`
- Compatibilidade retroativa preservada — código existente que usa `YExcel():New()` continua funcionando sem alteração
