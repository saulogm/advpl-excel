# Release Notes — advpl-excel

## [Unreleased]

### Fórmula em português e L1C1

- Novo método `oExcel:Formula(cFormula, lPtBr = .T., lL1C1 = .T.)` (classe `custom.excel.Formula`, `src/custom.excel.formula.tlpp`): fórmula escrita como no Excel em português (nomes das funções, `VERDADEIRO`/`FALSO`, erros, `;` e `,` decimal) e com referências L1C1 (`L2C3`, `L[-1]C`, linhas/colunas inteiras). A conversão para inglês/A1 é feita no `SetValue`/`Cell`, relativa à célula posicionada, permitindo reutilizar o mesmo objeto em um laço.
- `SetValue` e `Cell` aceitam o objeto no parâmetro `cFormula`.
- Correção: no `Open` (Windows com 7-Zip) a pasta de extração que sobrou de uma execução anterior com o mesmo nome temporário é apagada e o 7-Zip recebe `-y`; antes a leitura travava esperando a confirmação de substituição.
- Testes `U_XFMLPTTST` e desempenho `U_XFMLBENCH` (`exemplo/custom.excel.calc.teste.tlpp`) e exemplo na segunda planilha de `u_tstXlsx` (`exemplo/custom.excel.teste.tlpp`).

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
