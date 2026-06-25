# Release Notes — advpl-excel

## [Unreleased]

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
