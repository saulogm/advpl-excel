# Migração YExcel → custom.excel.Xlsx

`src/YEXCEL.prw` só contém subclasses de compatibilidade. Por exemplo, `Class YExcel From custom.excel.Xlsx` tem um `New` que apenas chama `_Super:New(...)`. **Todo método tem o mesmo nome e os mesmos parâmetros**, então a migração consiste principalmente em trocar o nome das classes e, opcionalmente, modernizar as chamadas da API v1.

## 1. Mapa de classes

| Legado | Novo |
|---|---|
| `YExcel():New(cNomeFile, cFileOpen, cTipo)` | `custom.excel.Xlsx():New(cNomeFile, cFileOpen, cTipo)` |
| `YExcel_Style():New(oPai, oExcel)` | `oExcel:NewStyle(oPai)`, que é a forma preferida (registra o estilo para limpeza), ou `custom.excel.Style():New(oPai, oExcel)` |
| `YExcel_StyleRules():New(oExcel)` | `oExcel:NewStyRules()` |
| `YExcel_RegraLinha():New(bBloco, aRegra, oExcel)` | `oExcel:NewRuleLine(bBloco, aRegra)` |
| `YExcel_Table():New(oExcel, nLin, nCol, cNome)` | `oExcel:AddTabela(cNome, nLin, nCol)`. **Atenção à ordem dos parâmetros** |
| `YExcel_DateTime():New(dData, cTime, ...)` | `oExcel:GetDateTime(dData, cTime)` ou `custom.excel.DateTime():New(dData, cTime)` |
| `YExcelTag` | `custom.excel.Tag` (uso interno) |
| `YExcelfunction` | `custom.excel.Formula` |
| `ValType(o) == "O" .And. o:ClassName() == "YEXCEL"` | Comparar com `"XLSX"`, ou melhor, não depender do ClassName |

## 2. API v1 → API atual (recomendado ao migrar)

| v1 | Atual |
|---|---|
| `oExcel:Cell(nLin, nCol, xValor, cFormula, nStyle)` | `oExcel:Pos(nLin, nCol):SetValue(xValor, cFormula):SetStyle(oStyle)` |
| `nSty := oExcel:AddStyles(numFmtId, fontId, fillId, nborderId, , aValores)` | `oSty := oExcel:NewStyle():SetnumFmt(numFmtId):Setfont(fontId):Setfill(fillId):Setborder(nborderId):SetaValores(aValores)` |
| `oExcel:CreateStyle(nStylePai, ...)` | `oExcel:NewStyle(oStylePai):Set...()` |
| `oExcel:Gravar(cLocal, lAbrir, lDelSrv)` | `cArq := oExcel:Save(cLocal)` → `If lAbrir; oExcel:OpenApp(); EndIf` → `oExcel:Destroy()` → `FreeObj(oExcel)` |
| `oExcel:Close()` | `oExcel:Destroy()` (equivalente) + `FreeObj(oExcel)` |
| `oExcel:OpenRead(cFile, nPlan)` + `oExcel:CellRead(nLin, nCol)` + `oExcel:adimension` + `oExcel:CloseRead()` | `oExcel:Open(cFile)` + `SetPlanAt(nPlan)` + `GetValue(nLin, nCol)` + `LinTam()`/`ColTam()` + `Destroy()` |
| `oExcel:SetDefRow(...)` | Obsoleto: remover |
| `cTipo` documentado como "BD/TMPDB/MEM" | Os valores reais são `"A"` (arquivo), `"M"` (memória) e `"B"` (banco temporário) |

`AddStyles` e `Cell` ainda funcionam. Se o usuário pedir uma migração mínima, basta trocar as classes. Se pedir modernização, ou se você estiver reescrevendo o trecho de qualquer forma, converta para a API atual.

## 3. `.prw` x `.tlpp`

- Um `.prw` pode instanciar `custom.excel.Xlsx():New(...)` pelo nome completo, mas **não pode usar argumento nomeado**. Mantenha as chamadas posicionais no `.prw`.
- Converter o fonte para `.tlpp` só quando o usuário pedir. Nesse caso, veja a skill `advpl-to-tlpp-migration` e adicione `#include "tlpp-core.th"`, `Namespace`, tipagem `As` e argumentos nomeados.
- Em `.tlpp`, `using namespace custom.excel` permite escrever `Xlsx():New()`. Mesmo assim, o nome completo deixa mais claro de onde a classe vem, e é o padrão dos exemplos.

## 4. Checklist de migração

1. Veja o `git status`/`git diff` do fonte. Se houver alterações do usuário ainda não commitadas, preserve-as e pergunte em caso de dúvida.
2. Confirme o encoding (`file <arq>` deve mostrar ISO-8859/não UTF-8) e edite **via PowerShell com CP-1252** (ver a seção 6 do SKILL.md), nunca com ferramentas que gravem UTF-8 em arquivo com acentos.
3. Troque os construtores e as classes seguindo a tabela 1. Procure por `YExcel` (case-insensitive) no fonte inteiro, inclusive em comentários e ProtheusDOC (`@see`).
4. Substitua `Gravar` e `Close` pelo ciclo `Save`/`OpenApp`/`Destroy`/`FreeObj`, garantindo o `Destroy` também nos caminhos de erro.
5. Se modernizar, converta `Cell`/`AddStyles` para `Pos`/`SetValue`/`NewStyle`, conferindo que a ordem de escrita continua sequencial no tipo "A".
6. Troque `YExcel_Table():New(oExcel, l, c, nome)` por `oExcel:AddTabela(nome, l, c)`, porque a ordem dos parâmetros muda.
7. Remova variáveis e liberações que não fazem mais sentido. Não dê `FreeObj` em estilos, tabelas ou gráficos criados pela instância, que já são liberados no `Destroy`.
8. Valide o arquivo: sem BOM, CRLF, `EF BF BD` = 0. Se a skill `advpl-tlpp-compile` estiver disponível, ofereça compilar.

## 5. Exemplo

Antes (`.prw`):
```advpl
oExcel := YExcel():New("Relatorio")
nFonte := oExcel:AddFont(11, "FFFFFFFF", "Calibri", "2", , .T.)
nCor   := oExcel:CorPreenc("FF4F81BD")
nSty   := oExcel:AddStyles(, nFonte, nCor, oExcel:Borda("ALL"))
oExcel:ADDPlan("Dados")
oExcel:Cell(1, 1, "Código", , nSty)
oExcel:Cell(2, 1, "000001")
oExcel:Gravar(GetTempPath(), .T., .T.)
```

Depois (`.prw`, posicional):
```advpl
oExcel := custom.excel.Xlsx():New("Relatorio")
nFonte := oExcel:AddFont(11, "FFFFFFFF", "Calibri", "2", , .T.)
nCor   := oExcel:CorPreenc("FF4F81BD")
oSty   := oExcel:NewStyle():Setfont(nFonte):Setfill(nCor):Setborder(oExcel:Borda("ALL"))
oExcel:ADDPlan("Dados")
oExcel:Pos(1, 1):SetValue("Código"):SetStyle(oSty)
oExcel:Pos(2, 1):SetValue("000001")
oExcel:Save(GetTempPath())
oExcel:OpenApp()
oExcel:Destroy()
FreeObj(oExcel)
```

Em `.tlpp`, a mesma lógica fica com argumentos nomeados: `oExcel:AddFont(nTamanho = 11, xCorRGB = "FFFFFFFF", cNome = "Calibri", cfamily = "2", lNegrito = .T.)`.
