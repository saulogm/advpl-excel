# Tabelas formatadas, Alias2Tab e Bulk

## Sumário
1. Tabela formatada do Excel (AddTabela)
2. Preencher por alias/query (Alias2Tab)
3. Definição de campos (NewFldTab)
4. Subtotais (DefSubTotal/AddSubTotal)
5. Estilo por tipo de linha (NewRuleLine)
6. Bulk — grande volume

---

## 1. Tabela formatada: `AddTabela(cNome, nLinha, nColuna)` → `custom.excel.Table`

| Método | Descrição |
|---|---|
| `AddStyle(cNome, lLinhaTiras, lColTiras, lFormPrimCol, lFormUltCol)` | Estilos nativos `TableStyleLight1..21`, `TableStyleMedium1..28`, `TableStyleDark1..11` |
| `AddFilter(lFiltro)` | Filtro da tabela. Não use o `AutoFilter` da planilha junto |
| `AddColumn(cNome, nStyle, cNameField)` | Cria uma coluna. O nome precisa ser único na tabela. `nStyle` (id numérico ou objeto `NewStyle`) é o estilo padrão da coluna, aplicado pelo `Cell` quando ele não recebe estilo. `cNameField` é um apelido opcional (ex.: o nome do campo), que pode ser usado no lugar do título em `Cell` e `AddTotal` |
| `AddLine(nQtd)` | Nova linha de dados |
| `Cell(xColuna, xValor, cFormula, nStyle)` | Preenche uma célula da linha atual. `xColuna` é o título, o `cNameField` ou o número da coluna. Sem `nStyle`, usa o estilo da coluna |
| `AddTotal(xColuna, xValor, cFunction, nStyle)` | Total da coluna (`xColuna` aceita título, `cNameField` ou número). `cFunction` pode ser `"SUM"`, `"COUNTA"`, uma fórmula completa (`"SUBTOTAL(103,Tabela1[Filial])"`) ou `""` para texto |
| `AddTotais()` | Cria a linha de totais (depois dos `AddTotal`) |
| `Alias2Tab(cAlias, oStyle, lSx3, jCab, lExibirCab, lCombo, aOnlyFieds, xRegraStyle, oStyleLinha, lRTrim, lFiltro)` | Preenche a tabela a partir de um alias. **Atenção:** os dois últimos parâmetros estão invertidos em relação ao `Xlsx:Alias2Tab` |
| `Finish()` | **Obrigatório**: fecha a edição da tabela |

```tlpp
oTab := oExcel:AddTabela(cNome = "Vendas", nLinha = 1, nColuna = 1)
oTab:AddStyle(cNome = "TableStyleMedium15", lLinhaTiras = .T.)
oTab:AddFilter()
oTab:AddColumn("Filial")
oTab:AddColumn(cNome = "Valor da Venda", nStyle = oStyMoeda, cNameField = "VALOR")	//estilo padrão + apelido
oTab:AddColumn("Data")
While (cAlias)->(!Eof())
	oTab:AddLine()
	oTab:Cell(xColuna = "Filial", xValor = (cAlias)->FILIAL)
	oTab:Cell(xColuna = "VALOR", xValor = (cAlias)->VALOR)	//pelo apelido; aplica oStyMoeda
	oTab:Cell(xColuna = 3, xValor = SToD((cAlias)->EMISSAO))
	nTotal += (cAlias)->VALOR
	(cAlias)->(DbSkip())
EndDo
oTab:AddTotal(xColuna = "Filial", xValor = "TOTAL", cFunction = "")
oTab:AddTotal(xColuna = "VALOR", xValor = nTotal, cFunction = "SUM")
oTab:AddTotais()
oTab:Finish()
```

O nome da tabela (`"Vendas"`) serve para referência estruturada em fórmulas. Nelas vale o **título** da coluna (`Vendas[Valor da Venda]`), não o `cNameField`.

## 2. `Alias2Tab`: preencher a planilha a partir de um alias ou query

`Alias2Tab(cAlias, oStyle, lSx3, jCab, lExibirCab, lCombo, aOnlyFieds, xRegraStyle, oStyleLinha, lFiltro, lRTrim, oTabela)`

| Parâmetro | Significado |
|---|---|
| `cAlias` | Alias da tabela ou da query (é percorrido do registro atual até o EOF) |
| `oStyle` | Estilo do cabeçalho |
| `lSx3` | Busca título, tamanho, picture e combo na SX3 (precisa de ambiente aberto) |
| `jCab` | Definições manuais dos campos (NewFldTab) |
| `lExibirCab` | Mostra a linha de cabeçalho |
| `lCombo` | Troca o código do combobox pela descrição |
| `aOnlyFieds` | Array com os campos a exibir, na ordem |
| `xRegraStyle` | Regra de estilo por linha (NewRuleLine) |
| `oStyleLinha` | Estilo padrão das linhas de dados |
| `lFiltro` | Aplica o autofiltro |
| `lRTrim` | Remove espaços à direita |

Com uma query, `MpSysOpenQuery` já devolve os tipos. Campos de data vindos de SQL chegam como caractere, a menos que você use `TCSetField`, e com `lSx3` a SX3 resolve o tipo. Feche o alias depois (`(cAlias)->(DbCloseArea())`).

## 3. `NewFldTab(jCab, cCampo, cDescricao, nTamanho, cPicture, cCombo, xStyle, nOrdem, cTipo, cDados, lNewCampo, lHidden, nTamDados, xStyleCab)` → JSON

Sobrescreve ou complementa a definição de um campo dentro de `jCab`. O `jCab` é um `JsonObject():New()` e é registrado na lib, então o `Destroy()` o libera.

```tlpp
jCab := JsonObject():New()
oExcel:NewFldTab(jCab = jCab, cCampo = "B1_PRV1", xStyle = oStyMoeda, xStyleCab = oStyCabMoeda)	//xStyle: células da coluna; xStyleCab: só o cabeçalho deste campo (sobrepõe o oStyle do Alias2Tab)
oExcel:NewFldTab(jCab = jCab, cCampo = "B1_GRUPO")
jCab["B1_GRUPO"]["ordem"] := 1			//muda a ordem
oExcel:NewFldTab(jCab = jCab, cCampo = "A1_COD", cDescricao = "Código do Cliente", nTamanho = 10)
oExcel:NewFldTab(jCab = jCab, cCampo = "A1_PESSOA", cCombo = "F=Física;J=Jurídica")
oExcel:NewFldTab(jCab = jCab, cCampo = "A1_DESC", cPicture = "@E 999,999,999.99")
//Ajuste direto no json (equivalente aos parâmetros): "style", "stylecab", "ordem", "descricao", "tamanho", "picture", "combo", "hidden"...
jCab["A1_COD"]["stylecab"] := oStyCabDestaque
//Campo calculado (não existe no alias)
oExcel:NewFldTab(jCab = jCab, cCampo = "REGISTRO", cDescricao = "Registro", nTamanho = 10, cPicture = "9", cTipo = "N", cDados = "Recno()", lNewCampo = .T.)
```

## 4. Subtotais: chame **antes** do `Alias2Tab`

- `DefSubTotal(cChave, lSubTotal, lTotalGeral, lAgrupar, nNivel)`: `cChave` é o campo de quebra, `lAgrupar` cria o agrupamento de linhas (outline) e `nNivel` define o nível (1 total geral, 2 subtotal, 3 linhas).
- `AddSubTotal(cCampo, cFuncao, bFormula, bAdvpl, bAdvplExib, xValorIni)`:
  - `cFuncao` pode ser `"1"` média, `"2"` contar números, `"3"` contar valores, `"4"` máximo, `"5"` mínimo, `"9"` soma (padrão), `"last"` (último conteúdo) ou `"formula"`.
  - `bFormula` recebe `{|cLinIni,cLinFim,cColuna,cChave,lTotalGeral| "FORMULA" }`.
  - `bAdvpl` recebe `{|aValAtu,xValor,cChave| ... }` e acumula o valor em `aValAtu[1]`.
  - `bAdvplExib` recebe `{|aValAtu| texto }` e define o que aparece na célula.

```tlpp
oExcel:DefSubTotal(cChave = "B1_GRUPO", lSubTotal = .T., lTotalGeral = .T., lAgrupar = .T., nNivel = 2)
oExcel:AddSubTotal(cCampo = "B1_GRUPO", cFuncao = "last")
oExcel:AddSubTotal(cCampo = "B1_PRV1", cFuncao = "9")
oExcel:AddSubTotal(cCampo = "B1_PICM", cFuncao = "1")
oExcel:AddSubTotal(cCampo = "B1_LOCPAD", cFuncao = "formula", ;
	bFormula = {|cLinIni,cLinFim,cColuna,cChave,lTotalGeral| "SUBTOTAL(3," + cColuna + cLinIni + ":" + cColuna + cLinFim + ")" }, ;
	bAdvpl = {|aValAtu,xValor| aValAtu[1] += 1 }, bAdvplExib = {|aValAtu,xValor| cValToChar(aValAtu[1]) }, xValorIni = 0)
oExcel:AddSubTotal(cCampo = "B1_COD", cFuncao = "formula", bAdvpl = {|aValAtu,xValor,cChave| aValAtu[1] := cChave }, ;
	bAdvplExib = {|aValAtu| If(oExcel:cSubTotalTp == "T", "Total Geral", "Total do Grupo: " + aValAtu[1]) }, xValorIni = "")
```

Durante a geração, `oExcel:cSubTotalTp` vale `"L"` (linha), `"S"` (subtotal) ou `"T"` (total geral).

## 5. `NewRuleLine(bBloco, aRegra)` → `custom.excel.RegraLinha`

O bloco retorna uma chave, e `AddRegra(cRegra, oStyle, lPrincipal)` associa um estilo a cada chave.

```tlpp
//Estilo nas linhas de subtotal e total
oRegra := oExcel:NewRuleLine({|| oExcel:cSubTotalTp })
oRegra:AddRegra(cRegra = "S", oStyle = oStySubTot, lPrincipal = .F.)
oRegra:AddRegra(cRegra = "T", oStyle = oStySubTot, lPrincipal = .F.)

//Estilo em uma única coluna, conforme o conteúdo: chave "CAMPO|valor"
oRegra2 := oExcel:NewRuleLine({|| oExcel:cCampo + "|" + (cAlias)->A1_RECCOFI })
oRegra2:AddRegra(cRegra = "A1_RECCOFI|S", oStyle = oStyDestaque, lPrincipal = .F.)
```

No bulk, a regra é passada como array: `oExcel:NewRuleLine(bBloco, {"chave", oStyle, .F.}):GetArray()`.

### Exemplo completo com Alias2Tab e subtotal

```tlpp
oExcel:ADDPlan(cNome = "Produtos", cCor = "E26B0A")
oExcel:Pos(1,1):SetValue("Cadastro de produtos")	//Título: o Alias2Tab continua na linha seguinte
SB1->(DbSetOrder(4))	//B1_FILIAL+B1_GRUPO+B1_COD: precisa estar ordenado pela chave do subtotal
SB1->(DbGoTop())
jCab := JsonObject():New()
oExcel:NewFldTab(jCab = jCab, cCampo = "B1_PRV1")
jCab["B1_PRV1"]["style"] := oStyMoeda
aCampos := {"B1_COD","B1_DESC","B1_GRUPO","B1_UM","B1_PRV1"}
oExcel:DefSubTotal(cChave = "B1_GRUPO", lSubTotal = .T., lTotalGeral = .T., lAgrupar = .T., nNivel = 2)
oExcel:AddSubTotal(cCampo = "B1_PRV1", cFuncao = "9")
oRegra := oExcel:NewRuleLine({|| oExcel:cSubTotalTp })
oRegra:AddRegra(cRegra = "S", oStyle = oStySubTot, lPrincipal = .F.)
oRegra:AddRegra(cRegra = "T", oStyle = oStySubTot, lPrincipal = .F.)
oExcel:Alias2Tab(cAlias = "SB1", oStyle = oStyCab, lSx3 = .T., jCab = jCab, lExibirCab = .T., lCombo = .T., aOnlyFieds = aCampos, xRegraStyle = oRegra, oStyleLinha = oStyBorda)
FwFreeArray(aCampos)
```

Numa tabela formatada use `oTab:Alias2Tab(cAlias = cAlias, lSx3 = .T., jCab = jCab, xRegraStyle = oRegra)`, seguido de `AddTotal`, `AddTotais()` e `Finish()`. O `Alias2Tab` cria cada coluna com o **nome do campo** como `cNameField` e com o estilo do campo (`jCab[campo]["style"]`/`oStyleLinha`). Por isso o `AddTotal` pode referenciar a coluna pelo nome do campo, sem depender do título vindo da SX3:

```tlpp
oTab:Alias2Tab(cAlias = cAlias, lSx3 = .T., jCab = jCab)
oTab:AddTotal(xColuna = "A1_COD", xValor = "TOTAL", cFunction = "")
oTab:AddTotal(xColuna = "A1_MSALDO", xValor = nSaldo, cFunction = "SUM")
oTab:AddTotais()
oTab:Finish()
```

## 6. Bulk: grande volume (use `cTipo = "B"`)

O bulk é a forma mais rápida de preencher dezenas ou centenas de milhares de linhas. Os valores são informados **na ordem das colunas definidas**.

1. `BulkNewField(nColuna, cTipo, nTamanho, cCombo, oStyle, lFormula, lDatetime, cCampo, cDados)` → JSON de definição de campo. `cTipo` é `"C"`, `"N"`, `"D"` ou `"L"`. Se a coluna for receber fórmula, use `lFormula = .T.`.
2. `DefBulkLine(aCampos, xRegraStyle, lMontarLin)`
3. Para cada linha: um `SetValueBulk(xValor, cFormula)` por coluna e depois `SetBulkLine()`.
4. `FlushBulk()` **antes** do `Save`.

```tlpp
oExcel := custom.excel.Xlsx():New(cNomeFile = "Grande", cTipo = "B")
oStyBorda := oExcel:NewStyle():Setborder(oExcel:Borda("ALL"))
oExcel:ADDPlan()
oExcel:Pos(1,1):SetValue("Linha"):SetStyle(oStyBorda)
oExcel:Pos(1,2):SetValue("Filial"):SetStyle(oStyBorda)
oExcel:Pos(1,3):SetValue("Venda"):SetStyle(oStyBorda)
oExcel:Pos(1,4):SetValue("Data"):SetStyle(oStyBorda)

aCampos := {}
aAdd(aCampos, oExcel:BulkNewField(nColuna = 1, cTipo = "N"))
aAdd(aCampos, oExcel:BulkNewField(nColuna = 2, cTipo = "C", nTamanho = 12))
aAdd(aCampos, oExcel:BulkNewField(nColuna = 3, cTipo = "N", oStyle = oStyMoeda))
aAdd(aCampos, oExcel:BulkNewField(nColuna = 4, cTipo = "D"))
oExcel:DefBulkLine(aCampos = aCampos, lMontarLin = .F.)
oExcel:nLinha := 2	//começa na linha 2
For nCont := 1 To 100000
	oExcel:SetValueBulk(nCont)
	oExcel:SetValueBulk("Filial" + cValToChar(nCont))
	oExcel:SetValueBulk(Randomize(1,100))
	oExcel:SetValueBulk(Date() + nCont)
	oExcel:SetBulkLine()
Next
oExcel:FlushBulk()
cArq := oExcel:Save(GetTempPath())
oExcel:Destroy()
FreeObj(oExcel)
FwFreeArray(aCampos)	//os JSON de campo já foram liberados pelo Destroy; aqui só o array
```
