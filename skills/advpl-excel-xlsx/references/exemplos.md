# Receitas completas

Os exemplos de referência completos estão no repositório advpl-excel:
- `exemplo/custom.excel.teste.tlpp`: `U_tstXlsx` demonstra quase todos os recursos (estilos, fórmulas, nomes, hyperlink, comentário, formatação condicional, imagens, tabela, Alias2Tab SB1/SA1, agrupamento, impressão, proteção). Também traz leitura (`YTstRW`), edição (`TstEdit`), memória (`tstMemoria`) e bulk (`tstBulk`).
- `exemplo/custom.excel.grafico.teste.tlpp`: todos os tipos de gráfico.
- `exemplo/custom.excel.crypto.teste.tlpp`: senha de abertura (`tstXlsxSenha`).

Se esses arquivos estiverem acessíveis no projeto, leia o trecho relevante antes de escrever algo pouco comum. Eles estão em CP-1252 e os acentos podem aparecer corrompidos na leitura, sem que isso afete o código.

---

## Receita 1 — Relatório de query com Alias2Tab, SX3, subtotal e zebra

```tlpp
#include "tlpp-core.th"
#include "protheus.ch"
Namespace custom.fat

/*/{Protheus.doc} RelProdExcel
Exporta produtos agrupados por grupo, com subtotal de preço, para Excel
@type function
@author ...
@since dd/mm/aaaa
@version 1.0
@return character, caminho do arquivo gerado
/*/
User Function RelProdExcel() As Character
	Local oExcel As Object
	Local oStyCab As Object
	Local oStyLinha As Object
	Local oStySub As Object
	Local oStyMoeda As Object
	Local oRegra As Object
	Local jCab As Json
	Local aCampos As Array
	Local nBorda As Numeric
	Local nFonteBr As Numeric
	Local cAlias As Character
	Local cArquivo := "" As Character

	cAlias	:= MpSysOpenQuery("SELECT B1_COD, B1_DESC, B1_GRUPO, B1_UM, B1_PRV1 FROM " + RetSqlName("SB1") + " WHERE B1_FILIAL = '" + xFilial("SB1") + "' AND D_E_L_E_T_ = ' ' ORDER BY B1_GRUPO, B1_COD")

	oExcel		:= custom.excel.Xlsx():New(cNomeFile = "Produtos", cTipo = "A")
	nBorda		:= oExcel:Borda("ALL")
	nFonteBr	:= oExcel:AddFont(nTamanho = 11, xCorRGB = "FFFFFFFF", cNome = "Calibri", cfamily = "2", lNegrito = .T.)
	oStyLinha	:= oExcel:NewStyle():Setborder(nBorda)
	oStyCab		:= oExcel:NewStyle(oStyLinha):Setfont(nFonteBr):Setfill(oExcel:CorPreenc("FF1F497D"))
	oStySub		:= oExcel:NewStyle(oStyLinha):Setfont(nFonteBr):Setfill(oExcel:CorPreenc("FF4F81BD"))
	oStyMoeda	:= oExcel:NewStyle(oStyLinha):SetnumFmt(44)

	oExcel:ADDPlan(cNome = "Produtos")
	oExcel:AddPane(nySplit = 1, nxSplit = 0)
	oExcel:SetPrintTitles(nLinha = 1, nLinha2 = 1)

	jCab	:= JsonObject():New()
	oExcel:NewFldTab(jCab = jCab, cCampo = "B1_PRV1", xStyle = oStyMoeda)
	aCampos	:= {"B1_GRUPO","B1_COD","B1_DESC","B1_UM","B1_PRV1"}

	oExcel:DefSubTotal(cChave = "B1_GRUPO", lSubTotal = .T., lTotalGeral = .T., lAgrupar = .T., nNivel = 2)
	oExcel:AddSubTotal(cCampo = "B1_PRV1", cFuncao = "9")
	oExcel:AddSubTotal(cCampo = "B1_GRUPO", cFuncao = "formula", bAdvpl = {|aValAtu,xValor,cChave| aValAtu[1] := cChave }, ;
		bAdvplExib = {|aValAtu| If(oExcel:cSubTotalTp == "T", "Total Geral", "Total " + AllTrim(aValAtu[1])) }, xValorIni = "")

	oRegra	:= oExcel:NewRuleLine({|| oExcel:cSubTotalTp })
	oRegra:AddRegra(cRegra = "S", oStyle = oStySub, lPrincipal = .F.)
	oRegra:AddRegra(cRegra = "T", oStyle = oStySub, lPrincipal = .F.)

	oExcel:Alias2Tab(cAlias = cAlias, oStyle = oStyCab, lSx3 = .T., jCab = jCab, lExibirCab = .T., lCombo = .T., ;
		aOnlyFieds = aCampos, xRegraStyle = oRegra, oStyleLinha = oStyLinha, lFiltro = .T., lRTrim = .T.)
	(cAlias)->(DbCloseArea())

	cArquivo	:= oExcel:Save(GetTempPath())
	If !IsBlind()
		oExcel:OpenApp()
	EndIf
	oExcel:Destroy()	//libera estilos, regra e jCab (registrados na instância)
	FreeObj(oExcel)
	FwFreeArray(aCampos)
Return cArquivo
```

## Receita 2 — Tabela formatada com totais e formatação condicional

```tlpp
oExcel	:= custom.excel.Xlsx():New(cNomeFile = "Vendas", cTipo = "M")	//"M": permite aplicar estilo em intervalo depois de escrever
oStyMoeda	:= oExcel:NewStyle():SetnumFmt(44)
nDxfNeg	:= oExcel:ADDdxf(aFont = oExcel:Font(cCorRGB = "9C0006"), aCorPreenc = oExcel:Preenc("FFC7CE"))
oExcel:ADDPlan(cNome = "Vendas", cCor = "1F497D")
oExcel:AddTamCol(nMin = 1, nMax = 3, nWidth = 18)
oTab	:= oExcel:AddTabela(cNome = "TabVendas", nLinha = 1, nColuna = 1)
oTab:AddStyle(cNome = "TableStyleMedium9", lLinhaTiras = .T.)
oTab:AddFilter()
oTab:AddColumn("Filial")
oTab:AddColumn("Valor")
oTab:AddColumn("Emissão")
nLin	:= 1
While (cAlias)->(!Eof())
	nLin++
	oTab:AddLine()
	oTab:Cell(xColuna = "Filial", xValor = (cAlias)->FILIAL)
	oTab:Cell(xColuna = "Valor", xValor = (cAlias)->VALOR)
	oTab:Cell(xColuna = "Emissão", xValor = SToD((cAlias)->EMISSAO))
	nTotal	+= (cAlias)->VALOR
	(cAlias)->(DbSkip())
EndDo
oTab:AddTotal(xColuna = "Filial", xValor = "TOTAL", cFunction = "")
oTab:AddTotal(xColuna = "Valor", xValor = nTotal, cFunction = "SUM")
oTab:AddTotais()
oTab:Finish()
oExcel:SetStyle(oStyMoeda, 2, 2, nLin + 1, 2)	//formato moeda na coluna Valor (inclui total)
oExcel:FormatCond(cRefDe = "B2", cRefAte = oExcel:Ref(nLin,2), nEstilo = nDxfNeg, cOperator = "<", xFormula = 0)
```

Em `cTipo = "A"` evite aplicar `SetStyle` em intervalo já escrito, porque a escrita é sequencial. Nesse caso aplique o estilo em cada célula ao escrever.

## Receita 3 — Dados + gráfico na mesma planilha

```tlpp
aMeses	:= {"Jan","Fev","Mar","Abr","Mai","Jun"}
aValor	:= {100,120,90,140,160,150}
oExcel	:= custom.excel.Xlsx():New(cNomeFile = "Grafico", cTipo = "A")
oExcel:ADDPlan(cNome = "Dados")
oExcel:Pos(1,1):SetValue("Mês")
oExcel:Pos(1,2):SetValue("Vendas")
For nCont := 1 To Len(aMeses)
	oExcel:Pos(nCont+1,1):SetValue(aMeses[nCont])
	oExcel:Pos(nCont+1,2):SetValue(aValor[nCont])
Next
cFim	:= cValToChar(Len(aMeses)+1)
oGraf	:= oExcel:NewChart("col")
oGraf:SetTitulo("Vendas por mês")
oGraf:SetLegenda(lExibir = .F.)
oGraf:AddSerie(cNome = "Vendas", cRefVal = "Dados!$B$2:$B$" + cFim, cRefCat = "Dados!$A$2:$A$" + cFim, aVal = aValor, aCat = aMeses)
oGraf:Insert(nLinha = 1, nColuna = 4, nLinha2 = 16, nColuna2 = 12)
cArq	:= oExcel:Save(GetTempPath())
oExcel:OpenApp()
oExcel:Destroy()	//libera o gráfico e as séries
FreeObj(oExcel)
FwFreeArray(aMeses)
FwFreeArray(aValor)
```

## Receita 4 — Ler um xlsx enviado pelo usuário e importar

```tlpp
cArq	:= tFileDialog("Arquivos Excel (*.xlsx)|*.xlsx", "Selecione o arquivo", , , .F., )
oExcel	:= custom.excel.Xlsx():New()
If !oExcel:Open(cArq)
	FWAlertError(oExcel:GetError(), "Importação")
Else
	oExcel:SetPlanAt(1)
	aLin	:= oExcel:LinTam()
	For nLin := 2 To aLin[2]	//pula o cabeçalho
		cCod	:= oExcel:Pos(nLin,1):GetString()
		nQtd	:= oExcel:Pos(nLin,2):GetNumber()
		dData	:= oExcel:Pos(nLin,3):GetDate()
		//... gravar
	Next
	FwFreeArray(aLin)
EndIf
oExcel:Destroy()
FreeObj(oExcel)
```

## Receita 5 — Job (sem interface) salvando no servidor

```tlpp
oExcel	:= custom.excel.Xlsx():New(cNomeFile = "Fechamento_" + DToS(Date()), cTipo = "B")
//... preencher
cArq	:= oExcel:Save("\relatorios")	//pasta do servidor, sem OpenApp
oExcel:Destroy()
FreeObj(oExcel)
//cArq pode seguir como anexo de e-mail (TMailManager) etc.
```
