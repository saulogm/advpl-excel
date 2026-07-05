#include "totvs.ch"

User Function ytst6()
	Local oExcel
	Local cAlias
	//Local oPosFill
	RpcSetEnv("99","01")
	oExcel	:= YExcel():new()
	oExcel:ADDPlan("SB1","1F497D")		//Adiciona nova planilha

	cAlias := MpSysOpenQuery("SELECT * FROM SB1990")

	//oExcel:Alias2Tab(cAlias,,.T.,,,,,)
	oTabela	:= oExcel:AddTabela("Tabela2")	//Cria uma tabela de estilos
	oTabela:Alias2Tab(cAlias,,.T.)
	(cAlias)->(DbCloseArea())

	oTabela:AddTotal("Codigo",0,"COUNTA")	//Usa função COUNTA(Contar Valores)
	oTabela:AddTotais()	//Adiciona linha de totais
	oTabela:Finish()	//Fecha a edição da tabela

	oExcel:Save(GetTempPath())
	oExcel:OpenApp()
	oExcel:Close()
Return

User Function ytst7()
	Local nLinha
	Local oExcel		:= YExcel():new("Pasta1")
	Local aCampos		:= {}
	Local nPosBorda2	:= oExcel:Borda("ALL")
	Local oBordas		:= oExcel:NewStyle():Setborder(nPosBorda2)		//Cria um estilo com bordas
	Local nPosCor		:= oExcel:CorPreenc("FF0000FF")					//Cor de Fundo Azul
	Local oStAzul		:= oExcel:NewStyle(oBordas):Setfill(nPosCor)	//Cria um estilo com herança de bordas e seta cor de fundo azul


	oExcel:ADDPlan(/*cNome*/)		//Adiciona uma planilha em branco
	//nColuna,cTipo,nTamanho,cCombo,oStyle,lFormula,lDatetime,cCampo,cDados
	AADD(aCampos,oExcel:BulkNewField(1,"C",,,oBordas))		//Tipo de conteudo C=Caracter
	AADD(aCampos,oExcel:BulkNewField(2,"N",,,oBordas))		//Tipo de conteudo N=Numero
	AADD(aCampos,oExcel:BulkNewField(3,"L",,,oBordas))		//Tipo de conteudo L=Logico
	AADD(aCampos,oExcel:BulkNewField(4,"D",,,oStAzul))		//Tipo de conteudo D=Data

	oExcel:DefBulkLine(aCampos)
	For nLinha:=1 to 1000
		oExcel:SetValueBulk("Linha "+cValtoChar(nLinha))
		oExcel:SetValueBulk(nLinha)
		oExcel:SetValueBulk(nLinha%2==0)
		oExcel:SetValueBulk(Date()+nLinha)
		oExcel:SetBulkLine()
	Next
	oExcel:SetBulkLine()
	oExcel:FlushBulk()
	oExcel:Save(GetTempPath())
	oExcel:OpenApp()
	oExcel:Close()
Return
