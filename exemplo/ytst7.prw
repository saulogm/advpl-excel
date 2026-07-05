#include "totvs.ch"
User Function ytst8()
	Local oExcel := YExcel():new()
	Local oSemBloq		:= oExcel:Cellprotection(.F.)		//Célula não Bloqueada
	Local oOculForm		:= oExcel:Cellprotection(,.T.)		//Célula Oculta formulas
	Local oProtecao1	:= oExcel:NewStyle():SetaValores({oSemBloq})
	Local oProtecao2	:= oExcel:NewStyle():SetaValores({oOculForm})

	oExcel:ADDPlan() //Cria Planinha
	oExcel:Pos(1,1):SetValue("A senha é 123") //Escreve
	oExcel:Pos(2,1):SetValue("Aqui pode editar"):SetStyle(oProtecao1)	//Permitir editar essa célula
	oExcel:Pos(3,1):SetValue(30,"15+15"):SetStyle(oProtecao2)	//Ocultar formula
	oExcel:SetsheetProtection("123") //Proteger a Planilha

	nFmtNum1 := oExcel:AddFmtNum(0,.T.,/*cPrefixo*/,/*cSufixo*/,"(",")",/*cValorZero*/,/*cCorPos*/,"Red"/*cCorNeg*/,/*nNumFmtId*/)
	oPosNum1 := oExcel:NewStyle():SetNumFmt(nFmtNum1)

	nFmtNum2 := oExcel:AddFmtNum(18,.T.,/*cPrefixo*/,/*cSufixo*/,"(",")",/*cValorZero*/,/*cCorPos*/,"Red"/*cCorNeg*/,/*nNumFmtId*/)
	oPosNum2 := oExcel:NewStyle():SetNumFmt(nFmtNum2)

	////////////////////////////////////////////////////////////////////////////////

	oExcel:AddPlan('Parâmetros')
	oExcel:SetPrintTitles(1,1)
	oExcel:ShowGridLines(.F.)
	//------------------------------------------------------------------------------------------------------------//
	oExcel:MergeCells(1,1,2,1)
	oExcel:MergeCells(1,2,2,2)
	//oExcel:Pos(1,1):SetStyle(oTitleStyle)
	oExcel:Pos(1,2):SetValue('Parâmetros')//:SetStyle(oPosNum1)

	
	oTabela	:= oExcel:AddTabela('Parâmetros',3,1)
	oTabela:AddStyle("TableStyleMedium2"/*'TableStyleMedium2'*//*cNome*/,.T./*lLinhaTiras*/,/*lColTiras*/,.T./*lFormPrimCol*/,/*lFormUltCol*/)
	oTabela:AddFilter()

	oTabela:AddColumn("Parâmetro:")
	oExcel:AddTamCol(1,1,30,.T.,.F.)
	oTabela:AddColumn("Conteúdo:")
	oExcel:AddTamCol(2,2,30,.T.,.F.)

	oTabela:AddLine()
	oTabela:Cell(1,'Emissão:',,)
	oTabela:Cell(2,cValToChar(MsDate()),,)

	oTabela:AddLine()
	oTabela:Cell(1,'Hora:',,)
	oTabela:Cell(2,cValToChar(Time()),,)

	oTabela:AddLine()
	oTabela:Cell(1,'Usuário:',,)
	oTabela:Cell(2,"teste",,)

	oTabela:Finish()

	oExcel:Save()    //Salvar
	oExcel:OpenApp() //Abrir Excel
	oExcel:Close()   //Fechar e limpar objeto
Return
