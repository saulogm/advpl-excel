#include 'totvs.ch'
#include "Fileio.ch"
#Include "ParmType.ch"

Static cAr7Zip	//Caminho do 7zip para compactar o arquivo
Static cRootPath
//CLASSE EXCEL
/*/{Protheus.doc} YExcel
Gerar Excel formato xlsx com menor consumo de memoria e mais otimizado possivel
@author Saulo Gomes Martins
@since 27/10/2014 17:51:57
@version P11
@obs As linhas e colunas devem ser informadas sempre de forma sequencial crecente
@obs Até a versão 7.00.131227 deve ser intalado no servidor o compactador 7zip
	Adicionar no appserver.ini
	[GENERAL]
	LOCAL7ZIP=C:\Program Files\7-Zip\7z.exe
@OBS
RECURSOS DISPONIVEIS
* Definir células String,Numerica,data,Logica,formula
* Adicionar novas planilhas(Nome,Cor)
* Cor de preenchimento(simples,efeito de preenchimento)
* Alinhamento(Horizontal,Vertical,Reduzir para Caber,Quebra Texto,Angulo de Rotação)
* Formato da celula
* Mesclar células
* Auto Filtro
* Congelar painéis(colunas e linhas)
* Definir tamanho da coluna
* Definir tamanho da linha
* Formatar numeros(casas decimais)
* Letra: Fonte,Tamanho,Cor,Negrito,Italico,Sublinhado,Tachado
* Bordas: (Left,Right,Top,Bottom),Cor,Estilo
* Formatação condicional:(operador,formula)(font,fundo,bordas)
* Formatar como tabela(Estilos Predefinidos,Filtros,Totalizadores)
* Cria nome para refencia de célula ou intervalo
* Agrupamento de linha

* Leitura simples dos dados
@type class
/*/
//Dummy Function
User Function YExcel()
Return .T.

CLASS YExcel From LongClassName
	Data aRelsWorkBook
	Data oString			//String compartilhadas
	Data nQtdString			//Quantidade de string conmpartilhadas
	Data adimension			//Dimensão da planilha
	Data cClassName			//Nome da Classe
	Data cName				//Nome da Classe
	Data oFonts				//Fontes
	Data osheetData			//Objeto com dados das linhas
	Data oSyles				//Objeto com dados dos estilos
	Data aCorPreenc			//Array com as cores de preenchimento
	Data cTmpFile			//Arquivo temporario criado no servidor
	Data cNomeFile			//Nome do arquivo para gerar
	Data nFileTmpRow		//nHeader do Arquivo temporario de linhas
	Data cPlanilhaAt
	Data aPlanilhas
	Data osheetViews
	Data oCols
	Data odefinedNames
	Data oAutoFilter
	Data oMergeCells
	Data aConditionalFormatting
	Data oBorders
	Data lRowDef
	Data aSpanRow
	Data nTamLinha
	Data nColunaAtual
	Data nPriodFormCond
	Data odxfs
	Data onumFmts
	Data otableParts
	Data atable
	Data aFiles
	Data nQtdTables
	Data nCont
	Data osheetPr
	Data oCell
	Data nNumFmtId
	//Agrupamento de linha
	Data nRowoutlineLevel
	Data lRowcollapsed
	Data lRowHidden

	METHOD New() CONSTRUCTOR
	METHOD ClassName()

	//Controle das planilhas
	METHOD ADDPlan()		//Adiciona nova planilha ao arquivo
	METHOD Gravar()			//Grava em disco
	METHOD Cell()			//Grava as células
	METHOD mergeCells()		//Mescla células
	METHOD NumToString()	//Algoritimo para converte numero em string A=1,B=2
	METHOD StringToNum()	//Algoritimo para converte string em numero 1=A,2=B
	METHOD Ref()			//Passa a localização numerica e transforma em referencia da celula
	METHOD LocRef()			//Retorna linha  e coluna de acordo com referencia enviada
	METHOD SetDefRow()		//Defini as colunas da linha. Habilita a gravação automatica de cada coluna. Importante para prover performace na gravação de varias linhas
	METHOD AddTamCol()		//Defini o tamanho de uma coluna ou varias colunas
	METHOD AddPane()		//Congelar Painéis
	METHOD AutoFilter()		//Cria os Filtros na planilha
	METHOD AddNome()		//Cria nome para refencia de célula ou intervalo
	METHOD NivelLinha()

	//Leitura de planilha
	METHOD OpenRead()
	METHOD CellRead()
	METHOD CloseRead()

	//Interno
	METHOD CriarFile()		//Cria arquivos temporarios
	METHOD GravaFile()		//Grava em arquivos temporarios
	METHOD GravaRow()		//Grava temporario de linhas
	METHOD AddBorda()		//Adiciona borda
	METHOD AddFormatCond()	//Formatação condicional(todos rercusos)
	METHOD Pane()			//Congelar Painéis
	METHOD AddAgrCol()		//Em Desenvolvimento

	//Estilo
	METHOD CorPreenc()		//Cria um nova cor para ser usada
	METHOD EfeitoPreenc()	//Cria um novo efeito de preenchimento
	METHOD AddFont()		//Cria objeto de font
	METHOD AddStyles()		//Adiciona Estilos
	METHOD Alinhamento()	//Adiciona alinhamento
	METHOD Borda()			//Adiciona borda(auxiliar)
	Method AddFmtNum()		//Cria formato para numeros

	//Formatação condicional
	METHOD FormatCond()		//Definir formatação condicional(auxiliar)
	METHOD Font()			//Cria objeto de font
	METHOD Preenc()			//Cria objeto de Preenchimento
	METHOD ObjBorda()		//Cria objeto de borda
	METHOD gradientFill()	//Cria objeto de efeito de preenchimento
	METHOD ADDdxf()			//Cria o estilo para formatação condicional
	//Tabela
	METHOD AddTabela()
	/*Tabela
		Method cell()
		Method AddStyle()
		Method AddLine()
		Method AddColumn()
		Method AddFilter()
		METHOD AddTotal()
		METHOD AddTotais()
		METHOD Finish()
	*/
ENDCLASS

/*/{Protheus.doc} AddNome
Cria nome para refencia de célula ou intervalo
@author Saulo Gomes Martins
@since 09/05/2017
@param cNome, characters, Nome
@param nLinha, numeric, Linha da referencia
@param nColuna, numeric, Coluna da referencia
@param [nLinha2], numeric, Linha final se intervalo
@param [nColuna2], numeric, Coluna final se intervalo
@param [cRefPar], characters, Rerefencia
@param [cPlanilha], characters, Planilha
@type function
/*/
METHOD AddNome(cNome,nLinha,nColuna,nLinha2,nColuna2,cRefPar,cPlanilha) CLASS YExcel
	Local odefinedName	:= yExcelTag():New("definedName",)
	Local cRef
	PARAMTYPE 0	VAR cNome			AS CHARACTER
	PARAMTYPE 1	VAR nLinha			AS NUMERIC			OPTIONAL
	PARAMTYPE 2	VAR nColuna	  		AS NUMERIC			OPTIONAL
	PARAMTYPE 3	VAR nLinha2	  		AS NUMERIC			OPTIONAL
	PARAMTYPE 4	VAR nColuna2  		AS NUMERIC			OPTIONAL
	PARAMTYPE 5	VAR cRefPar	 	 	AS CHARACTER		OPTIONAL
	PARAMTYPE 6	VAR cPlanilha	  	AS CHARACTER		OPTIONAL DEFAULT ::cPlanilhaAt

	odefinedName:SetAtributo("name",cNome)
	If ValType(cRefPar)=="U"
		cRef	:= "'"+cPlanilha+"'!"+::Ref(nLinha,nColuna,.T.,.T.)
		If Valtype(nLinha2)<>"U" .and. Valtype(nColuna2)<>"U"
			cRef	+= ":"+::Ref(nLinha2,nColuna2,.T.,.T.)
		Endif
	Else
		cRef	:= cRefPar
	EndiF
	odefinedName:SetValor(cRef)
	::odefinedNames:AddValor(odefinedName,cNome)
Return

METHOD ClassName() CLASS YExcel
Return "YExcel"

/*/{Protheus.doc} New
Construtor da classe
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param cNomeFile, characters, Nome do arquivo para gerar
@type function
/*/
METHOD New(cNomeFile) CLASS YExcel
	Local oPlanilha,oRelationship
	Local oMoeda
	If ValType(cAr7Zip)=="U"
		cAr7Zip := GetPvProfString("GENERAL", "LOCAL7ZIP" , "C:\Program Files\7-Zip\7z.exe" , GetAdv97() )
	Endif
	PARAMTYPE 0	VAR cNomeFile  AS CHARACTER 		OPTIONAL DEFAULT CriaTrab(,.F.)
	::cClassName	:= "YExcel"
	::cName			:= "YExcel"
	::oString		:= tHashMap():new()
	::oCell			:= tHashMap():new()	//Usado no leitura simples
	::oBorders		:= yExcelTag():New("borders",{})
	::odxfs			:= yExcelTag():New("dxfs",{})
	::odxfs:SetAtributo("count",0)
	::onumFmts		:= yExcelTag():New("numFmts",{})	//Formatos de numeros
	::onumFmts:SetAtributo("count",0)
	::odefinedNames	:= yExcelTag():New("definedNames",{})
	::Borda()	//Sem borda
	::nQtdString	:= 0
	::oFonts		:= YExcelFont():New()
	::AddFont(11,"FF000000","Calibri","2")
	::oSyles		:= yExcel_cellXfs():New()
	::nNumFmtId		:= 164
	::AddStyles(0/*numFmtId*/,/*fontId*/,/*fillId*/,/*borderId*/,/*xfId*/,/*aValores*/,/*aOutrosAtributos*/)	//Sem Formatação
	::AddStyles(14/*numFmtId*/,/*fontId*/,/*fillId*/,/*borderId*/,/*xfId*/,/*aValores*/,{{"applyNumberFormat","0"}}/*aOutrosAtributos*/)	//Formato Data padrão
	::aRelsWorkBook	:= {}
	::aPlanilhas	:= {}
	::aCorPreenc	:= {}	//yExcel_CorPreenc():New()
	::cTmpFile		:= CriaTrab(,.F.)
	::cNomeFile		:= cNomeFile
	::nFileTmpRow	:= 0
	::lRowDef		:= .F.
	::nColunaAtual	:= 0
	::aFiles	:= {}
	::nQtdTables	:= 0
	oMoeda	:= yExcelTag():New("numFmt")
	oMoeda:SetAtributo("formatCode","_-&quot;R$&quot;\ * #,##0.00_-;\-&quot;R$&quot;\ * #,##0.00_-;_-&quot;R$&quot;\ * &quot;-&quot;??_-;_-@_-")
	oMoeda:SetAtributo("numFmtId",44)
	::onumFmts:AddValor(oMoeda)
	::onumFmts:SetAtributo("count",1)

	AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\[Content_Types].xml")
	AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\_rels\.rels")
	AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\docProps\app.xml")
	AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\docProps\core.xml")
	AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\sharedStrings.xml")
	AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\styles.xml")
	AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\workbook.xml")
	AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\_rels\workbook.xml.rels")
	AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\theme\theme1.xml")

	oRelationship	:= yExcelTag():New("Relationship",)
	oRelationship:SetAtributo("Id","rId1")
	oRelationship:SetAtributo("Type","http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
	oRelationship:SetAtributo("Target","theme/theme1.xml")
	AADD(::aRelsWorkBook,oRelationship)
	oRelationship	:= yExcelTag():New("Relationship",)
	oRelationship:SetAtributo("Id","rId2")
	oRelationship:SetAtributo("Type","http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles")
	oRelationship:SetAtributo("Target","styles.xml")
	AADD(::aRelsWorkBook,oRelationship)
	oRelationship	:= yExcelTag():New("Relationship",)
	oRelationship:SetAtributo("Id","rId3")
	oRelationship:SetAtributo("Type","http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings")
	oRelationship:SetAtributo("Target","sharedStrings.xml")
	AADD(::aRelsWorkBook,oRelationship)


	AADD(::aCorPreenc,yExcel_CorPreenc():New("none"))
	AADD(::aCorPreenc,yExcel_CorPreenc():New("gray125"))
Return self

/*/{Protheus.doc} OpenRead
Abrir planilha e armazena conteudo para leitura
@author Saulo Gomes Martins
@since 03/03/2018
@version 1.0
@return lRet, Se conseguiu ler a planilha
@param cFile, characters, arquivo que será aberto
@param nPlanilha, numeric, numero(indexado em 1,2,3) da planilha a ser lida
@type function
/*/
METHOD OpenRead(cFile,nPlanilha) Class YExcel
	Local nRet
	Local cDrive, cDir, cNome, cExt
	Local nCont,nCont2,nCont3
	Local cTipo,cRef
	Local aChildren,aChildren2,aAtributos,oXml
	Local cCamSrv	:= ""
	PARAMTYPE 0	VAR cFile			AS CHARACTER
	PARAMTYPE 1	VAR nPlanilha	  	AS NUMERIC		OPTIONAL DEFAULT 1
	cFile	:= Alltrim(cFile)
	If !File(cFile)
		ConOut("Arquivo nao encontrado!")
		Return .F.
	EndIf
	If ValType(cRootPath)=="U"
		cRootPath	:= GetSrvProfString( "RootPath", "" )
	EndIf
	If !File("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\worksheets\sheet"+cValTochar(nPlanilha)+".xml")
		SplitPath( cFile, @cDrive, @cDir, @cNome, @cExt)
		FWMakeDir("\tmpxls\"+::cTmpFile+'\',.F.)
		FWMakeDir("\tmpxls\"+::cTmpFile+'\'+::cNomeFile+'\',.F.)
		If "C:" $ UPPER(cFile)
			CpyT2S(cFile,"\tmpxls\"+::cTmpFile+'\',,.F.)
			cCamSrv	:= cRootPath+"\tmpxls\"+::cTmpFile+'\'+cNome+cExt
		Else
			cCamSrv	:= cRootPath+cFile
		EndIf
		If !FindFunction("FZIP")
			WaitRunSrv('"'+cAr7Zip+'" x -tzip "'+cCamSrv+'" -o"'+cRootPath+'\tmpxls\'+::cTmpFile+'\'+::cNomeFile+'" * -r -y',.T.,"C:\")
			If !File("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\sharedStrings.xml")
				nRet	:= -1
				ConOut("Arquivo nao descompactado!")
				Return .F.
			Else
				nRet	:= 0
			EndIf
		Else
			nRet	:= FUnZip("\tmpxls\"+::cTmpFile+'\'+cNome+cExt,"\tmpxls\"+::cTmpFile+'\'+::cNomeFile+'\')
		EndIf
		If nRet!=0
			ConOut(Ferror())
			ConOut("Arquivo nao descompactado!")
			Return .F.
		EndIf
		oXml	:= TXmlManager():New()
		oXML:ParseFile("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\sharedStrings.xml")
		oXML:XPathRegisterNs( "ns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main" )
		aChildren := oXML:XPathGetChildArray( "/ns:sst" )
		For nCont:=1 to Len(aChildren)
			::oString:Set(::nQtdString,oXML:XPathGetNodeValue("/ns:sst/ns:si["+cValToChar(nCont)+"]/ns:t"))
			::nQtdString++
		Next
	EndIf
	oXml	:= TXmlManager():New()
	oXML:ParseFile("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\worksheets\sheet"+cValTochar(nPlanilha)+".xml")
	oXML:XPathRegisterNs( "ns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main" )

    aChildren := oXML:XPathGetChildArray("/ns:worksheet/ns:sheetData")
	::adimension	:= {{0,0},{999999,999999}}
    For nCont:=1 to Len(aChildren)
    	aChildren2	:= oXML:XPathGetChildArray( aChildren[nCont][2] )
    	For nCont2:=1 to Len(aChildren2)
    		cTipo		:= "N"
    		aAtributos	:= oXML:XPathGetAttArray(aChildren2[nCont2][2])						//Atributos do elemento
    		cRet		:= oXML:XPathGetNodeValue("/ns:worksheet/ns:sheetData/ns:row["+cValToChar(nCont)+"]/ns:c["+cValToChar(nCont2)+"]/ns:v")
    		For nCont3:=1 to Len(aAtributos)
    			If aAtributos[nCont3][1]=="r"
    				cRef	:= aAtributos[nCont3][2]
    				aPosicao	:= ::LocRef(cRef)	//Retorna linha e coluna
					If ::adimension[2][1]>aPosicao[1]	//Menor linha
						::adimension[2][1] := aPosicao[1]
					EndIf
					If ::adimension[2][2]>aPosicao[2]	//Menor Coluna
						::adimension[2][2]	:= aPosicao[2]
					EndIf
					If ::adimension[1][1]<aPosicao[1]	//Maior Linha
						::adimension[1][1]	:= aPosicao[1]
					EndIf
					If ::adimension[1][2]<aPosicao[2]	//Maior Coluna
						::adimension[1][2]	:= aPosicao[2]
					EndIf
    			ElseIf aAtributos[nCont3][1]=="t" .and. aAtributos[nCont3][2]=="s"
    				cRet	:= ""
    				cTipo	:= "C"
    				::oString:Get(Val(oXML:XPathGetNodeValue("/ns:worksheet/ns:sheetData/ns:row["+cValToChar(nCont)+"]/ns:c["+cValToChar(nCont2)+"]/ns:v")),@cRet)
    			ElseIf aAtributos[nCont3][1]=="t" .and. aAtributos[nCont3][2]=="b"
    				cTipo	:= "L"
    				cRet	:= oXML:XPathGetNodeValue("/ns:worksheet/ns:sheetData/ns:row["+cValToChar(nCont)+"]/ns:c["+cValToChar(nCont2)+"]/ns:v")=="1"
    			ElseIf aAtributos[nCont3][1]=="s" .and. aAtributos[nCont3][2]=="1"
    				cTipo	:= "D"
    				cRet	:= STOD("19000101")-2+Val(oXML:XPathGetNodeValue("/ns:worksheet/ns:sheetData/ns:row["+cValToChar(nCont)+"]/ns:c["+cValToChar(nCont2)+"]/ns:v"))
    			EndIf
    		Next
    		If cTipo=="N"
	    		::oCell:Set(cRef,Val(cRet))
    		Else
    			::oCell:Set(cRef,cRet)
    		EndIf
    	Next
    Next
Return nRet==0

/*/{Protheus.doc} CellRead
Retorna o valor de uma celula, após o uso do método OpenRead()
@author Saulo Gomes Martins
@since 03/03/2018
@version 1.0
@return xValor, Conteúdo da celula
@param nLinha, numeric, Linha da informação
@param nColuna, numeric, Coluna da informação
@param xDefault, naodefinido , Valor padrão caso não tenha a informação
@param lAchou, logical, passa por referencia se achou a informação da celula
@type function
/*/
Method CellRead(nLinha,nColuna,xDefault,lAchou) Class YExcel
	Local cRef	:= ::Ref(nLinha,nColuna)
	Local xValor:= Nil
	lAchou	:= .T.
	If !::oCell:Get(cRef,@xValor)
		xValor	:= xDefault
		lAchou	:= .F.
	EndIf
Return xValor

/*/{Protheus.doc} CloseRead
Limpa a pasta temporaria
@author Saulo Gomes Martins
@since 03/03/2018
@version 1.0

@type function
/*/
METHOD CloseRead() Class YExcel
	::oString:clean()
	::oCell:clean()
	::nQtdString := 0
	DelPasta("\tmpxls\"+::cTmpFile)
Return

/*/{Protheus.doc} ADDPlan
Adiciona nova planilha ao arquivo
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param cNome, characters, nome da planilha
@type function
/*/
METHOD ADDPlan(cNome,cCor) CLASS YExcel
	Local nPos	:= Len(::aRelsWorkBook)+1
	Local oPlanilha,nFile,oSelection
	Local nQtdPlanilhas	:= Len(::aPlanilhas)
	Local nCont
	Local oCorPlan
	Private oSelf	:= Self
	PARAMTYPE 0	VAR cNome		  	AS CHARACTER		OPTIONAL DEFAULT "Planilha"+cValToChar(nQtdPlanilhas+1)
	PARAMTYPE 1	VAR cCor			AS CHARACTER		OPTIONAL
	If Len(cNome)>31
		cNome	:= SubStr(cNome,1,31)
	EndIf
	If nQtdPlanilhas>0	//Grava a Planilha anterior
		If Empty(::oCols:GetValor())
			::oCols:AddValor(yExcelTag():New("col"))
			::oCols:GetValor(1):SetAtributo("min",::adimension[2][2])
			::oCols:GetValor(1):SetAtributo("max",::adimension[1][2])
			::oCols:GetValor(1):SetAtributo("width",12.00)
			::oCols:GetValor(1):SetAtributo("bestFit","1")
			::oCols:GetValor(1):SetAtributo("customWidth","1")
		EndIf
		//Grava a ultima linha
		::GravaRow(::adimension[1][1])
		::CriarFile("\"+::cNomeFile+"\xl\worksheets"	,"sheet"+cValToChar(nQtdPlanilhas)+".xml"			,""			,)
		GravaFile(@nFile,"","\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\worksheets","sheet"+cValToChar(nQtdPlanilhas)+".xml")
		h_xls_sheet(nFile)
		fClose(nFile)
		fErase("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\worksheets\tmprow.xml")

		If !Empty(::atable)
			::CriarFile("\"+::cNomeFile+"\xl\worksheets\_rels\"	,"sheet"+cValToChar(nQtdPlanilhas)+".xml.rels"		,h_xlsrelssheet()		,)
			AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\worksheets\_rels\sheet"+cValToChar(nQtdPlanilhas)+".xml.rels")
			For nCont:=1 to Len(::atable)
				::nCont	:= nCont
				::CriarFile("\"+::cNomeFile+"\xl\tables\"	,"table"+cValToChar(::nQtdTables-Len(::atable)+nCont)+".xml"		,h_xls_table()		,)
			AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\tables\table"+cValToChar(::nQtdTables-Len(::atable)+nCont)+".xml")
			Next
		EndIf
	EndIf
	AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\worksheets\sheet"+cValToChar(nQtdPlanilhas+1)+".xml")
	If ValType(cCor)=="C"
		If Len(cCor)==6
			cCor	:= "FF"+cCor
		EndIf
		oCorPlan		:= yExcelTag():New("tabColor",,{{"rgb",cCor}})
	EndIf
	::osheetPr		:= yExcelTag():New("sheetPr",{oCorPlan},{{"codeName",cNome}})
	::adimension	:= {{0,0},{999999,999999}}
	::osheetData	:= yExcelsheetData():New(self)
	::osheetViews	:= yExcelTag():New("sheetViews",yExcelTag():New("sheetView",{}))
	oSelection	:= yExcelTag():New("selection",nil)
	If nQtdPlanilhas==0
		::osheetViews:GetValor():SetAtributo("tabSelected",1)
		//oSelection:SetAtributo("activeCell","A1")
	EndIf
	oSelection:SetAtributo("sqref","A1")
	AADD(::osheetViews:GetValor():GetValor(),oSelection)
	::cPlanilhaAt				:= cNome
	::osheetViews:GetValor():SetAtributo("workbookViewId",0)
	::nColunaAtual				:= 0
	::oAutoFilter				:= nil
	::oMergeCells				:= yExcelTag():New("mergeCells",{})
	::aConditionalFormatting	:= {}
	::nPriodFormCond			:= 1
//	::osheetViews:GetValor():SetAtributo("tabSelected","1")
	::osheetViews:GetValor():SetAtributo("workbookViewId","0")
	::oCols			:= yExcelTag():New("cols",{})
	::otableParts	:= yExcelTag():New("tableParts",{})
	::atable		:= {}
	::nRowoutlineLevel	:= nil
	::lRowcollapsed		:= .F.
	::lRowHidden		:= .F.

	//Cria arquivo temporario de gravação das linhas
	::CriarFile("\"+::cNomeFile+"\xl\worksheets"	,"tmprow.xml"			,""			,)
	GravaFile(@::nFileTmpRow,"","\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\worksheets","tmprow.xml")
	//Cria nova planilha
	nQtdPlanilhas++
	oPlanilha		:= yExcelTag():New("sheet",)
	oPlanilha:SetAtributo("name",cNome)
	oPlanilha:SetAtributo("sheetId",nQtdPlanilhas)
	oPlanilha:SetAtributo("r:id","rId"+cValToChar(nPos))
	AADD(::aPlanilhas,oPlanilha)
	//Cria novo arquivo
	oRelationship	:= yExcelTag():New("Relationship",)
	oRelationship:SetAtributo("Id","rId"+cValToChar(nPos))
	oRelationship:SetAtributo("Type","http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
	oRelationship:SetAtributo("Target","worksheets/sheet"+cValToChar(nQtdPlanilhas)+".xml")
	AADD(::aRelsWorkBook,oRelationship)
Return nQtdPlanilhas

/*/{Protheus.doc} Cell
Grava o conteudo de uma célula
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param nLinha, numeric, Linha a ser gravada
@param nColuna, numeric, Coluna a ser gravada
@param xValor, qualquer, Valor a ser gravado(texto,numero,data,logico)
@param cFormula, characters, Formula da célula
@param nStyle, numeric, posição do estilo criado pelo metodo :AddStyles()
@type function
/*/
METHOD Cell(nLinha,nColuna,xValor,cFormula,nStyle) CLASS YExcel
	Local oExcelRow
	PARAMTYPE 0	VAR nLinha		  	AS NUMERIC
	PARAMTYPE 1	VAR nColuna		  	AS NUMERIC
	PARAMTYPE 3	VAR cFormula	  	AS CHARACTER	OPTIONAL
	PARAMTYPE 4	VAR nStyle		  	AS NUMERIC		OPTIONAL
	If ValType(nColuna)=="C"
		//nColuna		:= ::odefinedNames:GetValor(nColuna,""):GetAtributo("name")
		nColuna	:= StringToNum(UPPER(nColuna))
	EndIf
	If nColuna==0
		UserException("YExcel - O índice da coluna não pode iniciar no 0")
	EndIf
	If nLinha<::adimension[2][1] .and. ::adimension[2][1]<>999999
		UserException("YExcel - As linhas devem ser informadas sempre de forma sequencial crecente. Linha informada:"+cValToChar(nLinha)+" | Linha Atual:"+cValToChar(::adimension[2][1])+".")
	EndIf
	If ::adimension[1][1]==nLinha .AND. nColuna==::nColunaAtual
		UserException("YExcel - Não é possivel redefinir o valor da celula gravada.")
	ElseIf ::adimension[1][1]==nLinha .AND. nColuna<=::nColunaAtual
		UserException("YExcel - As colunas devem ser informadas sempre de forma sequencial crecente. Coluna informada:"+cValToChar(nColuna)+" | Coluna Atual:"+cValToChar(::nColunaAtual)+".")
	Endif
	::nColunaAtual	:= nColuna
	If ::adimension[2][1]>nLinha	//Menor linha
		::adimension[2][1]	:= nLinha
	EndIf
	If ::adimension[2][2]>nColuna	//Menor Coluna
		::adimension[2][2]	:= nColuna
	EndIf
	If ::adimension[1][1]<nLinha	//Maior Linha
		If ::adimension[1][1]>0		//Primeira vez não faz
			::GravaRow(::adimension[1][1])
		EndIf
		::adimension[1][1]	:= nLinha
		If ::lRowDef
			oExcelRow		:= ::osheetData:Add(nLinha)
			oExcelRow:GetTag(::nFileTmpRow,.F./*não finaliza a tag*/)
		EndIf
	EndIf

	::osheetData:SetVal(nLinha,nColuna,xValor,cFormula,nStyle)

	If ::adimension[1][2]<nColuna	//Maior Coluna
		::adimension[1][2]	:= nColuna
	EndIf
	If ::lRowDef
		If ::osheetData:GetValor():Get(nLinha,@oExcelRow)
			oExcelRow:GetTag(::nFileTmpRow,,.T.)	//Grava em temporario
			oExcelRow:GetValor():Del(nColuna)			//Já exclui da memoria
		EndIf
	EndIf
Return self

/*/{Protheus.doc} SetDefRow
Defini as colunas da linha. Habilita a gravação automatica de cada coluna. Importante para prover performace na gravação de varias linhas
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param lHabilitar, logical, Habilita a definição
@param aSpanRow, array, 1-Coluna inicial|2-Coluna Final
@type function
/*/
METHOD SetDefRow(lHabilitar,aSpanRow) CLASS YExcel
	Default	lHabilitar	:= .T.
	::lRowDef		:= lHabilitar
	::aSpanRow		:= aSpanRow
Return

METHOD NivelLinha(nNivel,lFechado,lOculto) CLASS YExcel
	Default lFechado	:= .F.
	Default lOculto		:= .F.
	::nRowoutlineLevel	:= nNivel
	::lRowcollapsed		:= lFechado
	::lRowHidden		:= lOculto
Return
/*/{Protheus.doc} AutoFilter
Cria os Filtros na planilha
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param nLinha, numeric, Linha inicial
@param nColuna, numeric, Coluna Inicial
@param nLinha2, numeric, Linha final
@param nColuna2, numeric, Coluna Final
@type function
/*/
Method AutoFilter(nLinha,nColuna,nLinha2,nColuna2) CLASS YExcel
	Local cColuna,cColuna2
	cColuna		:= NumToString(nColuna)
	cColuna2	:= NumToString(nColuna2)
	::oAutoFilter	:= yExcelTag():New("autoFilter",)
	::oAutoFilter:SetAtributo("ref",cColuna+cValToChar(nLinha)+":"+cColuna2+cValToChar(nLinha2))
Return

/*/{Protheus.doc} mergeCells
Mescla células
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param nLinha, numeric, Linha inicial
@param nColuna, numeric, Coluna Inicial
@param nLinha2, numeric, Linha final
@param nColuna2, numeric, Coluna Final
@type function
/*/
Method mergeCells(nLinha,nColuna,nLinha2,nColuna2) CLASS YExcel
	Local oMergeCell,cColuna,cColuna2,nPos
	Local oExcelRow
	If nLinha2<nLinha
		UserException("YExcel - metodo mergeCells. Linha final não pode ser menor que linha inicial.")
	EndIf
	If nColuna2<nColuna
		UserException("YExcel - metodo mergeCells. Coluna final não pode ser menor que Coluna inicial.")
	EndIf
	cColuna		:= NumToString(nColuna)
	cColuna2	:= NumToString(nColuna2)
	nPos		:= aScan(::oMergeCells:GetValor(),{|x| Replace(cColuna+cValToChar(nLinha),"$","") $ Replace(x:GetAtributo("ref"),"$","") .OR. Replace(cColuna2+cValToChar(nLinha2),"$","") $ Replace(x:GetAtributo("ref"),"$","") })
	If nPos>0
		UserException("YExcel - metodo mergeCells. Célula "+cColuna+cValToChar(nLinha)+":"+cColuna2+cValToChar(nLinha2)+" não pode ser mesclada, essa célula já foi mesclada!")
	EndIf
	oMergeCell	:= yExcelTag():New("mergeCell",)
	oMergeCell:SetAtributo("ref",cColuna+cValToChar(nLinha)+":"+cColuna2+cValToChar(nLinha2))
	::oMergeCells:AddValor(oMergeCell)
	::oMergeCells:SetAtributo("count",Len(::oMergeCells:GetValor()))
	If ::osheetData:GetValor():Get(nLinha,@oExcelRow)
		If oExcelRow:aspans[2]<nColuna2
			oExcelRow:aspans[2]	:= nColuna2
			oExcelRow:SetAtributo("spans","1:"+cValToChar(nColuna2))
		EndIf
	EndIf
Return

/*/{Protheus.doc} Font
Cria objeto de fonte para ser usado na criação de estilos para formatação condicional
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param [nTamanho], numeric, Tamanho da fonte
@param [cCorRGB], characters, Cor da fonte em Alpha+RGB
@param [cNome], characters, Nome da fonte
@param [cfamily], characters, Familia da fonte
@param [cScheme], characters, Schema
@param [lNegrito], logical, Negrito
@param [lItalico], logical, Italico
@param [lSublinhado], logical, Soblinhado
@param [lTachado], logical, Tachado
@type function
/*/
METHOD Font(nTamanho,cCorRGB,cNome,cfamily,cScheme,lNegrito,lItalico,lSublinhado,lTachado) CLASS YExcel
	oFont	:= YExcelFont():New()
	oFont:Add(nTamanho,cCorRGB,cNome,cfamily,cScheme,lNegrito,lItalico,lSublinhado,lTachado)
Return oFont:GetValor(1)

/*/{Protheus.doc} Preenc
Cria objeto de preenchimento para ser usado na criação de estilos para formatação condicional
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param [cBgCor], characters, Cor em Alpha+RGB do preenchimento
@param [cFgCor], characters, Cor em Aplha+RGB do fundo
@param [cType], characters, tipo de preenchimento(padrão solid)
@type function
/*/
METHOD Preenc(cBgCor,cFgCor,cType) CLASS YExcel
Default cType	:= "solid"
Return oCorPreenc	:= yExcel_CorPreenc():New(cType,cFgCor,cBgCor)

/*/{Protheus.doc} ObjBorda
Cria objeto de bordas para ser usado na criação de estilos para formatação condicional
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param cTipo, characters, "C"-Cima|"B"-Baixo|"E"-Esquerda|"D"-Direita|T-TODAS("CBED") OU "T"-TOP|"B"-Bottom|"L"-Left|"R"-Right|A-ALL("TBLR") OU "DIAGONAL"-Diagonal
@param cCor, characters, Cor em Aplha+RGB da borda
@param cModelo, characters, Modelo da borda
@type function
@Obs pode juntar os tipo. Exemplo "ED"-Esquerda e direita
/*/
METHOD ObjBorda(cTipo,cCor,cModelo) CLASS YExcel
	Local oBorda
	::Borda(cTipo,cCor,cModelo,@oBorda,.F.)
Return oBorda

/*/{Protheus.doc} ADDdxf
Cria estilo para formatação condicional
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param [oFont], object, objeto criado pelo metodo :Font() com fonte
@param [oCorPreenc], object, objeto com cor criado pelo metodo :Preench() de preenchimento
@param [oBorda], object, objeto criado pelo metodo :ObjBorda() com borda
@return posição do estilo
@type function
/*/
METHOD ADDdxf(oFont,oCorPreenc,oBorda) CLASS YExcel
	Local odxf
	Local nPos

	::odxfs:AddValor(yExcelTag():New("dxf",{}))
	nPos	:= Len(::odxfs:GetValor())
	::odxfs:SetAtributo("count",nPos)
	odxf	:= ::odxfs:GetValor(nPos)

	//Font
	If ValType(oFont)<>"U"
		odxf:AddValor(oFont)
	EndIf
	//Preenchimento
	If ValType(oCorPreenc)<>"U"
		odxf:AddValor(oCorPreenc)
	EndIf
	//Borda
	If ValType(oBorda)<>"U"
		odxf:AddValor(oBorda)
	EndIf
Return nPos-1

/*/{Protheus.doc} FormatCond
Cria uma regra para formatação condicional
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param cRefDe, characters, Rerefencia inicial (Exemplo "A1")
@param cRefAte, characters, Referencia final (Exemplo "A10")
@param nEstilo, numeric, posição do estilo criado pelo metodo :ADDdxf()
@param operator, object, Operado da regra. veja observações
@param xFormula, characters ou array, formula para uso
@type function
@obs operadores
	=		igual
	!=		diferente
	>		Maior que
	>=		Maior ou igual
	<		Menor que
	<=		Menor ou igual
	$		contem texto
	!$		não contem
	between	esta contido(enviar no parametro xformula um array com duas posições)
	FORMULA	enviar no parametro xformula a regra
/*/
METHOD FormatCond(cRefDe,cRefAte,nEstilo,operator,xFormula) CLASS YExcel
	//ST_ConditionalFormattingOperator Pag 2455
	/*
	beginsWith
	between
	containsText
	endsWith
	equal
	greaterThan
	greaterThanOrEqual
	lessThan
	lessThanOrEqual
	notBetween
	notContains
	notEqual
	*/
	If operator=="=" .or. operator=="=="
		operator	:= "equal"
	ElseIf operator=="!=" .or. operator=="<>"
		operator	:= "notEqual"
	ElseIf operator==">"
		operator	:= "greaterThan"
	ElseIf operator==">="
		operator	:= "greaterThanOrEqual"
	ElseIf operator=="<"
		operator	:= "lessThan"
	ElseIf operator=="<="
		operator	:= "lessThanOrEqual"
	ElseIf operator=="$"
		operator	:= "containsText"
	ElseIf operator=="!$"
		operator	:= "notContains"
	EndIf
	If operator=="between"	.and. (ValType(xFormula)<>"A" .or. Len(xFormula)<2)
		UserException("YExcel - operador between é necessario informar valor de, ate. Enviar parametro 5 xformula como array(2).")
	EndIf
	If operator=="FORMULA"
		::AddFormatCond(cRefDe,cRefAte,nEstilo,"expression",xFormula,,)
	Else
		::AddFormatCond(cRefDe,cRefAte,nEstilo,"cellIs",xFormula,operator)
	EndIf
	::nPriodFormCond++
Return
//NÃO DOCUMENTAR
METHOD AddFormatCond(cRefDe,cRefAte,nEstilo,cType,xFormula,operator,nPrioridade) CLASS YExcel
	Local cRef	:= cRefDe+If(!Empty(cRefAte),":"+cRefAte,"")
	Local oFormula,oRule,nCont
	PARAMTYPE 0	VAR cRefDe 		AS CHARACTER
	PARAMTYPE 1	VAR cRefAte  	AS CHARACTER				OPTIONAL
	PARAMTYPE 2	VAR nEstilo  	AS NUMERIC
	PARAMTYPE 3	VAR cType  		AS CHARACTER
	PARAMTYPE 4	VAR xFormula  	AS ARRAY,CHARACTER,NUMERIC
	PARAMTYPE 5	VAR operator  	AS CHARACTER				OPTIONAL
	PARAMTYPE 6	VAR nPrioridade	AS NUMERIC 					OPTIONAL DEFAULT ::nPriodFormCond
	/*	TYPES	(pag 2452)
	aboveAverage	-	abaixo da media
	beginsWith		-	inicia com
	cellIs			-	celula é(usar operador)
	colorScale		-	Estala de cor
	expression		-	Usar Formula
	top10			-
	...
	*/
	If ValType(xFormula)<>"U"
		If ValType(xFormula)=="A"
			oFormula	:= {}
			For nCont:=1 to Len(xFormula)
				AADD(oFormula,yExcelTag():New("formula",xFormula[nCont]))
			Next
		Else
			oFormula	:= yExcelTag():New("formula",xFormula)
		EndIf
	EndIf
	oRule	:= yExcelTag():New("cfRule",oFormula)
	oRule:SetAtributo("type",cType)
	oRule:SetAtributo("dxfId",nEstilo)
	oRule:SetAtributo("priority",nPrioridade)
	If ValType(operator)<>"U"
		oRule:SetAtributo("operator",operator)
	EndIf
	nPos	:= aScan(::aConditionalFormatting,{|x| x:GetAtributo("sqref")==cRef})
	If nPos==0
		AADD(::aConditionalFormatting,yExcelTag():New("conditionalFormatting",{oRule},{{"sqref",cRef}}))
	Else
		::aConditionalFormatting[nPos]:AddValor(oRule)
		aSort(::aConditionalFormatting[nPos]:GetValor(),,,{|x,y| x:GetAtributo("priority")>y:GetAtributo("priority")})
	EndIf
Return

/*/{Protheus.doc} AddFont
Adiciona fonte para ser usado no estilo das células
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param [nTamanho], numeric, Tamanho da fonte
@param [cCorRGB], characters, Cor da fonte em Alpha+RGB
@param [cNome], characters, Nome da fonte
@param [cfamily], characters, Familia da fonte
@param [cScheme], characters, Schema
@param [lNegrito], logical, Negrito
@param [lItalico], logical, Italico
@param [lSublinhado], logical, Soblinhado
@param [lTachado], logical, Tachado
@return posição da fonte
@type function
/*/
METHOD AddFont(nTamanho,cCorRGB,cNome,cfamily,cScheme,lNegrito,lItalico,lSublinhado,lTachado) CLASS YExcel
Return ::oFonts:Add(nTamanho,cCorRGB,cNome,cfamily,cScheme,lNegrito,lItalico,lSublinhado,lTachado)

/*/{Protheus.doc} CorPreenc
Adiciona cor de preenchimento para ser usado no estilo das células
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param [cBgCor], characters, Cor em Alpha+RGB do preenchimento
@param [cFgCor], characters, Cor em Aplha+RGB do fundo
@param [cType], characters, tipo de preenchimento(padrão solid)
@type function
/*/
METHOD CorPreenc(cFgCor,cBgCor,cType) CLASS YExcel
	Local nPos
	Default cType	:= "solid"
	AADD(::aCorPreenc,yExcel_CorPreenc():New(cType,cFgCor,cBgCor))
	nPos	:= Len(::aCorPreenc)-1
Return nPos

/*/{Protheus.doc} EfeitoPreenc
Adiciona cor com efeito de preenchimento
@author Saulo Gomes Martins
@since 17/05/2017
@version p11
@param nAngulo, numeric, Angulo para efeito de preenchimento
@param aCores, array, Cores de preenchimento {{CorRGB,nPerc},{"FF0000",0.5}}
@param [ctype], characters, Tipo de efeito (linear ou path)
@param [nleft], numeric, para efeito path posição esquerda
@param [nright], numeric, para efeito path posição direita
@param [ntop], numeric, para efeito path posição topo
@param [nbottom], numeric, para efeito path posição inferior
@return nPos, Posição para criação de estilo
@type function
/*/
METHOD EfeitoPreenc(nAngulo,aCores,ctype,nleft,nright,ntop,nbottom) CLASS YExcel
	Local nPos
	Local ogradientFill	:= ::gradientFill(nAngulo,aCores,ctype,nleft,nright,ntop,nbottom)
	AADD(::aCorPreenc,yExcelTag():New("fill",ogradientFill))
	nPos	:= Len(::aCorPreenc)-1
Return nPos

METHOD gradientFill(nAngulo,aCores,ctype,nleft,nright,ntop,nbottom) CLASS YExcel	//Pag 1779
	Local nPos,nCont
	Local ogradientFill	:= yExcelTag():New("gradientFill",{})
	Local ostop
	PARAMTYPE 0	VAR nAngulo 		AS NUMERIC		OPTIONAL
	PARAMTYPE 1	VAR aCores	 		AS ARRAY
	PARAMTYPE 2	VAR ctype	 		AS CHARACTER	OPTIONAL
	PARAMTYPE 3	VAR nleft	 		AS NUMERIC		OPTIONAL
	PARAMTYPE 4	VAR nright	 		AS NUMERIC		OPTIONAL
	PARAMTYPE 5	VAR ntop	 		AS NUMERIC		OPTIONAL
	PARAMTYPE 6	VAR nbottom	 		AS NUMERIC		OPTIONAL
	If ValType(ctype)!="U" .and. !(ctype $ "path|linear")
		UserException("YExcel - Tipo invalido para efeito de preenchimento.(path|linear)")
	EndIf
	ogradientFill:SetAtributo("type",ctype)
	If ValType(ctype)!="U" .and. ctype=="path"
		Default nleft	:= 0.5
		Default nright	:= 0.5
		Default ntop	:= 0.5
		Default nbottom	:= 0.5
		If ValType(nleft)!="N" .OR. !(nleft>=0 .and. nleft<=1)
			UserException("YExcel - definir posição left em 0 a 1. Valor informado:"+cValToChar(nleft))
		EndIf
		If ValType(nright)!="N" .OR. !(nright>=0 .and. nright<=1)
			UserException("YExcel - definir posição right em 0 a 1. Valor informado:"+cValToChar(nright))
		EndIf
		If ValType(ntop)!="N" .OR. !(ntop>=0 .and. ntop<=1)
			UserException("YExcel - definir posição top em 0 a 1. Valor informado:"+cValToChar(ntop))
		EndIf
		If ValType(nbottom)!="N" .OR. !(nbottom>=0 .and. nbottom<=1)
			UserException("YExcel - definir posição bottom em 0 a 1. Valor informado:"+cValToChar(nbottom))
		EndIf
		ogradientFill:SetAtributo("left"	,nleft)
		ogradientFill:SetAtributo("right"	,nright)
		ogradientFill:SetAtributo("top"		,ntop)
		ogradientFill:SetAtributo("bottom"	,nbottom)
	Else
		Default nAngulo	:= 90
		ogradientFill:SetAtributo("degree",nAngulo)
	EndIf
	For nCont:=1 to Len(aCores)
		If !(aCores[nCont][2]>=0 .and. aCores[nCont][2]<=1)
			UserException("YExcel - Definição de cor varia de 0 a 1. Valor informado:"+cValToChar(aCores[nCont][2]))
		EndIf
		If Len(aCores[nCont][1])==6
			aCores[nCont][1]	:= "FF"+aCores[nCont][1]
		EndIf
		ostop	:= yExcelTag():New("stop",yExcelTag():New("color",,{{"rgb",aCores[nCont][1]}}),{{"position",aCores[nCont][2]}})
		ogradientFill:AddValor(ostop)
	Next
Return ogradientFill

/*/{Protheus.doc} Borda
Cria borda para ser usado no estilo das células
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param cTipo, characters, "C"-Cima|"B"-Baixo|"E"-Esquerda|"D"-Direita|T-TODAS("CBED") OU "T"-TOP|"B"-Bottom|"L"-Left|"R"-Right|A-ALL("TBLR") OU "DIAGONAL"-Diagonal
@param cCor, characters, Cor em Aplha+RGB da borda
@param cModelo, characters, Modelo da borda

@param oBorder, object, Retorna o objeto criado de borda
@param lAdd, logical, deve criar o objeto como estilo de célula. Padrão .T.

@type function
@Obs pode juntar os tipo. Exemplo "ED"-Esquerda e direita

/*/
METHOD Borda(cTipo,cCor,cModelo,oBorder,lAdd) CLASS YExcel
	Local nPos
	Local cLeft,cRight,cTop,cBottom,cDiagonal
	Default cModelo	:= "thin"
	Default cCor	:= "FF000000"
	Default lAdd	:= .T.			//Adiciona o objeto ao estilo principal
	Default cTipo	:= ""
	If "E" $ cTipo .or. "L" $ cTipo
		cLeft	:= cModelo
	EndIf
	If "D" $ cTipo .or. "R" $ cTipo
		cRight	:= cModelo
	EndIf
	If "T" $ cTipo .or. "C" $ cTipo
		cTop	:= cModelo
	EndIf
	If "B" $ cTipo
		cBottom	:= cModelo
	EndIf
	If "DIAGONAL" $ cTipo
		cDiagonal	:= cModelo
	EndIf

	If cTipo=="T" .or. cTipo=="ALL" .or. cTipo=="A"	//Todas bordas
		oBorder	:= Border(cModelo,cModelo,cModelo,cModelo,,cCor,cCor,cCor,cCor,)
	Else
		oBorder	:= Border(cLeft,cRight,cTop,cBottom,cDiagonal,cCor,cCor,cCor,cCor,cCor)
	EndIf
	If lAdd
		nPos	:= ::AddBorda(oBorder)
	EndIf
Return nPos
//NÃO DOCUMENTAR
METHOD AddBorda(oBorder) CLASS YExcel	//Pag 1769
	Local nPos,nCont
	Local oColor,oBorder

	//(pag 2446)val single,double,dotted(serinhada),triple,thick(grosso)
	//
	::oBorders:AddValor(oBorder)
	nPos := Len(::oBorders:GetValor())
	::oBorders:SetAtributo("count",nPos)
Retur nPos-1

Static Function Border(cleft,cright,ctop,cbottom,cdiagonal,cCorleft,cCorright,cCortop,cCorbottom,cCordiagonal)
	Local oBorder	:= yExcelTag():New("border",{})
	oStyle	:= nil
	oColor	:= nil
	If ValType(cleft)<>"U"
		oStyle	:= tHashMap():new()
		oStyle:Set("style",cleft)
		If ValType(cCorleft)<>"U"
			oColor	:= yExcelTag():New("color",nil,oColor)
			If ValType(cCorleft)=="C"
				oColor:SetAtributo("rgb",cCorleft)
			ElseIf ValType(cCorleft)=="N"
				oColor:SetAtributo("indexed",cCorleft)
			EndIf
		EndIf
	EndIf
	oBorder:AddValor(yExcelTag():New("left",oColor,oStyle))

	oStyle	:= nil
	oColor	:= nil
	If ValType(cright)<>"U"
		oStyle	:= tHashMap():new()
		oStyle:Set("style",cright)
		If ValType(cCorright)<>"U"
			oColor	:= yExcelTag():New("color",nil,oColor)
			If ValType(cCorright)=="C"
				oColor:SetAtributo("rgb",cCorright)
			ElseIf ValType(cCorright)=="N"
				oColor:SetAtributo("indexed",cCorright)
			EndIf
		EndIf
	EndIf
	oBorder:AddValor(yExcelTag():New("right",oColor,oStyle))

	oStyle	:= nil
	oColor	:= nil
	If ValType(ctop)<>"U"
		oStyle	:= tHashMap():new()
		oStyle:Set("style",ctop)
		If ValType(cCortop)<>"U"
			oColor	:= yExcelTag():New("color",nil,oColor)
			If ValType(cCortop)=="C"
				oColor:SetAtributo("rgb",cCortop)
			ElseIf ValType(cCortop)=="N"
				oColor:SetAtributo("indexed",cCortop)
			EndIf
		EndIf
	EndIf
	oBorder:AddValor(yExcelTag():New("top",oColor,oStyle))

	oStyle	:= nil
	oColor	:= nil
	If ValType(cbottom)<>"U"
		oStyle	:= tHashMap():new()
		oStyle:Set("style",cbottom)
		If ValType(cCorbottom)<>"U"
			oColor	:= yExcelTag():New("color",nil,oColor)
			If ValType(cCorbottom)=="C"
				oColor:SetAtributo("rgb",cCorbottom)
			ElseIf ValType(cCorbottom)=="N"
				oColor:SetAtributo("indexed",cCorbottom)
			EndIf
		EndIf
	EndIf
	oBorder:AddValor(yExcelTag():New("bottom",oColor,oStyle))

	oStyle	:= nil
	oColor	:= nil
	If ValType(cdiagonal)<>"U"
		oStyle	:= tHashMap():new()
		oStyle:Set("style",cdiagonal)
		If ValType(cCordiagonal)<>"U"
			oColor	:= yExcelTag():New("color",nil,oColor)
			If ValType(cCordiagonal)=="C"
				oColor:SetAtributo("rgb",cCordiagonal)
			ElseIf ValType(cCordiagonal)=="N"
				oColor:SetAtributo("indexed",cCordiagonal)
			EndIf
		EndIf
	EndIf
	oBorder:AddValor(yExcelTag():New("diagonal",oColor,oStyle))

Return oBorder

/*/{Protheus.doc} AddFmtNum
Formatação para numeros
@author Saulo Gomes Martins
@since 04/03/2018
@version 1.0
@return nNumFmtId, Numero do formato criado/alterado
@param nDecimal, numeric, quantidade de casas decimais
@param lMilhar, logical, usa separador de 1000(.)
@param cPrefixo, characters, Prefixo para incluir no numero (Exemplo "R$ ")
@param cSufixo, characters, Sufixo para incluir no numero (Exemplo " %")
@param cNegINI, characters, simbolo para incluir no inicio de numeros negativos
@param cNegFim, characters, simbolo para incluir no fim de numeros negativos
@param cValorZero, characters, conteudo para sustituir valores zeros
@param cCor, characters, Cor para numero positivo
@param cCorNeg, characters, Cor para numero negativo
@param nNumFmtId, numeric, numFmtId para alteração
@type function
@example
:AddFmtNum(3,.T.)							//1234	1.234,000		| -1234	-1.234,000
:AddFmtNum(2,.T.,"R$ "," ",,,"-")			//1234	R$ 1.234,00		| -1234	-R$ 1.234,00	| 0	-
:AddFmtNum(2,.T.,," %")						//1234	1.234,00 %		| -1234	-1.234,00 %
:AddFmtNum(2,.T.,,"(",")")					//1234	1.234,00		| -1234	(1.234,00)
:AddFmtNum(2,.T.,,"(",")",,"Green","Red")	//1234	1.234,00 Verde	| -1234	(1.234,00) Vermelho

/*/
Method AddFmtNum(nDecimal,lMilhar,cPrefixo,cSufixo,cNegINI,cNegFim,cValorZero,cCor,cCorNeg,nNumFmtId) class YExcel
	Local nPos,cformatCode
	Local cDecimal
	Local cNumero	:= ""
	Local cNegINIAli:= ""
	Local cNegFIMAli:= ""
	Local oFormat
	Local aCores	:= {"Black","Blue","Cyan","Green","Magenta","Red","White","Yellow"}
	PARAMTYPE 0	VAR nDecimal			AS NUMERIC					OPTIONAL DEFAULT 0
	PARAMTYPE 1	VAR lMilhar			  	AS LOGICAL					OPTIONAL DEFAULT .F.
	PARAMTYPE 2	VAR cPrefixo		  	AS CHARACTER				OPTIONAL DEFAULT ""
	PARAMTYPE 3	VAR cSufixo			  	AS CHARACTER				OPTIONAL DEFAULT ""
	PARAMTYPE 4	VAR cNegINI			  	AS CHARACTER				OPTIONAL DEFAULT "-"
	PARAMTYPE 5	VAR cNegFIM			  	AS CHARACTER				OPTIONAL DEFAULT ""
	PARAMTYPE 6	VAR cValorZero		  	AS CHARACTER				OPTIONAL DEFAULT ""
	PARAMTYPE 7	VAR cCor			  	AS CHARACTER,NUMERIC		OPTIONAL DEFAULT ""
	PARAMTYPE 8	VAR cCorNeg			  	AS CHARACTER,NUMERIC		OPTIONAL DEFAULT ""
	PARAMTYPE 9	VAR nNumFmtId		  	AS NUMERIC					OPTIONAL

	If !Empty(cCor)
		If ValType(cCor)=="C"
			nPosCor	:= aScan(aCores,{|x| UPPER(x)==UPPER(cCor) })
			If nPosCor==0
				UserException("YExcel - Cor da formatação invalida ("+cCor+")")
			Else
				cCor	:= aCores[nPosCor]
			EndIf
		ElseIf ValType(cCor)=="N"
			If !(cCor>=1 .AND. cCor<=56)
				UserException("YExcel - Cor da formatação invalida ("+cValToChar(cCor)+"), Cores indexado valido de 1-56.")
			EndIf
			cCor	:= "Color"+cValToChar(cCor)
		EndIf
	EndIf
	If !Empty(cCorNeg)
		If ValType(cCorNeg)=="C"
			nPosCor	:= aScan(aCores,{|x| UPPER(x)==UPPER(cCorNeg) })
			If nPosCor==0
				UserException("YExcel - Cor da formatação invalida ("+cCorNeg+")")
			Else
				cCorNeg	:= aCores[nPosCor]
			EndIf
		ElseIf ValType(cCorNeg)=="N"
			If !(cCorNeg>=1 .AND. cCorNeg<=56)
				UserException("YExcel - Cor da formatação invalida ("+cValToChar(cCorNeg)+"), Cores indexado valido de 1-56.")
			EndIf
			cCorNeg	:= "Color"+cValToChar(cCorNeg)
		EndIf
	EndIf

	cDecimal	:= Replicate("0",nDecimal)
	If lMilhar
		cNumero	:= "#,##0"
	Else
		cNumero	:= "#"
	EndIf

	If !Empty(cDecimal)
		cNumero	:= cNumero+"."+cDecimal
	EndIf
	If !Empty(cPrefixo)
		cPrefixo	:= "&quot;"+cPrefixo+"&quot;
		cNumero		:= cPrefixo+cNumero
	EndIf
	If !Empty(cSufixo)
		cSufixo		:= "&quot;"+cSufixo+"&quot;
		cNumero		:= cNumero+cSufixo
	EndIf
	If !Empty(cNegINI)
		cNegINIAli	:= "_"+cNegINI
	EndIf
	If !Empty(cNegFIM)
		cNegFIMAli	:= "_"+cNegFIM
	EndIf
	If !Empty(cValorZero)
		cValorZero	:= "&quot;"+cValorZero+"&quot;
	Else
		cValorZero	:= cNumero
	EndIf
	If !Empty(cCor)
		cCor	:= "["+cCor+"]"
	EndIf
	If !Empty(cCorNeg)
		cCorNeg	:= "["+cCorNeg+"]"
	EndIf
	cformatCode	:= cCor+cNegINIAli+cNumero+cNegFIMAli+";"+cCorNeg+cNegINI+cNumero+cNegFIM+";"+cNegINIAli+cValorZero+cNegFIMAli+";@"
	If !Empty(nNumFmtId)
		nPos	:= aScan(::onumFmts:GetValor(),{|x| x:GetAtributo("numFmtId")==nNumFmtId })
		If nPos==0
			oFormat	:= yExcelTag():New("numFmt")
			oFormat:SetAtributo("numFmtId",nNumFmtId)
			nPos	:= nil
		Else
			oFormat	:= ::onumFmts:GetValor(nPos)
		EndIf
	Else
		oFormat	:= yExcelTag():New("numFmt")
		::nNumFmtId++
		oFormat:SetAtributo("numFmtId",::nNumFmtId)
	EndIf

	oFormat:SetAtributo("formatCode",cformatCode)
	::onumFmts:AddValor(oFormat,nPos)
	::onumFmts:SetAtributo("count",Len(::onumFmts:GetValor()))
Return If(!Empty(nNumFmtId),nNumFmtId,::nNumFmtId)

/*/{Protheus.doc} AddStyles
Cria um estilo para ser usado nas células
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param numFmtId, numeric, numero com formato da célula. ver em observações
@param fontId,numeric, posição da fonte criado pelo metodo :AddFont()
@param fillId, numeric, posição do preenchimento criado pelo metodo :CorPreenc()
@param borderId, numeric, posição da borda criado pelo metodo :Borda()
@param xfId, numeric, posição dos estilos padrões. não usado(uso futuro)
@param aValores, array, outros valores(alinhamento criado pelo metodo :Alinhamento())
@param aOutrosAtributos, array, Outros atributos do estilo
@type function

@obs
0 General
1 0
2 0.00
3 #,##0
4 #,##0.00
9 0%
10 0.00%
11 0.00E+00
12 # ?/?
13 # ??/??
14 mm-dd-yy
15 d-mmm-yy
16 d-mmm
17 mmm-yy
18 h:mm AM/PM
19 h:mm:ss AM/PM
20 h:mm
21 h:mm:ss
22 m/d/yy h:mm
37 #,##0 ;(#,##0)
38 #,##0 ;[Red](#,##0)
39 #,##0.00;(#,##0.00)
40 #,##0.00;[Red](#,##0.00)
45 mm:ss
46 [h]:mm:ss
47 mmss.0
48 ##0.0E+0
49 @

166 $#,##0.00
44 - Contabil R$  #.##0,00
/*/
METHOD AddStyles(numFmtId,fontId,fillId,borderId,xfId,aValores,aOutrosAtributos) CLASS YExcel
	PARAMTYPE 0	VAR numFmtId			AS NUMERIC 		OPTIONAL
	PARAMTYPE 1	VAR fontId				AS NUMERIC 		OPTIONAL
	PARAMTYPE 2	VAR fillId				AS NUMERIC 		OPTIONAL
	PARAMTYPE 3	VAR borderId			AS NUMERIC 		OPTIONAL
	PARAMTYPE 4	VAR xfId  				AS NUMERIC 		OPTIONAL
	PARAMTYPE 5	VAR aValores  			AS ARRAY 		OPTIONAL
	PARAMTYPE 6	VAR aOutrosAtributos	AS ARRAY 		OPTIONAL
	If ValType(fontId)=="N" .AND. (fontId+1)>Len(::oFonts:GetValor())
		UserException("YExcel - Fonte informada("+cValToChar(fontId)+") não definido. Utilize o indice informado pelo metodo :AddFont()")
	ElseIf ValType(fillId)=="N" .AND. (fillId+1)>Len(::aCorPreenc)
		UserException("YExcel - Cor Preenchimento informado("+cValToChar(fillId)+") não definido. Utilize o indice informado pelo metodo :CorPreenc()")
	ElseIf ValType(borderId)=="N" .AND. (borderId+1)>Len(::oBorders:GetValor())
		UserException("YExcel - Borda informada("+cValToChar(borderId)+") não definido. Utilize o indice informado pelo metodo :Borda()")
	EndIF
Return ::oSyles:Add(numFmtId,fontId,fillId,borderId,xfId,aValores,aOutrosAtributos)

/*/{Protheus.doc} Alinhamento
Cria objeto de alinhamento da célula para ser usado na criação de estilo
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param cHorizontal, characters, Alinhamento Horizontal
@param cVertical, characters, Alinhamento Vertical
@param lReduzCaber, logical, Reduz texto para caber
@param lQuebraTexto, logical, Quebra texto
@param ntextRotation, numeric, Graus para rotação
@type function
@obs 	HORIZONTAL
	center
	centerContinuous
	distributed
	fill				preencher
	general
	justify
	left
	right

	VERTICAL
	bottom
	center
	distributed
	justify
	top
/*/
METHOD Alinhamento(cHorizontal,cVertical,lReduzCaber,lQuebraTexto,ntextRotation) CLASS YExcel
	Local oAlinhamento	:= yExcelTag():New("alignment",)
	Default cVertical	:= "general"
	Default cHorizontal	:= "bottom"
	Default lReduzCaber	:= .F.
	Default lQuebraTexto	:= .F.
	oAlinhamento:SetAtributo("horizontal",cHorizontal)
	oAlinhamento:SetAtributo("vertical",cVertical)
	If ValType(ntextRotation)=="N" .and. ntextRotation>0
		oAlinhamento:SetAtributo("textRotation",ntextRotation)
	EndIf
	If lReduzCaber .and. !lQuebraTexto
		oAlinhamento:SetAtributo("shrinkToFit","1")	//Um valor booleano que indica se o texto exibido na célula deve ser encolhido para se ajustar à célula
	EndiF
	If lQuebraTexto
		oAlinhamento:SetAtributo("wrapText","1")	//Um valor booleano indicando se o texto em uma célula deve ser envolvido na linha dentro da célula.
	EndiF
Return oAlinhamento

/*/{Protheus.doc} AddPane
Congelar Painéis
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param nySplit, numeric, Quantidade de linhas congeladas
@param nxSplit, numeric, Quantidade de colunas congeladas
@type function
/*/
METHOD AddPane(nySplit,nxSplit) CLASS YExcel
	Local nPos
	Default nySplit	:= 0
	Default nxSplit	:= 0
	If nxSplit>0 .and. nySplit>0
		nPos	:= ::Pane("bottomRight","frozen",::Ref(nySplit+1,nxSplit+1),nySplit,nxSplit)
	ElseIf nxSplit==0 .and. nySplit>0
		nPos	:= ::Pane("bottomLeft","frozen",::Ref(nySplit+1,nxSplit+1),nySplit,)
	ElseIf nxSplit>0 .and. nySplit==0
		nPos	:= ::Pane("topRight","frozen",::Ref(nySplit+1,nxSplit+1),,nxSplit)
	Endif
Return nPos
//NÃO DOCUMENTAR
METHOD Pane(cActivePane,cState,cRef,nySplit,nxSplit) CLASS YExcel
	Local osheetView
	Local nPos
	Default cActivePane	:= "bottomLeft"
	osheetView	:= ::osheetViews:GetValor()
	osheetView:SetValor({})	//Limpa o osheetView
	osheetView:AddValor(yExcelTag():New("pane",))
	nPos	:= Len(osheetView:GetValor())
	/*
	bottomLeft	- Painel inferior esquerdo, quando ambos verticais e horizontais são aplicadas. Esse valor também é usado quando apenas uma divisão horizontal foi aplicada, dividindo o painel em superior e inferior. Nesse caso, esse valor especifica painel inferior
	bottomRight - Painel inferior direito, quando as divisões verticais e horizontais são aplicadas.
	topLeft		- Painel superior esquerdo, quando as divisões verticais e horizontais são aplicadas.
	topRight	- Painel superior direito, quando as divisões verticais e horizontais são aplicadas
	*/
	osheetView:GetValor(nPos):SetAtributo("activePane",cActivePane)
	/*
	frozen		- Panes são congelados, mas não foram divididos sendo congelados. Nesse estado, quando os painéis são desbloqueados novamente, um único painel resulta, sem divisão. Nesse estado, as barras de divisão não são ajustáveis.
	frozenSplit	- Os painéis são congelados e foram divididos antes de serem congelados. Neste estado, quando os painéis são desbloqueados novamente, a divisão permanece, mas é ajustável.
	split		- Os painéis são divididos, mas não congelados. Nesse estado, as barras de divisão são ajustáveis pelo usuário.
	*/
	osheetView:GetValor(nPos):SetAtributo("state",cState)
	//Localização da célula visível superior esquerda no painel inferior direito (quando no modo Esquerdo para Direito).
	osheetView:GetValor(nPos):SetAtributo("topLeftCell",cRef)
	//Posição horizontal da divisão, em 1/20º de um ponto; 0 (zero) se nenhum. Se o painel estiver congelado, este valor indica o número de colunas visíveis no painel superior
	osheetView:GetValor(nPos):SetAtributo("xSplit",nxSplit)
	//Posição vertical da divisão, em 1/20º de um ponto; 0 (zero) se nenhum. Se o painel estiver congelado, este valor indica o número de linhas visíveis no painel esquerdo.
	osheetView:GetValor(nPos):SetAtributo("ySplit",nySplit)

	osheetView:AddValor(yExcelTag():New("selection",,{{"pane",cActivePane}}))
	aSort(osheetView:xValor,,,{|x,y| If(x:getnome()=="pane",1,2)<If(y:getnome()=="pane",1,2) })
Return nPos

//NÃO DOCUMENTAR
METHOD GravaRow(nLinha) CLASS YExcel
	Local oExcelRow
	If ::osheetData:GetValor():Get(nLinha,@oExcelRow)
		If ::lRowDef
			FWRITE(::nFileTmpRow,"</row>")
		Else
			GravaFile(@::nFileTmpRow,oExcelRow:GetTag())
		EndIf
		oExcelRow:GetValor():Clean()				//Limpa o thashmap
		oExcelRow	:= nil
		::osheetData:GetValor():Del(nLinha)	//Exclui a chave
	EndIf
Return


/*/{Protheus.doc} Ref
Retorna a referencia do excel de acordo com posição da linha e coluna em formato numerico
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param nLinha, numeric, Linha
@param nColuna, numeric, Coluna
@Return cRef, Referencia da linha e coluna.
@type function
@obs
	oExcel:Ref(1,2)	//Retorno B1
	oExcel:Ref(3,3)	//Retorno C3
/*/
METHOD Ref(nLinha,nColuna,llinha,lColuna) CLASS YExcel
	Local cLinha	:= ""
	Local cColuna	:= ""
	Default llinha	:= .F.
	Default lColuna	:= .F.
	If llinha
		cLinha	:= "$"
	EndIf
	If lColuna
		cColuna	:= "$"
	EndIf
Return cColuna+NumToString(nColuna)+cLinha+cValToChar(nLinha)


/*/{Protheus.doc} LocRef
Retorna linha e coluna de acordo com informação da referencia
@author Saulo Gomes Martins
@since 03/03/2018
@version 1.0
@return aLinhaCol, Array com duas dimenções 1=Linha|2=Coluna
@param cRef, characters, Refencia da celula (exemplo A1)
@type function

@example
LocRef("A1")	//Retorno {1,1}
LocRef("C5")	//Retorno {5,3}
/*/
METHOD LocRef(cRef) CLASS YExcel
	Local nCont
	Local nTam	:= Len(cRef)
	Local cColuna	:= ""
	Local cLinha	:= ""
	For nCont:=1 to nTam
		If IsAlpha(SubStr(cRef,nCont,1))
			cColuna	+= SubStr(cRef,nCont,1)
		ElseIf IsDigit(SubStr(cRef,nCont,1))
			cLinha	+= SubStr(cRef,nCont,1)
		EndIf
	Next
Return {Val(cLinha),::StringToNum(cColuna)}


/*/{Protheus.doc} NumToString
Retorna a letra da coluna de acordo com a posição numerica
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param nNum, numeric, Numero da coluna
@type function
/*/
METHOD NumToString(nNum) Class YExcel
Return NumToString(nNum)

/*/{Protheus.doc} StringToNum
Retorna a posição da coluna de acordo com a letra da coluna
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param cString, characters, Letra da Coluna
@type function
/*/
METHOD StringToNum(cString) Class YExcel
Return StringToNum(cString)

/*/{Protheus.doc} AddTabela
Adiciona tabela com formatação
@author Saulo Gomes Martins
@since 08/05/2017
@param cNome, characters, Nome da tabel
@param nLinha, numeric, Linha inicial da tabela
@param nColuna, numeric, Coluna inicial da tabela
@type function
/*/
METHOD AddTabela(cNome,nLinha,nColuna) CLASS YExcel
	Local nCont,nPos
	Local oTable
//	Local nQtdPlan	:= Len(::aPlanilhas)
	PARAMTYPE 0	VAR cNome  AS CHARACTER 		OPTIONAL DEFAULT CriaTrab(,.F.)
	PARAMTYPE 1	VAR nLinha  AS NUMERIC 			OPTIONAL DEFAULT ::adimension[2][1]
	PARAMTYPE 2	VAR nColuna  AS NUMERIC
	::nQtdTables++
	nPos	:= ::nQtdTables
	::otableParts:AddValor(yExcelTag():New("tablePart",nil,{{"r:id","rId"+cValToChar(nPos)}}))
	::otableParts:SetAtributo("count",Len(::atable)+1)

	oTable	:= yExcel_Table():New(self,nLinha,nColuna,cNome) //yExcelTag():New("table",{},)
	oTable:SetAtributo("xmlns","http://schemas.openxmlformats.org/spreadsheetml/2006/main")
	oTable:SetAtributo("id",nPos)
	oTable:SetAtributo("name",cNome)
	oTable:SetAtributo("displayName",cNome)

	oTable:AddValor(yExcelTag():New("autoFilter",{}))

	oTable:oTableColumns	:= yExcelTag():New("tableColumns",{},{{"count",0}})	//Pag 1743
	oTable:AddValor(oTable:oTableColumns)

	oTable:otableStyleInfo	:= yExcelTag():New("tableStyleInfo",nil,)
	oTable:otableStyleInfo:SetAtributo("name","TableStyleMedium2")
	oTable:otableStyleInfo:SetAtributo("showFirstColumn",0)
	oTable:otableStyleInfo:SetAtributo("showLastColumn",0)
	oTable:otableStyleInfo:SetAtributo("showRowStripes",0)
	oTable:otableStyleInfo:SetAtributo("showColumnStripes",0)
	oTable:AddValor(oTable:otableStyleInfo)
	AADD(::atable,oTable)
Return oTable

/*/{Protheus.doc} Gravar
Grava o excel processado
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param cLocal, characters, Local para gerar o arquivo no client
@param lAbrir, logical, Abrir a planilha gerada
@param lDelSrv, logical, Deleta a planilha após copiar para o client
@return cArquivo, local do arquivo gerado
@type function
/*/
Method Gravar(cLocal,lAbrir,lDelSrv) Class YExcel
	Local nFile
	Local nCont,nQtdPlanilhas
	Local cArquivo	:= ""
	Local lServidor	:= !Empty(cLocal) .and. SubStr(cLocal,1,1)=="\"
	If !Empty(cLocal)
		cLocal	:= Alltrim(cLocal)
		If Right(cLocal,1)=="\"
			cLocal	:= SubStr(cLocal,1,Len(cLocal)-1)
		EndIf
	EndIf
	If ValType(cRootPath)=="U"
		cRootPath	:= GetSrvProfString( "RootPath", "" )
	EndIf
	Default lAbrir	:= .F.
	Default lDelSrv	:= .T.
	If Empty(::cNomeFile)
		Return
	EndIf
	Private oSelf			:= Self
	//Grava a ultima linha
	::GravaRow(::adimension[1][1])

	If Empty(::oCols:GetValor())
		::AddTamCol(::adimension[2][2],::adimension[1][2],12.00)
	EndIf

	::CriarFile("\"+::cNomeFile						,"[Content_Types].xml"	,h_xls_Content_Types()	,)
	::CriarFile("\"+::cNomeFile+"\_rels"			,".rels"				,h_xls_rels()			,)
	::CriarFile("\"+::cNomeFile+"\docProps"			,"app.xml"				,h_xls_app()			,)
	::CriarFile("\"+::cNomeFile+"\docProps"			,"core.xml"				,h_xls_core()			,)

	::CriarFile("\"+::cNomeFile+"\xl"				,"sharedStrings.xml"	,h_xls_sharedStrings()	,)
	::CriarFile("\"+::cNomeFile+"\xl"				,"styles.xml"			,h_xls_styles()			,)
	::CriarFile("\"+::cNomeFile+"\xl"				,"workbook.xml"			,h_xls_workbook()		,)
	::CriarFile("\"+::cNomeFile+"\xl\_rels"			,"workbook.xml.rels"	,h_xls_rworkbook()		,)
	::CriarFile("\"+::cNomeFile+"\xl\theme"			,"theme1.xml"			,h_xls_theme()			,)

	nQtdPlanilhas	:= Len(::aPlanilhas)
	::CriarFile("\"+::cNomeFile+"\xl\worksheets"	,"sheet"+cValToChar(nQtdPlanilhas)+".xml"			,""			,)
	GravaFile(@nFile,"","\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\worksheets","sheet"+cValToChar(nQtdPlanilhas)+".xml")
	h_xls_sheet(nFile)
	fClose(nFile)
	fErase("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\worksheets\tmprow.xml")

	If !Empty(::atable)
		::CriarFile("\"+::cNomeFile+"\xl\worksheets\_rels\"	,"sheet"+cValToChar(nQtdPlanilhas)+".xml.rels"		,h_xlsrelssheet()		,)
		AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\worksheets\_rels\sheet"+cValToChar(nQtdPlanilhas)+".xml.rels")
		For nCont:=1 to Len(::atable)
			::nCont	:= nCont
			::CriarFile("\"+::cNomeFile+"\xl\tables\"	,"table"+cValToChar(::nQtdTables-Len(::atable)+nCont)+".xml"		,h_xls_table()		,)
			AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\tables\table"+cValToChar(::nQtdTables-Len(::atable)+nCont)+".xml")
		Next
	EndIf

	If lServidor
		cArquivo	:= cLocal+'\'+::cNomeFile+'.xlsx'
		cLocal		:= ""
	Else
		cArquivo	:= '\tmpxls\'+::cTmpFile+'\'+::cNomeFile+'.xlsx'
	EndIf
	If !FindFunction("FZIP")
		WaitRunSrv('"'+cAr7Zip+'" a -tzip "'+cRootPath+cArquivo+'" "'+cRootPath+'\tmpxls\'+::cTmpFile+'\'+::cNomeFile+'\*"',.T.,"C:\")
	Else
		fZip(cArquivo,::aFiles,"\tmpxls\"+::cTmpFile+'\'+::cNomeFile+'\')
	EndIf

	For nCont:=1 to Len(::aFiles)
		If fErase(::aFiles[nCont])<>0
			ConOut(::aFiles[nCont])
			ConOut("Ferror:"+cValToChar(ferror()))
		EndIf
	Next

	DelPasta("\tmpxls\"+::cTmpFile+"\"+::cNomeFile)	//Apaga arquivos temporarios
	If substr(cArquivo,1,8)<>"\tmpxls\"
		DelPasta("\tmpxls\"+::cTmpFile)
	EndIf
	If !Empty(cLocal)
		If GetRemoteType() == REMOTE_HTML
			CpyS2TW(cArquivo, .T.)
		Else
			CpyS2T( cArquivo,cLocal)
			cArquivo	:= cLocal+'\'+::cNomeFile+'.xlsx'
			If lAbrir
				ShellExecute("open",cLocal+'\'+::cNomeFile+'.xlsx',"",cLocal+'\', 1 )
			EndIf
		EndIf
		If lDelSrv
			DelPasta("\tmpxls\"+::cTmpFile)	//Apaga o arquivo do servidor
		EndIf
	EndIf
Return cArquivo

/*/{Protheus.doc} DelPasta
Deleta uma pasta e qualquer arquivo ou pasta que esteja dentro dela
@author Saulo Gomes Martins
@since 02/05/2017
@version p11
@param cCaminho, characters, descricao
@type function
/*/
Static Function DelPasta(cCaminho)
	Local nCont
	Local aFiles	:= Directory(cCaminho+"\*","HSD")
	For nCont:=1 to Len(aFiles)
		If aFiles[nCont][1]=="." .or. aFiles[nCont][1]==".."
			Loop
		EndIf
		If aFiles[nCont][5] $ "D"
			DelPasta(cCaminho+"\"+aFiles[nCont][1])
		Else
//			ConOut("Deletando:"+cCaminho+"\"+aFiles[nCont][1])
			If fErase(cCaminho+"\"+aFiles[nCont][1])<>0
				ConOut(cCaminho+"\"+aFiles[nCont][1])
				ConOut("Ferror:"+cValToChar(ferror()))
			EndIf
		EndIf
	Next
//	ConOut("Apagando pasta:"+cCaminho)
	If !DirRemove(cCaminho)
		ConOut(cCaminho)
		ConOut("Ferror:"+cValToChar(ferror()))
	EndIf
Return
//NÃO DOCUMENTAR
METHOD CriarFile(cLocal,cNome,cString) Class YExcel
	Local cDirServ	:= "\tmpxls\"+::cTmpFile
	Local lOk			:= .T.
	Local nFile
	If ValType(cString)!="C"
		return lOk
	EndIf
	FWMakeDir(cDirServ+cLocal,.F.)
	oFile	:= FWFileIOBase():New(cDirServ+cLocal+"\"+cNome)
	If !oFile:Exists()
		oFile:Create()
	Else
		fErase(cDirServ+cLocal+"\"+cNome)
		oFile:Create()
	EndIf
	oFile:Close()
	nFile	:= FOPEN(cDirServ+cLocal+"\"+cNome, FO_READWRITE)
	cString	:= EncodeUTF8(cString)
	IF FWrite(nFile, cString, Len(cString)) < Len(cString)
	 	lOk	:= .F.
	EndIf
	fClose(nFile)
Return lOk
//NÃO DOCUMENTAR, USADO NA GRAVAÇÃO DO SHEET
METHOD GravaFile(nFile,cString,cLocal,cArquivo) Class YExcel
Return GravaFile(nFile,cString,cLocal,cArquivo)

Static Function GravaFile(nFile,cString,cLocal,cArquivo)
	Local lOk			:= .T.
	If ValType(cString)=="C"
	EndIf
	If !Empty(cArquivo)
		nFile	:= FOPEN(cLocal+"\"+cArquivo, FO_READWRITE)
	EndIf
	cString	:= EncodeUTF8(cString)
	FSeek(nFile, 0, FS_END)
		IF FWrite(nFile, cString, Len(cString)) < Len(cString)
		 	lOk	:= .F.
		EndIf
Return lOk

//EM DESENVOLVIMENTO
Method AddAgrCol(nMin,nMax,outlineLevel,collapsed) Class YExcel
	Local nPos
	Default outlineLevel	:= 1
	Default collapsed		:= 1
	::oCols:AddValor(yExcelTag():New("col"))
	nPos	:= Len(::oCols:GetValor())
	::oCols:GetValor(nPos):SetAtributo("min",nMin)
	::oCols:GetValor(nPos):SetAtributo("max",nMax)
	::oCols:GetValor(nPos):SetAtributo("outlineLevel",nWidth)
//	::oCols:GetValor(nPos):SetAtributo("collapsed",bestFit)
Return

/*/{Protheus.doc} AddTamCol
Defini o tamanho de uma coluna ou varias colunas
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param nMin, numeric, Coluna inicial
@param nMax, numeric, Coluna final
@param nWidth, numeric, descricao
@param lbestFit, logico, melhor ajuste numerico
@param lcustomWidth, logico, tamanho customizado
@type function
/*/
Method AddTamCol(nMin,nMax,nWidth,lbestFit,lcustomWidth) Class YExcel
	Local nPos
	PARAMTYPE 0	VAR nMin			AS NUMERIC
	PARAMTYPE 1	VAR nMax			AS NUMERIC		OPTIONAL DEFAULT nMin
	PARAMTYPE 2	VAR nWidth	  		AS NUMERIC		OPTIONAL DEFAULT "Calibri"
	PARAMTYPE 3	VAR lbestFit	  	AS LOGICAL		OPTIONAL DEFAULT .T.
	PARAMTYPE 4	VAR lcustomWidth	AS LOGICAL		OPTIONAL DEFAULT .T.

	::oCols:AddValor(yExcelTag():New("col"))
	nPos	:= Len(::oCols:GetValor())
	::oCols:GetValor(nPos):SetAtributo("min",nMin)
	::oCols:GetValor(nPos):SetAtributo("max",nMax)
	::oCols:GetValor(nPos):SetAtributo("width",nWidth)
	If lbestFit
		::oCols:GetValor(nPos):SetAtributo("bestFit","1")
	EndIf
	If lcustomWidth
		::oCols:GetValor(nPos):SetAtributo("customWidth","1")
	EndIf
Return

//----------------------------------------------------------------------
//CLASSE DAS CÉLULAS
//----------------------------------------------------------------------
Class yExcelc From yExcelTag
	Data nLinha
	Data nColuna
	Data oyExcel
	Method New() constructor
	Method SetVal()
	Method SetFormula()
	Method SetV()
EndClass

Method New(oyExcel,nLinha,nColuna) Class yExcelc
	_Super:New("c",tHashMap():New())
	::oyExcel	:= oyExcel
	::nLinha	:= nLinha
	::nColuna	:= nColuna
	::SetAtributo("r", NumToString(nColuna)+cValToChar(nLinha))
Return self

Method SetFormula(f) Class yExcelc
	::GetValor():Set("f",yExcelTag():New("f",f))
Return

Method SetV(v) Class yExcelc
	::GetValor():Set("v",yExcelTag():New("v",v))
Return

Method SetVal(v,f,nStyle) Class yExcelc
	Local cTipo	:= ValType(v)
	Local nPos
	Default ::oyExcel:oString	:= tHashMap():New()
	If !Empty(f)
		::SetFormula(f)
	EndIf
	If cTipo=="C"
		::SetAtributo("t","s")
		If ::oyExcel:oString:Get(v,@nPos)//nPos>0
			::SetV(nPos)
		Else
			::oyExcel:oString:Set(v,::oyExcel:nQtdString)
			::SetV(::oyExcel:nQtdString)
			::oyExcel:nQtdString++
		EndIf
	ElseIf cTipo=="L"
		::SetAtributo("t","b")
		::SetV(if(v,1,0))
	ElseIf cTipo=="N"
		::SetV(v)
	ElseIf cTipo=="D"
		::SetAtributo("s","1")	//Adiciona o estilo padrão de data
		//::SetAtributo("t","d")	//Adiciona o estilo padrão de data
		//::SetV(SUBSTR(DTOS(v),1,4)+"-"+SUBSTR(DTOS(v),5,2)+"-"+SUBSTR(DTOS(v),7,2))
		::SetV(v-STOD("19000101")+2)
	Else
		::SetV(v)
	EndIf
	If ValType(nStyle)=="N"
		If nStyle+1>Len(::oyExcel:oSyles:GetValor())
			UserException("YExcel - Estilo informado("+cValToChar(nStyle)+") não definido. Utilize o indice informado pelo metodo AddStyles")
		Else
			::SetAtributo("s",nStyle)
		EndIf
	EndIf
Return self

Class YExcelFont From YExcelTag
	Method New()	constructor
	Method Add()
EndClass

Method New() Class YExcelFont
	Local nTamFonts,nTamFont
	_Super:New("fonts",{})
Return

Method Add(nTamanho,cCorRGB,cNome,cfamily,cScheme,lNegrito,lItalico,lSublinhado,lTachado) Class YExcelFont
	PARAMTYPE 0	VAR nTamanho		AS NUMERIC				OPTIONAL DEFAULT 11
	PARAMTYPE 1	VAR cCorRGB			AS CHARACTER,NUMERIC	OPTIONAL DEFAULT "FF000000"
	PARAMTYPE 2	VAR cNome	  		AS CHARACTER			OPTIONAL DEFAULT "Calibri"
	PARAMTYPE 3	VAR cfamily	  		AS CHARACTER			OPTIONAL DEFAULT "2"
	PARAMTYPE 4	VAR cScheme	  		AS CHARACTER			OPTIONAL DEFAULT "minor"
	PARAMTYPE 5	VAR lNegrito	  	AS LOGICAL				OPTIONAL DEFAULT .F.
	PARAMTYPE 6	VAR lItalico	  	AS LOGICAL				OPTIONAL DEFAULT .F.
	PARAMTYPE 7	VAR lSublinhado	  	AS LOGICAL				OPTIONAL DEFAULT .F.
	PARAMTYPE 8	VAR lTachado	  	AS LOGICAL				OPTIONAL DEFAULT .F.

	If ValType(cCorRGB)=="C" .and. Len(cCorRGB)==6
		cCorRGB	:= "FF"+cCorRGB
	EndIf
	::AddValor(yExcelTag():New("font",{}))
	nTamFonts	:= Len(::GetValor())
	::SetAtributo("count",nTamFonts)
	::SetAtributo("x14ac:knownFonts",1)

	If lNegrito
		AADD(::xValor[nTamFonts]:xValor,yExcelTag():New("b"))
	EndIf
	If lItalico
		AADD(::xValor[nTamFonts]:xValor,yExcelTag():New("i"))
	EndIf
	If lTachado
		AADD(::xValor[nTamFonts]:xValor,yExcelTag():New("strike"))
	EndIf
	If lSublinhado
		AADD(::xValor[nTamFonts]:xValor,yExcelTag():New("u"))
	EndIf


	AADD(::xValor[nTamFonts]:xValor,yExcelTag():New("sz"))
	nTamFont	:= Len(::xValor[nTamFonts]:xValor)
	::xValor[nTamFonts]:xValor[nTamFont]:SetAtributo("val",nTamanho)

	AADD(::xValor[nTamFonts]:xValor,yExcelTag():New("color"))
	nTamFont	:= Len(::xValor[nTamFonts]:xValor)
	If ValType(cCorRGB)=="N"
		::xValor[nTamFonts]:xValor[nTamFont]:SetAtributo("indexed",cCorRGB)
	Else
		If ValType(cCorRGB)=="C" .and. Len(cCorRGB)==6
			cCorRGB	:= "FF"+cCorRGB
		EndIf
		::xValor[nTamFonts]:xValor[nTamFont]:SetAtributo("rgb",cCorRGB)
	EndIf

	AADD(::xValor[nTamFonts]:xValor,yExcelTag():New("name"))
	nTamFont	:= Len(::xValor[nTamFonts]:xValor)
	::xValor[nTamFonts]:xValor[nTamFont]:SetAtributo("val",cNome)

	AADD(::xValor[nTamFonts]:xValor,yExcelTag():New("family"))
	nTamFont	:= Len(::xValor[nTamFonts]:xValor)
	::xValor[nTamFonts]:xValor[nTamFont]:SetAtributo("val",cfamily)
	/* pag 2525
	0 Not applicable.
	1 Roman
	2 Swiss
	3 Modern
	4 Script
	5 Decorative
	*/
	AADD(::xValor[nTamFonts]:xValor,yExcelTag():New("scheme"))
	nTamFont	:= Len(::xValor[nTamFonts]:xValor)
	::xValor[nTamFonts]:xValor[nTamFont]:SetAtributo("val",cScheme)
return nTamFonts-1

//CLASSE DAS LINHAS
Class yExcelRow From yExcelTag
	Data	nLinha
	Data	aspans
	Data	oyExcel
	Method New() constructor
	Method SetVal()
EndClass

Method New(oyExcel,r,spans,x14ac_dyDescent) Class yExcelRow
	Default spans	:= {1,1}
	Default x14ac_dyDescent	:= 0.25
	Default oyExcel:lRowcollapsed	:= .F.
	Default oyExcel:lRowhidden		:= .F.
	_Super:New("row",tHashMap():New())
	::oyExcel		:= oyExcel
	::nLinha		:= r
	::SetAtributo("r",r)
	::SetAtributo("spans",cValToChar(spans[1])+":"+cValToChar(spans[2]))
	::aspans	:= spans
	::SetAtributo("x14ac:dyDescent",x14ac_dyDescent)
	If ValType(::oyExcel:nRowoutlineLevel)=="N"
		::SetAtributo("outlineLevel",::oyExcel:nRowoutlineLevel)
	EndIf
	If ::oyExcel:lRowcollapsed
		::SetAtributo("collapsed",1)
	EndIf
	If ::oyExcel:lRowhidden
		::SetAtributo("hidden",1)
	EndIf
Return self

Method SetVal(nColuna,xValor,cFormula,nStyle) Class yExcelRow
	Local oExcelC
	If !::GetValor():Get(nColuna,@oExcelC)
	EndIf
	oExcelC	:= yExcelC():New(::oyExcel,::nLinha,nColuna)
	::GetValor():Set(nColuna,oExcelC)
	oExcelC:SetVal(xValor,cFormula,nStyle)		//Passa alteração para a celula
	If nColuna>::aspans[2]
		::aspans[2]	:= nColuna
		::SetAtributo("spans","1:"+cValToChar(nColuna))
	EndIF
	If ValType(::oyExcel:nTamLinha)<>"U"
		::SetAtributo("customHeight",1)
		::SetAtributo("ht",::oyExcel:nTamLinha)
	Endif
Return

Class yExcel_CorPreenc From yExcelTag
	Method New() constructor
EndClass

/*
@see http://www.datypic.com/sc/ooxml/a-patternType-1.html
cType
	none
	solid
	mediumGray
	darkGray
	lightGray
	darkHorizontal
	darkVertical
	darkDown
	darkUp
	darkGrid
	darkTrellis
	lightHorizontal
	lightVertical
	lightDown
	lightUp
	lightGrid
	lightTrellis
	gray125
	gray0625
*/
Method New(cType,cFgCor,cBgCor) Class yExcel_CorPreenc
	Local ofgColor,obgColor,oPatternFill
	Local aValores	:= {}
	Default cType	:= "solid"
	ofgColor	:= yExcelTag():New("fgColor")
	If ValType(cFgCor)=="C"
		If Len(cFgCor)==6
			cFgCor	:= "FF"+cFgCor
		EndIf
		ofgColor:SetAtributo("rgb",cFgCor)
	Elseif ValType(cFgCor)=="N"
		ofgColor:SetAtributo("indexed",cFgCor)	//indexed="65" System Background n/a 	pag:1775
	Else
		ofgColor:SetAtributo("indexed",65)	//indexed="65" System Background n/a 	pag:1775
	EndiF
	obgColor	:= yExcelTag():New("bgColor")
	If ValType(cBgCor)=="C"
		If Len(cBgCor)==6
			cBgCor	:= "FF"+cBgCor
		EndIf
		obgColor:SetAtributo("rgb",cBgCor)
	Elseif ValType(cBgCor)=="N"
		obgColor:SetAtributo("indexed",cFgCor)
	Else
		obgColor:SetAtributo("indexed",64)	//indexed="64" System Foreground n/a
	EndIf
	If cType == "solid"
		AADD(aValores,ofgColor)
		AADD(aValores,obgColor)
	EndIf
	oPatternFill	:= yExcelTag():New("patternFill",aValores)
	oPatternFill:SetAtributo("patternType",cType)
	_Super:New("fill",oPatternFill)
Return


//----------------------------------------------------------
Class yExcel_cellXfs From yExcelTag
	Method New() constructor
	Method Add()
EndClass

Method New() Class yExcel_cellXfs
	_Super:New("cellXfs",{})
Return self

Method Add(numFmtId,fontId,fillId,borderId,xfId,aValores,aOutrosAtributos) Class yExcel_cellXfs
	Local nPos,nCont
	Local oAtrr	:= tHashMap():new()
	Default aOutrosAtributos		:= {}
	Default aValores				:= {}
	Default xfId	:= 0

	If ValType(numFmtId)<>"U"
		oAtrr:Set("numFmtId",numFmtId)
		oAtrr:Set("applyNumberFormat","1")
	Else
		oAtrr:Set("numFmtId",0)
	EndIf

	If ValType(fontId)<>"U"
		oAtrr:Set("fontId",fontId)
		oAtrr:Set("applyFont","1")
	Else
		oAtrr:Set("fontId",0)
	EndIf
	If ValType(fillId)<>"U"
		oAtrr:Set("fillId",fillId)
		oAtrr:Set("applyFill","1")
	Endif
	If ValType(borderId)<>"U"
		oAtrr:Set("borderId",borderId)
		oAtrr:Set("applyBorder","1")
	Else
		oAtrr:Set("borderId",0)
	EndIf

	oAtrr:Set("xfId",xfId)

	If aScan(aValores,{|x| x:GetNome()=="alignment"})>0
		oAtrr:Set("applyAlignment","1")
	EndIf

	For nCont:=1 to Len(aOutrosAtributos)
		oAtrr:Set(aOutrosAtributos[nCont][1],aOutrosAtributos[nCont][2])
	Next
	::AddValor(yExcelTag():New("xf",aValores,oAtrr))
	nPos	:= Len(::GetValor())
	::SetAtributo("count", nPos)
return nPos-1

//----------------------------------------------------------
Class yExcelsheetData From yExcelTag
	Data oyExcel
	Method New() constructor
	Method SetVal()
	Method Add()
EndClass

Method New(oyExcel) Class yExcelsheetData
	_Super:New("sheetData",tHashMap():New())
	::oyExcel	:= oyExcel
Return self

Method Add(nLinha,aSpan) Class yExcelsheetData
	Local oExcelRow
	If !::GetValor():Get(nLinha,@oExcelRow)
		oExcelRow	:= yExcelRow():New(::oyExcel,nLinha,::oyExcel:aSpanRow)
	EndIf
Return oExcelRow

Method SetVal(nLinha,nColuna,xValor,cFormula,nStyle) Class yExcelsheetData
	Local oExcelRow
	If !::GetValor():Get(nLinha,@oExcelRow)
		oExcelRow	:= yExcelRow():New(::oyExcel,nLinha,::oyExcel:aSpanRow)
		::GetValor():Set(nLinha,oExcelRow)
	EndIf
	oExcelRow:SetVal(nColuna,xValor,cFormula,nStyle)	//Passa a alteração para a linha
Return
//-----------------------------------------------------------
//ALGORITIMO PARA CONVERTE COLUNAS DA PLANILHA
Static Function NumToString(nNum)
	Local cRet	:= ""
	If nNum<=26
		cRet	:= ColunasIndex(nNum)
	ElseIf nNum<=702
		IF nNum % 26==0
			cRet	+= ColunasIndex(((nNum-(nNum % 26))/26)-1)
		Else
			cRet	+= ColunasIndex((nNum-(nNum % 26))/26)
		EndIf
		cRet	+= ColunasIndex(nNum % 26)
	Else
		IF nNum % 26==0
			cRet	+= NumToString(((nNum-(nNum % 26))/26)-1)
		Else
			cRet	+= NumToString((nNum-(nNum % 26))/26)
		EndIf
		cRet	+= ColunasIndex(nNum % 26)
	EndIf
Return cRet

Static Function StringToNum(cString)
	Local nTam	:= Len(cString)
	Local nRet
	If nTam==1
		nRet	:= ColunasIndex(cString,2)
	ElseIf nTam==2
		nRet	:= (ColunasIndex(SubStr(cString,1,1),2)*26)+ColunasIndex(SubStr(cString,2,1),2)
	ElseIf nTam==3
		nRet	:= (ColunasIndex(SubStr(cString,1,1),2)*676)+(ColunasIndex(SubStr(cString,2,1),2)*26)+ColunasIndex(SubStr(cString,3,1),2)
	EndIf
Return nRet

Static aColIdx	:= {{1,"A"},;
					{2,"B"},;
					{3,"C"},;
					{4,"D"},;
					{5,"E"},;
					{6,"F"},;
					{7,"G"},;
					{8,"H"},;
					{9,"I"},;
					{10,"J"},;
					{11,"K"},;
					{12,"L"},;
					{13,"M"},;
					{14,"N"},;
					{15,"O"},;
					{16,"P"},;
					{17,"Q"},;
					{18,"R"},;
					{19,"S"},;
					{20,"T"},;
					{21,"U"},;
					{22,"V"},;
					{23,"W"},;
					{24,"X"},;
					{25,"Y"},;
					{26,"Z"},;
					{0,"Z"},;
					}
Static Function ColunasIndex(xNum,nIdx)
	Local cRet		:= ""
	Default nIdx	:= 1
	nPos	:= aScan(aColIdx,{|x| x[nIdx]==xNum})
	If nPos>0
		If nIdx==1
			cRet	:= aColIdx[nPos][2]
		Else
			cRet	:= aColIdx[nPos][1]
		EndIf
	EndIf
Return cRet

//----------------------------------------------------------------------
//CLASSE DE TAGS
//----------------------------------------------------------------------
/*/{Protheus.doc} yExcelTag
Criação de Tag
@author Saulo Gomes Martins
@since 22/04/2017
@version p11.8

@type class
/*/
Class yExcelTag From LongClassName
	Data cNome
	Data cClassName
	Data oAtributos
	Data oIndice
	Data xValor
	Data oExcel			//Objeto referencia do yexcel
	Method New()			Constructor
	Method ClassName()
	Method GetNome()
	Method SetValor()
	Method AddValor()
	Method GetVAlor()
	Method AddAtributo()
	Method SetAtributo()
	Method GetAtributo()
//	Method GetPosAtributo()
	Method GetTag()
EndClass

Method New(cNome,xValor,oAtributo) Class yExcelTag
	Local nCont
	PARAMTYPE 0	VAR cNome  AS CHARACTER
	PARAMTYPE 1	VAR xValor  AS ARRAY, CHARACTER, DATE, NUMERIC, LOGICAL, OBJECT 		OPTIONAL DEFAULT Nil
	PARAMTYPE 2	VAR oAtributo  AS ARRAY,OBJECT		OPTIONAL DEFAULT tHashMap():new()
	::cNome			:= cNome
	::xValor		:= xValor
	::oIndice		:= tHashMap():new()
	If ValType(oAtributo)=="A"
		::oAtributos		:= tHashMap():new()
		For nCont:=1 to Len(oAtributo)
			::oAtributos:Set(oAtributo[nCont][1],oAtributo[nCont][2])
		Next
	ElseIf ValType(oAtributo)=="O"
		::oAtributos		:= oAtributo
	EndIf
	::cClassName	:= "YEXCELTAG"
Return self

Method GetNome() Class yExcelTag
Return ::cNome

Method ClassName() Class yExcelTag
Return "YEXCELTAG"

Method SetValor(xValor,xIndice) Class yExcelTag
	If ValType(xIndice)=="U"
		::xValor	:= xValor
	ElseIf ValType(xIndice)=="N"
		::xValor[xIndice]	:= xValor
	ElseIf ValType(xIndice)=="C" .and. ValType(::xValor)=="A"
		::AddValor(xValor,xIndice)
	Else
		::xValor	:= xValor
	EndIf
Return

Method GetValor(xIndice,xDefault) Class yExcelTag
	Local nPos
	If ValType(xIndice)=="U"
		xDefault	:=  ::xValor
	ElseIf ValType(xIndice)=="N"
		xDefault	:=  ::xValor[xIndice]
	ElseIf ValType(xIndice)=="C" .and. ValType(::xValor)=="A"
		If ::oIndice:Get(xIndice,@nPos)
			xDefault	:=  ::xValor[nPos]
		EndIf
	EndIf
Return xDefault

Method AddValor(xValor,xIndice) Class yExcelTag
	Local nPos
	If ValType(xIndice)=="C"
		If ::oIndice:Get(xIndice,@nPos)
			::xValor[nPos]	:= xValor
		Else
			AADD(::xValor,xValor)
			::oIndice:Set(xIndice,Len(::xValor))
		EndIf
	ElseIf ValType(xIndice)=="N"
		::xValor[xIndice]	:= xValor
	Else
		AADD(::xValor,xValor)
	EndIf
Return

Method AddAtributo(cAtributo,xValor) Class yExcelTag
	PARAMTYPE 0	VAR cAtributo  AS CHARACTER
	::oAtributos:Set(cAtributo,xValor)
Return

Method SetAtributo(cAtributo,xValor) Class yExcelTag
	Local nPos
	PARAMTYPE 0	VAR cAtributo  AS CHARACTER
	If ValType(xValor)=="U"
		::oAtributos:Del(cAtributo)
	Else
		::oAtributos:Set(cAtributo,xValor)
	EndIf
Return

Method GetAtributo(cAtributo,cDefault) Class yExcelTag
	Local xValor
	PARAMTYPE 0	VAR cAtributo  AS CHARACTER
	If ::oAtributos:Get(cAtributo,@xValor)
		Return xValor
	EndIf
Return cDefault

Method GetTag(nFile,lFechaTag,lSoValor) Class yExcelTag
	Local cRet	:= ""
	Local nCont,cString,lOk
	Local aListAtt
	Default lFechaTag	:= .T.
	Default lSoValor	:= .F.
	If ValType(nFile)<>"U"	//Gravação direto no arquivo
		If lSoValor
			GravaFile(nFile,VarTipo(::xValor,nFile))
		Else
			GravaFile(nFile,'<'+::cNome)
			::oAtributos:List(@aListAtt)
			For nCont:=1 to Len(aListAtt)
				GravaFile(nFile,' '+aListAtt[nCont][1]+'="'+Transform(aListAtt[nCont][2],"")+'"')
			Next
			If ValType(::xValor)=="U"
				If lFechaTag
					GravaFile(nFile,'/>')
				Else
					GravaFile(nFile,'>')
				EndIf
			Else
				GravaFile(nFile,'>')
				GravaFile(nFile,VarTipo(::xValor,nFile))
				If lFechaTag
					GravaFile(nFile,'</'+::cNome+'>')
				EndIf
			EndIf
		EndIf
	Else
		If lSoValor
			cRet	+= VarTipo(::xValor)
		Else
			cRet	:= '<'+::cNome
			::oAtributos:List(@aListAtt)
			For nCont:=1 to Len(aListAtt)
				cRet	+= ' '+aListAtt[nCont][1]+'="'+Transform(aListAtt[nCont][2],"")+'"'
			Next
			If ValType(::xValor)=="U"
				If lFechaTag
					cRet	+= '/>'
				Else
					cRet	+= '>'
				EndIf
			Else
				cRet	+= '>'
				cRet	+= VarTipo(::xValor)
				If lFechaTag
					cRet	+= '</'+::cNome+'>'
				EndIf
			EndIf
		EndIf
	EndIf
Return cRet

Static Function VarTipo(xValor,nFile)
	Local nCont,aList
	Local cRet	:= ""
	If ValType(xValor)=="A"
		For nCont:=1 to Len(xValor)
			cRet	+= VarTipo(xValor[nCont])
		Next
	ElseIf ValType(xValor)=="O"
		If GetClassName(xValor)=="THASHMAP"
			xValor:List(@aList)
			aSort(aList,,,{|x,y| x[1]<y[1] })
			For nCont:=1 to Len(aList)
				cRet	+= VarTipo(aList[nCont][2])
			Next
		Else
			cRet	+= xValor:GetTag(nFile)
		EndIf
	Else
		cRet	+= Transform(xValor,"")
	EndIf
Return cRet

//----------------------------------------------------------------------
//CLASSE DE TABELAS
//----------------------------------------------------------------------
/*/{Protheus.doc} yExcel_Table
CLASSE PARA CRIAÇÃO DE TABELAS PARA A CLASSE YEXCEL
@author Saulo Gomes Martins
@since 08/05/2017

@type class
/*/
Class yExcel_Table from yExcelTag
	Data oyExcel
	Data lAutoFilter
	Data aRef
	Data nPrimLinha
	Data oColunas
	Data aColunas
	Data nLinha
	Data oTableColumns
	Data otableStyleInfo
	Data cNomeTabela
	Method new() constructor
	Method cell()
	Method AddStyle()
	Method AddLine()
	Method AddColumn()
	Method AddFilter()
	METHOD AddTotal()
	METHOD AddTotais()
	METHOD Finish()
EndClass

Method new(oyExcel,nLinha,nColuna,cNome) Class yExcel_Table
	_Super:New("table",{})
	::oyExcel	:= oyExcel
	::aRef		:= {{nLinha,nColuna},{0,0}}
	::nPrimLinha:= nLinha
	::oColunas	:= tHashMap():New()
	::aColunas	:= {}
	::nLinha	:= 0
	::cNomeTabela	:= cNome
	::AddLine()
Return self

/*/{Protheus.doc} AddFilter
Adiciona filtro a tabela
@author Saulo Gomes Martins
@since 08/05/2017

@type function
/*/
Method AddFilter() Class yExcel_Table
	::lAutoFilter:= .T.
Return

/*/{Protheus.doc} Cell
Preenche informação da célula
@author Saulo Gomes Martins
@since 08/05/2017
@version undefined
@param cColuna, characters, Nome da coluna
@param xValor, , conteudo da célula
@param [cFormula], characters, Formula
@param [nStyle], numeric, posição da formatação
@type function
/*/
METHOD Cell(cColuna,xValor,cFormula,nStyle) CLASS yExcel_Table
	Local aColuna,nColuna
	If ValType(cColuna)=="C"
		If !::oColunas:Get(cColuna,@aColuna)
			UserException("YExcel - Coluna informada não encontrado. Utilize o metodo AddColumn para adicionar a coluna:"+cValToChar(cColuna))
		Endif
		nColuna	:= aColuna[2]
		If Empty(nStyle)
			nStyle	:= aColuna[5]
		EndIf
	Else
		nColuna	:= cColuna
	EndIf
	::oyExcel:Cell(::nLinha,nColuna,xValor,cFormula,nStyle)
Return

/*/{Protheus.doc} AddLine
Adiciona uma nova linha
@author Saulo Gomes Martins
@since 08/05/2017
@param nQtd, numeric, Quantidade de linhas para avançar
@type function
/*/
Method AddLine(nQtd) CLASS yExcel_Table
	Default nQtd	:= 1
	::nLinha		+= nQtd
	::aRef[2][1]	:= ::nLinha
return ::nLinha

/*/{Protheus.doc} AddColumn
Adiciona uma nova coluna a tabela
@author Saulo Gomes Martins
@since 08/05/2017
@version undefined
@param cNome, characters, descricao
@param nStyle, numeric, descricao
@type function
/*/
METHOD AddColumn(cNome,nStyle) CLASS yExcel_Table
	Local nPosCol		:= aScan(self:GetValor(),{|x| x:GetNome()=="tableColumns"})
	Local otableColumn
	::aRef[2][2]	+= 1

	nCont	:= Len(self:oTableColumns:GetValor())+1
	otableColumn	:= yExcelTag():New("tableColumn",{},)
	otableColumn:SetAtributo("id",nCont)
	otableColumn:SetAtributo("name",cNome)
	self:oTableColumns:SetAtributo("count",nCont)
	self:oTableColumns:AddValor(otableColumn)
	::oColunas:Set(cNome,{::aRef[1][1],::aRef[1][2]+Len(::aColunas),otableColumn,nil,nStyle})
	AADD(::aColunas,cNome)
	::Cell(cNome,cNome)
//	Varinfo("self:oTableColumns",self:oTableColumns,,.F.)
Return


/*/{Protheus.doc} AddTotal
Adiciona um totalizador na coluna
@author Saulo Gomes Martins
@since 08/05/2017
@version undefined
@param cColuna, characters, Nome da coluna
@param xValor, , Valor
@param cFunction, characters, Formula
@type function
@see https://support.office.com/pt-br/article/SUBTOTAL-Fun%C3%A7%C3%A3o-SUBTOTAL-7b027003-f060-4ade-9040-e478765b9939?NS=EXCEL&Version=16&SysLcid=1046&UiLcid=1046&AppVer=ZXL160&HelpId=xlmain11.chm60392&ui=pt-BR&rs=pt-BR&ad=BR
@obs PAG 2392
function-number				function-number					Function
(includes hidden values)	(excludes hidden values)
1 							101 							AVERAGE	MÉDIA
2 							102 							COUNT	CONTAR NUMERO
3 							103 							COUNTA	CONT.VALORES
4 							104 							MAX		MAX
5 							105 							MIN		MIN
6 							106 							PRODUCT	MULT
7 							107 							STDEV	DESVPAD
8 							108 							STDEVP	DESVPADP
9 							109 							SUM		SOMA
10 							110 							VAR		VAR
11 							111 							VARP	VARP
/*/
Method AddTotal(cColuna,xValor,cFunction,nStyle) CLASS yExcel_Table
	Local aColuna
	If ::oColunas:Get(cColuna,@aColuna)
		otableColumn	:= aColuna[3]
		aColuna[4]		:= xValor
		If ValType(nStyle)<>"U"
			aColuna[5]		:= nStyle
		EndIf
		//::oColunas:Set(cColuna,aColuna)
		If Empty(cFunction)
			otableColumn:SetAtributo("totalsRowLabel",xValor)
		Else
			If UPPER(cFunction)=="AVERAGE"
				otableColumn:SetAtributo("totalsRowFunction",lower("AVERAGE"))
			ElseIf UPPER(cFunction)=="COUNT"
				otableColumn:SetAtributo("totalsRowFunction",lower("COUNT"))
//			ElseIf UPPER(cFunction)=="COUNTA"
//				otableColumn:SetAtributo("totalsRowFunction",lower("COUNTA"))
			ElseIf UPPER(cFunction)=="MAX"
				otableColumn:SetAtributo("totalsRowFunction",lower("MAX"))
			ElseIf UPPER(cFunction)=="MIN"
				otableColumn:SetAtributo("totalsRowFunction",lower("MIN"))
//			ElseIf UPPER(cFunction)=="PRODUCT"
//				otableColumn:SetAtributo("totalsRowFunction",lower("PRODUCT"))
//			ElseIf UPPER(cFunction)=="STDEV"
//				otableColumn:SetAtributo("totalsRowFunction",lower("STDEV"))
//			ElseIf UPPER(cFunction)=="STDEVP"
//				otableColumn:SetAtributo("totalsRowFunction",lower("STDEVP"))
			ElseIf UPPER(cFunction)=="SUM"
				otableColumn:SetAtributo("totalsRowFunction",lower("SUM"))
			ElseIf UPPER(cFunction)=="VAR"
				otableColumn:SetAtributo("totalsRowFunction",lower("VAR"))
//			ElseIf UPPER(cFunction)=="VARP"
//				otableColumn:SetAtributo("totalsRowFunction",lower("VARP"))
			Else
				otableColumn:SetAtributo("totalsRowFunction","custom")
				otableColumn:AddValor(yExcelTag():New("totalsRowFormula",cFunction,),"totalsRowFormula")
			EndIf
		EndIf
	EndIf
Return

/*/{Protheus.doc} AddTotais
Inclui a linha de totalizador
@author Saulo Gomes Martins
@since 08/05/2017

@type function
/*/
Method AddTotais() CLASS yExcel_Table
	Local nCont,xValor,cFormula
	Local aColuna,cRef
	cRef		:= ::oyExcel:Ref(::aRef[1][1],::aRef[1][2])+":"+::oyExcel:Ref(::aRef[2][1]+1,::aRef[2][2])
	::SetAtributo("ref",cRef)
	::SetAtributo("totalsRowCount",1)
	//::SetAtributo("totalsRowShown",1)
	For nCont:=1 to Len(::aColunas)
		::oColunas:Get(::aColunas[nCont],@aColuna)
		xValor	:= aColuna[4]
		cFormula:= nil
		otableColumn	:= aColuna[3]
		If !Empty(otableColumn:GetAtributo("totalsRowLabel",""))
			xValor		:= otableColumn:GetAtributo("totalsRowLabel","")
		ElseIf !Empty(otableColumn:GetAtributo("totalsRowFunction",""))
			If UPPER(otableColumn:GetAtributo("totalsRowFunction",""))=="AVERAGE"
				cFormula		:= "SUBTOTAL(101,"+::cNomeTabela+"["+::aColunas[nCont]+"])"
			ElseIf UPPER(otableColumn:GetAtributo("totalsRowFunction",""))=="COUNT"
				cFormula		:= "SUBTOTAL(102,"+::cNomeTabela+"["+::aColunas[nCont]+"])"
			ElseIf UPPER(otableColumn:GetAtributo("totalsRowFunction",""))=="COUNTA"
				cFormula		:= "SUBTOTAL(103,"+::cNomeTabela+"["+::aColunas[nCont]+"])"
			ElseIf UPPER(otableColumn:GetAtributo("totalsRowFunction",""))=="MAX"
				cFormula		:= "SUBTOTAL(104,"+::cNomeTabela+"["+::aColunas[nCont]+"])"
			ElseIf UPPER(otableColumn:GetAtributo("totalsRowFunction",""))=="MIN"
				cFormula		:= "SUBTOTAL(105,"+::cNomeTabela+"["+::aColunas[nCont]+"])"
			ElseIf UPPER(otableColumn:GetAtributo("totalsRowFunction",""))=="PRODUCT"
				cFormula		:= "SUBTOTAL(106,"+::cNomeTabela+"["+::aColunas[nCont]+"])"
			ElseIf UPPER(otableColumn:GetAtributo("totalsRowFunction",""))=="STDEV"
				cFormula		:= "SUBTOTAL(107,"+::cNomeTabela+"["+::aColunas[nCont]+"])"
			ElseIf UPPER(otableColumn:GetAtributo("totalsRowFunction",""))=="STDEVP"
				cFormula		:= "SUBTOTAL(108,"+::cNomeTabela+"["+::aColunas[nCont]+"])"
			ElseIf UPPER(otableColumn:GetAtributo("totalsRowFunction",""))=="SUM"
				cFormula		:= "SUBTOTAL(109,"+::cNomeTabela+"["+::aColunas[nCont]+"])"
			ElseIf UPPER(otableColumn:GetAtributo("totalsRowFunction",""))=="VAR"
				cFormula		:= "SUBTOTAL(110,"+::cNomeTabela+"["+::aColunas[nCont]+"])"
			ElseIf UPPER(otableColumn:GetAtributo("totalsRowFunction",""))=="VARP"
				cFormula		:= "SUBTOTAL(111,"+::cNomeTabela+"["+::aColunas[nCont]+"])"
			Elseif otableColumn:GetAtributo("totalsRowFunction","")=="custom"
				cFormula		:= otableColumn:GetValor("totalsRowFormula"):GetValor()
			EndIf
		EndIf
		If ValType(xValor)=="U" .and. ValType(cFormula)=="U"
			Loop
		Else
			::oyExcel:Cell(::aRef[2][1]+1,aColuna[2],xValor,cFormula,aColuna[5])
		EndIf
	Next

Return

/*/{Protheus.doc} Finish
Finaliza a tabela criada
@author Saulo Gomes Martins
@since 03/05/2017
@version undefined
@type function

/*/
METHOD Finish() CLASS yExcel_Table
	cRef		:= ::oyExcel:Ref(::aRef[1][1],::aRef[1][2])+":"+::oyExcel:Ref(::aRef[2][1],::aRef[2][2])
	nPosCol		:= aScan(self:GetValor(),{|x| x:GetNome()=="autoFilter"})
	If ::lAutoFilter
		self:GetValor(nPosCol):SetAtributo("ref",cRef)
	Else
		aDel(self:GetValor(),nPosCol)
		aSize(self:GetValor(),Len(self:GetValor())-1)
	EndIf
	cRef		:= ::oyExcel:Ref(::aRef[1][1],::aRef[1][2])+":"+::oyExcel:Ref(::aRef[2][1]+::GetAtributo("totalsRowCount",0),::aRef[2][2])
	::SetAtributo("ref",cRef)
Return

/*/{Protheus.doc} AddStyle
Preenche os estilos da tabela
@author Saulo Gomes Martins
@since 08/05/2017
@param cNome, characters, Nome do estilo|ver Obs
@param lLinhaTiras, logical, Linhas em tiras
@param lColTiras, logical, Colunas em tiras
@param lFormPrimCol, logical, Exibir formato especial na primeira coluna da tabela
@param lFormUltCol, logical, Exibir formato especial na ultima coluna da tabela
@type function
@OBS Annex G. (normative) Predefined SpreadsheetML Style Definitions
PAG 4426
	TableStyleMedium2 	- AZUL|LINHA1-AZUL_CLARO|LINHA2-BRANCO|SEM BORDA
	TableStyleMedium9 	- AZUL|LINHA1-AZUL_ESCURO|LINHA2-AZUL_CLARO|SEM BORDA
	TableStyleMedium16	- AZUL|LINHA1-CINZA|LINHA2-BRANCO|SEM BORDA INTERNA
	TableStyleLight9	- AZUL|LINHA1-BRANCO|LINHA2-BRANCO|BORDA DE LINHA
	TableStyleMedium15	- PRETO|LINHA1-CINZA|LINHA2-BRANCO|COM BORDA
	TableStyleMedium1	- PRETO|LINHA1-CINZA|LINHA2-BRANCO|SEM BORDA
	TableStyleMedium8	- PRETO|LINHA1-CINZA_ESCURO|LINHA2-CINZA_CLARO|SEM BORDA
	TableStyleMedium22	- CINZA|LINHA1-CINZA_ESCURO|LINHA2-CINZA_CLARO|COM BORDA
	TableStyleLight16	- BRANCO|LINHA1-AZUL_CLARO|LINHA2-BRANCO|BORDA AZUL
	TableStyleLight15	- BRANCO|LINHA1-CINZA_CLARO|LINHA2-BRANCO|COM BORDA
	TableStyleLight1	- BRANCO|LINHA1-CINZA_CLARO|LINHA2-BRANCO|SEM BORDA
/*/
METHOD AddStyle(cNome,lLinhaTiras,lColTiras,lFormPrimCol,lFormUltCol) CLASS yExcel_Table
	Default cNome		:= nil
	Default lLinhaTiras	:= .F.
	Default lColTiras	:= .F.
	Default lFormPrimCol:= .F.
	Default lFormUltCol	:= .F.

	::otableStyleInfo:SetAtributo("name",cNome)
	If lLinhaTiras	//Linhas em tiras
		::otableStyleInfo:SetAtributo("showRowStripes",1)
	Else
		::otableStyleInfo:SetAtributo("showRowStripes",0)
	EndIf
	If lColTiras	//Colunas em tiras
		::otableStyleInfo:SetAtributo("showColumnStripes",1)
	Else
		::otableStyleInfo:SetAtributo("showColumnStripes",0)
	EndIf
	If lFormPrimCol	//Exibir formato especial na primeira coluna da tabela
		::otableStyleInfo:SetAtributo("showFirstColumn",1)
	Else
		::otableStyleInfo:SetAtributo("showFirstColumn",0)
	EndIf
	If lFormUltCol	//Exibir formato especial na ultima coluna da tabela
		::otableStyleInfo:SetAtributo("showLastColumn",1)
	Else
		::otableStyleInfo:SetAtributo("showLastColumn",0)
	EndIf
Return
