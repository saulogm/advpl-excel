#include "Totvs.ch"
#include "Fileio.ch"
#Include "ParmType.ch"

Static cAr7Zip	//Caminho do 7zip para compactar o arquivo
Static cRootPath
//CLASSE EXCEL
/*/{Protheus.doc} YExcel
Gerar Excel formato xlsx

@author Saulo Gomes Martins
@since 27/10/2014 17:51:57
@version 2.0
@type function
@OBS
RECURSOS DISPONIVEIS
* Definir células String,Numérica,data,DateTime,Logica,formula
* Adicionar novas planilhas(Nome,Cor)
* Cor de preenchimento(simples,efeito de preenchimento)
* Alinhamento(Horizontal,Vertical,Reduzir para Caber,Quebra Texto,Angulo de Rotação)
* Formato da célula
* Mesclar células
* Auto Filtro
* hyperlink dentro da planilha
* Comentário
* Congelar painéis(colunas e linhas)
* Definir tamanho da linha / largura da coluna
* Formatar numeros(casas decimais)
* Letra: Fonte,Tamanho,Cor,Negrito,Italico,Sublinhado,Tachado
* Bordas: (Left,Right,Top,Bottom),Cor,Estilo
* Formatação condicional:(operador,formula)(font,fundo,bordas)
* Formatar como tabela(Estilos Predefinidos,Filtros,Totalizadores)
* Cria nome para referência de célula ou intervalo
* Agrupamento de linha e colunas
* Imagens
* Exibir/Oculta linhas de Grade
* Definir linha para repetir na impressão
* Definir orientação da pagina na impressão
* Cabeçalho e Ropadé
* Leitura de dados já gravados

* Leitura simples dos dados
@type class
@see https://github.com/saulogm/advpl-excel
@obs Office Open XML File Formats, sumario 1533
/*/
//Dummy Function
User Function YExcel()
Return .T.

Class YExcel
	Data oString			//String compartilhadas
	Data nQtdString			//Quantidade de string conmpartilhadas
	Data adimension			//Dimensão da planilha
	Data cClassName			//Nome da Classe
	Data cName				//Nome da Classe
	Data cTmpFile			//Arquivo temporario criado no servidor
	Data cNomeFile			//Nome do arquivo para gerar
	Data cNomeFile2			//Nome do arquivo para gerar
	Data nFileTmpRow		//nHeader do Arquivo temporario de linhas
	Data cPlanilhaAt		//Nome da planilha atual
	Data nPlanilhaAt		//Indice da planilha atual
	Data aPlanilhas			//Dados das planilhas
	Data lRowDef			//deprecated
	Data nTamLinha			//Tamanho da linha atual
	Data nColunaAtual		//Ultima Coluna
	Data nQtdStyle			//Quantidade de styles
	Data nLinha
	Data nColuna
	Data nPriodFormCond
	Data atable
	Data aFiles
	Data nIdRelat
	Data nCont
	Data oCell
	Data nNumFmtId
	Data cPagOrientation
	Data aPadraoSty
	Data cDriver
	Data aTmpDB
	Data cAliasCol
	Data cAliasLin
	Data cAliasStr
	Data cAliasChv
	Data aworkdrawing	//arquivo drawing do worksheets
	Data odrawing		//tag drawing dentro do sheet
	Data aImagens		//Imagens adicionada
	Data aImgdraw		//Imagens usada no sheets(pode usar mais de uma vez a mesma imagem)
	Data nIDMedia		//Sequencial do id da imagem

	Data aRels			//Arquivos rels
	Data ocontent_types	//content_types.xml
	Data oapp			//app.xml
	Data ocore			//core.xml
	Data oworkbook		//workbook.xml
	Data aDraw			//Arquivos Draw
	Data oStyle			//styles.xml
	Data asheet			//arquivo sheet das planilhas
	Data aCleanObj
	Data lDelSrv
	Data cArqGrv
	Data cLocalFile
	//Agrupamento de linha
	Data nRowoutlineLevel
	Data lRowcollapsed
	Data lRowHidden
	Data osheet

	METHOD New() CONSTRUCTOR
	METHOD ClassName()

	//Controle das planilhas
	METHOD ADDPlan()		//Adiciona nova planilha ao arquivo
	METHOD Gravar()			//Grava em disco
	METHOD OpenApp()		//Abre arquivo gravado
	METHOD Save()			//Salva xlsx
	METHOD Close()			//fecha arquivo aberto e limpa temporario
	METHOD SetPlanName()	//Altera o nome da planilha
	METHOD SetPlanAt()		//Informa qual planilha está em edição
	METHOD GetPlanAt()		//Retorna qual planilha está em edição
	METHOD LenPlanAt()		//Quantidade de planilha
	//Controle de Células
	METHOD Cell()			//Grava as células
	METHOD Pos()			//posiciona na celula
	METHOD PosR()			//posiciona na celula de acordo com referência
	METHOD GetValue()		//Retorna conteudo da celulas posicionada
	METHOD SetValue()		//Grava conteudo da celulas posicionada
	METHOD SetDateTime()	//Grava conteudo com data e hora
	METHOD GetFormula()		//Retorna a formula da celula posicionada
	METHOD ColTam()			//Coluna Mínima e Máxima
	METHOD LinTam()			//Linha Mínima e Máxima
	METHOD mergeCells()		//Mescla células
	METHOD NumToString()	//Algoritimo para converte numero em string A=1,B=2
	METHOD StringToNum()	//Algoritimo para converte string em numero 1=A,2=B
	METHOD Ref()			//Passa a localização numerica e transforma em referencia da celula
	METHOD LocRef()			//Retorna linha  e coluna de acordo com referencia enviada
	METHOD AddTamCol()		//Defini o tamanho de uma coluna ou varias colunas
	METHOD AutoFilter()		//Cria os Filtros na planilha
	METHOD AddNome()		//Cria nome para refencia de célula ou intervalo
	METHOD Addhyperlink()	//Cria um hyperlink para uma referência da planilha
	METHOD AddComment()		//Cria um comentário para a celula posicionada
	
	METHOD InsertRowEmpty()	//Cria linhas vazia
	METHOD InsertCellEmpty()//Cria células vazias
	METHOD NivelLinha()
	METHOD SetsumBelow()	//Configurar linha resumo de agrupamento de linhas abaixo
	METHOD SetsumRight()	//Configurar coluna resumo a direita
	METHOD SetRowLevel()	//Configurar as linhas para agrupamento 
	METHOD SetRowH()		//Configurar tamanho das linhas
	METHOD SetColLevel()	//Configurar as linhas para agrupamento 
	METHOD AddRow()			//Adiciona linhas acima e move as demais para baixo
	METHOD AddCol()			//Adiciona Colunas a direita e move as demais para esquerda
	//Controle de layout, impressão e pagina
	METHOD AddPane()		//Congelar Painéis
	METHOD showGridLines()	//Exibir ou ocultar linhas de grade
	METHOD SetPrintTitles()	//Configurar linha para repetir na impressão
	METHOD SetPagOrientation()	//Configurar orientação da pagina na impressão
	Method SetHeader()		//Configurar Cabeçalho
	Method SetFooter()		//Configurar Rodapé

	METHOD GetDateTime()		//Cria dado para incluir celula data time

	//String Compartilhadas SQL
	METHOD GetStrComp()
	METHOD SetStrComp()

	//Leitura de planilha
	METHOD OpenRead()
	METHOD CellRead()
	METHOD CloseRead()
	METHOD LerPasta()

	//Interno
	METHOD CriarFile()		//Cria arquivos temporarios
	METHOD GravaFile()		//Grava em arquivos temporarios
	METHOD AddFormatCond()	//Formatação condicional(todos rercusos)
	METHOD Pane()			//Congelar Painéis
	METHOD CriaDB()			//Cria base de dados interna

	//Estilo
	METHOD CorPreenc()		//Cria um nova cor para ser usada
	METHOD EfeitoPreenc()	//Cria um novo efeito de preenchimento
	METHOD AddFont()		//Cria objeto de font
	METHOD CreateStyle()	//Cria estilos com herança
	METHOD NewStyle()		//Cria estilos orientado a objeto
	METHOD AddStyles()		//Adiciona Estilos
	METHOD SetStyle()		//Informa o estilo em uma ou várias célula
	METHOD GetStyle()		//Retorna o estilo em uma célula
	METHOD Alinhamento()	//Adiciona alinhamento
	METHOD Borda()			//Adiciona borda(auxiliar)
	METHOD Border()			//Cria Borda com todas opções
	Method AddFmt()			//Cria formato
	Method AddFmtNum()		//Cria formato para numeros
	Method SetStyFmt()		//Informa um formato de estilo
	Method SetStyFont()		//Informa uma fonte do estilo
	Method SetStyFill()		//Informa um preenchimento de fundo do estilo
	Method SetStyborder()	//Informa uma borda do estilo
	Method SetStyxf()		//Informa xf do estilo
	Method SetStyaValores()	//Informa aValores do estilo
	Method SetStyaOutrosAtributos()	//Informa aOutrosAtributos do estilo
	Method NewStyRules()	//Cria auxiliar
	Method StyleType()		//Retorna o tipo do estilo


	//Formatação condicional
	METHOD FormatCond()		//Definir formatação condicional(auxiliar)
	METHOD Font()			//Cria objeto de font
	METHOD Preenc()			//Cria objeto de Preenchimento
	METHOD ObjBorda()		//Cria objeto de borda
	METHOD gradientFill()	//Cria objeto de efeito de preenchimento
	METHOD ADDdxf()			//Cria o estilo para formatação condicional

	//Imagem
	METHOD ADDImg()			//Adiciona uma Imagem
	METHOD Img()			//Usa imagem

	//Inicializar TXmlManager
	METHOD new_content_types()
	METHOD new_rels()
	METHOD add_rels()
	METHOD Get_rels()
	METHOD FindRels()
	METHOD new_app()
	METHOD new_core()
	METHOD new_workbook()
	METHOD new_draw()
	METHOD xls_sheet()
	METHOD xls_table()
	METHOD xls_sharedStrings()
	METHOD Read_sharedStrings()
	METHOD new_styles()
	METHOD new_comment()
	METHOD new_vmlDrawing()

	//Tabela
	METHOD AddTabela()

	//deprecated
	METHOD SetDefRow()		//[deprecated]Defini as colunas da linha. Habilita a gravação automatica de cada coluna. Importante para prover performace na gravação de varias linhas
ENDCLASS

METHOD ClassName() Class YExcel
Return "YEXCEL"

/*/{Protheus.doc} New
Construtor da classe
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param cNomeFile, characters, Nome do arquivo para gerar
@type method
/*/
METHOD New(cNomeFile,cFileOpen) Class YExcel
	Local nPos
	Local aStruct,aIndex
	Local cDriver	:= "SQLITE_TMP"
	Local cNomTmp	:= lower(CriaTrab(,.F.))
	If ValType(cAr7Zip)=="U"
		cAr7Zip := GetPvProfString("GENERAL", "LOCAL7ZIP" , "C:\Program Files\7-Zip\7z.exe" , GetAdv97() )
	Endif
	PARAMTYPE 0	VAR cNomeFile	AS CHARACTER		OPTIONAL DEFAULT cNomTmp
	// PARAMTYPE 2	VAR cDriver		AS CHARACTER			OPTIONAL DEFAULT "SQLITE_TMP"
	::cClassName	:= "YEXCEL"
	::cName			:= "YEXCEL"
	::oString		:= tHashMap():new()
	::oCell			:= nil	//Usado no leitura simples
	::nQtdString	:= 0
	::nQtdStyle		:= 0
	::nNumFmtId		:= 166
	::aPlanilhas	:= {}
	::cTmpFile		:= cNomTmp
	::cNomeFile2	:= cNomeFile
	::cNomeFile		:= cNomTmp
	::nFileTmpRow	:= 0
	::lRowDef		:= .F.
	::nColunaAtual	:= 0
	::aFiles		:= {}
	::nIdRelat		:= 0
	::aworkdrawing	:= {}
	::aImagens		:= {}
	::aImgdraw		:= {}
	::nIDMedia		:= 0
	::aRels			:= {}
	::aDraw			:= {}
	::cPagOrientation	:= "landscape"
	::asheet		:= {}
	::aTmpDB		:= {}
	::aPadraoSty	:= {}
	::aCleanObj		:= {}
	::nLinha		:= 1
	::nColuna		:= 1
	::lDelSrv		:= .T.
	::cArqGrv		:= ""
	::cLocalFile	:= ""
	::nPriodFormCond:= 1
	::aImgdraw		:= {}
	AADD(::aCleanObj,::oString)

	//CRIAR ESTRUTURA DO BANCO
	If TYPE("__TTSInUse")=="U"
		CriaPublica()
	Endif
	::cDriver	:= cDriver
	If ::cDriver=="TMPDB"
		If !TCIsConnected()
			nH := TCLink()
			If nH < 0
				UserException("DBAccess - Erro de conexao "+cValToChar(nH))
			Endif
		Endif
	Endif

	//COLUNAS
	aStruct	:= {}
	aIndex	:= {}
	AADD(aStruct,{"PLA"		,	"N", 10		, 00})
	AADD(aStruct,{"LIN"		,	"N", 10		, 00})
	AADD(aStruct,{"COL"		,	"N", 10		, 00})
	AADD(aStruct,{"STY"		,	"N", 10		, 00})
	AADD(aStruct,{"TPSTY"	,	"C", 1		, 00})	//Tipo de estilo(texto,numero,data,datetime,logico)
	AADD(aStruct,{"TIPO"	,	"C", 1		, 00})
	AADD(aStruct,{"FORMULA"	,	"C", 200	, 00})
	AADD(aStruct,{"TPVLR"	,	"C", 1		, 00})	//Tipo campo usado, txt ou num
	AADD(aStruct,{"VLRTXT"	,	"C", 200	, 00})
	AADD(aStruct,{"VLRNUM"	,	"N", 20		, 08})
	AADD(aStruct,{"VLRDEC"	,	"N", 15		, 00})	//Decimal maior que oito decimais
	AADD(aIndex,{"PLA","LIN","COL"})
	::cAliasCol	:= ::CriaDB(aStruct,aIndex,"COL")
	
	//LINHAS
	aStruct	:= {}
	aIndex	:= {}
	AADD(aStruct,{"PLA"		,	"N", 10		, 00})
	AADD(aStruct,{"LIN"		,	"N", 10		, 00})
	AADD(aStruct,{"OLEVEL"	,	"C", 1		, 00})
	AADD(aStruct,{"COLLAP"	,	"C", 1		, 00})
	AADD(aStruct,{"CHIDDEN"	,	"C", 1		, 00})
	AADD(aStruct,{"CHEIGHT"	,	"C", 1		, 02})
	AADD(aStruct,{"HT"		,	"N", 8		, 02})
	AADD(aIndex,{"PLA","LIN"})
	::cAliasLin	:= ::CriaDB(aStruct,aIndex,"LIN")
	
	//STRING COMPARTILHAS
	aStruct	:= {}
	aIndex	:= {}
	AADD(aStruct,{"POS"			,"N", 10	, 00})
	AADD(aStruct,{"VLRTEXTO"	,"C", 200	, 00})
	AADD(aStruct,{"VLRMEMO"		,"M", 8		, 00})
	AADD(aIndex,{"VLRTEXTO","POS"})
	AADD(aIndex,{"POS"})
	::cAliasStr	:= ::CriaDB(aStruct,aIndex,"STR")
	
	//CHAVES E IDS
	aStruct	:= {}
	aIndex	:= {}
	AADD(aStruct,{"TIPO"		,"C", 10	, 00})
	AADD(aStruct,{"CHAVE"		,"C", 200	, 00})
	AADD(aStruct,{"ID"			,"N", 7		, 00})
	AADD(aIndex,{"TIPO","CHAVE"})
	AADD(aIndex,{"TIPO","ID"})
	::cAliasChv	:= ::CriaDB(aStruct,aIndex,"CHV")

	If !Empty(cFileOpen)
		FWMakeDir("\tmpxls\"+::cTmpFile+'\',.F.)
		FWMakeDir("\tmpxls\"+::cTmpFile+'\'+::cNomeFile+'\',.F.)
		
		cNome	:= SubStr(cFileOpen,Rat("\",cFileOpen)+1)

		__COPYFILE(cFileOpen,"\tmpxls\"+::cTmpFile+'\'+cNome,,,.F.)

		nRet	:= StartJob("FUnZip",GetEnvServer(), .T.,"\tmpxls\"+::cTmpFile+'\'+cNome,"\tmpxls\"+::cTmpFile+'\'+::cNomeFile+'\')
		//Problema de não descompactar direto no servidor linux
		If nRet==0 .and. !File("\tmpxls\"+::cTmpFile+'\'+::cNomeFile+"\_rels\.rels",,.F.)
			If ValType(cRootPath)=="U"
				cRootPath	:= GetSrvProfString( "RootPath", "" )
			Endif
			If IsSrvUnix()
				WaitRunSrv('unzip -a "'+cRootPath+'/tmpxls/'+::cTmpFile+'/'+cNome+'" -d "'+cRootPath+'/tmpxls/'+::cTmpFile+'/'+::cNomeFile+'/"',.T.,cRootPath+'/tmpxls/'+self:cTmpFile+'/'+self:cNomeFile+'/')
			Endif
			If !File("\tmpxls\"+::cTmpFile+'\'+::cNomeFile+"\_rels\.rels",,.F.)
				FWMakeDir(GetTempPath()+"tmpxls\"+::cTmpFile)
				CpyS2T("\tmpxls\"+::cTmpFile+'\'+cNome, GetTempPath()+"tmpxls\"+::cTmpFile,,.F.)
				StartJob("FUnZip",GetEnvServer(),.T.,GetTempPath()+"tmpxls\"+::cTmpFile+"\"+cNome,GetTempPath()+"\tmpxls\"+::cTmpFile)
				fErase(GetTempPath()+"tmpxls\"+::cTmpFile+"\"+cNome)
				CpyPasta(GetTempPath()+"tmpxls/"+::cTmpFile,"\tmpxls\"+::cTmpFile+'\'+::cNomeFile+'\')
				__COPYFILE(GetTempPath()+"tmpxls/"+::cTmpFile+"/_rels/.rels","\tmpxls\"+::cTmpFile+'\'+::cNomeFile+"\_rels\.rels",,,.F.)
			Endif
		Endif

		nPos	:= ::new_rels("\tmpxls\"+::cTmpFile+'\'+::cNomeFile+"\_rels\.rels","\_rels\.rels")	//Arquivo não é carregado pela função Directory
		fErase("\tmpxls\"+::cTmpFile+'\'+::cNomeFile+"\_rels\.rels")
		::LerPasta("\tmpxls\"+::cTmpFile+'\'+::cNomeFile,,".rels")	//Ler todos rels
		::LerPasta("\tmpxls\"+::cTmpFile+'\'+::cNomeFile)
		LerChvStys(::self)
		If aScan(::aFiles,{|x| lower(x)=="\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\sharedstrings.xml"})==0
			AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\sharedStrings.xml")
			nPos	:= aScan(::aRels,{|x| x[2]=="\xl\_rels\workbook.xml.rels"})
			::add_rels(nPos,"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings","sharedStrings.xml")
			::ocontent_types:XPathAddNode("/xmlns:Types","Override","")
			::ocontent_types:XPathAddAtt("/xmlns:Types/xmlns:Override[last()]","PartName","/xl/sharedStrings.xml")
			::ocontent_types:XPathAddAtt("/xmlns:Types/xmlns:Override[last()]","ContentType","application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml")
		EndIf
	else
		::new_app()
		::new_core()
		::new_workbook()
		::new_content_types()
		::new_styles()

		nPos	:= ::new_rels(,"\_rels\.rels")
		::add_rels(nPos,"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument","xl/workbook.xml")
		::add_rels(nPos,"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties","docProps/core.xml")
		::add_rels(nPos,"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties","docProps/app.xml")
		
		AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\sharedStrings.xml")
		AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\theme\theme1.xml")
		nPos	:= ::new_rels(,"\xl\_rels\workbook.xml.rels")
		::add_rels(nPos,"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme","theme/theme1.xml")
		::add_rels(nPos,"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles","styles.xml")
		::add_rels(nPos,"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings","sharedStrings.xml")
		//Defini formato Moeda padrão brasileiro
		::AddFmt('_-"R$"\ * #,##0.00_-;\-"R$"\ * #,##0.00_-;_-"R$"\ * "-"??_-;_-@_-',44)
	Endif

	//Cria conteudo padrão
	::AddFont(11,"FF000000","Calibri","2")
	::Borda()	//Sem borda
	::CorPreenc(,,"none")
	::CorPreenc(,,"gray125")
	AADD(::aPadraoSty,::AddStyles(0/*numFmtId*/,/*fontId*/,/*fillId*/,/*borderId*/,/*xfId*/,/*aValores*/,{{"applyNumberFormat","0"}}/*aOutrosAtributos*/))	//Sem Formatação
	AADD(::aPadraoSty,::AddStyles(14/*numFmtId*/,/*fontId*/,/*fillId*/,/*borderId*/,/*xfId*/,/*aValores*/,/*aOutrosAtributos*/))	//Formato Data padrão
	AADD(::aPadraoSty,::AddStyles(::AddFmt("dd/mm/yyyy\ hh:mm AM/PM;@")/*numFmtId*/,/*fontId*/,/*fillId*/,/*borderId*/,/*xfId*/,/*aValores*/,/*aOutrosAtributos*/))	//Formato Data time padrão
Return self

//Não Documentear
Method CriaDB(aStruct,aIndex,cPrefixo) Class YExcel
	Local oTabTmp
	Local cAliasRet
	Local nCont
	Default cPrefixo	:= ""
	If ::cDriver=="TMPDB"
		oTabTmp	:= FWTemporaryTable():New( )
		oTabTmp:SetFields( aStruct )
		For nCont:=1 to Len(aIndex)
			oTabTmp:AddIndex("indice"+cValToChar(nCont), aIndex[nCont] )
		Next
		oTabTmp:Create()
		cAliasRet	:= oTabTmp:GetAlias()
		AADD(::aTmpDB,oTabTmp)
	Else
		cAliasRet	:= cPrefixo+CriaTrab(,.F.)
		DBCreate( cAliasRet , aStruct, ::cDriver )
		DBUseArea( .T., ::cDriver, cAliasRet, cAliasRet, .F., .F. )
		CriaIndices(cAliasRet,aIndex,aStruct)
	Endif
Return cAliasRet
//Criar indices
Static Function CriaIndices(cAliasTMP,aIndex,aStruct)
	Local nCont,nCont2,cStringInd
	Local nPos
	For nCont:=1 to Len(aIndex)
		cStringInd	:= ""
		For nCont2:=1 to Len(aIndex[nCont])
			If nCont2>1
				cStringInd	+= "+"
			Endif
			nPos	:= aScan(aStruct,{|x| x[1]==aIndex[nCont][nCont2] })
			If nPos==0
				UserException("Erro na estrutura interna do servico YExcel")
			ElseIf aStruct[nPos][2]=="N"
				cStringInd	+= "Str("+aIndex[nCont][nCont2]+","+cValToChar(aStruct[nPos][3])+","+cValToChar(aStruct[nPos][4])+")"
			ElseIf aStruct[nPos][2]=="D"
				cStringInd	+= "DTOS("+aIndex[nCont][nCont2]+")"
			Else
				cStringInd	+= aIndex[nCont][nCont2]
			Endif
		Next
		(cAliasTMP)->(DBCreateIndex(cAliasTMP+'IDX'+cValToChar(nCont), cStringInd, &("{ || "+cStringInd+" }")))
	Next
	(cAliasTMP)->(dbClearIndex())
	For nCont:=1 to Len(aIndex)
		(cAliasTMP)->(dbSetIndex(cAliasTMP+'IDX'+cValToChar(nCont)))
	Next
Return

/*/{Protheus.doc} ADDPlan
Adiciona nova planilha ao arquivo
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param cNome, characters, nome da planilha
@type method
/*/
METHOD ADDPlan(cNome,cCor) Class YExcel
	Local cID
	Local nQtdPlanilhas	:= Len(::aPlanilhas)
	Local nPos
	Private oSelf	:= Self
	PARAMTYPE 0	VAR cNome			AS CHARACTER		OPTIONAL DEFAULT "Planilha"+cValToChar(nQtdPlanilhas+1)
	PARAMTYPE 1	VAR cCor			AS CHARACTER		OPTIONAL
	cNome	:= Replace(cNome,"\","")
	cNome	:= Replace(cNome,"/","")
	cNome	:= Replace(cNome,":","")
	cNome	:= Replace(cNome,"*","")
	cNome	:= Replace(cNome,"[","")
	cNome	:= Replace(cNome,"]","")
	cNome	:= Replace(cNome,"?","")
	cNome	:= Replace(cNome,">","&gt;")
	cNome	:= Replace(cNome,"<","&lt;")
	If Len(cNome)>31
		cNome	:= SubStr(cNome,1,31)
	Endif
	cNome	:= EncodeUTF8(cNome)
	nPos	:= aScan(::aPlanilhas,{|x| x[2]==cNome })
	If nPos>0
		UserException("Esse nome de planilha já foi usado!")
	Endif

	AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\worksheets\sheet"+cValToChar(nQtdPlanilhas+1)+".xml")

	::adimension				:= {{0,0},{999999,999999}}
	::cPlanilhaAt				:= cNome
	::nColunaAtual				:= 0
	::nPriodFormCond			:= 1
	::nRowoutlineLevel			:= nil
	::nTamLinha					:= nil
	::lRowcollapsed				:= .F.
	::lRowHidden				:= .F.

	//Cria nova planilha
	nQtdPlanilhas++
	::xls_sheet(,"sheet"+cValToChar(nQtdPlanilhas)+".xml")
	// ::asheet[nQtdPlanilhas][1]:XPathAddNode( "/xmlns:worksheet", "sheetPr", "" )
	::asheet[nQtdPlanilhas][1]:XPathAddAtt( "/xmlns:worksheet/xmlns:sheetPr", "codeName"	, cNome)
	::asheet[nQtdPlanilhas][1]:XPathAddNode( "/xmlns:worksheet/xmlns:sheetPr", "tabColor", "" )
	::asheet[nQtdPlanilhas][1]:XPathAddAtt( "/xmlns:worksheet/xmlns:sheetPr/xmlns:tabColor", "auto"	, "0")
	If ValType(cCor)=="C"
		If Len(cCor)==6
			cCor	:= "FF"+cCor
		Endif
		::asheet[nQtdPlanilhas][1]:XPathAddAtt( "/xmlns:worksheet/xmlns:sheetPr/xmlns:tabColor", "rgb"	, cCor)
	Endif
	::asheet[nQtdPlanilhas][1]:XPathAddNode( "/xmlns:worksheet/xmlns:sheetPr", "outlinePr", "" )
	::asheet[nQtdPlanilhas][1]:XPathAddNode( "/xmlns:worksheet/xmlns:sheetPr", "pageSetUpPr", "" )
	::asheet[nQtdPlanilhas][1]:XPathAddAtt( "/xmlns:worksheet/xmlns:sheetPr/xmlns:pageSetUpPr", "fitToPage"	, "1")	//Flag indicating whether the Fit to Page print option is enabled. pag 1675
	::asheet[nQtdPlanilhas][1]:XPathAddNode("/xmlns:worksheet","pageSetup","")
	SetAtrr(::asheet[nQtdPlanilhas][1],"/xmlns:worksheet/xmlns:pageSetup","paperSize","9")
	SetAtrr(::asheet[nQtdPlanilhas][1],"/xmlns:worksheet/xmlns:pageSetup","fitToWidth","1")
	SetAtrr(::asheet[nQtdPlanilhas][1],"/xmlns:worksheet/xmlns:pageSetup","fitToHeight","0")
	SetAtrr(::asheet[nQtdPlanilhas][1],"/xmlns:worksheet/xmlns:pageSetup","orientation",::cPagOrientation)


	//Adiciona dentro do workbooks o relacionamento na planilha
	cID	:= ::add_rels("\xl\_rels\workbook.xml.rels","http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet","worksheets/sheet"+cValToChar(nQtdPlanilhas)+".xml")

	AADD(::aPlanilhas,{cID,cNome,/*id draw*/,/*drawsID*/,{}/*atable*/,yExcelTag():New("tableParts",{},,self)/*tableParts*/})
	::nPlanilhaAt	:= nQtdPlanilhas
	::SetFooter("TOTVS","","Página &P/&N")

	::oworkbook:XPathAddNode( "/xmlns:workbook/xmlns:sheets", "sheet", "" )
	::oworkbook:XPathAddAtt( "/xmlns:workbook/xmlns:sheets/xmlns:sheet[last()]", "name"		, cNome)
	::oworkbook:XPathAddAtt( "/xmlns:workbook/xmlns:sheets/xmlns:sheet[last()]", "sheetId"	, cValToChar(nQtdPlanilhas))
	::oworkbook:XPathAddAtt( "/xmlns:workbook/xmlns:sheets/xmlns:sheet[last()]", "r:id"		, cID)

	//Adiciona um nova Planilha no content_types
	::ocontent_types:XPathAddNode( "/xmlns:Types", "Override", "" )
	::ocontent_types:XPathAddAtt( "/xmlns:Types/xmlns:Override[last()]", "PartName"	, "/xl/worksheets/sheet"+cValToChar(nQtdPlanilhas)+".xml" )
	::ocontent_types:XPathAddAtt( "/xmlns:Types/xmlns:Override[last()]", "ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" )
Return nQtdPlanilhas

Method LerPasta(cCaminho,cCamIni,cSufFiltro) Class YExcel
	Local nCont
	Local aFiles	:= Directory(cCaminho+"\*","HSD",,.F.)
	Local lDelete	:= .F.
	Local cCamSheet,cArqSheet
	Local cName
	Local cID
	Local nCont2
	Local nContLn
	Local nContCol
	Local cStyle
	Local cValor
	Local cNumero,fNumero
	Local cDecimal
	Local nPosE,nPosPonto
	Local nFator,fFator
	Local cTargetDraw
	Local cTarget
	// Local cIDTable
	// Local nCont3
	Default cCamIni	:= cCaminho
	For nCont:=1 to Len(aFiles)
		If aFiles[nCont][1]=="." .or. aFiles[nCont][1]==".."
			Loop
		Endif
		If aFiles[nCont][5] $ "D"
			FRename(cCaminho+"\"+aFiles[nCont][1],cCaminho+"\"+lower(aFiles[nCont][1]),,.F.)
			aFiles[nCont][1]	:= lower(aFiles[nCont][1])
			::LerPasta(cCaminho+"\"+aFiles[nCont][1],cCamIni,cSufFiltro)
		Else
			If !Empty(cSufFiltro) .AND. !(lower(right(cCaminho+"\"+aFiles[nCont][1],Len(cSufFiltro)))==lower(cSufFiltro))
				Loop
			Endif
			lDelete	:= .F.
			If lower(aFiles[nCont][1])=="app.xml"
				FRename(cCaminho+"\"+aFiles[nCont][1],cCaminho+"\"+lower(aFiles[nCont][1]),,.F.)
				aFiles[nCont][1]	:= lower(aFiles[nCont][1])
				::new_app(cCaminho+"\"+aFiles[nCont][1])
				lDelete	:= .T.
			ElseIf lower(aFiles[nCont][1])=="core.xml"
				FRename(cCaminho+"\"+aFiles[nCont][1],cCaminho+"\"+lower(aFiles[nCont][1]),,.F.)
				aFiles[nCont][1]	:= lower(aFiles[nCont][1])
				::new_core(cCaminho+"\"+aFiles[nCont][1])
				lDelete	:= .T.
			ElseIf lower(aFiles[nCont][1])=="workbook.xml"
				FRename(cCaminho+"\"+aFiles[nCont][1],cCaminho+"\"+lower(aFiles[nCont][1]),,.F.)
				aFiles[nCont][1]	:= lower(aFiles[nCont][1])
				::new_workbook(cCaminho+"\"+aFiles[nCont][1])
				lDelete	:= .T.
				::new_rels(cCaminho+"\_rels\workbook.xml.rels",+"\xl\_rels\workbook.xml.rels")
				For nCont2:=1 to ::oworkbook:XPathChildCount("/xmlns:workbook/xmlns:sheets")
					// nQtdPlanilhas++
					cName			:= ::oworkbook:XPathGetAtt("/xmlns:workbook/xmlns:sheets/xmlns:sheet["+cValToChar(nCont2)+"]","name")
					cID				:= ::oworkbook:XPathGetAtt("/xmlns:workbook/xmlns:sheets/xmlns:sheet["+cValToChar(nCont2)+"]","id")
					cCamSheet		:= Replace(::Get_rels(cCaminho+"\_rels\workbook.xml.rels",cID,"Target"),"/","\")
					cArqSheet		:= SubStr(cCamSheet,Rat("\",cCamSheet)+1)
					::xls_sheet(cCaminho+"\"+cCamSheet,cArqSheet)
					fErase(cCaminho+"\"+cCamSheet)

					AADD(::aPlanilhas,{cID,cName,/*id draw*/,/*drawsID*/,{}/*atable*/,yExcelTag():New("tableParts",{})/*tableParts*/})
					AADD(::aFiles,cCaminho+"\"+cCamSheet)
					::nPlanilhaAt	:= Len(::aPlanilhas)
					SetAtrr(::asheet[::nPlanilhaAt][1],"/xmlns:worksheet/xmlns:sheetPr", "codeName"	, cName)

					For nContLn:=1 to ::asheet[::nPlanilhaAt][1]:XPathChildCount("/xmlns:worksheet/xmlns:sheetData")
						::nLinha	:= Val(::asheet[::nPlanilhaAt][1]:XPathGetAtt("/xmlns:worksheet/xmlns:sheetData/xmlns:row["+cValToChar(nContLn)+"]", "r"))
						(::cAliasLin)->(RecLock(::cAliasLin,.T.))
						(::cAliasLin)->PLA		:= ::nPlanilhaAt
						(::cAliasLin)->LIN		:= ::nLinha
						(::cAliasLin)->OLEVEL	:= ::asheet[::nPlanilhaAt][1]:XPathGetAtt("/xmlns:worksheet/xmlns:sheetData/xmlns:row["+cValToChar(nContLn)+"]", "outlineLevel")
						(::cAliasLin)->COLLAP	:= ::asheet[::nPlanilhaAt][1]:XPathGetAtt("/xmlns:worksheet/xmlns:sheetData/xmlns:row["+cValToChar(nContLn)+"]", "collapsed")
						(::cAliasLin)->CHIDDEN	:= ::asheet[::nPlanilhaAt][1]:XPathGetAtt("/xmlns:worksheet/xmlns:sheetData/xmlns:row["+cValToChar(nContLn)+"]", "hidden")
						(::cAliasLin)->CHEIGHT	:= ::asheet[::nPlanilhaAt][1]:XPathGetAtt("/xmlns:worksheet/xmlns:sheetData/xmlns:row["+cValToChar(nContLn)+"]", "customHeight")
						(::cAliasLin)->HT		:= Val(::asheet[::nPlanilhaAt][1]:XPathGetAtt("/xmlns:worksheet/xmlns:sheetData/xmlns:row["+cValToChar(nContLn)+"]", "ht"))
						(::cAliasLin)->(MsUnLock())
						For nContCol:=1 to ::asheet[::nPlanilhaAt][1]:XPathChildCount("/xmlns:worksheet/xmlns:sheetData/xmlns:row["+cValToChar(nContLn)+"]")
							::nColuna	:= ::LocRef(::asheet[::nPlanilhaAt][1]:XPathGetAtt("/xmlns:worksheet/xmlns:sheetData/xmlns:row["+cValToChar(nContLn)+"]/xmlns:c["+cValToChar(nContCol)+"]", "r"))[2]
							cStyle	:= ::asheet[::nPlanilhaAt][1]:XPathGetAtt("/xmlns:worksheet/xmlns:sheetData/xmlns:row["+cValToChar(nContLn)+"]/xmlns:c["+cValToChar(nContCol)+"]", "s")
							cTipoCol:= ::asheet[::nPlanilhaAt][1]:XPathGetAtt("/xmlns:worksheet/xmlns:sheetData/xmlns:row["+cValToChar(nContLn)+"]/xmlns:c["+cValToChar(nContCol)+"]", "t")
							cValor	:= ::asheet[::nPlanilhaAt][1]:XPathGetNodeValue("/xmlns:worksheet/xmlns:sheetData/xmlns:row["+cValToChar(nContLn)+"]/xmlns:c["+cValToChar(nContCol)+"]/xmlns:v")
							(::cAliasCol)->(RecLock(::cAliasCol,.T.))
							(::cAliasCol)->PLA		:= ::nPlanilhaAt
							(::cAliasCol)->LIN		:= ::nLinha
							(::cAliasCol)->COL		:= ::nColuna
							(::cAliasCol)->TIPO		:= cTipoCol
							If Empty(cStyle)
								(::cAliasCol)->STY		:= -1
							Else
								(::cAliasCol)->STY		:= Val(cStyle)
							Endif

							(::cAliasCol)->TPVLR	:= "N"	//Numeros
							(::cAliasCol)->TPSTY	:= " "
							If Empty(cValor)
								(::cAliasCol)->TPVLR		:= "U"
							ElseIf cTipoCol=="s"
								(::cAliasCol)->TPSTY		:= "S"
								(self:cAliasCol)->VLRNUM	:= Val(cValor)
							ElseIf cTipoCol=="b"
								(::cAliasCol)->TPSTY	:= "B"
								(self:cAliasCol)->VLRNUM	:= Val(cValor)
							ElseIf cTipoCol=="d"	//date and time UTF
								(::cAliasCol)->TPSTY		:= "D"
								(::cAliasCol)->VLRTXT		:= cValor
								(::cAliasCol)->TPVLR		:= "C"
							ElseIf cTipoCol==""
								(::cAliasCol)->TPSTY		:= "N"
								If "E" $ cValor
									nPosE	:= At("E",cValor)
									cNumero	:= SubStr(cValor,1,nPosE-1)
									nFator	:= Val(SubStr(cValor,nPosE+2))
									fNumero	:= DEC_CREATE(cNumero,21,20)
									fFator	:= DEC_CREATE("1"+Replicate("0",nFator),21,20)
									If "E-" $ cValor
										fNumero	:= DEC_DIV(fNumero,fFator)
									Else
										fNumero	:= DEC_MUL(fNumero,fFator)
									Endif
									(self:cAliasCol)->VLRNUM	:= DEC_TO_DBL(fNumero)	//&(Replace(cValor,"E-","/(10^")+")")
									(self:cAliasCol)->VLRDEC	:= Int(DEC_TO_DBL( DEC_MUL(DEC_CREATE("1000000000000000",21,20),DEC_SUB(fNumero,DEC_RESCALE(fNumero,8,2))) ))
								Else
									nPosPonto	:= At(".",cValor)
									If nPosPonto>0
										cDecimal:= SubStr(cValor,nPosPonto+1)
										If Len(cDecimal)>8
											(self:cAliasCol)->VLRDEC	:= Val(SubStr(cDecimal,9))
										Endif
									Endif
									(self:cAliasCol)->VLRNUM	:= Val(cValor)
								Endif
							Else
								(::cAliasCol)->TPVLR	:= "C"
								(::cAliasCol)->VLRTXT	:= cValor
							Endif

							If ::asheet[::nPlanilhaAt][1]:XPathHasNode("/xmlns:worksheet/xmlns:sheetData/xmlns:row["+cValToChar(nContLn)+"]/xmlns:c["+cValToChar(nContCol)+"]/xmlns:f")
								(::cAliasCol)->FORMULA	:= ::asheet[::nPlanilhaAt][1]:XPathGetNodeValue("/xmlns:worksheet/xmlns:sheetData/xmlns:row["+cValToChar(nContLn)+"]/xmlns:c["+cValToChar(nContCol)+"]/xmlns:f")
							Endif
							(::cAliasCol)->(MsUnLock())
						Next
					Next
					//Deleta as linhas adicionadas ao banco de dados
					While ::asheet[::nPlanilhaAt][1]:XPathHasNode( "/xmlns:worksheet/xmlns:sheetData/xmlns:row[1]")
						::asheet[::nPlanilhaAt][1]:XPathDelNode("/xmlns:worksheet/xmlns:sheetData/xmlns:row[1]")
					EndDo
					If ::asheet[::nPlanilhaAt][1]:XPathHasNode("/xmlns:worksheet/xmlns:drawing")
						::aPlanilhas[::nPlanilhaAt][3]	:= ::asheet[::nPlanilhaAt][1]:XPathGetAtt("/xmlns:worksheet/xmlns:drawing","id")
						cTargetDraw	:= ::FindRels("\xl\worksheets\_rels\"+cArqSheet+".rels","Target",;
							::aPlanilhas[::nPlanilhaAt][3],)
						//::Get_rels(cCaminho+"\worksheets\_rels\"+cArqSheet+".rels",::aPlanilhas[::nPlanilhaAt][3],"Target")
						cTargetDraw	:= SubStr(cTargetDraw,RAt("/drawing",cTargetDraw)+8)
						cTargetDraw	:= SubStr(cTargetDraw,1,Len(cTargetDraw)-4)
						::aPlanilhas[::nPlanilhaAt][4]	:= Val(cTargetDraw)
						AADD(::aworkdrawing,::aPlanilhas[::nPlanilhaAt][4])	//Cria o arquivo
						::new_draw(cCaminho+"\drawings\drawing"+cValToChar(::aPlanilhas[::nPlanilhaAt][4])+".xml","\xl\drawings\drawing"+cValToChar(::aPlanilhas[::nPlanilhaAt][4])+".xml")
						fErase(cCaminho+"\drawings\drawing"+cValToChar(::aPlanilhas[::nPlanilhaAt][4])+".xml")
					Endif
					cTarget	:= ::FindRels("\xl\worksheets\_rels\"+cArqSheet+".rels","Target",,"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments")
					If !Empty(cTarget)
						cTarget	:= "/tmpxls/"+::cTmpFile+"/"+::cNomeFile+"/xl"+Replace(cTarget,"..","")
						::asheet[::nPlanilhaAt][3]	:= ::new_comment(cTarget)
						::asheet[::nPlanilhaAt][4]	:= cTarget
						fErase(cTarget)
					EndIf
					If ::asheet[::nPlanilhaAt][1]:XPathHasNode( "/xmlns:worksheet/xmlns:legacyDrawing")
						cTarget	:= ::FindRels("\xl\worksheets\_rels\"+cArqSheet+".rels","Target",;
							::asheet[::nPlanilhaAt][1]:XPathGetAtt( "/xmlns:worksheet/xmlns:legacyDrawing","id"),)
						cTarget	:= "/tmpxls/"+::cTmpFile+"/"+::cNomeFile+"/xl"+Replace(cTarget,"..","")
						::asheet[::nPlanilhaAt][5]	:= ::new_vmlDrawing(cTarget)
						::asheet[::nPlanilhaAt][6]	:= cTarget
						fErase(cTarget)
					EndIf
					
					// If ::asheet[::nPlanilhaAt][1]:XPathHasNode("/xmlns:worksheet/xmlns:tableParts")
					// 	For nCont3:=1 to ::asheet[::nPlanilhaAt][1]:XPathChildCount("/xmlns:worksheet/xmlns:tableParts")
					// 		cIDTable	:= ::asheet[::nPlanilhaAt][1]:XPathGetAtt("/xmlns:worksheet/xmlns:tableParts/xmlns:tablePart["+cValToChar(nCont3)+"]","id")
					// 		AADD(::aPlanilhas[::nPlanilhaAt][5],cIDTable)
					// 		// cTargetDraw	:= ::Get_rels(cCaminho+"\worksheets\_rels\"+cArqSheet+".rels",cIDTable,"Target")
					// 		// cTargetDraw	:= SubStr(cTargetDraw,RAt("/table",cTargetDraw)+8)
					// 		// cTargetDraw	:= SubStr(cTargetDraw,1,Len(cTargetDraw)-4)
					// 		::aPlanilhas[::nPlanilhaAt][6]:AddValor(yExcelTag():New("tablePart",nil,{{"r:id",cIDTable}},self))
					// 	Next
					// 	::asheet[::nPlanilhaAt][1]:XPathDelNode("/xmlns:worksheet/xmlns:tableParts")
					// Endif
					// If ::asheet[::nPlanilhaAt][1]:XPathHasNode("/xmlns:worksheet/xmlns:tableParts")
					// 	::aPlanilhas[::nPlanilhaAt][5]	:= ::asheet[::nPlanilhaAt][1]:XPathGetAtt("/xmlns:worksheet/xmlns:tableParts","id")
					// 	::aPlanilhas[::nPlanilhaAt][6]	:= YExcelTag():New("tableParts",{})
					// 	::aPlanilhas[::nPlanilhaAt][6]:LoadTagXml(::asheet[::nPlanilhaAt][1],"/xmlns:worksheet/xmlns:tableParts")
					// 	::asheet[::nPlanilhaAt][1]:XPathDelNode("/xmlns:worksheet/xmlns:tableParts")
					// Endif
				Next
				::nPlanilhaAt	:= 1
			ElseIf right(cCaminho,12)=="\xl\drawings".and. right(lower(aFiles[nCont][1]),3)=="vml"
				lDelete	:= .F.
			ElseIf right(cCaminho,12)=="\xl\drawings".and."drawing" $ lower(aFiles[nCont][1])
				lDelete	:= .F.
			ElseIf right(cCaminho,3)=="\xl".and."comments" $ lower(aFiles[nCont][1])
				lDelete	:= .F.
			ElseIf lower(aFiles[nCont][1])=="[content_types].xml"
				FRename(cCaminho+"\"+aFiles[nCont][1],cCaminho+"\"+replace(replace(lower(aFiles[nCont][1]),"[",""),"]",""),,.F.)
				aFiles[nCont][1]	:= replace(replace(lower(aFiles[nCont][1]),"[",""),"]","")
				::new_content_types(cCaminho+"\"+aFiles[nCont][1])
				lDelete	:= .T.
			ElseIf lower(aFiles[nCont][1])=="styles.xml"
				FRename(cCaminho+"\"+aFiles[nCont][1],cCaminho+"\"+lower(aFiles[nCont][1]),,.F.)
				aFiles[nCont][1]	:= lower(aFiles[nCont][1])
				::new_styles(cCaminho+"\"+aFiles[nCont][1])
				lDelete	:= .T.
			ElseIf lower(aFiles[nCont][1])=="sharedstrings.xml"
				FRename(cCaminho+"\"+aFiles[nCont][1],cCaminho+"\"+lower(aFiles[nCont][1]),,.F.)
				aFiles[nCont][1]	:= lower(aFiles[nCont][1])
				::Read_sharedStrings(cCaminho+"\"+aFiles[nCont][1])
				lDelete	:= .T.
			ElseIf right(cCaminho,14)=="\xl\worksheets"
				Loop
				// lDelete	:= .T.
			ElseIf lower(aFiles[nCont][1])=="calcchain.xml"
				lDelete	:= .T.
			ElseIf right(aFiles[nCont][1],17)=="workbook.xml.rels"	//vai ser deletado posteriomente
				lDelete	:= .F.
			ElseIf right(aFiles[nCont][1],5)==".rels"
				FRename(cCaminho+"\"+aFiles[nCont][1],cCaminho+"\"+lower(aFiles[nCont][1]),,.F.)
				aFiles[nCont][1]	:= lower(aFiles[nCont][1])
				::new_rels(cCaminho+"\"+aFiles[nCont][1],Replace(cCaminho,cCamIni,"")+"\"+aFiles[nCont][1])
				lDelete	:= .T.
			ElseIf right(cCaminho,9)=="\xl\media"
				::nIDMedia++
				AADD(::aFiles,cCaminho+"\"+aFiles[nCont][1])
				AADD(::aImagens,{::nIDMedia,aFiles[nCont][1]})
				lDelete	:= .F.
			Else
				AADD(::aFiles,cCaminho+"\"+aFiles[nCont][1])
			Endif
			If lDelete
				If fErase(cCaminho+"\"+aFiles[nCont][1],,.F.)<>0
					ConOut(cCaminho+"\"+aFiles[nCont][1])
					ConOut("Ferror:"+cValToChar(ferror()))
				Endif
			Endif
		Endif
	Next
Return

Static Function LerChvStys(oSelf)
	Local cTipoStyle
	Local nCont,nCont2,nCont3
	Local aCores
	Local aChildren
	Local cChave 	:= ""
	Local cLocal	:= "/xmlns:styleSheet"

	For nCont:=1 to oSelf:oStyle:XPathChildCount(cLocal+"/xmlns:cellXfs")
		cTipoStyle := oSelf:StyleType(nCont-1)
		If cTipoStyle=="D"
			If !DbSqlExec(oSelf:cAliasCol,"UPDATE "+oSelf:cAliasCol+" SET TIPO='d',TPSTY='D' WHERE TPVLR='N' AND STY="+cValToChar(nCont-1),oSelf:cDriver)
				UserException("YExcel - Erro ao inserrir celulas. "+TCSqlError())
			Endif
		ElseIf cTipoStyle=="DT" .or. cTipoStyle=="H"
			If !DbSqlExec(oSelf:cAliasCol,"UPDATE "+oSelf:cAliasCol+" SET TIPO='d',TPSTY='T' WHERE TPVLR='N' AND STY="+cValToChar(nCont-1),oSelf:cDriver)
				UserException("YExcel - Erro ao inserrir celulas. "+TCSqlError())
			Endif
		Endif
		cChave	:= oSelf:oStyle:XPathGetAtt( cLocal+"/xmlns:cellXfs/xmlns:xf["+cValToChar(nCont)+"]", "numFmtId")
		If oSelf:oStyle:XPathGetAtt( cLocal+"/xmlns:cellXfs/xmlns:xf["+cValToChar(nCont)+"]", "applyFont")=="1"
			cChave	+= "|"+oSelf:oStyle:XPathGetAtt( cLocal+"/xmlns:cellXfs/xmlns:xf["+cValToChar(nCont)+"]", "fontId")
		Else
			cChave	+= "|"
		Endif
		If oSelf:oStyle:XPathGetAtt( cLocal+"/xmlns:cellXfs/xmlns:xf["+cValToChar(nCont)+"]", "applyFill")=="1"
			cChave	+= "|"+oSelf:oStyle:XPathGetAtt( cLocal+"/xmlns:cellXfs/xmlns:xf["+cValToChar(nCont)+"]", "fillId")
		Else
			cChave	+= "|"
		Endif
		If oSelf:oStyle:XPathGetAtt( cLocal+"/xmlns:cellXfs/xmlns:xf["+cValToChar(nCont)+"]", "applyBorder")=="1"
			cChave	+= "|"+oSelf:oStyle:XPathGetAtt( cLocal+"/xmlns:cellXfs/xmlns:xf["+cValToChar(nCont)+"]", "borderId")
		Else
			cChave	+= "|"
		Endif
		cChave	+= "|"+oSelf:oStyle:XPathGetAtt( cLocal+"/xmlns:cellXfs/xmlns:xf["+cValToChar(nCont)+"]", "xfId")
		For nCont2:=1 to oSelf:oStyle:XPathChildCount(cLocal+"/xmlns:cellXfs/xmlns:xf["+cValToChar(nCont)+"]")
			If oSelf:oStyle:XPathHasNode( cLocal+"/xmlns:cellXfs/xmlns:xf["+cValToChar(nCont)+"]/xmlns:alignment")
			cChave	+= "{"
			cChave	+= '<alignment'
			aChildren	:= oSelf:oStyle:XPathGetAttArray(cLocal+"/xmlns:cellXfs/xmlns:xf["+cValToChar(nCont)+"]/xmlns:alignment")
			For nCont3:=1 to Len(aChildren)
				cChave	+= " "+aChildren[nCont3][1]+'="'+aChildren[nCont3][2]+'"'
			Next
			cChave	+= "/>"
			cChave	+= "}"
			Endif
		Next
		cChave	+= "|"
		cChave	+= "{"
		aChildren	:= oSelf:oStyle:XPathGetAttArray(cLocal+"/xmlns:cellXfs/xmlns:xf["+cValToChar(nCont)+"]")
		For nCont2:=1 to Len(aChildren)
			If !("|"+aChildren[nCont2][1]+"|"	$ "|applyNumberFormat|numFmtId|applyFont|fontId|applyFill|fillId|applyBorder|borderId|xfId|applyAlignment|" )
				If Right(cChave,1)=="}"
					cChave	+= ","
				Endif
				cChave	+= "{"
				cChave	+= '"'+aChildren[nCont2][1]+'","'+aChildren[nCont2][2]+'"'
				cChave	+= "}"
			Endif
		Next
		cChave	+= "}"
	Next

	For nCont:=1 to oSelf:oStyle:XPathChildCount(cLocal+"/xmlns:fills")
		If oSelf:oStyle:XPathHasNode( cLocal+"/xmlns:fills/xmlns:fill["+cValToChar(nCont)+"]/xmlns:patternFill")
			cChave	:= ""
			cChave	+= oSelf:oStyle:XPathGetAtt( cLocal+"/xmlns:fills/xmlns:fill["+cValToChar(nCont)+"]/xmlns:patternFill/xmlns:fgColor", "rgb")
			cChave	+= oSelf:oStyle:XPathGetAtt( cLocal+"/xmlns:fills/xmlns:fill["+cValToChar(nCont)+"]/xmlns:patternFill/xmlns:fgColor", "indexed")
			cChave	+= "|"
			cChave	+= oSelf:oStyle:XPathGetAtt( cLocal+"/xmlns:fills/xmlns:fill["+cValToChar(nCont)+"]/xmlns:patternFill/xmlns:bgColor", "rgb")
			cChave	+= oSelf:oStyle:XPathGetAtt( cLocal+"/xmlns:fills/xmlns:fill["+cValToChar(nCont)+"]/xmlns:patternFill/xmlns:bgColor", "indexed")
			cChave	+= "|"
			cChave	+= oSelf:oStyle:XPathGetAtt( cLocal+"/xmlns:fills/xmlns:fill["+cValToChar(nCont)+"]/xmlns:patternFill", "patternType")
			cChave	+= "|"
			cChave	+= cLocal+"/xmlns:fills"
			RecLock(oSelf:cAliasChv,.T.)
			(oSelf:cAliasChv)->TIPO		:= "CORPREENC "
			(oSelf:cAliasChv)->CHAVE	:= cChave
			(oSelf:cAliasChv)->ID		:= nCont-1
			MsUnLock()
		ElseIf oSelf:oStyle:XPathHasNode( cLocal+"/xmlns:fill["+cValToChar(nCont)+"]/xmlns:gradientFill")
			aCores	:= {}
			For nCont2:=1 to oSelf:oStyle:XPathChildCount(cLocal+"/xmlns:fills/xmlns:gradientFill["+cValToChar(nCont)+"]")
				AADD(aCores,{;
							oSelf:oStyle:XPathGetAtt( cLocal+"/xmlns:fill["+cValToChar(nCont)+"]/xmlns:gradientFill/xmlns:stop["+cValToChar(nCont2)+"]/xmlns:color","rgb");
							,oSelf:oStyle:XPathGetAtt( cLocal+"/xmlns:fill["+cValToChar(nCont)+"]/xmlns:gradientFill/xmlns:stop["+cValToChar(nCont2)+"]/xmlns:color","position");
							})
			Next
			If oSelf:oStyle:XPathGetAtt( cLocal+"/xmlns:fill["+cValToChar(nCont)+"]/xmlns:gradientFill","type")=="path"
				cChave	:= ""
				cChave	+= "|"+Var2Chr(aCores)
				cChave	+= "|path"
				cChave	+= "|"+oSelf:oStyle:XPathGetAtt( cLocal+"/xmlns:fill["+cValToChar(nCont)+"]/xmlns:gradientFill","left")
				cChave	+= "|"+oSelf:oStyle:XPathGetAtt( cLocal+"/xmlns:fill["+cValToChar(nCont)+"]/xmlns:gradientFill","right")
				cChave	+= "|"+oSelf:oStyle:XPathGetAtt( cLocal+"/xmlns:fill["+cValToChar(nCont)+"]/xmlns:gradientFill","top")
				cChave	+= "|"+oSelf:oStyle:XPathGetAtt( cLocal+"/xmlns:fill["+cValToChar(nCont)+"]/xmlns:gradientFill","bottom")
				cChave	+= "|"+cLocal+"/xmlns:fills"
			ElseIf oSelf:oStyle:XPathGetAtt( cLocal+"/xmlns:fill["+cValToChar(nCont)+"]/xmlns:gradientFill","type")=="linear"
				cChave	:= ""+oSelf:oStyle:XPathGetAtt( cLocal+"/xmlns:fill["+cValToChar(nCont)+"]/xmlns:gradientFill","degree")
				cChave	+= "|"+Var2Chr(aCores)
				cChave	+= "|linear"
				cChave	+= "|"
				cChave	+= "|"
				cChave	+= "|"
				cChave	+= "|"
				cChave	+= "|"+cLocal+"/xmlns:fills"
			Endif
		Endif
	Next
	For nCont:=1 to oSelf:oStyle:XPathChildCount("/xmlns:styleSheet/xmlns:numFmts")
		oSelf:nNumFmtId	:= Max(oSelf:nNumFmtId,Val(oSelf:oStyle:XPathGetAtt("/xmlns:styleSheet/xmlns:numFmts/xmlns:numFmt["+cValToChar(nCont)+"]","numFmtId")))
	Next
	For nCont:=1 to oSelf:oStyle:XPathChildCount(cLocal+"/xmlns:fonts")
		cChave	:= ""
		cChave	+= oSelf:oStyle:XPathGetAtt( cLocal+"/xmlns:fonts/xmlns:font["+cValToChar(nCont)+"]/xmlns:sz", "val")
		cChave	+= "|"
		cChave	+= oSelf:oStyle:XPathGetAtt( cLocal+"/xmlns:fonts/xmlns:font["+cValToChar(nCont)+"]/xmlns:color", "indexed")
		cChave	+= oSelf:oStyle:XPathGetAtt( cLocal+"/xmlns:fonts/xmlns:font["+cValToChar(nCont)+"]/xmlns:color", "rgb")
		cChave	+= "|"
		cChave	+= oSelf:oStyle:XPathGetAtt( cLocal+"/xmlns:fonts/xmlns:font["+cValToChar(nCont)+"]/xmlns:name", "val")
		cChave	+= "|"
		cChave	+= oSelf:oStyle:XPathGetAtt( cLocal+"/xmlns:fonts/xmlns:font["+cValToChar(nCont)+"]/xmlns:family", "val")
		cChave	+= "|"
		cChave	+= oSelf:oStyle:XPathGetAtt( cLocal+"/xmlns:fonts/xmlns:font["+cValToChar(nCont)+"]/xmlns:scheme", "val")
		cChave	+= "|"
		cChave	+= If(oSelf:oStyle:XPathHasNode( cLocal+"/xmlns:fonts/xmlns:font["+cValToChar(nCont)+"]/xmlns:b"),".T.",".F.")
		cChave	+= "|"
		cChave	+= If(oSelf:oStyle:XPathHasNode( cLocal+"/xmlns:fonts/xmlns:font["+cValToChar(nCont)+"]/xmlns:i"),".T.",".F.")
		cChave	+= "|"
		cChave	+= If(oSelf:oStyle:XPathHasNode( cLocal+"/xmlns:fonts/xmlns:font["+cValToChar(nCont)+"]/xmlns:u"),".T.",".F.")
		cChave	+= "|"
		cChave	+= If(oSelf:oStyle:XPathHasNode( cLocal+"/xmlns:fonts/xmlns:font["+cValToChar(nCont)+"]/xmlns:strike"),".T.",".F.")
		cChave	+= "|"
		cChave	+= cLocal+"/xmlns:fonts"
		RecLock(oSelf:cAliasChv,.T.)
		(oSelf:cAliasChv)->TIPO		:= "FONTE     "
		(oSelf:cAliasChv)->CHAVE	:= cChave
		(oSelf:cAliasChv)->ID		:= nCont-1
		MsUnLock()
	Next

Return

METHOD SetPlanName(cNome) Class YExcel
	Local cOldName
	PARAMTYPE 0	VAR cNome			AS CHARACTER		OPTIONAL DEFAULT "Planilha"+cValToChar(nQtdPlanilhas+1)
	cNome	:= Replace(cNome,"\","")
	cNome	:= Replace(cNome,"/","")
	cNome	:= Replace(cNome,":","")
	cNome	:= Replace(cNome,"*","")
	cNome	:= Replace(cNome,"[","")
	cNome	:= Replace(cNome,"]","")
	cNome	:= Replace(cNome,"?","")
	cNome	:= Replace(cNome,">","&gt;")
	cNome	:= Replace(cNome,"<","&lt;")
	If Len(cNome)>31
		cNome	:= SubStr(cNome,1,31)
	Endif
	cNome	:= EncodeUTF8(cNome)
	nPos	:= aScan(::aPlanilhas,{|x| x[2]==cNome })
	If nPos>0
		UserException("Esse nome de planilha já foi usado!")
	Endif
	cOldName						:= ::aPlanilhas[::nPlanilhaAt][2]
	::aPlanilhas[::nPlanilhaAt][2]	:= cNome
	::asheet[::nPlanilhaAt][1]:XPathSetAtt( "/xmlns:worksheet/xmlns:sheetPr", "codeName"	, cNome)
	::oworkbook:XPathSetAtt( "/xmlns:workbook/xmlns:sheets/xmlns:sheet[@name='"+cOldName+"']", "name"		, cNome)
Return

/*/{Protheus.doc} SetPlanAt
Infoma a planilha de alteração
@author Saulo Gomes Martins
@since 13/06/2020
@param xPlan, variadic, (characters|numeric) Indice da planilha ou nome da planilha
@type method
/*/
METHOD SetPlanAt(xPlan) Class YExcel
	Local nPos
	Local lOk	:= .T.
	If ValType(xPlan)=="N"
		If xPlan>Len(::aPlanilhas)
			lOk := .F.
		Endif
		::nPlanilhaAt	:= xPlan
		::cPlanilhaAt	:= ::aPlanilhas[::nPlanilhaAt][2]
	elseif ValType(xPlan)=="C"
		nPos	:= aScan(::aPlanilhas,{|x| x[2]==xPlan })
		If nPos==0
			lOk := .F.
		Else
			::nPlanilhaAt	:= nPos
			::cPlanilhaAt	:= ::aPlanilhas[::nPlanilhaAt][2]
		Endif
	Endif
Return lOk
/*/{Protheus.doc} GetPlanAt
Retorna Indice da planilha ou nome da planilha
@author Saulo Gomes Martins
@since 13/06/2020
@param cRet, characters, 1=Indice da planilha | 2=nome da planilha
@type method
/*/
METHOD GetPlanAt(cRet) Class YExcel
	Default cRet := "1"
Return If(cRet=="1",::nPlanilhaAt,::aPlanilhas[::nPlanilhaAt][2])

/*/{Protheus.doc} LenPlanAt
Quantidade de planilha
@author Saulo Gomes Martins
@since 13/06/2020
@type method
/*/
METHOD LenPlanAt(xPlan) Class YExcel
Return Len(::aPlanilhas)

/*/{Protheus.doc} AddNome
Cria nome para refencia de célula ou intervalo
@author Saulo Gomes Martins
@since 09/05/2017
@param cNome, characters, Nome
@param nLinha, numeric, Linha da referencia
@param nColuna, numeric, Coluna da referencia
@param nLinha2, numeric, (opcional) Linha final se intervalo
@param nColuna2, numeric, (opcional) Coluna final se intervalo
@param cRefPar, characters, (opcional) Rerefencia
@param cPlanilha, characters, (opcional) Planilha
@param cEscopo, characters, (opcional) Planilha de escopo
@type method
/*/
METHOD AddNome(cNome,nLinha,nColuna,nLinha2,nColuna2,cRefPar,cPlanilha,cEscopo) Class YExcel
	Local cRef			:= ""
	Local nPos			:= 0
	PARAMTYPE 0	VAR cNome			AS CHARACTER
	PARAMTYPE 1	VAR nLinha			AS NUMERIC			OPTIONAL
	PARAMTYPE 2	VAR nColuna			AS NUMERIC			OPTIONAL
	PARAMTYPE 3	VAR nLinha2			AS NUMERIC			OPTIONAL
	PARAMTYPE 4	VAR nColuna2		AS NUMERIC			OPTIONAL
	PARAMTYPE 5	VAR cRefPar			AS CHARACTER		OPTIONAL
	PARAMTYPE 6	VAR cPlanilha		AS CHARACTER		OPTIONAL DEFAULT ::cPlanilhaAt
	PARAMTYPE 7	VAR cEscopo			AS CHARACTER		OPTIONAL

	If ValType(cRefPar)=="U"
		If !Empty(cPlanilha)
			cRef	:= "'"+cPlanilha+"'!"
		Endif
		cRef	+= ::Ref(nLinha,nColuna,.T.,.T.)
		If Valtype(nLinha2)<>"U" .OR. Valtype(nColuna2)<>"U"
			cRef	+= ":"+::Ref(nLinha2,nColuna2,.T.,.T.)
		Endif
	Else
		cRef	:= cRefPar
	Endif
	If ValType(cEscopo)=="C"
		nPos	:= aScan(::aPlanilhas,{|x| x[2]==cEscopo })
	Endif
	::oworkbook:XPathAddNode( "/xmlns:workbook/xmlns:definedNames"						, "definedName"			, cRef )
	::oworkbook:XPathAddAtt( "/xmlns:workbook/xmlns:definedNames/xmlns:definedName[last()]", "name"				, cNome)
	If nPos>0
		::oworkbook:XPathAddAtt( "/xmlns:workbook/xmlns:definedNames/xmlns:definedName[last()]", "localSheetId"		, cValToChar(nPos-1))
	Endif
Return

/*/{Protheus.doc} SetPrintTitles
Repetir linhas na impressão
@author Saulo Gomes Martins
@since 12/12/2019
@version 1.0
@param nLinha, numeric, Linha inicial
@param nLinha2, numeric, Linha final
@param cRefPar, characters, Referencia
@param cPlanilha, characters, Planilha
@type method
@obs pag 1566
/*/
METHOD SetPrintTitles(nLinha,nLinha2,cRefPar,cPlanilha) Class YExcel
	Default nLinha2	:= nLinha
	::AddNome("_xlnm.Print_Titles",nLinha,,nLinha2,,cRefPar,cPlanilha,::cPlanilhaAt)
Return

/*/{Protheus.doc} SetPagOrientation
Informa a orientação do papel na impressão
@author Saulo Gomes Martins
@since 12/12/2019
@version 1.0
@param cOrientation, characters, descricao
@type method
@obs pag 1667
/*/
METHOD SetPagOrientation(cOrientation) Class YExcel
	Default cOrientation := "default"
	If lower(cOrientation)+"|" $ "default|landscape|portrait|"
		::cPagOrientation	:= cOrientation	//Proximas planilhas segue mesma orientação
	Endif
	If !::asheet[::nPlanilhaAt][1]:XPathHasNode("/xmlns:worksheet/xmlns:pageSetup")
		::asheet[::nPlanilhaAt][1]:XPathAddNode("/xmlns:worksheet","pageSetup","")
	Endif
	SetAtrr(::asheet[::nPlanilhaAt][1],"/xmlns:worksheet/xmlns:pageSetup","orientation",::cPagOrientation)
Return ::cPagOrientation
/*/{Protheus.doc} YExcel::SetHeader
Configurar o Cabeçalho
@type method
@version 1.0
@author Saulo Gomes Martins
@since 17/03/2021
@param cLeft, character, Texto seção da esquerda
@param cCenter, character, Texto seção central
@param cRight, character, Texto seção direita 
/*/
METHOD SetHeader(cLeft,cCenter,cRight) Class YExcel
	Local cValor	:= ""
	If !::asheet[::nPlanilhaAt][1]:XPathHasNode("/xmlns:worksheet/xmlns:headerFooter")
		::asheet[::nPlanilhaAt][1]:XPathAddNode("/xmlns:worksheet","headerFooter","")
	Endif
	If !Empty(cLeft)
		cValor	+= "&L"+cLeft
	Endif
	If !Empty(cCenter)
		cValor	+= "&C"+cCenter
	Endif
	If !Empty(cRight)
		cValor	+= "&R"+cRight
	Endif
	cValor	:= EncodeUTF8(Replace(cValor,"&","&amp;"))
	If !::asheet[::nPlanilhaAt][1]:XPathHasNode("/xmlns:worksheet/xmlns:headerFooter/xmlns:oddHeader")
		::asheet[::nPlanilhaAt][1]:XPathAddNode("/xmlns:worksheet/xmlns:headerFooter","oddHeader",cValor)
	Else
		::asheet[::nPlanilhaAt][1]:XPathSetNode("/xmlns:worksheet/xmlns:headerFooter/xmlns:oddHeader","oddHeader",cValor)
	EndIf
Return
/*/{Protheus.doc} YExcel::SetFooter
Configurar o Rodapé
@type method
@version 1.0
@author Saulo Gomes Martins
@since 17/03/2021
@param cLeft, character, Texto seção da esquerda
@param cCenter, character, Texto seção central
@param cRight, character, Texto seção direita 
/*/
METHOD SetFooter(cLeft,cCenter,cRight) Class YExcel
	Local cValor	:= ""
	If !::asheet[::nPlanilhaAt][1]:XPathHasNode("/xmlns:worksheet/xmlns:headerFooter")
		::asheet[::nPlanilhaAt][1]:XPathAddNode("/xmlns:worksheet","headerFooter","")
	Endif
	If !Empty(cLeft)
		cValor	+= "&L"+cLeft
	Endif
	If !Empty(cCenter)
		cValor	+= "&C"+cCenter
	Endif
	If !Empty(cRight)
		cValor	+= "&R"+cRight
	Endif
	cValor	:= EncodeUTF8(Replace(cValor,"&","&amp;"))
	If !::asheet[::nPlanilhaAt][1]:XPathHasNode("/xmlns:worksheet/xmlns:headerFooter/xmlns:oddFooter")
		::asheet[::nPlanilhaAt][1]:XPathAddNode("/xmlns:worksheet/xmlns:headerFooter","oddFooter",cValor)
	Else
		::asheet[::nPlanilhaAt][1]:XPathSetNode("/xmlns:worksheet/xmlns:headerFooter/xmlns:oddFooter","oddFooter",cValor)
	EndIf
Return

/*/{Protheus.doc} ADDImg
Adiciona imagem para ser usado
@author Saulo Gomes Martins
@since 06/01/2019
@version 1.0
@return numeric, ID da imagem
@param cImg, characters, Localização da imagem
@type method
/*/
METHOD ADDImg(cImg) Class YExcel
	Local cDrive, cDir, cNome, cExt
	Local cDirImg	:= "\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\media\"
	PARAMTYPE 0	VAR cImg		AS CHARACTER

	If !File(cImg,,.F.)
		UserException("YExcel - Imagem não encontrada ("+cImg+")")
	Endif

	::nIDMedia++
	FWMakeDir(cDirImg,.F.)
	SplitPath( cImg, @cDrive, @cDir, @cNome, @cExt)
	cNome	:= SubStr(cImg,Rat("\",cImg)+1)
	If ":" $ UPPER(cImg)
		CpyT2S(cImg,cDirImg,,.F.)
		FRename(cDirImg+cNome,cDirImg+"image"+cValToChar(::nIDMedia)+cExt,,.F.)
	Else
		__COPYFILE(cImg,cDirImg+"image"+cValToChar(::nIDMedia)+cExt,,,.F.)
	Endif
	AADD(::aFiles,cDirImg+"image"+cValToChar(::nIDMedia)+cExt)
	AADD(::aImagens,{::nIDMedia,"image"+cValToChar(::nIDMedia)+cExt})
	cExt	:= replace(lower(cExt),".","")
	If !::ocontent_types:XPathHasNode("/xmlns:Types/xmlns:Default[@Extension='"+cExt+"']")
		::ocontent_types:XPathAddNode("/xmlns:Types","Default","")
		::ocontent_types:XPathAddAtt("/xmlns:Types/xmlns:Default[last()]","Extension",cExt)
		if cExt=="jpg"
			cExt	:= "jpeg"
		Endif
		::ocontent_types:XPathAddAtt("/xmlns:Types/xmlns:Default[last()]","ContentType","image/"+cExt)
	Endif
Return ::nIDMedia

/*/{Protheus.doc} Img
Usa imagem
@author Saulo Gomes Martins
@since 06/01/2019
@version 1.0
@param nID, numeric, ID da imagem
@param nLinha, numeric, Linha para adicionar a imagem
@param nColuna, numeric, Coluna para adicionar a imagem
@param nX, numeric, Largura da imagem
@param nY, numeric, Altura da imagem
@param cUnidade, characters, (opcional) Unidade da dimensão da imagem. padrão em pixel
@param nRot, numeric, rotação da imagem
@type method
@OBS pag 3166
/*/
METHOD Img(nID,nLinha,nColuna,nX,nY,cUnidade,nRot) Class YExcel
	Local nPos
	Local cCellType
	Local cID
	Local cIdDraw
	PARAMTYPE 0	VAR nID			AS NUMERIC
	PARAMTYPE 1	VAR nLinha		AS NUMERIC		OPTIONAL DEFAULT ::nLinha
	PARAMTYPE 2	VAR nColuna		AS NUMERIC		OPTIONAL DEFAULT ::nColuna
	PARAMTYPE 3	VAR nY			AS NUMERIC
	PARAMTYPE 4	VAR nX			AS NUMERIC
	PARAMTYPE 5	VAR cUnidade	AS CHARACTER	OPTIONAL DEFAULT "px"
	PARAMTYPE 6	VAR nRot		AS NUMERIC		OPTIONAL DEFAULT 0

	If aScan(::aImagens,{|x| x[1]==nID })==0
		UserException("YExcel - Imagem não cadastrada, usar metodo ADDImg. ID("+cValToChar(nID)+")")
	Endif

	cUnidade	:= lower(cUnidade)
	//Converte para  EMUs (English Metric Units)
	If cUnidade=="px"
		nX	:= nX*36000*0.2645833333
		nY	:= nY*36000*0.2645833333
	ElseIf cUnidade=="cm"
		nX	:= nX*36000
		nY	:= nY*36000
	Endif
	Default cCellType	:= "oneCellAnchor"
	//absolute	- Não mover ou redimensionar com linhas / colunas subjacentes
	//oneCell	- Mova-se com células, mas não redimensione
	//twoCell	- Mover e redimensionar com células âncoras

	If !::asheet[::nPlanilhaAt][1]:XPathHasNode("/xmlns:worksheet/xmlns:drawing")
		::asheet[::nPlanilhaAt][1]:XPathAddNode("/xmlns:worksheet","drawing","")
		::nIdRelat++
		nPos	:= ::nIdRelat
		cID		:= ::add_rels("\xl\worksheets\_rels\sheet"+cValToChar(::nPlanilhaAt)+".xml.rels","http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing","../drawings/drawing"+cValToChar(nPos)+".xml")
		::aPlanilhas[::nPlanilhaAt][3]	:= ::new_draw(,"\xl\drawings\drawing"+cValToChar(nPos)+".xml")	//Cria o arquivo \xl\drawings\drawing1
		::asheet[::nPlanilhaAt][1]:XPathAddAtt("/xmlns:worksheet/xmlns:drawing","r:id",cID)
		AADD(::aworkdrawing,nPos)	//Cria o arquivo
		::aPlanilhas[::nPlanilhaAt][4]	:= nPos
		//Adiciona um nova drawing no content_types
		::ocontent_types:XPathAddNode( "/xmlns:Types", "Override", "" )
		::ocontent_types:XPathAddAtt( "/xmlns:Types/xmlns:Override[last()]", "PartName"	, "/xl/drawings/drawing"+cValToChar(Len(::aworkdrawing))+".xml" )
		::ocontent_types:XPathAddAtt( "/xmlns:Types/xmlns:Override[last()]", "ContentType", "application/vnd.openxmlformats-officedocument.drawing+xml" )
	EndIf
	// If Empty(::aPlanilhas[::nPlanilhaAt][4])
	// 	::nIdRelat++
	// 	nPos	:= ::nIdRelat
	// 	cID		:= ::add_rels("\xl\worksheets\_rels\sheet"+cValToChar(::nPlanilhaAt)+".xml.rels","http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing","../drawings/drawing"+cValToChar(nPos)+".xml")

	// 	::aPlanilhas[::nPlanilhaAt][3]	:= ::new_draw(,"\xl\drawings\drawing"+cValToChar(nPos)+".xml")	//Cria o arquivo \xl\drawings\drawing1

	// 	::aPlanilhas[::nPlanilhaAt][4] := yExcelTag():New("drawing",,,self)		
	// 	::aPlanilhas[::nPlanilhaAt][4]:SetAtributo("r:id",cID)
	// 	::aPlanilhas[::nPlanilhaAt][4]:xDados	:= nPos
	// 	AADD(::aworkdrawing,nPos)	//Cria o arquivo
	// 	//Adiciona um nova drawing no content_types
	// 	::ocontent_types:XPathAddNode( "/xmlns:Types", "Override", "" )
	// 	::ocontent_types:XPathAddAtt( "/xmlns:Types/xmlns:Override[last()]", "PartName"	, "/xl/drawings/drawing"+cValToChar(Len(::aworkdrawing))+".xml" )
	// 	::ocontent_types:XPathAddAtt( "/xmlns:Types/xmlns:Override[last()]", "ContentType", "application/vnd.openxmlformats-officedocument.drawing+xml" )
	// Endif
	nPos	:= ::aPlanilhas[::nPlanilhaAt][3]
	cIdDraw	:= ::add_rels("\xl\drawings\_rels\drawing"+cValToChar(::aPlanilhas[::nPlanilhaAt][4])+".xml.rels","http://schemas.openxmlformats.org/officeDocument/2006/relationships/image","../media/"+::aImagens[nID][2])

	::aDraw[nPos][1]:XPathAddNode( "/xdr:wsDr", cCellType, "" )
	If cCellType!="oneCellAnchor"
		::aDraw[nPos][1]:XPathAddAtt( "/xdr:wsDr/xdr:"+cCellType+"[last()]", "editAs"	, "oneCell" )
	EndIf

	::aDraw[nPos][1]:XPathAddNode( "/xdr:wsDr/xdr:"+cCellType+"[last()]", "from", "" )
	::aDraw[nPos][1]:XPathAddNode( "/xdr:wsDr/xdr:"+cCellType+"[last()]/xdr:from", "col", cValToChar(nColuna-1) )
	::aDraw[nPos][1]:XPathAddNode( "/xdr:wsDr/xdr:"+cCellType+"[last()]/xdr:from", "colOff", cValToChar(0) )
	::aDraw[nPos][1]:XPathAddNode( "/xdr:wsDr/xdr:"+cCellType+"[last()]/xdr:from", "row", cValToChar(nLinha-1) )
	::aDraw[nPos][1]:XPathAddNode( "/xdr:wsDr/xdr:"+cCellType+"[last()]/xdr:from", "rowOff", cValToChar(0) )


	::aDraw[nPos][1]:XPathAddNode( "/xdr:wsDr/xdr:"+cCellType+"[last()]", "ext", "" )
	::aDraw[nPos][1]:XPathAddAtt( "/xdr:wsDr/xdr:"+cCellType+"[last()]/xdr:ext", "cx"	, cValToChar(Round(nX,0)) )
	::aDraw[nPos][1]:XPathAddAtt( "/xdr:wsDr/xdr:"+cCellType+"[last()]/xdr:ext", "cy"	, cValToChar(Round(nY,0)) )

	::aDraw[nPos][1]:XPathAddNode(	"/xdr:wsDr/xdr:"+cCellType+"[last()]", "pic", "" )
	//nvPicPr
	::aDraw[nPos][1]:XPathAddNode(	"/xdr:wsDr/xdr:"+cCellType+"[last()]/xdr:pic", "nvPicPr", "" )
	//cNvPr
	::aDraw[nPos][1]:XPathAddNode(	"/xdr:wsDr/xdr:"+cCellType+"[last()]/xdr:pic/xdr:nvPicPr", "cNvPr", "" )
	::aDraw[nPos][1]:XPathAddAtt(	"/xdr:wsDr/xdr:"+cCellType+"[last()]/xdr:pic/xdr:nvPicPr/xdr:cNvPr","id", cValToChar(nID) )
	::aDraw[nPos][1]:XPathAddAtt(	"/xdr:wsDr/xdr:"+cCellType+"[last()]/xdr:pic/xdr:nvPicPr/xdr:cNvPr","name", "Imagem "+cValToChar(nID) )

	//cNvPicPr
	::aDraw[nPos][1]:XPathAddNode(	"/xdr:wsDr/xdr:"+cCellType+"[last()]/xdr:pic/xdr:nvPicPr", "cNvPicPr", "" )
	::aDraw[nPos][1]:XPathAddNode(	"/xdr:wsDr/xdr:"+cCellType+"[last()]/xdr:pic/xdr:nvPicPr/xdr:cNvPicPr", "a:picLocks", "" )
	ajustNS(::aDraw[nPos][1],"<xdr:a:","<a:")
	::aDraw[nPos][1]:XPathAddAtt(	"/xdr:wsDr/xdr:"+cCellType+"[last()]/xdr:pic/xdr:nvPicPr/xdr:cNvPicPr/a:picLocks", "noChangeAspect", "1" )

	//blipFill
	::aDraw[nPos][1]:XPathAddNode(	"/xdr:wsDr/xdr:"+cCellType+"[last()]/xdr:pic", "blipFill", "" )
	::aDraw[nPos][1]:XPathAddNode(	"/xdr:wsDr/xdr:"+cCellType+"[last()]/xdr:pic/xdr:blipFill", "a:blip", "" )
	ajustNS(::aDraw[nPos][1],"<xdr:a:","<a:")
	::aDraw[nPos][1]:XPathAddNs(	"/xdr:wsDr/xdr:"+cCellType+"[last()]/xdr:pic/xdr:blipFill/a:blip", "r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships" )
	::aDraw[nPos][1]:XPathAddAtt(	"/xdr:wsDr/xdr:"+cCellType+"[last()]/xdr:pic/xdr:blipFill/a:blip", "r:embed", cIdDraw )
	::aDraw[nPos][1]:XPathAddNode(	"/xdr:wsDr/xdr:"+cCellType+"[last()]/xdr:pic/xdr:blipFill", "a:stretch", "" )
	ajustNS(::aDraw[nPos][1],"<xdr:a:","<a:")
	::aDraw[nPos][1]:XPathAddNode(	"/xdr:wsDr/xdr:"+cCellType+"[last()]/xdr:pic/xdr:blipFill/a:stretch", "fillRect", "" )

	//spPr
	::aDraw[nPos][1]:XPathAddNode(	"/xdr:wsDr/xdr:"+cCellType+"[last()]/xdr:pic", "spPr", "" )
	::aDraw[nPos][1]:XPathAddNode(	"/xdr:wsDr/xdr:"+cCellType+"[last()]/xdr:pic/xdr:spPr", "a:xfrm", "" )
	ajustNS(::aDraw[nPos][1],"<xdr:a:","<a:")
	::aDraw[nPos][1]:XPathAddAtt(	"/xdr:wsDr/xdr:"+cCellType+"[last()]/xdr:pic/xdr:spPr/a:xfrm", "rot", cValToChar(nRot*60000) )
	::aDraw[nPos][1]:XPathAddNode(	"/xdr:wsDr/xdr:"+cCellType+"[last()]/xdr:pic/xdr:spPr", "a:prstGeom", "" )
	ajustNS(::aDraw[nPos][1],"<xdr:a:","<a:")
	::aDraw[nPos][1]:XPathAddAtt(	"/xdr:wsDr/xdr:"+cCellType+"[last()]/xdr:pic/xdr:spPr/a:prstGeom", "prst", "rect" )
	::aDraw[nPos][1]:XPathAddNode(	"/xdr:wsDr/xdr:"+cCellType+"[last()]/xdr:pic/xdr:spPr/a:prstGeom", "avLst", "" )


	::aDraw[nPos][1]:XPathAddNode(	"/xdr:wsDr/xdr:"+cCellType+"[last()]", "clientData", "" )
	::aDraw[nPos][3]++

	AADD(::aImgdraw,Len(::aImgdraw)+1)

Return

/*/{Protheus.doc} Cell
Grava o conteudo de uma célula
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param nLinha, numeric, Linha a ser gravada
@param nColuna, numeric, Coluna a ser gravada
@param xValor, variadic, Valor a ser gravado(texto,numero,data,logico)
@param cFormula, characters, Formula da célula
@param nStyle, numeric, posição do estilo criado pelo metodo :AddStyles()
@type method
/*/
METHOD Cell(nLinha,nColuna,xValor,cFormula,xStyle) Class YExcel
	Local cTipo		:= ValType(xValor)
	Local cTpStyle	:= ValType(xStyle)
	PARAMTYPE 0	VAR nLinha			AS NUMERIC			OPTIONAL DEFAULT ::nLinha
	PARAMTYPE 1	VAR nColuna			AS NUMERIC			OPTIONAL DEFAULT ::nColuna
	PARAMTYPE 3	VAR cFormula		AS CHARACTER		OPTIONAL
	PARAMTYPE 4	VAR xStyle			AS NUMERIC,OBJECT	OPTIONAL

	If ValType(nColuna)=="C"
		//nColuna		:= ::odefinedNames:GetValor(nColuna,""):GetAtributo("name")
		nColuna	:= StringToNum(UPPER(nColuna))
	Endif
	If nColuna==0
		UserException("YExcel - O índice da coluna não pode iniciar no 0")
	Endif

	::nColunaAtual	:= nColuna
	If ::adimension[2][1]>nLinha	//Menor linha
		::adimension[2][1]	:= nLinha
	Endif
	If ::adimension[2][2]>nColuna	//Menor Coluna
		::adimension[2][2]	:= nColuna
	Endif
	If ::adimension[1][1]<nLinha	//Maior Linha
		::adimension[1][1]	:= nLinha
	Endif
	::Pos(nLinha,nColuna)
	::SetValue(xValor,cFormula)
	If cTpStyle=="U".AND.cTipo=="D"
	ElseIf cTpStyle=="U" .AND. cTipo=="O".AND.GetClassName(xValor)=="YEXCEL_DATETIME" 	//Data e Date time, deixa formato padrão, não limpa formato
	Else
		::SetStyle(xStyle,nLinha,nColuna)
	EndIf

	If ::adimension[1][2]<nColuna	//Maior Coluna
		::adimension[1][2]	:= nColuna
	Endif
	
Return self
/*/{Protheus.doc} YExcel::AddCol
Adiciona colunas vazias e move para esquerdas as demais
@type method
@version 1.0
@author Saulo Gomes Martins
@since 06/07/2020
@param nQtd, numeric, Quantidade de colunas a incluir
@param nColunaIni, numeric, Coluna inicial a ser movida para esquerda
@param nLinhaIni, numeric, Linha inicial a ser deslocada
@param nLinhaFim, numeric, Linha final a ser deslocada
/*/
Method AddCol(nQtd,nColunaIni,nLinhaIni,nLinhaFim) Class YExcel
	Local cUpdate	:= ""
	PARAMTYPE 0	VAR nQtd			AS NUMERIC			OPTIONAL DEFAULT 1
	PARAMTYPE 1	VAR nColunaIni		AS NUMERIC			OPTIONAL DEFAULT ::nColuna
	cUpdate	:= "UPDATE "+::cAliasCol+" SET COL=COL+"+cValToChar(nQtd)+" WHERE PLA="+cValToChar(::nPlanilhaAt)+" AND COL>="+cValToChar(nColunaIni)
	If ValType(nLinhaIni)=="N"
		cUpdate	+= " AND LIN>="+cValToChar(nLinhaIni)
	Endif
	If ValType(nLinhaFim)=="N"
		cUpdate	+= " AND LIN<="+cValToChar(nLinhaFim)
	Endif
	If !DbSqlExec(::cAliasCol,cUpdate,::cDriver)
		UserException("YExcel - Erro ao incluir linhas. "+TCSqlError())
	Endif	
Return

Method AddRow(nQtd,nLinhaIni,nColunaIni,nColunaFim) Class YExcel
	Local cUpdate	:= ""

	PARAMTYPE 0	VAR nQtd			AS NUMERIC			OPTIONAL DEFAULT 1
	PARAMTYPE 1	VAR nLinhaIni		AS NUMERIC			OPTIONAL DEFAULT ::nLinha
	// If ValType(nColunaFim)=="N" .AND. ValType(nColunaIni)!="N"
	// 	nColunaIni	:= 1
	// Endif
	//Atualiza as células para nova posição
	cUpdate	:= "UPDATE "+::cAliasCol+" SET LIN=LIN+"+cValToChar(nQtd)+" WHERE PLA="+cValToChar(::nPlanilhaAt)+" AND LIN>="+cValToChar(nLinhaIni)
	If ValType(nColunaIni)=="N"
		cUpdate	+= " AND COL>="+cValToChar(nColunaIni)
	Endif
	If ValType(nColunaFim)=="N"
		cUpdate	+= " AND COL<="+cValToChar(nColunaFim)
	Endif
	If !DbSqlExec(::cAliasCol,cUpdate,::cDriver)
		UserException("YExcel - Erro ao incluir linhas. "+TCSqlError())
	Endif

	//Atualiza as linhas para nova posição
	If ValType(nColunaFim)=="N" .OR. (ValType(nColunaIni)=="N" .AND. nColunaIni>1)
		//Movendo apenas as celulas para baixo
		//Cria linhas as células que foram movidas para baixo
		If !DbSqlExec(::cAliasLin,"INSERT INTO "+::cAliasLin+" (PLA,LIN) SELECT DISTINCT C.PLA,C.LIN FROM "+::cAliasCol+" C LEFT JOIN "+::cAliasLin+" L ON C.PLA=L.PLA AND C.LIN=L.LIN WHERE L.LIN IS NULL",::cDriver)
			UserException("YExcel - Erro ao incluir linhas. "+TCSqlError())
		Endif
	else
		//Movendo a linha inteira
		cUpdate	:= "UPDATE "+::cAliasLin+" SET LIN=LIN+"+cValToChar(nQtd)+" WHERE PLA="+cValToChar(::nPlanilhaAt)+" AND LIN>="+cValToChar(nLinhaIni)
		If !DbSqlExec(::cAliasCol,cUpdate,::cDriver)
			UserException("YExcel - Erro ao incluir linhas. "+TCSqlError())
		Endif
	Endif

Return

/*/{Protheus.doc} YExcel::SetRowH
Definir altura das llinhas. Se não enviado linha de/ate, considerar as próximas linhas a ser criadas
@type method
@version 1.0
@author Saulo Gomes Martins
@since 06/07/2020
@param nHeight, numeric, Altura da linha
@param nLinha, numeric, Opcional se vai alterar linha inicial
@param nLinha2, numeric, Opcional linha final para alteração
/*/
Method SetRowH(nHeight,nLinha,nLinha2) Class YExcel
	Local cCHeight	:= "1"
	PARAMTYPE 0	VAR nHeight			AS NUMERIC			OPTIONAL
	PARAMTYPE 1	VAR nLinha			AS NUMERIC			OPTIONAL
	PARAMTYPE 2	VAR nLinha2			AS NUMERIC			OPTIONAL	DEFAULT nLinha
	If ValType(nLinha)=="N"	//Altera apenas o range
		If ValType(nHeight)=="U"
			cCHeight	:= " "
			nHeight		:= 0
		Endif
		
		//Inserir linhas que não existe para definir o tamanho
		::InsertRowEmpty(nLinha,nLinha2)

		If !DbSqlExec(::cAliasLin,"UPDATE "+::cAliasLin+" SET CHEIGHT='"+cCHeight+"',HT="+cValToChar(nHeight)+" WHERE PLA="+cValToChar(::nPlanilhaAt)+" AND LIN>="+cValToChar(nLinha)+" AND LIN<="+cValToChar(nLinha2)+" ",::cDriver)
			UserException("YExcel - 2 Erro ao atualiza tamanho das linhas. "+TCSqlError())
		Endif
	Else
		::nTamLinha	:= nHeight
	Endif
Return

/*/{Protheus.doc} YExcel::InsertRowEmpty
Inserir linhas vazias
@type method
@version 1.0
@author Saulo Gomes Martins
@since 06/07/2020
@param nLinha, numeric, Linha inicial
@param nLinha2, numeric, Linha final
/*/
Method InsertRowEmpty(nLinha,nLinha2) Class YExcel
	Default nLinha	:= nLinha2
	If !DbSqlExec(::cAliasLin,"INSERT INTO "+::cAliasLin+" (PLA,LIN)"+;
		" WITH RECURSIVE lin(x) AS (VALUES("+cValToChar(nLinha)+") UNION ALL SELECT x+1 FROM lin WHERE x<"+cValToChar(nLinha2)+")"+;
		" SELECT "+cValToChar(::nPlanilhaAt)+",x FROM lin LEFT JOIN "+::cAliasLin+" TAB ON lin.x=TAB.LIN"+;
		" WHERE TAB.LIN is null",::cDriver)
		UserException("YExcel - Erro ao inserrir linhas. "+TCSqlError())
	Endif
Return

/*/{Protheus.doc} YExcel::InsertCellEmpty
Inserir células vazias
@type method
@version 1.0
@author Saulo Gomes Martins
@since 06/07/2020
@param nLinha, numeric, Linha inicial
@param nColuna, numeric, Coluna inicial
@param nLinha2, numeric, Linha final
@param nColuna2, numeric, Coluna final
/*/
Method InsertCellEmpty(nLinha,nColuna,nLinha2,nColuna2) Class YExcel
	Local nPos
	Local lAchou
	Default nLinha2		:= nLinha
	Default nColuna2	:= nColuna
	nPos	:= ::GetStrComp("",@lAchou)
	If !lAchou
		nPos	:= ::SetStrComp("")
	Endif
	::InsertRowEmpty(nLinha,nLinha2)
	If !DbSqlExec(::cAliasCol,"INSERT INTO "+::cAliasCol+" (PLA,LIN,COL,TIPO,TPSTY,TPVLR)"+;
		" WITH RECURSIVE lin(x) AS (VALUES("+cValToChar(nLinha)+") UNION ALL SELECT x+1 FROM lin WHERE x<"+cValToChar(nLinha2)+")"+;
		" ,col(y) AS (VALUES("+cValToChar(nColuna)+") UNION ALL SELECT y+1 FROM col WHERE y<"+cValToChar(nColuna2)+")"+;
		" SELECT "+cValToChar(::nPlanilhaAt)+",x,y,'s','S','U' FROM lin INNER JOIN col on 1=1 LEFT JOIN "+::cAliasCol+" TAB ON lin.x=TAB.LIN AND col.y=TAB.COL"+;
		" WHERE TAB.LIN is null",::cDriver)
		UserException("YExcel - Erro ao inserrir celulas. "+TCSqlError())
	Endif
Return

/*/{Protheus.doc} YExcel::SetValue
Alteração de valores da célula posicionada
@type method
@version 1.0
@author Saulo Gomes Martins
@since 06/07/2020
@param xValor, variadic, valor a ser gravado
@param cFormula, character, formula a ser gravada
@return object, self
/*/
Method SetValue(xValor,cFormula) Class YExcel
	Local cTipo	:= ValType(xValor)
	Local lAchou
	Local nPos
	Default cFormula	:= ""
	If !(::cAliasLin)->(DbSeek(Str(::nPlanilhaAt,10)+Str(::nLinha,10)))
		(::cAliasLin)->(RecLock(::cAliasLin,.T.))
		(::cAliasLin)->PLA		:= ::nPlanilhaAt
		(::cAliasLin)->LIN		:= ::nLinha
		(::cAliasLin)->OLEVEL	:= ""
		(::cAliasLin)->COLLAP	:= ""
		(::cAliasLin)->CHIDDEN	:= ""
		(::cAliasLin)->CHEIGHT	:= ""
		(::cAliasLin)->HT		:= 0
		If ValType(::nRowoutlineLevel)=="N"
			(::cAliasLin)->OLEVEL	:= cValToChar(::nRowoutlineLevel)
		Endif
		If ::lRowcollapsed
			(::cAliasLin)->COLLAP	:= "1"
		Endif
		If ::lRowhidden
			(::cAliasLin)->CHIDDEN	:= "1"
		Endif
		If ValType(::nTamLinha)=="N"
			(::cAliasLin)->CHEIGHT	:= "1"
			(::cAliasLin)->HT	:= ::nTamLinha
		Endif
		(::cAliasLin)->(MsUnLock())
	Endif

	If !(::cAliasCol)->(DbSeek(Str(::nPlanilhaAt,10)+Str(::nLinha,10)+Str(::nColuna,10)))
		(::cAliasCol)->(RecLock(::cAliasCol,.T.))
		(::cAliasCol)->STY		:= -1
		(::cAliasCol)->TPSTY	:= " "
	else
		(::cAliasCol)->(RecLock(::cAliasCol,.F.))
	Endif
	(::cAliasCol)->PLA		:= ::nPlanilhaAt
	(::cAliasCol)->LIN		:= ::nLinha
	(::cAliasCol)->COL		:= ::nColuna
	If cTipo=="C"
		(::cAliasCol)->TIPO		:= "s"
		(::cAliasCol)->TPSTY	:= "S"
	ElseIf cTipo=="L"
		(::cAliasCol)->TIPO		:= "b"
		(::cAliasCol)->TPSTY	:= "B"
	ElseIf cTipo=="N"
		(::cAliasCol)->TIPO		:= "n"
		(::cAliasCol)->TPSTY	:= "N"
	ElseIf cTipo=="D"
		(::cAliasCol)->TIPO		:= "d"
		If (::cAliasCol)->STY>=0 .AND. !::oStyle:XPathGetAtt("/xmlns:styleSheet/xmlns:cellXfs/xmlns:xf["+cValToChar((::cAliasCol)->STY+1)+"]","applyNumberFormat")=="1"//!("D" $ ::StyleType((::cAliasCol)->STY))
			//Se tem estilo e ele não tem NumFmtId aplicado
			(::cAliasCol)->STY		:= ::CreateStyle((::cAliasCol)->STY,14)	//Cria um estilo com Numfmt data
			(::cAliasCol)->TPSTY	:= "D"
		ElseIf (::cAliasCol)->STY<0	//Não tem estilo
			(::cAliasCol)->STY		:= ::aPadraoSty[2]	//Estilo padrão de data
			(::cAliasCol)->TPSTY	:= "D"
		Endif
	ElseIf cTipo=="O" .and. GetClassName(xValor)=="YEXCEL_DATETIME"
		(::cAliasCol)->TIPO		:= "d"
		If (::cAliasCol)->STY>=0 .AND. ::oStyle:XPathGetAtt("/xmlns:styleSheet/xmlns:cellXfs/xmlns:xf["+cValToChar((::cAliasCol)->STY+1)+"]","applyNumberFormat")!="1"//!("D" $ ::StyleType((::cAliasCol)->STY))
			//Se tem estilo e ele não tem NumFmtId aplicado
			(::cAliasCol)->STY		:= ::CreateStyle((::cAliasCol)->STY,::AddFmt("dd/mm/yyyy\ hh:mm AM/PM;@"))	//Cria estilo com Numfmt datetime
			(::cAliasCol)->TPSTY	:= "T"
		ElseIf (::cAliasCol)->STY<0	//Não tem estilo
			(::cAliasCol)->STY		:= ::aPadraoSty[3]	//Estilo padrão de datetime
			(::cAliasCol)->TPSTY	:= "T"
		Endif
	Endif
	If ValType(cFormula)=="C"
		(::cAliasCol)->FORMULA	:= cFormula
	Endif
	(::cAliasCol)->TPVLR	:= "N"	//Numeros
	If cTipo=="C"
		nPos	:= ::GetStrComp(xValor,@lAchou)
		If lAchou
			(self:cAliasCol)->VLRNUM	:= nPos
		Else
			nPos	:= ::SetStrComp(xValor)
			(self:cAliasCol)->VLRNUM	:= nPos
		Endif
	ElseIf cTipo=="L"
		(::cAliasCol)->VLRNUM	:= if(xValor,1,0)
	ElseIf cTipo=="N"
		(::cAliasCol)->VLRNUM	:= xValor
	ElseIf cTipo=="D"
		(::cAliasCol)->TPVLR	:= "N"	//Numerico
		If !Empty(xValor)
			// (::cAliasCol)->VLRTXT	:= cValToChar(xValor-STOD("19000101")+2)
			(::cAliasCol)->VLRNUM	:= xValor-STOD("19000101")+2
		Else
			(::cAliasCol)->TPVLR	:= "U"	//Nulo
			// (::cAliasCol)->VLRTXT	:= ""
		Endif
	ElseIf cTipo=="O" .and. GetClassName(xValor)=="YEXCEL_DATETIME"
		(::cAliasCol)->TPVLR	:= "N"	//Numerico
		(::cAliasCol)->VLRNUM	:= xValor:nNumero
		(::cAliasCol)->VLRDEC	:= xValor:nDecimal
		// (::cAliasCol)->VLRTXT	:= cValToChar(xValor:GetStrNumber())
		// (::cAliasCol)->TPVLR	:= "C"	//Caracteres
	ElseIf cTipo=="U"
		(::cAliasCol)->VLRTXT	:= ""
		(::cAliasCol)->TPVLR	:= "U"	//Nulo
	Else
		(::cAliasCol)->VLRTXT	:= cValToChar(xValor)
		(::cAliasCol)->TPVLR	:= "C"	//Caracteres
	Endif
	(::cAliasCol)->(MsUnLock())
Return self
/*/{Protheus.doc} YExcel::SetDateTime
Alteração de data e hora da célula posicionada
@type method
@version 1.0
@author Saulo Gomes Martins
@since 30/03/2021
@param dDate, date, Data a ser gravada
@param cTime, character, Hora a ser gravada
@return object, self
/*/
Method SetDateTime(dDate,cTime) Class YExcel
	Local oDateTime
	Default dDate	:= CTOD("")
	Default cTime	:= "00:00:00"
	oDateTime	:= ::GetDateTime(dDate,cTime)
	::SetValue(oDateTime)
	FreeObj(oDateTime)
Return self

//NÃO DOCUMENTAR
Method GetStrComp(xTexto,lAchou) Class YExcel
	Local xRet
	Local cTxtMd5
	lAchou	:= .F.
	If ValType(xTexto)=="C"
		(::cAliasStr)->(DbSetOrder(1))
		cTxtMd5	:= Md5(xTexto)
		If (::cAliasStr)->(DbSeek(PadR(cTxtMd5,200)))
			xRet	:= (::cAliasStr)->POS
			lAchou	:= .T.
		Endif
	ElseIf ValType(xTexto)=="N"
		(::cAliasStr)->(DbSetOrder(2))
		If (::cAliasStr)->(DbSeek(Str(xTexto,10)))
			lAchou	:= .T.
			xRet	:= &((::cAliasStr)->VLRMEMO)
		Endif
	Endif
Return xRet

//NÃO DOCUMENTAR
Method SetStrComp(xTexto) Class YExcel
	Local nPos	:= ::nQtdString
	(::cAliasStr)->(RecLock(::cAliasStr,.T.))
	(::cAliasStr)->POS		:= nPos
	(::cAliasStr)->VLRTEXTO	:= Md5(xTexto)
	(::cAliasStr)->VLRMEMO	:= "'"+Replace(xTexto,"'","'+chr(39)+'")+"'"
	(::cAliasStr)->(MsUnLock())
	::nQtdString++
Return nPos

/*/{Protheus.doc} YExcel::GetCell
Posiciona em uma celula
@type method
@version 1.0
@author Saulo Gomes Martins
@since 03/07/2020
@param nLinha, numeric, Linha para posicionamento
@param nColuna, numeric, Coluna para posicionamento
@param nPlanilha, numeric, Planilha para posicionamento
@return object, self
/*/
Method Pos(nLinha,nColuna,nPlanilha) Class YExcel
	PARAMTYPE 0	VAR nLinha		AS NUMERIC
	PARAMTYPE 1	VAR nColuna		AS NUMERIC
	PARAMTYPE 2	VAR nPlanilha	AS NUMERIC OPTIONAL DEFAULT ::nPlanilhaAt
	::nLinha		:= nLinha
	::nColuna		:= nColuna
	::nPlanilhaAt	:= nPlanilha
Return self

/*/{Protheus.doc} YExcel::GetCell
Posiciona pela referência
@type method
@version 1.0
@author Saulo Gomes Martins
@since 03/07/2020
@param nLinha, numeric, Linha para posicionamento
@param nColuna, numeric, Coluna para posicionamento
@param nPlanilha, numeric, Planilha para posicionamento
@return object, self
/*/
Method PosR(cRef) Class YExcel
	Local aRef := ::LocRef(cRef)
	::nLinha		:= aRef[1]
	::nColuna		:= aRef[2]
Return self

/*/{Protheus.doc} YExcel::Getformula
Retorna a Formula de uma celula
@type method
@version 1.0
@author Saulo Gomes Martins
@since 26/06/2020
@param nLinha, numeric, Linha a ser lida
@param nColuna, numeric, Coluna a ser lida
@param lAchou, logical, se achou o conteudo
@return character, formula gravada na celula
/*/
Method Getformula() Class YExcel
	If (::cAliasCol)->(DbSeek(Str(::nPlanilhaAt,10)+Str(::nLinha,10)+Str(::nColuna,10)))	//Coluna
		Return (::cAliasCol)->FORMULA
	Endif
Return ""

METHOD ColTam(nLinha,nLinha2)	Class YExcel		//Coluna Mínima e Máxima
	Local cQuery
	Local cAliasQry := GetNextAlias()
	Local nMin		:= 0
	Local nMax		:= 0
	Default nLinha2	:= nLinha
	cQuery	:= "SELECT MIN(COL) COL_INI,MAX(COL) COL_FIM FROM "+::cAliasCol+" WHERE"
	If ValType(nLinha)=="N"
		cQuery	+= " LIN>="+cValToChar(nLinha)+" AND"
		cQuery	+= " LIN<="+cValToChar(nLinha2)+" AND"
	Endif
	cQuery	+= " PLA="+cValToChar(::nPlanilhaAt)+" AND TPVLR<>'U' AND D_E_L_E_T_=' '"
	If !DbSqlExec(cAliasQry,cQuery,::cDriver)
		UserException("YExcel - Erro ao obter dados max e min colunas. "+TCSqlError())
	Endif
	If (cAliasQry)->(!EOF())
		nMin	:= (cAliasQry)->COL_INI
		nMax	:= (cAliasQry)->COL_FIM
	Endif
	(cAliasQry)->(DbCloseArea())
Return {nMin,nMax}

METHOD LinTam(nColuna,nColuna2)	Class YExcel		//Linha Mínima e Máxima
	Local cQuery
	Local cAliasQry 	:= GetNextAlias()
	Local nMin			:= 1
	Local nMax			:= 1
	Default nColuna2	:= nColuna
	cQuery	:= "SELECT MIN(LIN) LIN_INI,MAX(LIN) LIN_FIM FROM "+::cAliasCol+" WHERE"
	If ValType(nColuna)=="N"
		cQuery	+= " COL>="+cValToChar(nColuna)+" AND"
		cQuery	+= " COL<="+cValToChar(nColuna2)+" AND"
	Endif
	cQuery	+= " PLA="+cValToChar(::nPlanilhaAt)+" AND TPVLR<>'U' AND D_E_L_E_T_=' '"
	If !DbSqlExec(cAliasQry,cQuery,::cDriver)
		UserException("YExcel - Erro ao obter dados max e min linhas. "+TCSqlError())
	Endif
	nMin	:= (cAliasQry)->LIN_INI
	nMax	:= (cAliasQry)->LIN_FIM
	(cAliasQry)->(DbCloseArea())
Return {nMin,nMax}

/*/{Protheus.doc} GetValue
Retorna o conteudo de uma célula
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param nLinha, numeric, Linha a ser lida
@param nColuna, numeric, Coluna a ser lida
@param xDefault, variadic, Valor padrão
@param lAchou, logical, se achou o conteudo
@type method
@obs apenas se modo sql ativo
/*/
Method GetValue(nLinha,nColuna,xDefault,lAchou) Class YExcel
	Local xRet := xDefault
	Local cTmp
	Default nLinha	:= ::nLinha
	Default nColuna	:= ::nColuna
	lAchou	:= .F.
	If (::cAliasCol)->(DbSeek(Str(::nPlanilhaAt,10)+Str(nLinha,10)+Str(nColuna,10)))	//Coluna
		If (::cAliasCol)->TPVLR=="C"
			xRet	:= Alltrim((::cAliasCol)->VLRTXT)
			If (::cAliasCol)->TPSTY="D"
				xRet	:= yExcel_DateTime():New(,,,,xRet)//STOD("19000101")+Val(xRet)-2
			Endif
		ElseIf (::cAliasCol)->TPVLR=="U"
			Return nil
		Else
			xRet	:= (::cAliasCol)->VLRNUM
			If (::cAliasCol)->TIPO=="s"
				cTmp	:= ::GetStrComp(xRet,@lAchou)
				If lAchou
					xRet	:= cTmp
				Endif
			ElseIf (::cAliasCol)->TIPO=="d" .AND. (::cAliasCol)->TPSTY="D"
				xRet	:= STOD("19000101")-2+(::cAliasCol)->VLRNUM
			ElseIf (::cAliasCol)->TIPO=="d"
				xRet	:= yExcel_DateTime():New(,,(::cAliasCol)->VLRNUM,(::cAliasCol)->VLRDEC)
			ElseIf (::cAliasCol)->TIPO=="b"
				If xRet==0
					xRet	:= .F.
				Else
					xRet	:= .T.
				Endif
			Endif
		Endif
	Endif
Return xRet


/*/{Protheus.doc} SetDefRow
Defini as colunas da linha. Habilita a gravação automatica de cada coluna. Importante para prover performace na gravação de varias linhas
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param lHabilitar, logical, Habilita a definição
@param aSpanRow, array, 1-Coluna inicial|2-Coluna Final
@type method
/*/
METHOD SetDefRow(lHabilitar,aSpanRow) Class YExcel
	Default	lHabilitar	:= .T.
	::lRowDef		:= lHabilitar
	ConOut("[Warning] Method deprecated")
Return
/*/{Protheus.doc} YExcel::NivelLinha
Controla nível da estrutura de tópicos das próximas linhas criadas
@type method
@version 1.0
@author Saulo Gomes Martins
@since 06/07/2020
@param nNivel, numeric, Nível da linha
@param lFechado, logical, Se tem nível
@param lOculto, logical, se vai ficar aculta a linha
/*/
METHOD NivelLinha(nNivel,lFechado,lOculto) Class YExcel
	PARAMTYPE 0	VAR nNivel		AS NUMERIC	OPTIONAL
	PARAMTYPE 1	VAR lFechado	AS LOGICAL	OPTIONAL	DEFAULT	.F.
	PARAMTYPE 2	VAR lOculto		AS LOGICAL	OPTIONAL	DEFAULT .F.
	::nRowoutlineLevel	:= nNivel
	::lRowcollapsed		:= lFechado
	::lRowHidden		:= lOculto
Return
/*/{Protheus.doc} YExcel::SetsumBelow
Indica se as linhas de resumo aparece abaixo das linhas agrupadas
@type method
@version 1.0
@author Saulo Gomes Martins
@since 16/03/2021
@param lsummaryBelow, logical, defini resumo abaixo .T. ou acima .F.
/*/
Method SetsumBelow(lsummaryBelow) Class YExcel
	If !::asheet[::nPlanilhaAt][1]:XPathHasNode("/xmlns:worksheet/xmlns:sheetPr/xmlns:outlinePr")
		::asheet[::nPlanilhaAt][1]:XPathAddNode("/xmlns:worksheet/xmlns:sheetPr","outlinePr","")
	EndIf
	If lsummaryBelow
		If !Empty(::asheet[::nPlanilhaAt][1]:XPathGetAtt("/xmlns:worksheet/xmlns:sheetPr/xmlns:outlinePr","summaryBelow"))
			::asheet[::nPlanilhaAt][1]:XPathSetAtt("/xmlns:worksheet/xmlns:sheetPr/xmlns:outlinePr","summaryBelow","1")
		Else
			::asheet[::nPlanilhaAt][1]:XPathAddAtt("/xmlns:worksheet/xmlns:sheetPr/xmlns:outlinePr","summaryBelow","1")
		Endif
	Else
		If !Empty(::asheet[::nPlanilhaAt][1]:XPathGetAtt("/xmlns:worksheet/xmlns:sheetPr/xmlns:outlinePr","summaryBelow"))
			::asheet[::nPlanilhaAt][1]:XPathSetAtt("/xmlns:worksheet/xmlns:sheetPr/xmlns:outlinePr","summaryBelow","0")
		Else
			::asheet[::nPlanilhaAt][1]:XPathAddAtt("/xmlns:worksheet/xmlns:sheetPr/xmlns:outlinePr","summaryBelow","0")
		Endif
	Endif
Return
/*/{Protheus.doc} YExcel::SetsumRight
Indica se as colunas de resumo aparece a direita das colunas agrupadas
@type method
@version 1.0
@author Saulo Gomes Martins
@since 16/03/2021
@param lsummaryRight, logical, .T. Coluna resumo a direita .F. coluna a esquerda
/*/
Method SetsumRight(lsummaryRight) Class YExcel
	If !::asheet[::nPlanilhaAt][1]:XPathHasNode("/xmlns:worksheet/xmlns:sheetPr/xmlns:outlinePr")
		::asheet[::nPlanilhaAt][1]:XPathAddNode("/xmlns:worksheet/xmlns:sheetPr","outlinePr","")
	EndIf
	If lsummaryRight
		If !Empty(::asheet[::nPlanilhaAt][1]:XPathGetAtt("/xmlns:worksheet/xmlns:sheetPr/xmlns:outlinePr","summaryRight"))
			::asheet[::nPlanilhaAt][1]:XPathSetAtt("/xmlns:worksheet/xmlns:sheetPr/xmlns:outlinePr","summaryRight","1")
		Else
			::asheet[::nPlanilhaAt][1]:XPathAddAtt("/xmlns:worksheet/xmlns:sheetPr/xmlns:outlinePr","summaryRight","1")
		Endif
	Else
		If !Empty(::asheet[::nPlanilhaAt][1]:XPathGetAtt("/xmlns:worksheet/xmlns:sheetPr/xmlns:outlinePr","summaryRight"))
			::asheet[::nPlanilhaAt][1]:XPathSetAtt("/xmlns:worksheet/xmlns:sheetPr/xmlns:outlinePr","summaryRight","0")
		Else
			::asheet[::nPlanilhaAt][1]:XPathAddAtt("/xmlns:worksheet/xmlns:sheetPr/xmlns:outlinePr","summaryRight","0")
		Endif
	Endif
Return
/*/{Protheus.doc} YExcel::SetRowLevel
Defini o nível das linhas informadas (agrupamento de linhas)
@type method
@version 1.0
@author Saulo Gomes Martins
@since 16/03/2021
@param nLinha, numeric, Linha inicial
@param nLinha2, numeric, Linha final
@param nNivel, numeric, Nível
@param lFechado, logical, Se esse nível está fechado
/*/
Method SetRowLevel(nLinha,nLinha2,nNivel,lFechado) Class YExcel
	Local nCont			:= nLinha-1
	Local lsummaryBelow	:= .T.		//Resumo abaixo
	Local csummaryBelow	:= ::asheet[::nPlanilhaAt][1]:XPathGetAtt("/xmlns:worksheet/xmlns:sheetPr/xmlns:outlinePr","summaryBelow")
	PARAMTYPE 0	VAR nNivel		AS NUMERIC	OPTIONAL
	PARAMTYPE 1	VAR lFechado	AS LOGICAL	OPTIONAL	DEFAULT	.F.
	If !Empty(csummaryBelow) .AND. csummaryBelow=="0"
		lsummaryBelow	:= .F.
	Endif
	If ValType(nNivel)!="N"
		lFechado	:= .F.
	Endif
	If !lsummaryBelow .AND. lFechado .AND. nCont>0
		If !(::cAliasLin)->(DbSeek(Str(::nPlanilhaAt,10)+Str(nCont,10)))
			(::cAliasLin)->(RecLock(::cAliasLin,.T.))
			(::cAliasLin)->PLA		:= ::nPlanilhaAt
			(::cAliasLin)->LIN		:= nCont
		Else
			(::cAliasLin)->(RecLock(::cAliasLin,.F.))
		Endif
		If lFechado
			(::cAliasLin)->COLLAP	:= "1"
		Else
			(::cAliasLin)->COLLAP	:= ""
		Endif
		(::cAliasLin)->(MsUnLock())
	Endif
	For nCont:=nLinha to nLinha2
		If !(::cAliasLin)->(DbSeek(Str(::nPlanilhaAt,10)+Str(nCont,10)))
			(::cAliasLin)->(RecLock(::cAliasLin,.T.))
			(::cAliasLin)->PLA		:= ::nPlanilhaAt
			(::cAliasLin)->LIN		:= nCont
		Else
			(::cAliasLin)->(RecLock(::cAliasLin,.F.))
		Endif
		If ValType(nNivel)=="N"
			If nNivel>Val((::cAliasLin)->OLEVEL)
				(::cAliasLin)->OLEVEL	:= cValToChar(nNivel)
				If !lFechado
					(::cAliasLin)->CHIDDEN	:= ""
				EndIf
			Endif
		Else
			(::cAliasLin)->OLEVEL	:= ""
			If !lFechado
				(::cAliasLin)->CHIDDEN	:= ""
			EndIf
		Endif
		If lFechado
			(::cAliasLin)->CHIDDEN	:= "1"
		EndIf
		(::cAliasLin)->(MsUnLock())
	Next
	If lsummaryBelow .AND. lFechado
		If !(::cAliasLin)->(DbSeek(Str(::nPlanilhaAt,10)+Str(nCont,10)))
			(::cAliasLin)->(RecLock(::cAliasLin,.T.))
			(::cAliasLin)->PLA		:= ::nPlanilhaAt
			(::cAliasLin)->LIN		:= nCont
		Else
			(::cAliasLin)->(RecLock(::cAliasLin,.F.))
		Endif
		If lFechado
			(::cAliasLin)->COLLAP	:= "1"
		Else
			(::cAliasLin)->COLLAP	:= ""
		Endif
		(::cAliasLin)->(MsUnLock())
	Endif
Return

/*/{Protheus.doc} showGridLines
Se vai exibir ou ocultar linhas de grade na planilha
@author Saulo Gomes Martins
@since 11/12/2019
@version 1.0
@param lView, logical, Se falso oculta linhas de grade
@type method
@obs pag 1709
/*/
METHOD showGridLines(lView) Class YExcel
	If Empty(::asheet[::nPlanilhaAt][1]:XPathGetAtt("/xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView","showGridLines"))
		::asheet[::nPlanilhaAt][1]:XPathAddAtt("/xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView", "showGridLines"	, If(lView,"1","0") )
	Else
		::asheet[::nPlanilhaAt][1]:XPathSetAtt("/xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView", "showGridLines"	, If(lView,"1","0") )
	Endif
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
@type method
@obs pag 1601 - 18.3.1.2
/*/
Method AutoFilter(nLinha,nColuna,nLinha2,nColuna2) Class YExcel
	Local cColuna,cColuna2
	cColuna		:= NumToString(nColuna)
	cColuna2	:= NumToString(nColuna2)
	If !::asheet[::nPlanilhaAt][1]:XPathHasNode( "/xmlns:worksheet/xmlns:autoFilter" )
		::asheet[::nPlanilhaAt][1]:XPathAddNode( "/xmlns:worksheet", "autoFilter", "" )
	Endif
	SetAtrr(::asheet[::nPlanilhaAt][1],"/xmlns:worksheet/xmlns:autoFilter", "ref"	, cColuna+cValToChar(nLinha)+":"+cColuna2+cValToChar(nLinha2) )
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
@type method
/*/
Method mergeCells(nLinha,nColuna,nLinha2,nColuna2) Class YExcel
	Local cColuna,cColuna2,nCont,cAtrr
	Local aChildren
	Local nPos	:= 0
	If nLinha2<nLinha
		UserException("YExcel - metodo mergeCells. Linha final não pode ser menor que linha inicial.")
	Endif
	If nColuna2<nColuna
		UserException("YExcel - metodo mergeCells. Coluna final não pode ser menor que Coluna inicial.")
	Endif
	cColuna		:= NumToString(nColuna)
	cColuna2	:= NumToString(nColuna2)
	aChildren	:= ::asheet[::nPlanilhaAt][1]:XPathGetChildArray( "/xmlns:worksheet/xmlns:mergeCells" )
	For nCont:=1 to Len(aChildren)
		cAtrr	:= ::asheet[::nPlanilhaAt][1]:XPathGetAtt(aChildren[nCont][2],"ref")
		If Replace(cColuna+cValToChar(nLinha),"$","") $ Replace(cAtrr,"$","") .OR. Replace(cColuna2+cValToChar(nLinha2),"$","") $ Replace(cAtrr,"$","")
			nPos	:= nCont
			Exit
		Endif
	Next
	If nPos>0
		UserException("YExcel - metodo mergeCells. Célula "+cColuna+cValToChar(nLinha)+":"+cColuna2+cValToChar(nLinha2)+" não pode ser mesclada, essa célula já foi mesclada!")
	Endif
	// If Empty(aChildren)
	// 	::asheet[::nPlanilhaAt][1]:XPathAddNode( "/xmlns:worksheet","mergeCells", "" )
	// 	::asheet[::nPlanilhaAt][1]:XPathAddAtt( "/xmlns:worksheet/xmlns:mergeCells","count","0")
	// Endif
	
	SetAtrr(::asheet[::nPlanilhaAt][1],"/xmlns:worksheet/xmlns:mergeCells","count", cValToChar(Val(::asheet[::nPlanilhaAt][1]:XPathGetAtt("xmlns:worksheet/xmlns:mergeCells","count"))+1) )
	::asheet[::nPlanilhaAt][1]:XPathAddNode( "/xmlns:worksheet/xmlns:mergeCells","mergeCell", "" )
	::asheet[::nPlanilhaAt][1]:XPathAddAtt( "/xmlns:worksheet/xmlns:mergeCells/xmlns:mergeCell[last()]","ref", cColuna+cValToChar(nLinha)+":"+cColuna2+cValToChar(nLinha2) )

	If (::cAliasCol)->(DbSeek(Str(::nPlanilhaAt,10)+Str(nLinha,10)+Str(nColuna,10)))
		If (::cAliasCol)->STY>=0	//Replicar estilo da primeira célula
			::SetStyle((::cAliasCol)->STY,nLinha,nColuna,nLinha2,nColuna2)
		Endif
	Endif

Return

/*/{Protheus.doc} Font
Cria objeto de fonte para ser usado na criação de estilos para formatação condicional
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param nTamanho, numeric, (opcional) Tamanho da fonte
@param cCorRGB, characters, (opcional) Cor da fonte em Alpha+RGB
@param cNome, characters, (opcional) Nome da fonte
@param cfamily, characters, (opcional) Familia da fonte
@param cScheme, characters, (opcional) Schema
@param lNegrito, logical, (opcional) Negrito
@param lItalico, logical, (opcional) Italico
@param lSublinhado, logical, (opcional) Soblinhado
@param lTachado, logical, (opcional) Tachado
@type method
/*/
METHOD Font(nTamanho,cCorRGB,cNome,cfamily,cScheme,lNegrito,lItalico,lSublinhado,lTachado) Class YExcel

Return {nTamanho,cCorRGB,cNome,cfamily,cScheme,lNegrito,lItalico,lSublinhado,lTachado}

/*/{Protheus.doc} Preenc
Cria objeto de preenchimento para ser usado na criação de estilos para formatação condicional
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param cBgCor, characters, (Opcional) Cor em Alpha+RGB do preenchimento
@param cFgCor, characters, (Opcional) Cor em Aplha+RGB do fundo
@param cType, characters, (Opcional) tipo de preenchimento(padrão solid)
@type method
/*/
METHOD Preenc(cBgCor,cFgCor,cType) Class YExcel
Default cType	:= "solid"
Return {cFgCor,cBgCor,cType}

/*/{Protheus.doc} ObjBorda
Cria objeto de bordas para ser usado na criação de estilos para formatação condicional
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param cTipo, characters, "C"-Cima|"B"-Baixo|"E"-Esquerda|"D"-Direita|T-TODAS("CBED") OU "T"-TOP|"B"-Bottom|"L"-Left|"R"-Right|A-ALL("TBLR") OU "DIAGONAL"-Diagonal
@param cCor, characters, Cor em Aplha+RGB da borda
@param cModelo, characters, Modelo da borda
@type method
@Obs pode juntar os tipo. Exemplo "ED"-Esquerda e direita
/*/
METHOD ObjBorda(cTipo,cCor,cModelo) Class YExcel
Return {cTipo,cCor,cModelo}

/*/{Protheus.doc} ADDdxf
Cria estilo para formatação condicional
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param aFont, array, (opcional) objeto criado pelo metodo :Font() com fonte
@param aCorPreenc, array, (opcional) objeto com cor criado pelo metodo :Preench() de preenchimento
@param aBorda, object, (opcional) objeto criado pelo metodo :ObjBorda() com borda
@return numeric, posição do estilo
@type method
/*/
METHOD ADDdxf(aFont,aCorPreenc,aBorda) Class YExcel
	Local nTamdxfs

	::oStyle:XPathAddNode( "xmlns:styleSheet/xmlns:dxfs","dxf", "" )
	nTamdxfs	:= Val(::oStyle:XPathGetAtt("xmlns:styleSheet/xmlns:dxfs","count"))+1
	::oStyle:XPathSetAtt("xmlns:styleSheet/xmlns:dxfs","count",cValToChar(nTamdxfs))

	//Font
	If ValType(aFont)=="A"
		::AddFont(aFont[1],aFont[2],aFont[3],aFont[4],aFont[5],aFont[6],aFont[7],aFont[8],aFont[9],"xmlns:styleSheet/xmlns:dxfs/xmlns:dxf[last()]")
	Endif
	//Preenchimento
	If ValType(aCorPreenc)=="A"
		::CorPreenc(aCorPreenc[1],aCorPreenc[2],aCorPreenc[3],"xmlns:styleSheet/xmlns:dxfs/xmlns:dxf[last()]")
	Endif
	//Borda
	If ValType(aBorda)=="A"
		::Borda(aBorda[1],aBorda[2],aBorda[3],"xmlns:styleSheet/xmlns:dxfs/xmlns:dxf[last()]")
	Endif
Return nTamdxfs-1

/*/{Protheus.doc} FormatCond
Cria uma regra para formatação condicional
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param cRefDe, characters, Rerefencia inicial (Exemplo "A1")
@param cRefAte, characters, Referencia final (Exemplo "A10")
@param nEstilo, numeric, posição do estilo criado pelo metodo :ADDdxf()
@param operator, object, Operado da regra. veja observações
@param xFormula, variadic, (characters ou array) formula para uso
@type method
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
METHOD FormatCond(cRefDe,cRefAte,nEstilo,operator,xFormula) Class YExcel
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
	Endif
	If operator=="between"	.and. (ValType(xFormula)<>"A" .or. Len(xFormula)<2)
		UserException("YExcel - operador between é necessario informar valor de, ate. Enviar parametro 5 xformula como array(2).")
	Endif
	If operator=="FORMULA"
		::AddFormatCond(cRefDe,cRefAte,nEstilo,"expression",xFormula,,)
	Else
		::AddFormatCond(cRefDe,cRefAte,nEstilo,"cellIs",xFormula,operator,)
	Endif
	::nPriodFormCond++
Return
//NÃO DOCUMENTAR
METHOD AddFormatCond(cRefDe,cRefAte,nEstilo,cType,xFormula,operator,nPrioridade) Class YExcel
	Local cRef	:= cRefDe+If(!Empty(cRefAte),":"+cRefAte,"")
	Local nCont
	Local aChildren
	Local cPos := ""
	Local nPos	:= 0
	PARAMTYPE 0	VAR cRefDe		AS CHARACTER
	PARAMTYPE 1	VAR cRefAte		AS CHARACTER				OPTIONAL
	PARAMTYPE 2	VAR nEstilo		AS NUMERIC
	PARAMTYPE 3	VAR cType		AS CHARACTER
	PARAMTYPE 4	VAR xFormula	AS ARRAY,CHARACTER,NUMERIC
	PARAMTYPE 5	VAR operator	AS CHARACTER				OPTIONAL
	PARAMTYPE 6	VAR nPrioridade	AS NUMERIC					OPTIONAL DEFAULT ::nPriodFormCond
	/*	TYPES	(pag 2452)
	aboveAverage	-	abaixo da media
	beginsWith		-	inicia com
	cellIs			-	celula é(usar operador)
	colorScale		-	Estala de cor
	expression		-	Usar Formula
	top10			-
	...
	*/
	aChildren	:= ::asheet[::nPlanilhaAt][1]:XPathGetChildArray( "/xmlns:worksheet/xmlns:conditionalFormatting" )
	For nCont:=1 to Len(aChildren)
		cAtrr	:= ::asheet[::nPlanilhaAt][1]:XPathGetAtt(aChildren[nCont][2],"ref")
		If cAtrr==cRef
			nPos	:= nCont
			Exit
		Endif
	Next
	If nPos==0
		::asheet[::nPlanilhaAt][1]:XPathAddNode( "/xmlns:worksheet", "conditionalFormatting", "" )
		::asheet[::nPlanilhaAt][1]:XPathAddAtt( "/xmlns:worksheet/xmlns:conditionalFormatting[last()]", "sqref"	, cRef)
		cPos	:= "last()"
	Else
		cPos	:= cValToChar(nPos)
	Endif

	::asheet[::nPlanilhaAt][1]:XPathAddNode( "/xmlns:worksheet/xmlns:conditionalFormatting["+cPos+"]", "cfRule", "" )
	::asheet[::nPlanilhaAt][1]:XPathAddAtt( "/xmlns:worksheet/xmlns:conditionalFormatting["+cPos+"]/xmlns:cfRule[last()]", "type"	, cType)
	::asheet[::nPlanilhaAt][1]:XPathAddAtt( "/xmlns:worksheet/xmlns:conditionalFormatting["+cPos+"]/xmlns:cfRule[last()]", "dxfId"	, cValToChar(nEstilo))
	::asheet[::nPlanilhaAt][1]:XPathAddAtt( "/xmlns:worksheet/xmlns:conditionalFormatting["+cPos+"]/xmlns:cfRule[last()]", "priority", cValToChar(nPrioridade))
	If ValType(operator)<>"U"
		::asheet[::nPlanilhaAt][1]:XPathAddAtt( "/xmlns:worksheet/xmlns:conditionalFormatting["+cPos+"]/xmlns:cfRule[last()]", "operator", operator)
	Endif

	If ValType(xFormula)<>"U"
		If ValType(xFormula)=="A"
			For nCont:=1 to Len(xFormula)
				::asheet[::nPlanilhaAt][1]:XPathAddNode( "/xmlns:worksheet/xmlns:conditionalFormatting["+cPos+"]/xmlns:cfRule[last()]", "formula", cValToChar(xFormula[nCont]) )
				If nCont==3	//maxOccurs="3" pag 3936
					Exit
				EndIf
			Next
		Else
			::asheet[::nPlanilhaAt][1]:XPathAddNode( "/xmlns:worksheet/xmlns:conditionalFormatting["+cPos+"]/xmlns:cfRule[last()]", "formula", cValToChar(xFormula) )
		Endif
	Endif
Return

/*/{Protheus.doc} AddFont
Adiciona fonte para ser usado no estilo das células
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param nTamanho, numeric, (Opcional) Tamanho da fonte
@param cCorRGB, characters, (Opcional) Cor da fonte em Alpha+RGB
@param cNome, characters, (Opcional) Nome da fonte
@param cfamily, characters, (Opcional) Familia da fonte
@param cScheme, characters, (Opcional) Schema
@param lNegrito, logical, (Opcional) Negrito
@param lItalico, logical, (Opcional) Italico
@param lSublinhado, logical, (Opcional) Soblinhado
@param lTachado, logical, (Opcional) Tachado
@return numeric, posição da fonte
@type method
/*/
METHOD AddFont(nTamanho,cCorRGB,cNome,cfamily,cScheme,lNegrito,lItalico,lSublinhado,lTachado,cLocal) Class YExcel
	Local nTamFonts := 0
	Local cChave	:= ""
	PARAMTYPE 0	VAR nTamanho		AS NUMERIC				OPTIONAL DEFAULT 11
	PARAMTYPE 1	VAR cCorRGB			AS CHARACTER,NUMERIC	OPTIONAL DEFAULT "FF000000"
	PARAMTYPE 2	VAR cNome			AS CHARACTER			OPTIONAL DEFAULT "Calibri"
	PARAMTYPE 3	VAR cfamily			AS CHARACTER			OPTIONAL
	PARAMTYPE 4	VAR cScheme			AS CHARACTER			OPTIONAL
	PARAMTYPE 5	VAR lNegrito		AS LOGICAL				OPTIONAL DEFAULT .F.
	PARAMTYPE 6	VAR lItalico		AS LOGICAL				OPTIONAL DEFAULT .F.
	PARAMTYPE 7	VAR lSublinhado		AS LOGICAL				OPTIONAL DEFAULT .F.
	PARAMTYPE 8	VAR lTachado		AS LOGICAL				OPTIONAL DEFAULT .F.
	PARAMTYPE 9	VAR cLocal			AS CHARACTER			OPTIONAL DEFAULT "/xmlns:styleSheet/xmlns:fonts"	///xmlns:styleSheet/xmlns:dxfs/xmlns:dxf[last()]

	If ValType(cCorRGB)=="C" .and. Len(cCorRGB)==6
		cCorRGB	:= "FF"+cCorRGB
	Endif
	//Busca se já existe chave
	cChave	:= PadR(cValToChar(nTamanho)+"|"+cValToChar(cCorRGB)+"|"+cValToChar(cNome)+"|"+cValToChar(cfamily)+"|"+cValToChar(cScheme)+"|"+cValToChar(lNegrito)+"|"+cValToChar(lItalico)+"|"+cValToChar(lSublinhado)+"|"+cValToChar(lTachado)+"|"+cValToChar(cLocal),200)
	If (::cAliasChv)->(DbSeek("FONTE     "+cChave))
		return (::cAliasChv)->ID
	Endif

	If cLocal=="/xmlns:styleSheet/xmlns:fonts"
		nTamFonts	:= Val(::oStyle:XPathGetAtt(cLocal,"count"))+1
		::oStyle:XPathSetAtt(cLocal,"count",cValToChar(nTamFonts))
	Endif
	RecLock(::cAliasChv,.T.)
	(::cAliasChv)->TIPO		:= "FONTE     "
	(::cAliasChv)->CHAVE	:= cChave
	(::cAliasChv)->ID		:= nTamFonts-1
	MsUnLock()

	::oStyle:XPathAddNode( cLocal, "font", "" )
	If lNegrito
		::oStyle:XPathAddNode( cLocal+"/xmlns:font[last()]", "b", "" )
	Endif
	If lItalico
		::oStyle:XPathAddNode( cLocal+"/xmlns:font[last()]", "i", "" )
	Endif
	If lTachado
		::oStyle:XPathAddNode( cLocal+"/xmlns:font[last()]", "strike", "" )
	Endif
	If lSublinhado
		::oStyle:XPathAddNode( cLocal+"/xmlns:font[last()]", "u", "" )
	Endif

	::oStyle:XPathAddNode( cLocal+"/xmlns:font[last()]", "sz", "" )
	::oStyle:XPathAddAtt( cLocal+"/xmlns:font[last()]/xmlns:sz", "val"	, cValToChar(nTamanho) )

	::oStyle:XPathAddNode( cLocal+"/xmlns:font[last()]", "color", "" )
	If ValType(cCorRGB)=="N"
		::oStyle:XPathAddAtt( cLocal+"/xmlns:font[last()]/xmlns:color", "indexed"	, cValToChar(cCorRGB) )
	Else
		If ValType(cCorRGB)=="C" .and. Len(cCorRGB)==6
			cCorRGB	:= "FF"+cCorRGB
		Endif
		::oStyle:XPathAddAtt( cLocal+"/xmlns:font[last()]/xmlns:color", "rgb"	, cCorRGB )
	Endif

	::oStyle:XPathAddNode( cLocal+"/xmlns:font[last()]", "name", "" )
	::oStyle:XPathAddAtt( cLocal+"/xmlns:font[last()]/xmlns:name", "val"	, cNome )

	If !Empty(cfamily)
		::oStyle:XPathAddNode( cLocal+"/xmlns:font[last()]", "family", "" )
		::oStyle:XPathAddAtt( cLocal+"/xmlns:font[last()]/xmlns:family", "val"	, cfamily )
	Endif
	/* pag 2525
	0 Not applicable.
	1 Roman
	2 Swiss
	3 Modern
	4 Script
	5 Decorative
	*/
	If !Empty(cScheme)
		::oStyle:XPathAddNode( cLocal+"/xmlns:font[last()]", "scheme", "" )
		::oStyle:XPathAddAtt( cLocal+"/xmlns:font[last()]/xmlns:scheme", "val"	, cScheme )
	Endif
return nTamFonts-1

/*/{Protheus.doc} CorPreenc
Adiciona cor de preenchimento para ser usado no estilo das células
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param cBgCor, characters, (Opcional) Cor em Alpha+RGB do preenchimento
@param cFgCor, characters, (Opcional) Cor em Aplha+RGB do fundo
@param cType, characters, (Opcional) tipo de preenchimento(padrão solid)
@type method

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
/*/
METHOD CorPreenc(cFgCor,cBgCor,cType,cLocal) Class YExcel
	Local nPos
	Local cChave
	Default cType	:= "solid"
	PARAMTYPE 3	VAR cLocal			AS CHARACTER			OPTIONAL DEFAULT "/xmlns:styleSheet/xmlns:fills"	///xmlns:styleSheet/xmlns:dxfs/xmlns:dxf[last()]
	If ValType(cFgCor)=="C"
		If Len(cFgCor)==6
			cFgCor	:= "FF"+cFgCor
		Endif
	Else
		cFgCor	:= 64	//indexed="64" System Foreground n/a
	Endif
	If ValType(cBgCor)=="C"
		If Len(cBgCor)==6
			cBgCor	:= "FF"+cBgCor
		Endif
	Else
		cBgCor	:= 65	//indexed="65" System Background n/a	pag:1775
	Endif

	//Busca se já existe chave
	cChave	:= PadR(cValToChar(cFgCor)+"|"+cValToChar(cBgCor)+"|"+cValToChar(cType)+"|"+cValToChar(cLocal),200)
	If (::cAliasChv)->(DbSeek("CORPREENC "+cChave))
		return (::cAliasChv)->ID
	Endif

	::oStyle:XPathAddNode( cLocal, "fill", "" )
	nPos	:= Val(::oStyle:XPathGetAtt(cLocal,"count"))+1
	If cLocal=="/xmlns:styleSheet/xmlns:fills"
		::oStyle:XPathSetAtt(cLocal,"count",cValToChar(nPos))
	Endif

	RecLock(::cAliasChv,.T.)
	(::cAliasChv)->TIPO		:= "CORPREENC "
	(::cAliasChv)->CHAVE	:= cChave
	(::cAliasChv)->ID		:= nPos-1
	MsUnLock()

	::oStyle:XPathAddNode( cLocal+"/xmlns:fill[last()]", "patternFill", "" )
	::oStyle:XPathAddAtt( cLocal+"/xmlns:fill[last()]/xmlns:patternFill", "patternType"	, cType )
	If cType != "none"
		::oStyle:XPathAddNode( cLocal+"/xmlns:fill[last()]/xmlns:patternFill", "fgColor", "" )
		If ValType(cFgCor)=="C"
			::oStyle:XPathAddAtt( cLocal+"/xmlns:fill[last()]/xmlns:patternFill/xmlns:fgColor", "rgb"	, cFgCor )
		Elseif ValType(cFgCor)=="N"
			::oStyle:XPathAddAtt( cLocal+"/xmlns:fill[last()]/xmlns:patternFill/xmlns:fgColor", "indexed"	, cValToChar(cFgCor) )
		Else
			::oStyle:XPathAddAtt( cLocal+"/xmlns:fill[last()]/xmlns:patternFill/xmlns:fgColor", "indexed"	, "64" )	
		Endif
		::oStyle:XPathAddNode( cLocal+"/xmlns:fill[last()]/xmlns:patternFill", "bgColor", "" )
		If ValType(cBgCor)=="C"
			::oStyle:XPathAddAtt( cLocal+"/xmlns:fill[last()]/xmlns:patternFill/xmlns:bgColor", "rgb"	, cBgCor )
		Elseif ValType(cBgCor)=="N"
			::oStyle:XPathAddAtt( cLocal+"/xmlns:fill[last()]/xmlns:patternFill/xmlns:bgColor", "indexed"	, cValToChar(cBgCor) )
		Else
			::oStyle:XPathAddAtt( cLocal+"/xmlns:fill[last()]/xmlns:patternFill/xmlns:bgColor", "indexed"	, "65" )	
		Endif
	Endif

Return nPos-1

/*/{Protheus.doc} EfeitoPreenc
Adiciona cor com efeito de preenchimento
@author Saulo Gomes Martins
@since 17/05/2017
@version p11
@param nAngulo, numeric, Angulo para efeito de preenchimento
@param aCores, array, Cores de preenchimento {{CorRGB,nPerc},{"FF0000",0.5}}
@param ctype, characters, (Opcional) Tipo de efeito (linear ou path)
@param nleft, numeric, (Opcional) para efeito path posição esquerda
@param nright, numeric, (Opcional) para efeito path posição direita
@param ntop, numeric, (Opcional) para efeito path posição topo
@param nbottom, numeric, (Opcional) para efeito path posição inferior
@return numeric, Posição para criação de estilo
@type method
/*/
METHOD EfeitoPreenc(nAngulo,aCores,ctype,nleft,nright,ntop,nbottom,cLocal) Class YExcel
	Local nPos
	Local cChave
	PARAMTYPE 2	VAR ctype			AS CHARACTER			OPTIONAL DEFAULT "linear"
	PARAMTYPE 7	VAR cLocal			AS CHARACTER			OPTIONAL DEFAULT "/xmlns:styleSheet/xmlns:fills"	///xmlns:styleSheet/xmlns:dxfs/xmlns:dxf[last()]
	If ctype=="path"
		nAngulo	:= nil
	Else
		nleft	:= nil
		nright	:= nil
		ntop	:= nil
		nbottom	:= nil
	Endif
	//Busca se já existe chave
	cChave	:= PadR(cValToChar(nAngulo)+"|"+Var2Chr(aCores)+"|"+cValToChar(ctype)+"|"+cValToChar(nleft)+"|"+cValToChar(nright)+"|"+cValToChar(ntop)+"|"+cValToChar(nbottom)+"|"+cValToChar(cLocal),200)
	If (::cAliasChv)->(DbSeek("EFEITOPREE"+cChave))
		return (::cAliasChv)->ID
	Endif
	
	::oStyle:XPathAddNode( cLocal, "fill", "" )
	nPos	:= Val(::oStyle:XPathGetAtt(cLocal,"count"))+1
	If cLocal=="/xmlns:styleSheet/xmlns:fills"
		::oStyle:XPathSetAtt(cLocal,"count",cValToChar(nPos))
	Endif
	::gradientFill(nAngulo,aCores,ctype,nleft,nright,ntop,nbottom,cLocal+"/xmlns:fill[last()]")
	RecLock(::cAliasChv,.T.)
	(::cAliasChv)->TIPO		:= "EFEITOPREE"
	(::cAliasChv)->CHAVE	:= cChave
	(::cAliasChv)->ID		:= nPos-1
	MsUnLock()
Return nPos-1

METHOD gradientFill(nAngulo,aCores,ctype,nleft,nright,ntop,nbottom,cLocal) Class YExcel	//Pag 1779
	Local nCont
	PARAMTYPE 0	VAR nAngulo			AS NUMERIC		OPTIONAL
	PARAMTYPE 1	VAR aCores			AS ARRAY
	PARAMTYPE 2	VAR ctype			AS CHARACTER	OPTIONAL
	PARAMTYPE 3	VAR nleft			AS NUMERIC		OPTIONAL
	PARAMTYPE 4	VAR nright			AS NUMERIC		OPTIONAL
	PARAMTYPE 5	VAR ntop			AS NUMERIC		OPTIONAL
	PARAMTYPE 6	VAR nbottom			AS NUMERIC		OPTIONAL
	PARAMTYPE 7	VAR cLocal			AS CHARACTER	OPTIONAL DEFAULT "/xmlns:styleSheet/xmlns:fills/xmlns:fill[last()]"	///xmlns:styleSheet/xmlns:dxfs/xmlns:dxf[last()]

	::oStyle:XPathAddNode( cLocal, "gradientFill", "" )

	If ValType(ctype)!="U" .and. !(ctype $ "path|linear")
		UserException("YExcel - Tipo invalido para efeito de preenchimento.(path|linear)")
	Endif

	If ValType(ctype)!="U" .and. ctype=="path"
		::oStyle:XPathAddAtt( cLocal+"/xmlns:gradientFill[last()]", "type"	, ctype )
		Default nleft	:= 0.5
		Default nright	:= 0.5
		Default ntop	:= 0.5
		Default nbottom	:= 0.5
		If ValType(nleft)!="N" .OR. !(nleft>=0 .and. nleft<=1)
			UserException("YExcel - definir posição left em 0 a 1. Valor informado:"+cValToChar(nleft))
		Endif
		If ValType(nright)!="N" .OR. !(nright>=0 .and. nright<=1)
			UserException("YExcel - definir posição right em 0 a 1. Valor informado:"+cValToChar(nright))
		Endif
		If ValType(ntop)!="N" .OR. !(ntop>=0 .and. ntop<=1)
			UserException("YExcel - definir posição top em 0 a 1. Valor informado:"+cValToChar(ntop))
		Endif
		If ValType(nbottom)!="N" .OR. !(nbottom>=0 .and. nbottom<=1)
			UserException("YExcel - definir posição bottom em 0 a 1. Valor informado:"+cValToChar(nbottom))
		Endif
		::oStyle:XPathAddAtt( cLocal+"/xmlns:gradientFill[last()]", "left"		, cValToChar(nleft) )
		::oStyle:XPathAddAtt( cLocal+"/xmlns:gradientFill[last()]", "right"		, cValToChar(nright) )
		::oStyle:XPathAddAtt( cLocal+"/xmlns:gradientFill[last()]", "top"		, cValToChar(ntop) )
		::oStyle:XPathAddAtt( cLocal+"/xmlns:gradientFill[last()]", "bottom"	, cValToChar(nbottom) )
	Else
		Default nAngulo	:= 90
		::oStyle:XPathAddAtt( cLocal+"/xmlns:gradientFill[last()]", "degree"	, cValToChar(nAngulo) )
	Endif
	For nCont:=1 to Len(aCores)
		If !(aCores[nCont][2]>=0 .and. aCores[nCont][2]<=1)
			UserException("YExcel - Definição de cor varia de 0 a 1. Valor informado:"+cValToChar(aCores[nCont][2]))
		Endif
		If Len(aCores[nCont][1])==6
			aCores[nCont][1]	:= "FF"+aCores[nCont][1]
		Endif
		::oStyle:XPathAddNode( cLocal+"/xmlns:gradientFill[last()]", "stop", "" )
		::oStyle:XPathAddNode( cLocal+"/xmlns:gradientFill[last()]/xmlns:stop[last()]", "color", "" )
		::oStyle:XPathAddAtt( cLocal+"/xmlns:gradientFill[last()]/xmlns:stop[last()]/xmlns:color", "rgb"	, aCores[nCont][1] )
		::oStyle:XPathAddAtt( cLocal+"/xmlns:gradientFill[last()]/xmlns:stop[last()]", "position"	, cValToChar(aCores[nCont][2]) )
	Next
Return

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

@type method
@Obs pode juntar os tipo. Exemplo "ED"-Esquerda e direita

/*/
METHOD Borda(cTipo,cCor,cModelo,cLocal) Class YExcel
	Local nPos
	Local cLeft,cRight,cTop,cBottom,cDiagonal
	PARAMTYPE 0	VAR cTipo			AS CHARACTER			OPTIONAL DEFAULT ""
	PARAMTYPE 1	VAR cCor			AS CHARACTER			OPTIONAL DEFAULT "FF000000"
	PARAMTYPE 2	VAR cModelo			AS CHARACTER			OPTIONAL DEFAULT "thin"
	PARAMTYPE 3	VAR cLocal			AS CHARACTER			OPTIONAL DEFAULT "/xmlns:styleSheet/xmlns:borders"	///xmlns:styleSheet/xmlns:dxfs/xmlns:dxf[last()]
	If "E" $ cTipo .or. "L" $ cTipo
		cLeft	:= cModelo
	Endif
	If "D" $ cTipo .or. "R" $ cTipo
		cRight	:= cModelo
	Endif
	If "T" $ cTipo .or. "C" $ cTipo
		cTop	:= cModelo
	Endif
	If "B" $ cTipo
		cBottom	:= cModelo
	Endif
	If "DIAGONAL" $ cTipo
		cDiagonal	:= cModelo
	Endif

	If cTipo=="T" .or. cTipo=="ALL" .or. cTipo=="A"	//Todas bordas
		nPos	:= ::Border(cModelo,cModelo,cModelo,cModelo,,cCor,cCor,cCor,cCor,,cLocal)
	Else
		nPos	:= ::Border(cLeft,cRight,cTop,cBottom,cDiagonal,cCor,cCor,cCor,cCor,cCor,cLocal)
	Endif
Return nPos

METHOD Border(cleft,cright,ctop,cbottom,cdiagonal,cCorleft,cCorright,cCortop,cCorbottom,cCordiagonal,cLocal) Class YExcel
	Local nPos
	Local cChave
	PARAMTYPE 10	VAR cLocal			AS CHARACTER			OPTIONAL DEFAULT "/xmlns:styleSheet/xmlns:borders"	///xmlns:styleSheet/xmlns:dxfs/xmlns:dxf[last()]

	//Busca se já existe chave
	cChave	:= PadR(cValToChar(cleft)+"|"+cValToChar(cright)+"|"+cValToChar(ctop)+"|"+cValToChar(cbottom)+"|"+cValToChar(cdiagonal)+"|"+cValToChar(cCorleft)+"|"+cValToChar(cCorright)+"|"+cValToChar(cCortop)+"|"+cValToChar(cCorbottom)+"|"+cValToChar(cCordiagonal)+"|"+cValToChar(cLocal),200)
	If (::cAliasChv)->(DbSeek("BORDER    "+cChave))
		return (::cAliasChv)->ID
	Endif

	::oStyle:XPathAddNode( cLocal, "border", "" )
	nPos	:= Val(::oStyle:XPathGetAtt(cLocal,"count"))+1
	RecLock(::cAliasChv,.T.)
	(::cAliasChv)->TIPO		:= "BORDER    "
	(::cAliasChv)->CHAVE	:= cChave
	(::cAliasChv)->ID		:= nPos-1
	MsUnLock()
	If cLocal=="/xmlns:styleSheet/xmlns:borders"
		::oStyle:XPathSetAtt(cLocal,"count",cValToChar(nPos))
	Endif

	::oStyle:XPathAddNode( cLocal+"/xmlns:border[last()]", "left", "" )
	If ValType(cleft)<>"U"
		::oStyle:XPathAddAtt( cLocal+"/xmlns:border[last()]/xmlns:left", "style"	, cleft )
		If ValType(cCorleft)<>"U"
			::oStyle:XPathAddNode( cLocal+"/xmlns:border[last()]/xmlns:left", "color", "" )
			If ValType(cCorleft)=="C"
				::oStyle:XPathAddAtt( cLocal+"/xmlns:border[last()]/xmlns:left/xmlns:color", "rgb"	,cCorleft )
			ElseIf ValType(cCorleft)=="N"
				::oStyle:XPathAddAtt( cLocal+"/xmlns:border[last()]/xmlns:left/xmlns:color", "indexed"	,cValToChar(cCorleft) )
			Endif
		Endif
	Endif

	::oStyle:XPathAddNode( cLocal+"/xmlns:border[last()]", "right", "" )
	If ValType(cright)<>"U"
		::oStyle:XPathAddAtt( cLocal+"/xmlns:border[last()]/xmlns:right", "style"	, cright )
		If ValType(cCorright)<>"U"
			::oStyle:XPathAddNode( cLocal+"/xmlns:border[last()]/xmlns:right", "color", "" )
			If ValType(cCorright)=="C"
				::oStyle:XPathAddAtt( cLocal+"/xmlns:border[last()]/xmlns:right/xmlns:color", "rgb"	,cCorright )
			ElseIf ValType(cCorright)=="N"
				::oStyle:XPathAddAtt( cLocal+"/xmlns:border[last()]/xmlns:right/xmlns:color", "indexed"	,cValToChar(cCorright) )
			Endif
		Endif
	Endif

	::oStyle:XPathAddNode( cLocal+"/xmlns:border[last()]", "top", "" )
	If ValType(ctop)<>"U"
		::oStyle:XPathAddAtt( cLocal+"/xmlns:border[last()]/xmlns:top", "style"	, ctop )
		If ValType(cCortop)<>"U"
			::oStyle:XPathAddNode( cLocal+"/xmlns:border[last()]/xmlns:top", "color", "" )
			If ValType(cCortop)=="C"
				::oStyle:XPathAddAtt( cLocal+"/xmlns:border[last()]/xmlns:top/xmlns:color", "rgb"	,cCortop )
			ElseIf ValType(cCortop)=="N"
				::oStyle:XPathAddAtt( cLocal+"/xmlns:border[last()]/xmlns:top/xmlns:color", "indexed"	,cValToChar(cCortop) )
			Endif
		Endif
	Endif

	::oStyle:XPathAddNode( cLocal+"/xmlns:border[last()]", "bottom", "" )
	If ValType(cbottom)<>"U"
		::oStyle:XPathAddAtt( cLocal+"/xmlns:border[last()]/xmlns:bottom", "style"	, cbottom )
		If ValType(cCorbottom)<>"U"
			::oStyle:XPathAddNode( cLocal+"/xmlns:border[last()]/xmlns:bottom", "color", "" )
			If ValType(cCorbottom)=="C"
				::oStyle:XPathAddAtt( cLocal+"/xmlns:border[last()]/xmlns:bottom/xmlns:color", "rgb"	,cCorbottom )
			ElseIf ValType(cCorbottom)=="N"
				::oStyle:XPathAddAtt( cLocal+"/xmlns:border[last()]/xmlns:bottom/xmlns:color", "indexed"	,cValToChar(cCorbottom) )
			Endif
		Endif
	Endif

	::oStyle:XPathAddNode( cLocal+"/xmlns:border[last()]", "diagonal", "" )
	If ValType(cdiagonal)<>"U"
		::oStyle:XPathAddAtt( cLocal+"/xmlns:border[last()]/xmlns:diagonal", "style"	, cdiagonal )
		If ValType(cCordiagonal)<>"U"
			::oStyle:XPathAddNode( cLocal+"/xmlns:border[last()]/xmlns:diagonal", "color", "" )
			If ValType(cCordiagonal)=="C"
				::oStyle:XPathAddAtt( cLocal+"/xmlns:border[last()]/xmlns:diagonal/xmlns:color", "rgb"	,cCordiagonal )
			ElseIf ValType(cCordiagonal)=="N"
				::oStyle:XPathAddAtt( cLocal+"/xmlns:border[last()]/xmlns:diagonal/xmlns:color", "indexed"	,cValToChar(cCordiagonal) )
			Endif
		Endif
	Endif
Return nPos-1

/*/{Protheus.doc} AddFmtNum
Formatação para numeros
@author Saulo Gomes Martins
@since 04/03/2018
@version 1.0
@return numeric, nNumFmtId Numero do formato criado/alterado
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
@type method
@example |
:AddFmtNum(3,.T.)							//1234	1.234,000		| -1234	-1.234,000
:AddFmtNum(2,.T.,"R$ "," ",,,"-")			//1234	R$ 1.234,00		| -1234	-R$ 1.234,00	| 0	-
:AddFmtNum(2,.T.,," %")						//1234	1.234,00 %		| -1234	-1.234,00 %
:AddFmtNum(2,.T.,,"(",")")					//1234	1.234,00		| -1234	(1.234,00)
:AddFmtNum(2,.T.,,"(",")",,"Green","Red")	//1234	1.234,00 Verde	| -1234	(1.234,00) Vermelho

/*/
Method AddFmtNum(nDecimal,lMilhar,cPrefixo,cSufixo,cNegINI,cNegFim,cValorZero,cCor,cCorNeg,nNumFmtId) Class YExcel
	Local cformatCode
	Local cDecimal
	Local cNumero	:= ""
	Local cNegINIAli:= ""
	Local cNegFIMAli:= ""
	Local nPosCor
	Local aCores	:= {"Black","Blue","Cyan","Green","Magenta","Red","White","Yellow"}
	PARAMTYPE 0	VAR nDecimal			AS NUMERIC					OPTIONAL DEFAULT 0
	PARAMTYPE 1	VAR lMilhar				AS LOGICAL					OPTIONAL DEFAULT .F.
	PARAMTYPE 2	VAR cPrefixo			AS CHARACTER				OPTIONAL DEFAULT ""
	PARAMTYPE 3	VAR cSufixo				AS CHARACTER				OPTIONAL DEFAULT ""
	PARAMTYPE 4	VAR cNegINI				AS CHARACTER				OPTIONAL DEFAULT "-"
	PARAMTYPE 5	VAR cNegFIM				AS CHARACTER				OPTIONAL DEFAULT ""
	PARAMTYPE 6	VAR cValorZero			AS CHARACTER				OPTIONAL DEFAULT ""
	PARAMTYPE 7	VAR cCor				AS CHARACTER,NUMERIC		OPTIONAL DEFAULT ""
	PARAMTYPE 8	VAR cCorNeg				AS CHARACTER,NUMERIC		OPTIONAL DEFAULT ""
	PARAMTYPE 9	VAR nNumFmtId			AS NUMERIC					OPTIONAL

	If !Empty(cCor)
		If ValType(cCor)=="C"
			nPosCor	:= aScan(aCores,{|x| UPPER(x)==UPPER(cCor) })
			If nPosCor==0
				UserException("YExcel - Cor da formatação invalida ("+cCor+")")
			Else
				cCor	:= aCores[nPosCor]
			Endif
		ElseIf ValType(cCor)=="N"
			If !(cCor>=1 .AND. cCor<=56)
				UserException("YExcel - Cor da formatação invalida ("+cValToChar(cCor)+"), Cores indexado valido de 1-56.")
			Endif
			cCor	:= "Color"+cValToChar(cCor)
		Endif
	Endif
	If !Empty(cCorNeg)
		If ValType(cCorNeg)=="C"
			nPosCor	:= aScan(aCores,{|x| UPPER(x)==UPPER(cCorNeg) })
			If nPosCor==0
				UserException("YExcel - Cor da formatação invalida ("+cCorNeg+")")
			Else
				cCorNeg	:= aCores[nPosCor]
			Endif
		ElseIf ValType(cCorNeg)=="N"
			If !(cCorNeg>=1 .AND. cCorNeg<=56)
				UserException("YExcel - Cor da formatação invalida ("+cValToChar(cCorNeg)+"), Cores indexado valido de 1-56.")
			Endif
			cCorNeg	:= "Color"+cValToChar(cCorNeg)
		Endif
	Endif

	cDecimal	:= Replicate("0",nDecimal)
	If lMilhar
		cNumero	:= "#,##0"
	Else
		cNumero	:= "#"
	Endif

	If !Empty(cDecimal)
		cNumero	:= cNumero+"."+cDecimal
	Endif
	If !Empty(cPrefixo)
		cPrefixo	:= '"'+cPrefixo+'"'
		cNumero		:= cPrefixo+cNumero
	Endif
	If !Empty(cSufixo)
		cSufixo		:= '"'+cSufixo+'"'
		cNumero		:= cNumero+cSufixo
	Endif
	If !Empty(cNegINI)
		cNegINIAli	:= "_"+cNegINI
	Endif
	If !Empty(cNegFIM)
		cNegFIMAli	:= "_"+cNegFIM
	Endif
	If !Empty(cValorZero)
		cValorZero	:= '"'+cValorZero+'"'
	Else
		cValorZero	:= cNumero
	Endif
	If !Empty(cCor)
		cCor	:= "["+cCor+"]"
	Endif
	If !Empty(cCorNeg)
		cCorNeg	:= "["+cCorNeg+"]"
	Endif
	cformatCode	:= cCor+cNegINIAli+cNumero+cNegFIMAli+";"+cCorNeg+cNegINI+cNumero+cNegFIM+";"+cNegINIAli+cValorZero+cNegFIMAli+";@"

	nNumFmtId	:= ::AddFmt(cformatCode,nNumFmtId)

Return nNumFmtId	//Não retorna a posição, mas o atributo numFmtId

/*/{Protheus.doc} AddFmt
Cria um formatos personalizado ou altera um existente
@author Saulo Gomes Martins
@since 20/06/2020
@version 2.0
@param cformatCode, characters, formato personalizado
@param nNumFmtId, numeric, numFmtId para alteração
@type method
/*/
Method AddFmt(cformatCode,nNumFmtId) Class YExcel
	Local cChave
	PARAMTYPE 0	VAR cformatCode			AS CHARACTER
	PARAMTYPE 1	VAR nNumFmtId			AS NUMERIC					OPTIONAL

	cChave	:= PadR(cValToChar(cformatCode),200)
	If Empty(nNumFmtId)
		//Busca se já existe chave
		If ::oStyle:XPathHasNode("/xmlns:styleSheet/xmlns:numFmts/xmlns:numFmt[@formatCode='"+cformatCode+"']")
			Return Val(::oStyle:XPathGetAtt("/xmlns:styleSheet/xmlns:numFmts/xmlns:numFmt[@formatCode='"+cformatCode+"']","numFmtId"))
		Endif
		nNumFmtId	:= ::nNumFmtId++
	Endif
	If !::oStyle:XPathHasNode( "/xmlns:styleSheet/xmlns:numFmts/xmlns:numFmt[@numFmtId='"+cValToChar(nNumFmtId)+"']")	//Se não existe o ID
		::oStyle:XPathAddNode( "/xmlns:styleSheet/xmlns:numFmts", "numFmt", "" )
		::oStyle:XPathAddAtt( "/xmlns:styleSheet/xmlns:numFmts/xmlns:numFmt[last()]", "numFmtId"	, cValToChar(nNumFmtId) )
		::oStyle:XPathSetAtt("/xmlns:styleSheet/xmlns:numFmts","count",cValToChar(Val(::oStyle:XPathGetAtt("/xmlns:styleSheet/xmlns:numFmts","count"))+1))
		::oStyle:XPathAddAtt( "/xmlns:styleSheet/xmlns:numFmts/xmlns:numFmt[@numFmtId='"+cValToChar(nNumFmtId)+"']", "formatCode"	, "" )
	Endif
	::oStyle:XPathSetAtt( "/xmlns:styleSheet/xmlns:numFmts/xmlns:numFmt[@numFmtId='"+cValToChar(nNumFmtId)+"']", "formatCode"	, cformatCode )
Return nNumFmtId

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
@type method

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
METHOD AddStyles(numFmtId,fontId,fillId,borderId,xfId,aValores,aOutrosAtributos) Class YExcel
	Local cLocal	:= "/xmlns:styleSheet/xmlns:cellXfs"
	Local nPos
	Local cChave
	PARAMTYPE 0	VAR numFmtId			AS NUMERIC		OPTIONAL
	PARAMTYPE 1	VAR fontId				AS NUMERIC		OPTIONAL
	PARAMTYPE 2	VAR fillId				AS NUMERIC		OPTIONAL
	PARAMTYPE 3	VAR borderId			AS NUMERIC		OPTIONAL
	PARAMTYPE 4	VAR xfId				AS NUMERIC		OPTIONAL DEFAULT 0
	PARAMTYPE 5	VAR aValores			AS ARRAY		OPTIONAL DEFAULT {}
	PARAMTYPE 6	VAR aOutrosAtributos	AS ARRAY		OPTIONAL DEFAULT {}
	If ValType(fontId)=="N" .AND. (fontId+1)>Val(::oStyle:XPathGetAtt("/xmlns:styleSheet/xmlns:fonts","count"))
		UserException("YExcel - Fonte informada("+cValToChar(fontId)+") não definido. Utilize o indice informado pelo metodo :AddFont()")
	ElseIf ValType(fillId)=="N" .AND. (fillId+1)>Val(::oStyle:XPathGetAtt("/xmlns:styleSheet/xmlns:fills","count"))
		UserException("YExcel - Cor Preenchimento informado("+cValToChar(fillId)+") não definido. Utilize o indice informado pelo metodo :CorPreenc()")
	ElseIf ValType(borderId)=="N" .AND. (borderId+1)>Val(::oStyle:XPathGetAtt("/xmlns:styleSheet/xmlns:borders","count"))
		UserException("YExcel - Borda informada("+cValToChar(borderId)+") não definido. Utilize o indice informado pelo metodo :Borda()")
	Endif

	// Busca se já existe chave
	cChave	:= PadR(cValToChar(numFmtId)+"|"+cValToChar(fontId)+"|"+cValToChar(fillId)+"|"+cValToChar(borderId)+"|"+cValToChar(xfId)+"|"+Var2Chr(aValores)+"|"+Var2Chr(aOutrosAtributos),200)
	If (::cAliasChv)->(DbSeek("STYLE     "+cChave))
		return (::cAliasChv)->ID
	Endif

	::oStyle:XPathAddNode( cLocal, "xf", "" )
	nPos	:= Val(::oStyle:XPathGetAtt(cLocal,"count"))+1
	::oStyle:XPathSetAtt(cLocal,"count",cValToChar(nPos))
	::nQtdStyle	:= nPos

	RecLock(::cAliasChv,.T.)
	(::cAliasChv)->TIPO		:= "STYLE     "
	(::cAliasChv)->CHAVE	:= cChave
	(::cAliasChv)->ID		:= nPos-1
	MsUnLock()

	::SetStyFmt(nPos-1,numFmtId)

	::SetStyFont(nPos-1,fontId)

	::SetStyFill(nPos-1,fillId)

	::SetStyborder(nPos-1,borderId)

	::SetStyxf(nPos-1,xfId)

	::SetStyaValores(nPos-1,aValores)

	::SetStyaOutrosAtributos(nPos-1,aOutrosAtributos)

return nPos-1

/*/{Protheus.doc} YExcel_Style
Cria estilos orientado a objeto
@type class
@version 1.0
@author Saulo Gomes Martins
@since 27/06/2020
/*/
Class YExcel_Style
	Data oPai
	Data cClassName
	Data aFilhos
	Data oExcel
	Data nStyle
	Data numFmtId
	Data fontId
	Data fillId
	Data borderId
	Data xfId
	Data aValores
	Data aOutrosAtributos
	Method New()
	Method ClassName()
	Method SetnumFmt()
	Method Setfont()
	Method Setfill()
	Method Setborder()
	Method SetxfId()
	Method SetaValores()
	Method SetAtt()
	Method GetnumFmt()
	Method Getfont()
	Method Getfill()
	Method Getborder()
	Method GetxfId()
	Method GetaValores()
	Method GetAtt()
	Method GetId()
EndClass

Method New(oPai,oExcel) Class YExcel_Style
	PARAMTYPE 0	VAR oPai					AS OBJECT	OPTIONAL
	PARAMTYPE 0	VAR oExcel					AS OBJECT
	::oPai			:= oPai
	::aFilhos		:= {}
	::oExcel		:= oExcel
	If ValType(::oPai)=="O"
		AADD(::oPai:aFilhos,self)
	Endif
	::cClassName	:= "YEXCEL_STYLE"
Return Self

METHOD ClassName() Class YExcel_Style
Return "YEXCEL_STYLE"

Method SetnumFmt(numFmtId,lProprio) Class YExcel_Style
	Local nCont
	Default lProprio	:= .T.	//Se .T. é um estilo proprio, se não é do pai 
	If lProprio
		::numFmtId	:= numFmtId
	Endif
	If ValType(::nStyle)!="U"	//Se já criou o estilo, modifica
		// ::oExcel:SetStyFmt(::nStyle,numFmtId)
		::nStyle	:= nil
	Endif
	For nCont:=1 to Len(::aFilhos)	//Passa a herança para frente se foi herdada
		If ValType(::aFilhos[nCont]:numFmtId)=="U" .AND.ValType(::aFilhos[nCont]:nStyle)!="U"	//Se o filho não tem formato proprio(herdou do pai)
			::aFilhos[nCont]:SetnumFmt(numFmtId,.F.)
		Endif
	Next
Return self

Method Setfont(fontId,lProprio) Class YExcel_Style
	Local nCont
	Default lProprio	:= .T.	//Se .T. é um estilo proprio, se não é do pai 
	If lProprio
		::fontId	:= fontId
	Endif
	If ValType(::nStyle)!="U"	//Se já criou o estilo, modifica
		// ::oExcel:SetStyFont(::nStyle,fontId)
		::nStyle	:= nil
	Endif
	For nCont:=1 to Len(::aFilhos)	//Passa a herança para frente se foi herdada
		If ValType(::aFilhos[nCont]:fontId)=="U" .AND.ValType(::aFilhos[nCont]:nStyle)!="U"	//Se o filho não tem formato proprio(herdou do pai)
			::aFilhos[nCont]:Setfont(fontId,.F.)
		Endif
	Next
Return self

Method Setfill(fillId,lProprio) Class YExcel_Style
	Local nCont
	Default lProprio	:= .T.	//Se .T. é um estilo proprio, se não é do pai 
	If lProprio
		::fillId	:= fillId
	Endif
	If ValType(::nStyle)!="U"	//Se já criou o estilo, modifica
		// ::oExcel:SetStyFill(::nStyle,fillId)
		::nStyle	:= nil
	Endif
	For nCont:=1 to Len(::aFilhos)	//Passa a herança para frente se foi herdada
		If ValType(::aFilhos[nCont]:fillId)=="U" .AND.ValType(::aFilhos[nCont]:nStyle)!="U"	//Se o filho não tem formato proprio(herdou do pai)
			::aFilhos[nCont]:Setfill(fillId,.F.)
		Endif
	Next
Return self

Method Setborder(borderId,lProprio) Class YExcel_Style
	Local nCont
	Default lProprio	:= .T.	//Se .T. é um estilo proprio, se não é do pai 
	If lProprio
		::borderId	:= borderId
	Endif
	If ValType(::nStyle)!="U"	//Se já criou o estilo, modifica
		// ::oExcel:SetStyborder(::nStyle,borderId)
		::nStyle	:= nil
	Endif
	For nCont:=1 to Len(::aFilhos)	//Passa a herança para frente se foi herdada
		If ValType(::aFilhos[nCont]:borderId)=="U" .AND.ValType(::aFilhos[nCont]:nStyle)!="U"	//Se o filho não tem formato proprio(herdou do pai)
			::aFilhos[nCont]:Setborder(borderId,.F.)
		Endif
	Next
Return self

Method SetxfId(xfId,lProprio) Class YExcel_Style
	Local nCont
	Default lProprio	:= .T.	//Se .T. é um estilo proprio, se não é do pai 
	If lProprio
		::xfId	:= xfId
	Endif
	If ValType(::nStyle)!="U"	//Se já criou o estilo, modifica
		// ::oExcel:SetStyxf(::nStyle,xfId)
		::nStyle	:= nil
	Endif
	For nCont:=1 to Len(::aFilhos)	//Passa a herança para frente se foi herdada
		If ValType(::aFilhos[nCont]:xfId)=="U" .AND.ValType(::aFilhos[nCont]:nStyle)!="U"	//Se o filho não tem formato proprio(herdou do pai)
			::aFilhos[nCont]:SetxfId(xfId,.F.)
		Endif
	Next
Return self

Method SetaValores(aValores,lProprio) Class YExcel_Style
	Local nCont
	Default lProprio	:= .T.	//Se .T. é um estilo proprio, se não é do pai 
	If lProprio
		::aValores	:= aValores
	Endif
	If ValType(::nStyle)!="U"	//Se já criou o estilo, modifica
		// ::oExcel:SetStyaValores(::nStyle,aValores)
		::nStyle	:= nil
	Endif
	For nCont:=1 to Len(::aFilhos)	//Passa a herança para frente se foi herdada
		If ValType(::aFilhos[nCont]:aValores)=="U" .AND.ValType(::aFilhos[nCont]:nStyle)!="U"	//Se o filho não tem formato proprio(herdou do pai)
			::aFilhos[nCont]:SetaValores(aValores,.F.)
		Endif
	Next
Return self

Method SetAtt(aOutrosAtributos,lProprio) Class YExcel_Style
	Local nCont
	Default lProprio	:= .T.	//Se .T. é um estilo proprio, se não é do pai 
	If lProprio
		::aOutrosAtributos	:= aOutrosAtributos
	Endif
	If ValType(::nStyle)!="U"	//Se já criou o estilo, modifica
		// ::oExcel:SetStyaOutrosAtributos(::nStyle,aOutrosAtributos)
		::nStyle	:= nil
	Endif
	For nCont:=1 to Len(::aFilhos)	//Passa a herança para frente se foi herdada
		If ValType(::aFilhos[nCont]:aOutrosAtributos)=="U" .AND.ValType(::aFilhos[nCont]:nStyle)!="U"	//Se o filho não tem formato proprio(herdou do pai)
			::aFilhos[nCont]:SetAtt(aOutrosAtributos,.F.)
		Endif
	Next
Return self

Method GetnumFmt() Class YExcel_Style
	If ValType(::numFmtId)=="U" .AND. ValType(::oPai)=="O"
		Return ::oPai:GetnumFmt()
	Endif
Return ::numFmtId

Method Getfont() Class YExcel_Style
	If ValType(::fontId)=="U" .AND. ValType(::oPai)=="O"
		Return ::oPai:Getfont()
	Endif
Return ::fontId

Method Getfill() Class YExcel_Style
	If ValType(::fillId)=="U" .AND. ValType(::oPai)=="O"
		Return ::oPai:Getfill()
	Endif
Return ::fillId

Method Getborder() Class YExcel_Style
	If ValType(::borderId)=="U" .AND. ValType(::oPai)=="O"
		Return ::oPai:Getborder()
	Endif
Return ::borderId

Method GetxfId() Class YExcel_Style
	If ValType(::xfId)=="U" .AND. ValType(::oPai)=="O"
		Return ::oPai:GetxfId()
	Endif
Return ::xfId

Method GetaValores() Class YExcel_Style
	If ValType(::aValores)=="U" .AND. ValType(::oPai)=="O"
		Return ::oPai:GetaValores()
	Endif
Return ::aValores

Method GetAtt() Class YExcel_Style
	If ValType(::aOutrosAtributos)=="U" .AND. ValType(::oPai)=="O"
		Return ::oPai:GetAtt()
	Endif
Return ::aOutrosAtributos

Method GetId() Class YExcel_Style
	If ValType(::nStyle)=="U"
		::nStyle	:= ::oExcel:CreateStyle(,::GetnumFmt(),::Getfont(),::Getfill() ,::Getborder(),::GetxfId(),::GetaValores(),::GetAtt())
	Endif
Return  ::nStyle

/*/{Protheus.doc} YExcel::NewStyle
Adiciona estilo com herança de outro estilo
@type method
@version 1.0
@author Saulo Gomes Martins
@since 26/06/2020
@param nStyle, numeric, Estilo para ser usado como herança
@return numeric, nPos posição do estilo criado
/*/
Method NewStyle(oStyleClone) Class YExcel
	Local oObj := YExcel_Style():New(oStyleClone,self)
	AADD(::aCleanObj,oObj)
Return oObj

// Method GetStyFmt(nStyle) Class YExcel
// 	Local cLocal	:= "/xmlns:styleSheet/xmlns:cellXfs"
// 	Local cXPath	:= cLocal+"/xmlns:xf["+cValToChar(nStyle+1)+"]"
// 	PARAMTYPE 0	VAR nStyle				AS NUMERIC
// 	if ::oStyle:XPathGetAtt(cXPath,"applyNumberFormat")=="1"
// 		Return ::oStyle:XPathGetAtt(cXPath,"numFmtId")
// 	Endif
// Return ""

/*/{Protheus.doc} YExcel::SetStyFmt
Alterar Fmt do estilo já criado
@type method
@version 1.0
@author Saulo Gomes Martins
@since 27/06/2020
@param nStyle, numeric, id do estilo
@param numFmtId, numeric, id do novo fmt
@return object, self
/*/
Method SetStyFmt(nStyle,numFmtId) Class YExcel
	Local cLocal	:= "/xmlns:styleSheet/xmlns:cellXfs"
	Local cXPath	:= cLocal+"/xmlns:xf["+cValToChar(nStyle+1)+"]"
	Local aChave,cChave
	PARAMTYPE 0	VAR nStyle				AS NUMERIC
	PARAMTYPE 1	VAR numFmtId			AS NUMERIC		OPTIONAL
	
	If nStyle+1>::nQtdStyle
		UserException("YExcel - Estilo informado("+cValToChar(nStyle)+") não definido. Utilize o indice informado pelo metodo AddStyles")
	Endif

	::oStyle:XPathDelAtt(cXPath,"applyNumberFormat")
	::oStyle:XPathDelAtt(cXPath,"numFmtId")
	If ValType(numFmtId)=="U"
		::oStyle:XPathAddAtt(cXPath,"applyNumberFormat","0")
		::oStyle:XPathAddAtt(cXPath,"numFmtId","0")
	Else
		::oStyle:XPathAddAtt(cXPath,"applyNumberFormat","1")
		::oStyle:XPathAddAtt(cXPath,"numFmtId",cValToChar(numFmtId))
	Endif
	// Altera chave
	(::cAliasChv)->(DbSetOrder(2))
	If (::cAliasChv)->(DbSeek("STYLE     "+Str(nStyle,7)))
		aChave	:= Separa(Alltrim((::cAliasChv)->CHAVE),"|")
		cChave	:= PadR(cValToChar(numFmtId)+"|"+aChave[2]+"|"+aChave[3]+"|"+aChave[4]+"|"+aChave[5]+"|"+aChave[6]+"|"+aChave[7],200)
		RecLock(::cAliasChv,.F.)
		(::cAliasChv)->CHAVE	:= cChave
		MsUnLock()
	Endif
	(::cAliasChv)->(DbSetOrder(1))
Return self

/*/{Protheus.doc} YExcel::SetStyFont
Alterar a fonte do estilo já criado
@type method
@version 1.0
@author Saulo Gomes Martins
@since 27/06/2020
@param nStyle, numeric, id do estilo
@param fontId, numeric, id do novo fontId
@return object, self
/*/
Method SetStyFont(nStyle,fontId) Class YExcel
	Local cLocal	:= "/xmlns:styleSheet/xmlns:cellXfs"
	Local cXPath	:= cLocal+"/xmlns:xf["+cValToChar(nStyle+1)+"]"
	Local aChave,cChave
	PARAMTYPE 0	VAR nStyle				AS NUMERIC
	PARAMTYPE 1	VAR fontId			AS NUMERIC		OPTIONAL
	
	If nStyle+1>::nQtdStyle
		UserException("YExcel - Estilo informado("+cValToChar(nStyle)+") não definido. Utilize o indice informado pelo metodo AddStyles")
	ElseIf ValType(fontId)=="N" .AND. (fontId+1)>Val(::oStyle:XPathGetAtt("/xmlns:styleSheet/xmlns:fonts","count"))
		UserException("YExcel - Fonte informada("+cValToChar(fontId)+") não definido. Utilize o indice informado pelo metodo :AddFont()")
	Endif

	::oStyle:XPathDelAtt(cXPath,"applyFont")
	::oStyle:XPathDelAtt(cXPath,"fontId")
	If ValType(fontId)=="U"
		::oStyle:XPathAddAtt(cXPath,"applyFont","0")
		::oStyle:XPathAddAtt(cXPath,"fontId","0")
	Else
		::oStyle:XPathAddAtt(cXPath,"applyFont","1")
		::oStyle:XPathAddAtt(cXPath,"fontId",cValToChar(fontId))
	Endif
	// Altera chave
	(::cAliasChv)->(DbSetOrder(2))
	If (::cAliasChv)->(DbSeek("STYLE     "+Str(nStyle,7)))
		aChave	:= Separa(Alltrim((::cAliasChv)->CHAVE),"|")
		cChave	:= PadR(aChave[1]+"|"+cValToChar(fontId)+"|"+aChave[3]+"|"+aChave[4]+"|"+aChave[5]+"|"+aChave[6]+"|"+aChave[7],200)
		RecLock(::cAliasChv,.F.)
		(::cAliasChv)->CHAVE	:= cChave
		MsUnLock()
	Endif
	(::cAliasChv)->(DbSetOrder(1))
Return self

/*/{Protheus.doc} YExcel::SetStyFill
Alterar preenchimento de fundo do estilo já criado
@type method
@version 1.0
@author Saulo Gomes Martins
@since 27/06/2020
@param nStyle, numeric, id do estilo
@param fontId, numeric, id do novo fontId
@return object, self
/*/
Method SetStyFill(nStyle,fillId) Class YExcel
	Local cLocal	:= "/xmlns:styleSheet/xmlns:cellXfs"
	Local cXPath	:= cLocal+"/xmlns:xf["+cValToChar(nStyle+1)+"]"
	Local aChave,cChave
	PARAMTYPE 0	VAR nStyle			AS NUMERIC
	PARAMTYPE 1	VAR fillId			AS NUMERIC		OPTIONAL
	
	If nStyle+1>::nQtdStyle
		UserException("YExcel - Estilo informado("+cValToChar(nStyle)+") não definido. Utilize o indice informado pelo metodo AddStyles")
	ElseIf ValType(fillId)=="N" .AND. (fillId+1)>Val(::oStyle:XPathGetAtt("/xmlns:styleSheet/xmlns:fills","count"))
		UserException("YExcel - Cor Preenchimento informado("+cValToChar(fillId)+") não definido. Utilize o indice informado pelo metodo :CorPreenc()")
	Endif

	::oStyle:XPathDelAtt(cXPath,"applyFill")
	::oStyle:XPathDelAtt(cXPath,"fillId")
	If ValType(fillId)=="U"
		::oStyle:XPathAddAtt(cXPath,"applyFill","0")
		::oStyle:XPathAddAtt(cXPath,"fillId","0")
	Else
		::oStyle:XPathAddAtt(cXPath,"applyFill","1")
		::oStyle:XPathAddAtt(cXPath,"fillId",cValToChar(fillId))
	Endif
	// Altera chave
	(::cAliasChv)->(DbSetOrder(2))
	If (::cAliasChv)->(DbSeek("STYLE     "+Str(nStyle,7)))
		aChave	:= Separa(Alltrim((::cAliasChv)->CHAVE),"|")
		cChave	:= PadR(aChave[1]+"|"+aChave[2]+"|"+cValToChar(fillId)+"|"+aChave[4]+"|"+aChave[5]+"|"+aChave[6]+"|"+aChave[7],200)
		RecLock(::cAliasChv,.F.)
		(::cAliasChv)->CHAVE	:= cChave
		MsUnLock()
	Endif
	(::cAliasChv)->(DbSetOrder(1))
Return self

/*/{Protheus.doc} YExcel::SetStyborder
Alterar a borda do estilo já criado
@type method
@version 1.0
@author Saulo Gomes Martins
@since 27/06/2020
@param nStyle, numeric, id do estilo
@param fontId, numeric, id do novo fontId
@return object, self
/*/
Method SetStyborder(nStyle,borderId) Class YExcel
	Local cLocal	:= "/xmlns:styleSheet/xmlns:cellXfs"
	Local cXPath	:= cLocal+"/xmlns:xf["+cValToChar(nStyle+1)+"]"
	Local aChave,cChave
	PARAMTYPE 0	VAR nStyle			AS NUMERIC
	PARAMTYPE 1	VAR borderId		AS NUMERIC		OPTIONAL
	
	If nStyle+1>::nQtdStyle
		UserException("YExcel - Estilo informado("+cValToChar(nStyle)+") não definido. Utilize o indice informado pelo metodo AddStyles")
	ElseIf ValType(borderId)=="N" .AND. (borderId+1)>Val(::oStyle:XPathGetAtt("/xmlns:styleSheet/xmlns:borders","count"))
		UserException("YExcel - Borda informada("+cValToChar(borderId)+") não definido. Utilize o indice informado pelo metodo :Borda()")
	Endif

	::oStyle:XPathDelAtt(cXPath,"applyBorder")
	::oStyle:XPathDelAtt(cXPath,"borderId")
	If ValType(borderId)=="U"
		::oStyle:XPathAddAtt(cXPath,"applyBorder","0")
		::oStyle:XPathAddAtt(cXPath,"borderId","0")
	Else
		::oStyle:XPathAddAtt(cXPath,"applyBorder","1")
		::oStyle:XPathAddAtt(cXPath,"borderId",cValToChar(borderId))
	Endif
	// Altera chave
	(::cAliasChv)->(DbSetOrder(2))
	If (::cAliasChv)->(DbSeek("STYLE     "+Str(nStyle,7)))
		aChave	:= Separa(Alltrim((::cAliasChv)->CHAVE),"|")
		cChave	:= PadR(aChave[1]+"|"+aChave[2]+"|"+aChave[3]+"|"+cValToChar(borderId)+"|"+aChave[5]+"|"+aChave[6]+"|"+aChave[7],200)
		RecLock(::cAliasChv,.F.)
		(::cAliasChv)->CHAVE	:= cChave
		MsUnLock()
	Endif
	(::cAliasChv)->(DbSetOrder(1))
Return self

/*/{Protheus.doc} YExcel::SetStyborder
Alterar xf do estilo já criado
@type method
@version 1.0
@author Saulo Gomes Martins
@since 27/06/2020
@param nStyle, numeric, id do estilo
@param fontId, numeric, id do novo fontId
@return object, self
/*/
Method SetStyxf(nStyle,xfId) Class YExcel
	Local cLocal	:= "/xmlns:styleSheet/xmlns:cellXfs"
	Local cXPath	:= cLocal+"/xmlns:xf["+cValToChar(nStyle+1)+"]"
	Local aChave,cChave
	PARAMTYPE 0	VAR nStyle		AS NUMERIC
	PARAMTYPE 1	VAR xfId		AS NUMERIC		OPTIONAL
	
	If nStyle+1>::nQtdStyle
		UserException("YExcel - Estilo informado("+cValToChar(nStyle)+") não definido. Utilize o indice informado pelo metodo AddStyles")
	Endif

	::oStyle:XPathDelAtt(cXPath,"xfId")
	If ValType(xfId)=="U"
		::oStyle:XPathAddAtt(cXPath,"xfId","0")
	Else
		::oStyle:XPathAddAtt(cXPath,"xfId",cValToChar(xfId))
	Endif
	// Altera chave
	(::cAliasChv)->(DbSetOrder(2))
	If (::cAliasChv)->(DbSeek("STYLE     "+Str(nStyle,7)))
		aChave	:= Separa(Alltrim((::cAliasChv)->CHAVE),"|")
		cChave	:= PadR(aChave[1]+"|"+aChave[2]+"|"+aChave[3]+"|"+aChave[4]+"|"+cValToChar(xfId)+"|"+aChave[6]+"|"+aChave[7],200)
		RecLock(::cAliasChv,.F.)
		(::cAliasChv)->CHAVE	:= cChave
		MsUnLock()
	Endif
	(::cAliasChv)->(DbSetOrder(1))
Return self

/*/{Protheus.doc} YExcel::SetStyborder
Alterar aValores estilo já criado
@type method
@version 1.0
@author Saulo Gomes Martins
@since 27/06/2020
@param nStyle, numeric, id do estilo
@param fontId, numeric, id do novo fontId
@return object, self
/*/
Method SetStyaValores(nStyle,aValores) Class YExcel
	Local cLocal	:= "/xmlns:styleSheet/xmlns:cellXfs"
	Local cXPath	:= cLocal+"/xmlns:xf["+cValToChar(nStyle+1)+"]"
	Local nCont,nCont2
	Local aListAtt
	Local aChave,cChave
	PARAMTYPE 0	VAR nStyle			AS NUMERIC
	PARAMTYPE 1	VAR aValores		AS ARRAY		OPTIONAL
	
	If nStyle+1>::nQtdStyle
		UserException("YExcel - Estilo informado("+cValToChar(nStyle)+") não definido. Utilize o indice informado pelo metodo AddStyles")
	Endif

	If ValType(aValores)=="U"
		::oStyle:XPathDelNode(cXPath)
	Else
		While ::oStyle:XPathHasNode(cXPath+"/xmlns:alignment[1]")
			::oStyle:XPathDelNode(cXPath+"/xmlns:alignment[1]")
		EndDo
		For nCont:=1 to Len(aValores)
			::oStyle:XPathAddNode( cXPath, aValores[nCont]:GetNome(), "" )
			If aValores[nCont]:GetNome()=="alignment"
				If aScan(self:oStyle:XPathGetAttArray(cXPath),{|x| x[1]=="applyAlignment"})>0
					::oStyle:XPathSetAtt( cXPath, "applyAlignment"		, "1" )
				Else
					::oStyle:XPathAddAtt( cXPath, "applyAlignment"		, "1" )
				Endif
			Endif
			aValores[nCont]:oAtributos:List(@aListAtt)
			For nCont2:=1 to Len(aListAtt)
				If aScan(self:oStyle:XPathGetAttArray(cXPath+"/xmlns:"+aValores[nCont]:GetNome()),{|x| x[1]==aListAtt[nCont2][1] })>0
					::oStyle:XPathSetAtt( cXPath+"/xmlns:"+aValores[nCont]:GetNome(), aListAtt[nCont2][1]			, cValToChar(aListAtt[nCont2][2]) )
				Else
					::oStyle:XPathAddAtt( cXPath+"/xmlns:"+aValores[nCont]:GetNome(), aListAtt[nCont2][1]			, cValToChar(aListAtt[nCont2][2]) )
				Endif
			Next
		Next
	Endif
	// Altera chave
	(::cAliasChv)->(DbSetOrder(2))
	If (::cAliasChv)->(DbSeek("STYLE     "+Str(nStyle,7)))
		aChave	:= Separa(Alltrim((::cAliasChv)->CHAVE),"|")
		cChave	:= PadR(aChave[1]+"|"+aChave[2]+"|"+aChave[3]+"|"+aChave[4]+"|"+aChave[5]+"|"+Var2Chr(aValores)+"|"+aChave[7],200)
		RecLock(::cAliasChv,.F.)
		(::cAliasChv)->CHAVE	:= cChave
		MsUnLock()
	Endif
	(::cAliasChv)->(DbSetOrder(1))
Return self

/*/{Protheus.doc} YExcel::SetStyborder
Alterar aOutrosAtributos estilo já criado
@type method
@version 1.0
@author Saulo Gomes Martins
@since 27/06/2020
@param nStyle, numeric, id do estilo
@param fontId, numeric, id do novo fontId
@return object, self
/*/
Method SetStyaOutrosAtributos(nStyle,aOutrosAtributos) Class YExcel
	Local cLocal	:= "/xmlns:styleSheet/xmlns:cellXfs"
	Local cXPath	:= cLocal+"/xmlns:xf["+cValToChar(nStyle+1)+"]"
	Local nCont
	Local aChave,cChave
	PARAMTYPE 0	VAR nStyle					AS NUMERIC
	PARAMTYPE 1	VAR aOutrosAtributos		AS ARRAY		OPTIONAL
	
	If nStyle+1>::nQtdStyle
		UserException("YExcel - Estilo informado("+cValToChar(nStyle)+") não definido. Utilize o indice informado pelo metodo AddStyles")
	Endif

	If ValType(aOutrosAtributos)!="U"
		For nCont:=1 to Len(aOutrosAtributos)
			::oStyle:XPathDelAtt(cXPath,aOutrosAtributos[nCont][1])
			If ValType(aOutrosAtributos[nCont][2])=="U"
			Else
				::oStyle:XPathAddAtt( cXPath, aOutrosAtributos[nCont][1]	, cValToChar(aOutrosAtributos[nCont][2]) )
			Endif
		Next
	Endif
	// Altera chave
	(::cAliasChv)->(DbSetOrder(2))
	If (::cAliasChv)->(DbSeek("STYLE     "+Str(nStyle,7)))
		aChave	:= Separa(Alltrim((::cAliasChv)->CHAVE),"|")
		cChave	:= PadR(aChave[1]+"|"+aChave[2]+"|"+aChave[3]+"|"+aChave[4]+"|"+aChave[5]+"|"+aChave[6]+"|"+Var2Chr(aOutrosAtributos),200)
		RecLock(::cAliasChv,.F.)
		(::cAliasChv)->CHAVE	:= cChave
		MsUnLock()
	Endif
	(::cAliasChv)->(DbSetOrder(1))
Return self

/*/{Protheus.doc} YExcel::CreateStyle
Adiciona estilo com herança de outro estilo
@type method
@version 1.0
@author Saulo Gomes Martins
@since 26/06/2020
@param nStyle, numeric, Estilo para ser usado como herança
@return numeric, nPos posição do estilo criado
/*/
Method CreateStyle(nStyle,numFmtId,fontId,fillId,borderId,xfId,aValores,aOutrosAtributos) Class YExcel
	Local cLocal	:= "/xmlns:styleSheet/xmlns:cellXfs"
	Local cXPath
	Local aChildren
	Local nCont
	If ValType(nStyle)=="U"	//Cria estilo sem herança
		Return ::AddStyles(numFmtId,fontId,fillId,borderId,xfId,aValores,aOutrosAtributos)
	Endif
	PARAMTYPE 0	VAR nStyle				AS NUMERIC
	PARAMTYPE 1	VAR numFmtId			AS NUMERIC		OPTIONAL
	PARAMTYPE 2	VAR fontId				AS NUMERIC		OPTIONAL
	PARAMTYPE 3	VAR fillId				AS NUMERIC		OPTIONAL
	PARAMTYPE 4	VAR borderId			AS NUMERIC		OPTIONAL
	PARAMTYPE 5	VAR xfId				AS NUMERIC		OPTIONAL
	PARAMTYPE 6	VAR aValores			AS ARRAY		OPTIONAL
	PARAMTYPE 7	VAR aOutrosAtributos	AS ARRAY		OPTIONAL DEFAULT {}
	
	If nStyle+1>::nQtdStyle
		UserException("YExcel - Estilo informado("+cValToChar(nStyle)+") não definido. Utilize o indice informado pelo metodo AddStyles")
	Endif

	cXPath	:= cLocal+"/xmlns:xf["+cValToChar(nStyle+1)+"]"

	If ValType(fontId)=="N" .AND. (fontId+1)>Val(::oStyle:XPathGetAtt("/xmlns:styleSheet/xmlns:fonts","count"))
		UserException("YExcel - Fonte informada("+cValToChar(fontId)+") não definido. Utilize o indice informado pelo metodo :AddFont()")
	ElseIf ValType(fillId)=="N" .AND. (fillId+1)>Val(::oStyle:XPathGetAtt("/xmlns:styleSheet/xmlns:fills","count"))
		UserException("YExcel - Cor Preenchimento informado("+cValToChar(fillId)+") não definido. Utilize o indice informado pelo metodo :CorPreenc()")
	ElseIf ValType(borderId)=="N" .AND. (borderId+1)>Val(::oStyle:XPathGetAtt("/xmlns:styleSheet/xmlns:borders","count"))
		UserException("YExcel - Borda informada("+cValToChar(borderId)+") não definido. Utilize o indice informado pelo metodo :Borda()")
	Endif
	
	If ValType(numFmtId)=="U" .AND. ::oStyle:XPathGetAtt(cXPath,"applyNumberFormat")=="1"
		numFmtId	:= Val(::oStyle:XPathGetAtt(cXPath,"numFmtId"))
	Endif

	If ValType(fontId)=="U" .AND. ::oStyle:XPathGetAtt(cXPath,"applyFont")=="1"
		fontId	:= Val(::oStyle:XPathGetAtt(cXPath,"fontId"))
	Endif

	If ValType(fillId)=="U" .AND. ::oStyle:XPathGetAtt(cXPath,"applyFill")=="1"
		fillId	:= Val(::oStyle:XPathGetAtt(cXPath,"fillId"))
	Endif
	If ValType(borderId)=="U" .AND. ::oStyle:XPathGetAtt(cXPath,"applyBorder")=="1"
		borderId	:= Val(::oStyle:XPathGetAtt(cXPath,"borderId"))
	Endif

	If ValType(xfId)=="U"
		xfId	:= Val(::oStyle:XPathGetAtt(cXPath,"xfId"))
	Endif

	If ValType(aValores)=="U"
		aValores	:= {}
		aChildren	:= ::oStyle:XPathGetChildArray(cXPath)
		For nCont:=1 to Len(aChildren)
			AADD(aValores,yExcelTag():New(;
							aChildren[nCont][1];
							,aChildren[nCont][3];
							,::oStyle:XPathGetAttArray(aChildren[nCont][2]);
							,self);
				)
		Next
	Endif

	aAtributos	:= ::oStyle:XPathGetAttArray(cXPath)

	For nCont:=1 to Len(aAtributos)
		If !("|"+aAtributos[nCont][1]+"|" $ "|numFmtId|fontId|fillId|borderId|xfId|applyFont|applyFill|applyBorder|applyAlignment|applyNumberFormat|") .AND. aScan(aOutrosAtributos,{|x| x[1]==aAtributos[nCont][1] })==0
			AADD(aOutrosAtributos,aClone(aAtributos[nCont]))
		Endif
	Next

Return ::AddStyles(numFmtId,fontId,fillId,borderId,xfId,aValores,aOutrosAtributos)

/*/{Protheus.doc} YExcel::NewStyRules
Cria regras de estilo e formatação
@type method
@version 1.0
@author Saulo Gomes Martins
@since 04/07/2020
@return object, Objeto para criar regras de formatação
/*/
Method NewStyRules() Class YExcel
	Local oStyRules	:= YExcel_StyleRules():New(self)
	AADD(::aCleanObj,oStyRules)
Return oStyRules

/*/{Protheus.doc} YExcel_StyleRules
Regras de estilo e formatação
@type class
@version 1.0
@author Saulo Gomes Martins
@since 04/07/2020
/*/
Class YExcel_StyleRules
	Data cClassName
	Data oExcel
	Data aStyles
	Data aFmtNum
	Data afont
	Data afill
	Data aborder
	Data aRValores
	Method ClassName()
	Method New()
	Method GetStyle()
	Method GetId()
	Method AddStyle()		//Regra de estilo
	Method AddnumFmt()		//Regra de formato numero
	Method AddFont()		//Regra de fonte
	Method Addfill()		//Regra de preenchimento de fundo
	Method Addborder()		//Regra de borda
	Method AddValores()		//Regra de Valores
EndClass

Method ClassName() Class YExcel_StyleRules
Return "YEXCEL_STYLERULES"

Method New(oExcel) Class YExcel_StyleRules
	::cClassName	:= "YEXCEL_STYLERULES"
	::oExcel		:= oExcel
	::aStyles		:= {}
	::aFmtNum		:= {}
	::afont			:= {}
	::afill			:= {}
	::aborder		:= {}
	::aRValores		:= {}
	AADD(::oExcel:aCleanObj,self)
Return self

Method GetStyle(nLinha,nColuna) Class YExcel_StyleRules
	Local nCont
	Local oStyle
	For nCont:=1 to Len(::aStyles)
		If Eval(::aStyles[nCont][1],nLinha,nColuna,::oExcel)
			oStyle	:= ::aStyles[nCont][2]
			Exit
		Endif
	Next
	oStyle	:= YExcel_Style():New(oStyle,::oExcel)	//::oExcel:NewStyle(oStyle)	//Cria um estilo novo com herança para evitar modificação do principal
	For nCont:=1 to Len(::aFmtNum)
		If Eval(::aFmtNum[nCont][1],nLinha,nColuna,::oExcel)
			oStyle:SetnumFmt(::aFmtNum[nCont][2])
			Exit
		Endif
	Next
	For nCont:=1 to Len(::afont)
		If Eval(::afont[nCont][1],nLinha,nColuna,::oExcel)
			oStyle:Setfont(::afont[nCont][2])
			Exit
		Endif
	Next
	For nCont:=1 to Len(::afill)
		If Eval(::afill[nCont][1],nLinha,nColuna,::oExcel)
			oStyle:Setfill(::afill[nCont][2])
			Exit
		Endif
	Next
	For nCont:=1 to Len(::aborder)
		If Eval(::aborder[nCont][1],nLinha,nColuna,::oExcel)
			oStyle:Setborder(::aborder[nCont][2])
			Exit
		Endif
	Next
	For nCont:=1 to Len(::aRValores)
		If Eval(::aRValores[nCont][1],nLinha,nColuna,::oExcel)
			oStyle:Setborder(::aRValores[nCont][2])
			Exit
		Endif
	Next
Return oStyle

Method  GetId(nLinha,nColuna) Class YExcel_StyleRules
	Local oStyTmp	:= ::GetStyle(nLinha,nColuna)	//Classe YExcel_Style
	Local cId		:= oStyTmp:GetId()
	//Após pegar o ID libera objeto criado da memoria
	FreeObj(oStyTmp)	//Limpa obj da memoria
Return cId

Method AddStyle(bRule,xStyle) Class YExcel_StyleRules
	PARAMTYPE 0	VAR bRule			AS BLOCK
	PARAMTYPE 1	VAR xStyle			AS NUMERIC,OBJECT
	If ValType(xStyle)=="N" .and. xStyle>=0 .AND. xStyle+1>::oExcel:nQtdStyle
		UserException("YExcel - Estilo informado("+cValToChar(xStyle)+") não definido. Utilize o indice informado pelo metodo AddStyles")
	ElseIf ValType(xStyle)=="O" .AND. xStyle:ClassName()!="YEXCEL_STYLE"
		UserException("YExcel - Objeto de estilo deve ser inicializado pelo metodo NewStyle")
	Endif
	AADD(::aStyles,{bRule,xStyle})
Return

Method AddnumFmt(bRule,nFmtNum) Class YExcel_StyleRules
	PARAMTYPE 0	VAR bRule			AS BLOCK
	PARAMTYPE 1	VAR nFmtNum			AS NUMERIC
	AADD(::aFmtNum,{bRule,nFmtNum})
Return

Method AddFont(bRule,fontId) Class YExcel_StyleRules
	PARAMTYPE 0	VAR bRule			AS BLOCK
	PARAMTYPE 1	VAR fontId			AS NUMERIC
	If (fontId+1)>Val(::oExcel:oStyle:XPathGetAtt("/xmlns:styleSheet/xmlns:fonts","count"))
		UserException("YExcel - Fonte informada("+cValToChar(fontId)+") não definido. Utilize o indice informado pelo metodo :AddFont()")
	Endif
	AADD(::afont,{bRule,fontId})
Return

Method Addfill(bRule,fillId) Class YExcel_StyleRules
	PARAMTYPE 0	VAR bRule			AS BLOCK
	PARAMTYPE 1	VAR fillId			AS NUMERIC
	If (fillId+1)>Val(::oExcel:oStyle:XPathGetAtt("/xmlns:styleSheet/xmlns:fills","count"))
		UserException("YExcel - Cor Preenchimento informado("+cValToChar(fillId)+") não definido. Utilize o indice informado pelo metodo :CorPreenc()")
	Endif
	AADD(::afill,{bRule,fillId})
Return

Method Addborder(bRule,borderId) Class YExcel_StyleRules
	PARAMTYPE 0	VAR bRule			AS BLOCK
	PARAMTYPE 1	VAR borderId		AS NUMERIC
	If ValType(borderId)=="N" .AND. (borderId+1)>Val(::oExcel:oStyle:XPathGetAtt("/xmlns:styleSheet/xmlns:borders","count"))
		UserException("YExcel - Borda informada("+cValToChar(borderId)+") não definido. Utilize o indice informado pelo metodo :Borda()")
	Endif
	AADD(::aborder,{bRule,borderId})
Return

Method AddValores(bRule,aValores) Class YExcel_StyleRules
	PARAMTYPE 0	VAR bRule			AS BLOCK
	PARAMTYPE 1	VAR aValores		AS ARRAY
	AADD(::aRValores,{bRule,aClone(aValores)})
Return

//Se estilo é data
//pag
Method StyleType(nStyle) Class YExcel
	Local cLocal	:= "/xmlns:styleSheet/xmlns:cellXfs"
	Local cXPath	:= cLocal+"/xmlns:xf["+cValToChar(nStyle+1)+"]"
	Local cFmtId	:= ::oStyle:XPathGetAtt(cXPath,"numFmtId")
	Local cFmtData	:= "|14|15|16|17|"
	Local cFmtHora	:= "|18|19|20|21|45|46|47|"
	Local cFmtDtHr	:= "|22|"
	Local cFmtNum	:= "|1|2|3|4|9|10|11|12|13|37|38|39|40|48|"
	Local cFmtTxt	:= "|49|"
	Local cformatCode:= ""
	Local nPosAsp,nPos
	// Local aFormaCode
	If "|"+cFmtId+"|" $ cFmtData
		Return "D"
	ElseIf "|"+cFmtId+"|" $ cFmtNum	//Numeros
		Return "N"
	ElseIf "|"+cFmtId+"|" $ cFmtHora	//Numeros
		Return "H"
	ElseIf "|"+cFmtId+"|" $ cFmtDtHr
		Return "DT"
	ElseIf "|"+cFmtId+"|" $ cFmtTxt
		Return "C"
	Else
		cformatCode	:= ::oStyle:XPathGetAtt("/xmlns:styleSheet/xmlns:numFmts/xmlns:numFmt[@numFmtId='"+cFmtId+"']","formatCode")
		nPos	:= At(';',cformatCode)
		nPosAsp	:= At('"',cformatCode)
		If nPosAsp==0
			nPosAsp	:= At("'",cformatCode)
		Endif
		If nPosAsp==0
		Endif
		If nPos>0 .and. nPosAsp==0
			// aFormaCode	:= Separa(cformatCode,";")
			If At("mm",cformatCode)>0 .OR. At("dd",cformatCode)>0 .OR. At("d/",cformatCode)>0 .OR. At("m/",cformatCode)>0 .OR. At("/a",cformatCode)>0 .OR. At("/y",cformatCode)>0 .OR. At(":m",cformatCode)>0 .OR. At("h:",cformatCode)>0 .OR. At(":s",cformatCode)>0
				Return "DT"
			Endif
		Endif
		//TODO como identificar o tipo pela mascara
	Endif
Return "X"

/*/{Protheus.doc} SetStyle
Altera o estilo de uma ou várias células
@author Saulo Gomes Martins
@since 18/06/2020
@version 2.0
@param xStyle, variadic, (numeric/Array) posição do estilo criado pelo metodo :AddStyles()
@param nLinha, numeric, Linha inicial a ser alterada
@param nColuna, numeric, Coluna inicial a ser a ser alterada
@param nLinha2, numeric, (Opcional) Linha final a ser alterada
@param nColuna2, numeric, (Opcional) Coluna final a ser a ser alterada
@type method
/*/
METHOD SetStyle(xStyle,nLinha,nColuna,nLinha2,nColuna2) Class YExcel
	Local nLin,nCol
	Local nStyle
	Local cTpAlte
	Local lNumFmtId	//Estilo enviado no parametro tem FmtId
	Local nNumFmtId
	Local nStyletmp
	Local cAliasQry
	PARAMTYPE 0	VAR xStyle			AS NUMERIC,ARRAY,OBJECT		OPTIONAL DEFAULT -1
	PARAMTYPE 1	VAR nLinha			AS NUMERIC					OPTIONAL DEFAULT ::nLinha
	PARAMTYPE 2	VAR nColuna			AS NUMERIC					OPTIONAL DEFAULT ::nColuna
	PARAMTYPE 3	VAR nLinha2			AS NUMERIC					OPTIONAL DEFAULT nLinha
	PARAMTYPE 4	VAR nColuna2		AS NUMERIC					OPTIONAL DEFAULT nColuna
	cTpAlte		:= ValType(xStyle)

	If ValType(xStyle)=="O" .AND. xStyle:ClassName()=="YEXCEL_STYLE"
		xStyle	:= xStyle:GetId()
		cTpAlte	:= "N"	
	Endif
	
	If cTpAlte=="N" .AND. xStyle>=0 .AND. xStyle+1>::nQtdStyle
		UserException("YExcel - Estilo informado("+cValToChar(xStyle)+") não definido. Utilize o indice informado pelo metodo AddStyles")
	Endif
	
	::InsertCellEmpty(nLinha,nColuna,nLinha2,nColuna2)	//Inserir as celulas vazias que não tem dados para preencher o estilo
	If cTpAlte=="N"	//Alteração direta de estilo
		lNumFmtId := ::oStyle:XPathGetAtt("/xmlns:styleSheet/xmlns:cellXfs/xmlns:xf["+cValToChar(xStyle+1)+"]","applyNumberFormat")=="1"	//Tem formatação definida
		If lNumFmtId
			//Altera todas a celulas com o mesmo estilo
			If !DBSqlExec(::cAliasCol, "UPDATE "+::cAliasCol+" SET STY="+cValToChar(xStyle)+" WHERE PLA="+cValToChar(::nPlanilhaAt)+" AND LIN>="+cValToChar(nLinha)+" AND LIN<="+cValToChar(nLinha2)+" AND COL>="+cValToChar(nColuna)+" AND COL<="+cValToChar(nColuna2)+" AND D_E_L_E_T_=' '", ::cDriver)
				UserException("YExcel - Erro ao atualiza estilo ("+cValToChar(xStyle)+"). "+TCSqlError())
			Endif
		Else
			cAliasQry := GetNextAlias()
			//Verifica se tem celula com tipo datatime definido
			cQuery	:= "SELECT DISTINCT STY FROM "+::cAliasCol+" WHERE PLA="+cValToChar(::nPlanilhaAt)+" AND LIN>="+cValToChar(nLinha)+" AND LIN<="+cValToChar(nLinha2)+" AND COL>="+cValToChar(nColuna)+" AND COL<="+cValToChar(nColuna2)+" "	
			cQuery	+= " AND D_E_L_E_T_=' '"
			If !DbSqlExec(cAliasQry,cQuery,::cDriver)
				UserException("YExcel - Erro ao atualiza estilo. "+TCSqlError())
			Endif
			While (cAliasQry)->(!EOF())
				lnumFmtId	:= ::oStyle:XPathGetAtt("/xmlns:styleSheet/xmlns:cellXfs/xmlns:xf["+cValToChar((cAliasQry)->STY+1)+"]","applyNumberFormat")=="1"
				If lnumFmtId	//Tem fmtid
					nNumFmtId	:= Val(::oStyle:XPathGetAtt("/xmlns:styleSheet/xmlns:cellXfs/xmlns:xf["+cValToChar((cAliasQry)->STY+1)+"]","numFmtId"))
					nStyletmp	:= ::CreateStyle(xStyle,nNumFmtId)	//Cria outro com base no atual com mesmo fmtid
				Else	//Para os que não tem fmtid definido, aplica o estilo enviado
					nStyletmp	:= xStyle
				EndIf
				If !DBSqlExec(::cAliasCol, "UPDATE "+::cAliasCol+" SET STY="+cValToChar(nStyletmp)+" WHERE PLA="+cValToChar(::nPlanilhaAt)+" AND LIN>="+cValToChar(nLinha)+" AND LIN<="+cValToChar(nLinha2)+" AND COL>="+cValToChar(nColuna)+" AND COL<="+cValToChar(nColuna2)+" AND STY='"+cValToChar((cAliasQry)->STY)+"' AND D_E_L_E_T_=' '", ::cDriver)
					UserException("YExcel - Erro ao atualiza estilo ("+cValToChar(nStyletmp)+"). "+TCSqlError())
				Endif
				(cAliasQry)->(DbSkip())
			EndDo
			(cAliasQry)->(DbCloseArea())
		EndIf
	ElseIf cTpAlte=="O"	//Se enviado objeto vai avaliar estilo a ser usado
		For nLin:=nLinha to nLinha2				//Ler as Linhas
			For nCol:=nColuna to nColuna2		//Ler as colunas
				If cTpAlte=="N"
					nStyle		:= xStyle
				Else				//Rules
					nStyle		:= xStyle:GetId(nLin,nCol)
				Endif
				If (::cAliasCol)->(DbSeek(Str(::nPlanilhaAt,10)+Str(nLin,10)+Str(nCol,10)))
					//Se não tem fmrid no estilo da regra e tem no estilo atual 
					If ::oStyle:XPathGetAtt("/xmlns:styleSheet/xmlns:cellXfs/xmlns:xf["+cValToChar(nStyle+1)+"]","applyNumberFormat")!="1" .AND. ::oStyle:XPathGetAtt("/xmlns:styleSheet/xmlns:cellXfs/xmlns:xf["+cValToChar((::cAliasCol)->STY+1)+"]","applyNumberFormat")=="1" 
						nNumFmtId	:= Val(::oStyle:XPathGetAtt("/xmlns:styleSheet/xmlns:cellXfs/xmlns:xf["+cValToChar((::cAliasCol)->STY+1)+"]","numFmtId"))
						nStyle		:= ::CreateStyle(nStyle,nNumFmtId)	//Cria outro com base no atual com formato data
					Endif
					(::cAliasCol)->(RecLock(::cAliasCol,.F.))
					(::cAliasCol)->STY	:= nStyle
					(::cAliasCol)->(MsUnLock())
				else
					::Cell(nLin,nCol,nil,,nStyle)
				Endif
			Next
		Next
	Endif
Return self

/*/{Protheus.doc} GetStyle
Retorna o estilo de uma células
@author Saulo Gomes Martins
@since 18/06/2020
@version 2.0
@param nLinha, numeric, Linha inicial a ser alterada
@param nColuna, numeric, Coluna inicial a ser a ser alterada
@type method
/*/
METHOD GetStyle(nLinha,nColuna) Class YExcel
	Local nStyle	:= -1
	Default nLinha	:= ::nLinha
	Default nColuna	:= ::nColuna
	If (::cAliasCol)->(DbSeek(Str(::nPlanilhaAt,10)+Str(nLinha,10)+Str(nColuna,10)))
		nStyle	:= (::cAliasCol)->STY
	Endif
Return nStyle

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
@type method
@obs	HORIZONTAL
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
METHOD Alinhamento(cHorizontal,cVertical,lReduzCaber,lQuebraTexto,ntextRotation) Class YExcel
	Local oAlinhamento	:= yExcelTag():New("alignment",,,self)
	Default cVertical	:= "general"
	Default cHorizontal	:= "bottom"
	Default lReduzCaber	:= .F.
	Default lQuebraTexto	:= .F.
	oAlinhamento:SetAtributo("horizontal",cHorizontal)
	oAlinhamento:SetAtributo("vertical",cVertical)
	If ValType(ntextRotation)=="N" .and. ntextRotation>0
		oAlinhamento:SetAtributo("textRotation",ntextRotation)
	Endif
	If lReduzCaber .and. !lQuebraTexto
		oAlinhamento:SetAtributo("shrinkToFit","1")	//Um valor booleano que indica se o texto exibido na célula deve ser encolhido para se ajustar à célula
	Endif
	If lQuebraTexto
		oAlinhamento:SetAtributo("wrapText","1")	//Um valor booleano indicando se o texto em uma célula deve ser envolvido na linha dentro da célula.
	Endif
Return oAlinhamento

/*/{Protheus.doc} AddPane
Congelar Painéis
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param nySplit, numeric, Quantidade de linhas congeladas
@param nxSplit, numeric, Quantidade de colunas congeladas
@type method
@obs pag 1712
/*/
METHOD AddPane(nySplit,nxSplit) Class YExcel
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
METHOD Pane(cActivePane,cState,cRef,nySplit,nxSplit) Class YExcel
	Local nPos
	Default cActivePane	:= "bottomLeft"
	
	//::asheet[::nPlanilhaAt][1]:XPathDelNode( "/xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView" )
	//::asheet[::nPlanilhaAt][1]:XPathAddNode( "/xmlns:worksheet/xmlns:sheetViews", "sheetView", "" )
	// ::asheet[::nPlanilhaAt][1]:XPathSetAtt("/xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView", "workbookViewId"	, "0" )

	If !::asheet[::nPlanilhaAt][1]:XPathHasNode("/xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView/xmlns:pane")
		::asheet[::nPlanilhaAt][1]:XPathAddNode( "/xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView", "pane", "" )
	EndIf
	/*
	bottomLeft	- Painel inferior esquerdo, quando ambos verticais e horizontais são aplicadas. Esse valor também é usado quando apenas uma divisão horizontal foi aplicada, dividindo o painel em superior e inferior. Nesse caso, esse valor especifica painel inferior
	bottomRight - Painel inferior direito, quando as divisões verticais e horizontais são aplicadas.
	topLeft		- Painel superior esquerdo, quando as divisões verticais e horizontais são aplicadas.
	topRight	- Painel superior direito, quando as divisões verticais e horizontais são aplicadas
	*/
	SetAtrr(::asheet[::nPlanilhaAt][1],"/xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView/xmlns:pane[last()]", "activePane"	, cActivePane )
	/*
	frozen		- Panes são congelados, mas não foram divididos sendo congelados. Nesse estado, quando os painéis são desbloqueados novamente, um único painel resulta, sem divisão. Nesse estado, as barras de divisão não são ajustáveis.
	frozenSplit	- Os painéis são congelados e foram divididos antes de serem congelados. Neste estado, quando os painéis são desbloqueados novamente, a divisão permanece, mas é ajustável.
	split		- Os painéis são divididos, mas não congelados. Nesse estado, as barras de divisão são ajustáveis pelo usuário.
	*/
	SetAtrr(::asheet[::nPlanilhaAt][1],"/xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView/xmlns:pane[last()]", "state"	, cState )
	//Localização da célula visível superior esquerda no painel inferior direito (quando no modo Esquerdo para Direito).
	SetAtrr(::asheet[::nPlanilhaAt][1],"/xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView/xmlns:pane[last()]", "topLeftCell"	, cRef )
	//Posição horizontal da divisão, em 1/20º de um ponto; 0 (zero) se nenhum. Se o painel estiver congelado, este valor indica o número de colunas visíveis no painel superior
	SetAtrr(::asheet[::nPlanilhaAt][1],"/xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView/xmlns:pane[last()]", "xSplit"	, cValToChar(nxSplit) )
	//Posição vertical da divisão, em 1/20º de um ponto; 0 (zero) se nenhum. Se o painel estiver congelado, este valor indica o número de linhas visíveis no painel esquerdo.
	SetAtrr(::asheet[::nPlanilhaAt][1],"/xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView/xmlns:pane[last()]", "ySplit"	, cValToChar(nySplit) )
Return nPos

/*/{Protheus.doc} YExcel::Addhyperlink
Cria um hyperlink para uma referência da planilha
@type method
@version 1.0
@author Saulo Gomes Martins
@since 18/03/2021
@param cLocation, character, Referência, pode ser simple (A1) ou intervalo (A1:C3)
@param ctooltip, character, Texto de dica ao passar mouse por cima
@param cDisplay, character, Texto de exibição na celula
@return object, self
/*/
METHOD Addhyperlink(cLocation,ctooltip,cDisplay) Class YExcel
	Local cRef	:= ::Ref(::nLinha,::nColuna)
	PARAMTYPE 0	VAR cLocation		AS CHARACTER		OPTIONAL
	PARAMTYPE 1	VAR ctooltip		AS CHARACTER		OPTIONAL
	PARAMTYPE 2	VAR cDisplay		AS CHARACTER		OPTIONAL
	If !Empty(cLocation)
		If !::asheet[::nPlanilhaAt][1]:XPathHasNode("/xmlns:worksheet/xmlns:hyperlinks")
			::asheet[::nPlanilhaAt][1]:XPathAddNode( "/xmlns:worksheet", "hyperlinks", "" )
		EndIf
		If !::asheet[::nPlanilhaAt][1]:XPathHasNode("/xmlns:worksheet/xmlns:hyperlinks/xmlns:hyperlink[@ref='"+cRef+"']")
			::asheet[::nPlanilhaAt][1]:XPathAddNode( "/xmlns:worksheet/xmlns:hyperlinks", "hyperlink", "" )
			::asheet[::nPlanilhaAt][1]:XPathAddAtt( "/xmlns:worksheet/xmlns:hyperlinks/xmlns:hyperlink[last()]", "ref", cRef )
		EndIf
		SetAtrr(::asheet[::nPlanilhaAt][1],"/xmlns:worksheet/xmlns:hyperlinks/xmlns:hyperlink[@ref='"+cRef+"']", "location", cLocation)
		SetAtrr(::asheet[::nPlanilhaAt][1],"/xmlns:worksheet/xmlns:hyperlinks/xmlns:hyperlink[@ref='"+cRef+"']", "tooltip", ctooltip)
		SetAtrr(::asheet[::nPlanilhaAt][1],"/xmlns:worksheet/xmlns:hyperlinks/xmlns:hyperlink[@ref='"+cRef+"']", "display", cDisplay)
	Else
		::asheet[::nPlanilhaAt][1]:XPathDelNode("/xmlns:worksheet/xmlns:hyperlinks/xmlns:hyperlink[@ref='"+cRef+"']")
		If ::asheet[::nPlanilhaAt][1]:XPathChildCount("/xmlns:worksheet/xmlns:hyperlinks")==0
			::asheet[::nPlanilhaAt][1]:XPathDelNode("/xmlns:worksheet/xmlns:hyperlinks")
		Endif
	EndIf
Return self
/*/{Protheus.doc} YExcel::AddComment
Adicionar comentário
@type method
@version 1.0
@author Saulo Gomes Martins
@since 18/03/2021
@param cText, character, Texto do comentário
@param cAutor, character, Autor do comentário
@return object, self
@obs pag 4682
/*/
METHOD AddComment(cText,cAutor) Class YExcel
	Local aChildren
	Local nPos
	Local cPos
	Local cAliasQry
	Local cQuery
	Local cValor
	Local ntop
	Local nleft
	Local nCont
	Local nQtdLinhas
	Local cAuthorId
	Local cRef	:= ::Ref(::nLinha,::nColuna)
	PARAMTYPE 0	VAR cText		AS CHARACTER		OPTIONAL Default ""
	PARAMTYPE 1	VAR cAutor		AS CHARACTER		OPTIONAL Default ""
	
	If Empty(::asheet[::nPlanilhaAt][3])
		::asheet[::nPlanilhaAt][3]	:= ::new_comment()
		::add_rels("\xl\worksheets\_rels\sheet"+cValToChar(::nPlanilhaAt)+".xml.rels","http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments","../comments"+cValToChar(::nPlanilhaAt)+".xml")
		::asheet[::nPlanilhaAt][4]	:= "\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\comments"+cValToChar(::nPlanilhaAt)+".xml"
		::ocontent_types:XPathAddNode( "/xmlns:Types", "Override", "" )
		::ocontent_types:XPathAddAtt( "/xmlns:Types/xmlns:Override[last()]", "PartName"	, "/xl/comments"+cValToChar(::nPlanilhaAt)+".xml" )
		::ocontent_types:XPathAddAtt( "/xmlns:Types/xmlns:Override[last()]", "ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml" )

		cId	:= ::add_rels("\xl\worksheets\_rels\sheet"+cValToChar(::nPlanilhaAt)+".xml.rels","http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing","../drawings/vmlDrawing"+cValToChar(::nPlanilhaAt)+".vml")
		::asheet[::nPlanilhaAt][1]:XPathAddNode("/xmlns:worksheet","legacyDrawing","")
		::asheet[::nPlanilhaAt][1]:XPathAddAtt("/xmlns:worksheet/xmlns:legacyDrawing","r:id",cId)
		If !::ocontent_types:XPathHasNode("/xmlns:Types/xmlns:Default[@Extension='vml']")
			::ocontent_types:XPathAddNode( "/xmlns:Types", "Default", "" )
			::ocontent_types:XPathAddAtt( "/xmlns:Types/xmlns:Default[last()]", "Extension"	, "vml" )
			::ocontent_types:XPathAddAtt( "/xmlns:Types/xmlns:Default[last()]", "ContentType", "application/vnd.openxmlformats-officedocument.vmlDrawing" )
		EndIf
		
		::asheet[::nPlanilhaAt][5]	:= ::new_vmlDrawing()
		::asheet[::nPlanilhaAt][6]	:= "\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\drawings\vmlDrawing"+cValToChar(::nPlanilhaAt)+".vml"

	EndIf
	//Docs vmlDrawing
	//ECMA-376 Part 4 pag 597
	If Empty(cText)
		//Deleta forma
		If ::asheet[::nPlanilhaAt][5]:XPathHasNode('/xml/v:shape[x:ClientData/x:Row="'+cValToChar(::nLinha-1)+'" and x:ClientData/x:Column="'+cValToChar(::nColuna-1)+'"]')
			::asheet[::nPlanilhaAt][5]:XPathDelNode('/xml/v:shape[x:ClientData/x:Row="'+cValToChar(::nLinha-1)+'" and x:ClientData/x:Column="'+cValToChar(::nColuna-1)+'"]')
		Endif
		//Deleta comentário
		If ::asheet[::nPlanilhaAt][3]:XPathHasNode('/xmlns:comments/xmlns:commentList/xmlns:comment[@ref="'+cRef+'"]')
			cAuthorId	:= ::asheet[::nPlanilhaAt][3]:XPathGetAtt('/xmlns:comments/xmlns:commentList/xmlns:comment[@ref="'+cRef+'"]',"authorId")
			::asheet[::nPlanilhaAt][3]:XPathDelNode('/xmlns:comments/xmlns:commentList/xmlns:comment[@ref="'+cRef+'"]')
			If !Empty(cAuthorId) .AND. !::asheet[::nPlanilhaAt][3]:XPathHasNode('/xmlns:comments/xmlns:commentList/xmlns:comment[@authorId="'+cAuthorId+'"]')
				::asheet[::nPlanilhaAt][3]:XPathDelNode('/xmlns:comments/xmlns:authors/xmlns:author['+cValToChar(Val(cAuthorId)+1)+']')
			EndIf
		Endif
	Else
		//Cria a forma do comentário se não existe
		If !::asheet[::nPlanilhaAt][5]:XPathHasNode('/xml/o:shapelayout')
			::asheet[::nPlanilhaAt][5]:XPathAddNode('/xml',"o:shapelayout","")
			::asheet[::nPlanilhaAt][5]:XPathAddAtt('/xml/*[last()]',"v:ext","edit")
			::asheet[::nPlanilhaAt][5]:XPathAddNode('/xml/*[last()]',"o:idmap","")
			::asheet[::nPlanilhaAt][5]:XPathAddAtt('/xml/*[last()]/*[last()]',"v:ext","edit")
			::asheet[::nPlanilhaAt][5]:XPathAddAtt('/xml/*[last()]/*[last()]',"data","1")
		Endif
		If !::asheet[::nPlanilhaAt][5]:XPathHasNode('/xml/v:shapetype[@id="_x0000_t202"]')	//Tipo de forma
			::asheet[::nPlanilhaAt][5]:XPathAddNode('/xml',"v:shapetype","")
			::asheet[::nPlanilhaAt][5]:XPathAddAtt('/xml/*[last()]',"id","_x0000_t202")
			::asheet[::nPlanilhaAt][5]:XPathAddAtt('/xml/*[last()]',"coordsize","21600,21600")
			::asheet[::nPlanilhaAt][5]:XPathAddAtt('/xml/*[last()]',"o:spt","202")
			::asheet[::nPlanilhaAt][5]:XPathAddAtt('/xml/*[last()]',"path","m,l,21600r21600,l21600,xe")
			::asheet[::nPlanilhaAt][5]:XPathAddNode('/xml/*[last()]',"v:stroke","")
			::asheet[::nPlanilhaAt][5]:XPathAddAtt('/xml/*[last()]/*[last()]',"joinstyle","miter")
			::asheet[::nPlanilhaAt][5]:XPathAddNode('/xml/*[last()]',"v:path","")
			::asheet[::nPlanilhaAt][5]:XPathAddAtt('/xml/*[last()]/*[last()]',"gradientshapeok","t")
			::asheet[::nPlanilhaAt][5]:XPathAddAtt('/xml/*[last()]/*[last()]',"o:connecttype","rect")
		Endif

		If !::asheet[::nPlanilhaAt][5]:XPathHasNode('/xml/v:shape[x:ClientData/x:Row="'+cValToChar(::nLinha-1)+'" and x:ClientData/x:Column="'+cValToChar(::nColuna-1)+'"]')
			//posição da forma de acordo com posição da célula
			nRowHeight	:= Val(::asheet[::nPlanilhaAt][1]:XPathGetAtt("/xmlns:worksheet/xmlns:sheetFormatPr","defaultRowHeight"))
			cQuery	:= "SELECT COUNT(*) QTD,SUM(1.000005*CASE WHEN CHEIGHT='1' THEN HT ELSE "+cValToChar(nRowHeight)+" END) VALOR FROM "+::cAliasLin+" WHERE"
			cQuery	+= " LIN<"+cValToChar(::nLinha)+" AND"
			cQuery	+= " PLA="+cValToChar(::nPlanilhaAt)+" AND D_E_L_E_T_=' '"
			cAliasQry := GetNextAlias()
			If !DbSqlExec(cAliasQry,cQuery,::cDriver)
				UserException("YExcel - Erro ao tamanho das linhas. "+TCSqlError())
			Endif
			ntop	:= 0
			If (cAliasQry)->(!EOF())
				nQtdLinhas	:= (cAliasQry)->QTD		//Quantidade de linhas que tem conteudo
				ntop		:= (cAliasQry)->VALOR
				If (::nLinha-1)>nQtdLinhas
					//Adiciona as linhas que não tem conteudo o tamanho padrão
					ntop	+= (::nLinha-1-nQtdLinhas)*1.000005*nRowHeight
				EndIf
				ntop		:= Round(ntop,0)
			Endif
			(cAliasQry)->(DbCloseArea())
			nleft		:= 0
			For nCont:=1 to ::nColuna
				cValor	:= ::asheet[::nPlanilhaAt][1]:XPathGetAtt( '/xmlns:worksheet/xmlns:cols/xmlns:col['+cValToChar(nCont)+'>=@min and '+cValToChar(nCont)+'<=@max and @customWidth="1"]',"width")
				If Empty(cValor)
					cValor	:= "9.28"
				Endif
				cValor	:= Val(cValor)
				nleft	+= cValor*5.25
			Next
			nleft	:= Round(nleft,0)

			::asheet[::nPlanilhaAt][5]:XPathAddNode('/xml',"v:shape","")
			::asheet[::nPlanilhaAt][5]:XPathAddAtt('/xml/*[last()]',"id","_x0000_r"+cValToChar(::nLinha-1)+"c"+cValToChar(::nColuna-1))
			::asheet[::nPlanilhaAt][5]:XPathAddAtt('/xml/*[last()]',"type","#_x0000_t202")
			::asheet[::nPlanilhaAt][5]:XPathAddAtt('/xml/*[last()]',"style","position:absolute;margin-left:"+cValToChar(nleft)+"pt;margin-top:"+cValToChar(ntop)+"pt;width:96pt;height:64.5pt;z-index:1;visibility:hidden")
			::asheet[::nPlanilhaAt][5]:XPathAddAtt('/xml/*[last()]',"fillcolor","#ffffc0")
			::asheet[::nPlanilhaAt][5]:XPathAddAtt('/xml/*[last()]',"o:insetmode","auto")
			::asheet[::nPlanilhaAt][5]:XPathAddNode('/xml/*[last()]',"v:fill","")
			::asheet[::nPlanilhaAt][5]:XPathAddAtt('/xml/*[last()]/*[last()]',"color2","#ffffc0")
			::asheet[::nPlanilhaAt][5]:XPathAddNode('/xml/*[last()]',"v:shadow","")
			::asheet[::nPlanilhaAt][5]:XPathAddAtt('/xml/*[last()]/*[last()]',"on","t")
			::asheet[::nPlanilhaAt][5]:XPathAddAtt('/xml/*[last()]/*[last()]',"color","black")
			::asheet[::nPlanilhaAt][5]:XPathAddAtt('/xml/*[last()]/*[last()]',"obscured","t")
			::asheet[::nPlanilhaAt][5]:XPathAddNode('/xml/*[last()]',"v:path","")
			::asheet[::nPlanilhaAt][5]:XPathAddAtt('/xml/*[last()]/*[last()]',"o:connecttype","none")
			::asheet[::nPlanilhaAt][5]:XPathAddNode('/xml/*[last()]',"v:textbox","")
			::asheet[::nPlanilhaAt][5]:XPathAddAtt('/xml/*[last()]/*[last()]',"style","mso-direction-alt:auto")
			::asheet[::nPlanilhaAt][5]:XPathAddNode('/xml/*[last()]/*[last()]',"div","")
			::asheet[::nPlanilhaAt][5]:XPathAddAtt('/xml/*[last()]/*[last()]/div',"style","text-align:left")
			::asheet[::nPlanilhaAt][5]:XPathAddNode('/xml/*[last()]',"x:ClientData","")
			::asheet[::nPlanilhaAt][5]:XPathAddAtt('/xml/*[last()]/*[last()]',"ObjectType","Note")
			::asheet[::nPlanilhaAt][5]:XPathAddNode('/xml/*[last()]/*[last()]',"x:MoveWithCells","")
			::asheet[::nPlanilhaAt][5]:XPathAddNode('/xml/*[last()]/*[last()]',"x:SizeWithCells","")
			::asheet[::nPlanilhaAt][5]:XPathAddNode('/xml/*[last()]/*[last()]',"x:AutoFill","False")
			::asheet[::nPlanilhaAt][5]:XPathAddNode('/xml/*[last()]/*[last()]',"x:Row",cValToChar(::nLinha-1))
			::asheet[::nPlanilhaAt][5]:XPathAddNode('/xml/*[last()]/*[last()]',"x:Column",cValToChar(::nColuna-1))
			::asheet[::nPlanilhaAt][5]	:= AjustXML(::asheet[::nPlanilhaAt][5])
		EndIf

		If !::asheet[::nPlanilhaAt][3]:XPathHasNode("/xmlns:comments/xmlns:authors")
			::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments","authors","")
		EndIf
		If !::asheet[::nPlanilhaAt][3]:XPathHasNode("/xmlns:comments/xmlns:commentList")
			::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments","commentList","")
		EndIf
		aChildren	:= ::asheet[::nPlanilhaAt][3]:XPathGetChildArray("/xmlns:comments/xmlns:authors")
		nPos		:= aScan(aChildren,{|x| x[3]==cAutor })
		If nPos==0
			::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments/xmlns:authors","author",EncodeUTF8(cAutor))
			cPos	:= cValToChar(Len(aChildren))
		else
			cPos	:= cValToChar(nPos-1)
		EndIf
		
		::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments/xmlns:commentList","comment","")
		::asheet[::nPlanilhaAt][3]:XPathAddAtt("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]","ref",cRef)
		::asheet[::nPlanilhaAt][3]:XPathAddAtt("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]","authorId",cPos)
		::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]","text","")
		::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:text[last()]","r","")
		::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:text[last()]/xmlns:r[last()]","rPr","")
		::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:text[last()]/xmlns:r[last()]/xmlns:rPr","rFont","")
		::asheet[::nPlanilhaAt][3]:XPathAddAtt("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:text[last()]/xmlns:r[last()]/xmlns:rPr/xmlns:rFont","val","Calibri")
		::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:text[last()]/xmlns:r[last()]/xmlns:rPr","b","")
		::asheet[::nPlanilhaAt][3]:XPathAddAtt("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:text[last()]/xmlns:r[last()]/xmlns:rPr/xmlns:b","val","true")
		::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:text[last()]/xmlns:r[last()]/xmlns:rPr","strike","")
		::asheet[::nPlanilhaAt][3]:XPathAddAtt("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:text[last()]/xmlns:r[last()]/xmlns:rPr/xmlns:strike","val","false")
		::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:text[last()]/xmlns:r[last()]/xmlns:rPr","sz","")
		::asheet[::nPlanilhaAt][3]:XPathAddAtt("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:text[last()]/xmlns:r[last()]/xmlns:rPr/xmlns:sz","val","11")
		// ::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:text[last()]/xmlns:r[last()]/xmlns:rPr","color","")
		// ::asheet[::nPlanilhaAt][3]:XPathAddAtt("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:text[last()]/xmlns:r[last()]/xmlns:rPr/xmlns:color","indexed","81")
		// ::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:text[last()]/xmlns:r[last()]/xmlns:rPr","charset","")
		// ::asheet[::nPlanilhaAt][3]:XPathAddAtt("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:text[last()]/xmlns:r[last()]/xmlns:rPr/xmlns:charset","val","1")
		// ::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:text[last()]/xmlns:r[last()]/xmlns:rPr","scheme","")
		// ::asheet[::nPlanilhaAt][3]:XPathAddAtt("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:text[last()]/xmlns:r[last()]/xmlns:rPr/xmlns:scheme","val","minor")
		::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:text[last()]/xmlns:r[last()]","t",EncodeUTF8(cText))
		::asheet[::nPlanilhaAt][3]:XPathAddAtt("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:text[last()]/xmlns:r[last()]/xmlns:t","xml:space","preserve")
		// ::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:text[last()]","r","")
		// ::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:text[last()]/xmlns:r[last()]","rPr","")
		// ::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:text[last()]/xmlns:r[last()]/xmlns:rPr","sz","")
		// ::asheet[::nPlanilhaAt][3]:XPathAddAtt("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:text[last()]/xmlns:r[last()]/xmlns:rPr/xmlns:sz","val","8")
		// ::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:text[last()]/xmlns:r[last()]/xmlns:rPr","color","")
		// ::asheet[::nPlanilhaAt][3]:XPathAddAtt("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:text[last()]/xmlns:r[last()]/xmlns:rPr/xmlns:color","indexed","81")
		// ::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:text[last()]/xmlns:r[last()]/xmlns:rPr","rFont","")
		// ::asheet[::nPlanilhaAt][3]:XPathAddAtt("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:text[last()]/xmlns:r[last()]/xmlns:rPr/xmlns:rFont","val","Calibri")
		// ::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:text[last()]/xmlns:r[last()]/xmlns:rPr","charset","")
		// ::asheet[::nPlanilhaAt][3]:XPathAddAtt("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:text[last()]/xmlns:r[last()]/xmlns:rPr/xmlns:charset","val","1")
		// ::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:text[last()]/xmlns:r[last()]/xmlns:rPr","scheme","")
		// ::asheet[::nPlanilhaAt][3]:XPathAddAtt("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:text[last()]/xmlns:r[last()]/xmlns:rPr/xmlns:scheme","val","minor")
		// ::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:text[last()]/xmlns:r[last()]","t",EncodeUTF8(cText))
		// ::asheet[::nPlanilhaAt][3]:XPathAddAtt("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:text[last()]/xmlns:r[last()]/xmlns:t","xml:space","preserve")


		// ::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]","commentPr","")
		// ::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:commentPr","anchor","")
		// ::asheet[::nPlanilhaAt][3]:XPathAddAtt("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:commentPr/xmlns:anchor","sizeWithCells","true")
		// ::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:commentPr/xmlns:anchor","from","")
		// ::asheet[::nPlanilhaAt][3]::XPathAddNs(	"/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:commentPr/xmlns:anchor/xmlns:from", "xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" )
		// ::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:commentPr/xmlns:anchor/xmlns:from","col","0")
		// ::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:commentPr/xmlns:anchor/xmlns:from","colOff","0")
		// ::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:commentPr/xmlns:anchor/xmlns:from","row","0")
		// ::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:commentPr/xmlns:anchor/xmlns:from","rowOff","0")
		// ::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:commentPr/xmlns:anchor","to","")
		// ::asheet[::nPlanilhaAt][3]::XPathAddNs(	"/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:commentPr/xmlns:anchor/xmlns:to", "xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" )
		// ::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:commentPr/xmlns:anchor/xmlns:to","col","4")
		// ::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:commentPr/xmlns:anchor/xmlns:to","colOff","182880")
		// ::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:commentPr/xmlns:anchor/xmlns:to","row","10")
		// ::asheet[::nPlanilhaAt][3]:XPathAddNode("/xmlns:comments/xmlns:commentList/xmlns:comment[last()]/xmlns:commentPr/xmlns:anchor/xmlns:to","rowOff","0")
	EndIf
	

Return self


/*/{Protheus.doc} Ref
Retorna a referencia do excel de acordo com posição da linha e coluna em formato numerico
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param nLinha, numeric, Linha
@param nColuna, numeric, Coluna
@Return character, cRef Referencia da linha e coluna.
@type method
@obs
	oExcel:Ref(1,2)	//Retorno B1
	oExcel:Ref(3,3)	//Retorno C3
/*/
METHOD Ref(nLinha,nColuna,llinha,lColuna) Class YExcel
	Local cLinha	:= ""
	Local cColuna	:= ""
	Local cRet		:= ""
	Default nLinha	:= ::nLinha
	Default nColuna	:= ::nColuna
	Default llinha	:= .F.
	Default lColuna	:= .F.
	If llinha
		cLinha	:= "$"
	Endif
	If lColuna
		cColuna	:= "$"
	Endif
	If ValType(nColuna)!="U"
		cRet	+= cColuna+NumToString(nColuna)
	Endif
	If ValType(nLinha)!="U"
		cRet	+= cLinha+cValToChar(nLinha)
	Endif
Return cRet


/*/{Protheus.doc} LocRef
Retorna linha e coluna de acordo com informação da referencia
@author Saulo Gomes Martins
@since 03/03/2018
@version 1.0
@return array, aLinhaCol Array com duas dimenções 1=Linha|2=Coluna
@param cRef, characters, Refencia da celula (exemplo A1)
@type method

@example
LocRef("A1")	//Retorno {1,1}
LocRef("C5")	//Retorno {5,3}
/*/
METHOD LocRef(cRef) Class YExcel
	Local nCont
	Local nTam	:= Len(cRef)
	Local cColuna	:= ""
	Local cLinha	:= ""
	For nCont:=1 to nTam
		If IsAlpha(SubStr(cRef,nCont,1))
			cColuna	+= SubStr(cRef,nCont,1)
		ElseIf IsDigit(SubStr(cRef,nCont,1))
			cLinha	+= SubStr(cRef,nCont,1)
		Endif
	Next
Return {Val(cLinha),If(!Empty(cColuna),::StringToNum(cColuna),0)}


/*/{Protheus.doc} NumToString
Retorna a letra da coluna de acordo com a posição numerica
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param nNum, numeric, Numero da coluna
@type method
/*/
METHOD NumToString(nNum) Class YExcel
Return NumToString(nNum)

/*/{Protheus.doc} StringToNum
Retorna a posição da coluna de acordo com a letra da coluna
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param cString, characters, Letra da Coluna
@type method
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
@type method
/*/
METHOD AddTabela(cNome,nLinha,nColuna) Class YExcel
	Local nPos
	Local oTable
	Local cID
	PARAMTYPE 0	VAR cNome  AS CHARACTER			OPTIONAL DEFAULT lower(CriaTrab(,.F.))
	PARAMTYPE 1	VAR nLinha  AS NUMERIC			OPTIONAL DEFAULT ::adimension[2][1]
	PARAMTYPE 2	VAR nColuna  AS NUMERIC
	::nIdRelat++
	nPos	:= ::nIdRelat

	oTable	:= yExcel_Table():New(self,nLinha,nColuna,cNome) //yExcelTag():New("table",{},)
	oTable:nIdRelat	:= nPos
	oTable:SetAtributo("xmlns","http://schemas.openxmlformats.org/spreadsheetml/2006/main")
	oTable:SetAtributo("id",nPos)
	oTable:SetAtributo("name",cNome)
	oTable:SetAtributo("displayName",cNome)

	oTable:AddValor(yExcelTag():New("autoFilter",{},,self))

	oTable:oTableColumns	:= yExcelTag():New("tableColumns",{},{{"count",0}},self)	//Pag 1743
	oTable:AddValor(oTable:oTableColumns)

	oTable:otableStyleInfo	:= yExcelTag():New("tableStyleInfo",nil,,self)
	oTable:otableStyleInfo:SetAtributo("name","TableStyleMedium2")
	oTable:otableStyleInfo:SetAtributo("showFirstColumn",0)
	oTable:otableStyleInfo:SetAtributo("showLastColumn",0)
	oTable:otableStyleInfo:SetAtributo("showRowStripes",0)
	oTable:otableStyleInfo:SetAtributo("showColumnStripes",0)
	oTable:AddValor(oTable:otableStyleInfo)
	AADD(::aPlanilhas[::nPlanilhaAt][5],oTable)

	cID		:= ::add_rels("\xl\worksheets\_rels\sheet"+cValToChar(::nPlanilhaAt)+".xml.rels","http://schemas.openxmlformats.org/officeDocument/2006/relationships/table","../tables/table"+cValToChar(oTable:nIdRelat)+".xml")
	::aPlanilhas[::nPlanilhaAt][6]:AddValor(yExcelTag():New("tablePart",nil,{{"r:id",cID}},self))
	::aPlanilhas[::nPlanilhaAt][6]:SetAtributo("count",Len(::aPlanilhas[::nPlanilhaAt][5]))

	//Adiciona um nova Tabela
	::ocontent_types:XPathAddNode( "/xmlns:Types", "Override", "" )
	::ocontent_types:XPathAddAtt( "/xmlns:Types/xmlns:Override[last()]", "PartName"	, "/xl/tables/table"+cValToChar(oTable:nIdRelat)+".xml" )
	::ocontent_types:XPathAddAtt( "/xmlns:Types/xmlns:Override[last()]", "ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml" )
	AADD(::aCleanObj,oTable)
Return oTable

Method Gravar(cLocal,lAbrir,lDelSrv) Class YExcel
	Default lAbrir	:= .F.
	Default lDelSrv	:= .T.
	::lDelSrv	:= lDelSrv
	::Save(cLocal,lDelSrv)
	If lAbrir
		::OpenApp()
	Endif
	::Close(lDelSrv)
Return

Method OpenApp() Class YExcel
	If !Empty(::cArqGrv)
		ShellExecute("open",::cArqGrv,"",::cLocalFile+'\', 1 )
	Endif
Return

Method Close() Class YExcel
	If ::cDriver=="TMPDB"
		aEval(::aTmpDB, {|x| x:Delete(),FreeObj(x) })
	Else
		(::cAliasCol)->(DbCloseArea())
		(::cAliasLin)->(DbCloseArea())
		(::cAliasStr)->(DbCloseArea())
		(::cAliasChv)->(DbCloseArea())
		DBSqlExec(::cAliasCol, 'DROP TABLE ' + ::cAliasCol , ::cDriver)
		DBSqlExec(::cAliasCol, 'DROP TABLE ' + ::cAliasLin , ::cDriver)
		DBSqlExec(::cAliasCol, 'DROP TABLE ' + ::cAliasStr , ::cDriver)
		DBSqlExec(::cAliasCol, 'DROP TABLE ' + ::cAliasChv , ::cDriver)
	Endif
	aEval(::aCleanObj, {|x| FreeObj(x) })
	If ::lDelSrv
		DelPasta("\tmpxls\"+::cTmpFile)
	Endif
Return

/*/{Protheus.doc} Gravar
Grava o excel processado
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param cLocal, characters, Local para gerar o arquivo no client
@param lAbrir, logical, Abrir a planilha gerada
@param lDelSrv, logical, Deleta a planilha após copiar para o client
@return characters, cArquivo local do arquivo gerado
@type method
/*/
Method Save(cLocal) Class YExcel
	Local nFile
	Local nCont,nQtdPlanilhas
	Local nCont2
	Local cArquivo	:= ""	//Nome do arquivo em minusculo
	Local cArquivo2	:= ""	//Nome do arquivo no case original
	Local cPath
	Local cDrive,cNome,cExtensao
	Local nHDestino,nHOrigem
	Local nTamArquivo,nBytesFalta,nBytesLidos,cBuffer,nBytesLer,nBytesSalvo,cBuffer2
	Local nPos
	Local oXmlSheet
	Local lServidor
	Local cNumero,nPosPonto,nQtdTmp
	Default cLocal := GetTempPath()
	lServidor	:= !Empty(cLocal) .and. SubStr(cLocal,1,1)=="\"
	::cLocalFile	:= cLocal

	If !Empty(::cLocalFile)
		::cLocalFile	:= Alltrim(::cLocalFile)
		If Right(::cLocalFile,1)=="\"
			::cLocalFile	:= SubStr(::cLocalFile,1,Len(::cLocalFile)-1)
		Endif
	Endif
	If ValType(cRootPath)=="U"
		cRootPath	:= GetSrvProfString( "RootPath", "" )
	Endif

	If Empty(::cNomeFile)
		Return
	Endif
	Private oSelf			:= Self

	FWMakeDir("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\")
	FWMakeDir("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\docprops")
	FWMakeDir("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl")
	::ocontent_types:Save2File("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\[content_types].xml")
	FRename( "\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\[content_types].xml", "\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\[Content_Types].xml", , .F. )
	::oapp:Save2File("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\docprops\app.xml")
	::ocore:Save2File("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\docprops\core.xml")
	::oworkbook:Save2File("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\workbook.xml")
	::oStyle:Save2File("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\styles.xml")

	For nCont:=1 to Len(::aRels)
		If !Empty(::aRels[nCont][3])
			FWMakeDir("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+SubStr(::aRels[nCont][2],1,rAt("\",::aRels[nCont][2])-1),.F.)
			::aRels[nCont][1]:Save2File("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+::aRels[nCont][2])
			AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+::aRels[nCont][2])
		Endif
	Next
	For nCont:=1 to Len(::aDraw)
		If !Empty(::aDraw[nCont][3])
			FWMakeDir("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+SubStr(::aDraw[nCont][2],1,rAt("\",::aDraw[nCont][2])-1),.F.)
			::aDraw[nCont][1]:Save2File("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+::aDraw[nCont][2])
			AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+::aDraw[nCont][2])
		Endif
	Next

	::CriarFile("\"+::cNomeFile+"\xl"				,"sharedStrings.xml"	,""						,)
	GravaFile(@nFile,"","\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl","sharedStrings.xml")
	::xls_sharedStrings(nFile)
	fClose(nFile)
	nFile	:= nil

	::CriarFile("\"+::cNomeFile+"\xl\theme"			,"theme1.xml"			,u_yxlsthe2()			,)

	nQtdPlanilhas	:= Len(::aPlanilhas)
	For nCont:=1 to Len(::asheet)
		FWMakeDir("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\worksheets\",.F.)
		If !::asheet[nCont][1]:XPathHasNode("/xmlns:worksheet/xmlns:cols/xmlns:col[1]")
			::nPlanilhaAt	:= nCont
			::AddTamCol(1,1,12,.T.,.F.)
		Endif
		
		If !Empty(::aPlanilhas[nCont][5])	//tableParts
			::aPlanilhas[nCont][6]:PutTxml(::asheet[nCont][1],"/xmlns:worksheet")
		Endif

		//Cria o sheet na ordem obrigatoria
		oXmlSheet	:= SheetTmp()
		//Ordenar os nodes de acordo com enviado no array
		//aOrdem {{patch,tags}}
		aOrdem	:= {;
					{;
						"/xmlns:worksheet";
						,{"sheetPr","dimension","sheetViews","sheetFormatPr","cols","sheetData","sheetCalcPr","sheetProtection","protectedRanges","scenarios","autoFilter","sortState","dataConsolidate","customSheetViews","mergeCells","phoneticPr","conditionalFormatting","dataValidations","hyperlinks","printOptions","pageMargins","pageSetup","headerFooter","rowBreaks","colBreaks","customProperties","cellWatches","ignoredErrors","smartTags","drawing","drawingHF","picture","oleObjects","controls","webPublishItems","tableParts","extLst"};
					};
					,{;
						"/xmlns:worksheet/xmlns:headerFooter";
						,{"oddHeader","oddFooter","evenHeader","evenFooter","evenFooter","firstHeader","firstFooter"};
					};
					,{;
						"/xmlns:worksheet/xmlns:sheetPr";
						,{"tabColor","outlinePr","pageSetUpPr"};
					};
					,{;
						"/xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView";
						,{"pane","selection","rowBreaks","colBreaks","pageMargins","printOptions","pageSetup","headerFooter","autoFilter","extLst"};
					};
					}
		Xml2Xml(oXmlSheet,::asheet[nCont][1],"/xmlns:worksheet",,,,,aOrdem)

		If Empty(oXmlSheet:XPathGetAtt("/xmlns:worksheet/xmlns:autoFilter","ref"))
			oXmlSheet:XPathDelNode("/xmlns:worksheet/xmlns:autoFilter")
		EndIf
		If Empty(oXmlSheet:XPathGetChildArray("/xmlns:worksheet/xmlns:mergeCells"))
			oXmlSheet:XPathDelNode("/xmlns:worksheet/xmlns:mergeCells")
		EndIf

		oXmlSheet:Save2File("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\worksheets\tmp"+::asheet[nCont][2])
		FreeObj(oXmlSheet)

		nHDestino	:= FCreate("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\worksheets\"+::asheet[nCont][2],FO_READWRITE + FO_SHARED,,.F.)
		nHOrigem	:= FOpen("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\worksheets\tmp"+::asheet[nCont][2],FO_READWRITE + FO_SHARED)
		nTamArquivo := Fseek(nHOrigem,0,2)	//Determina o tamanho do arquivo de origem
		Fseek(nHOrigem,0)					//Move o ponteiro do arquivo de origem para o inicio do arquivo
		nBytesFalta := nTamArquivo			//Define que a quantidade que falta copiar é o próprio tamanho do Arquivo

		cBuffer := SPACE(1024)
		While nBytesFalta > 0
			nBytesLer := Min(nBytesFalta, 1024 )
			nBytesLidos := FREAD(nHOrigem, @cBuffer, nBytesLer )
			If nBytesLidos < nBytesLer
				UserException("Erro de Leitura da Origem. " + Str(nBytesLer,8,2) +;
				" bytes a LER." + Str(nBytesLidos,8,2) + " bytes Lidos." + "Ferror = " + str(ferror(),4))
				Exit
			Endif
			nPos	:= At("<sheetData/>",cBuffer)
			If nPos>0
				nBytesSalvo := FWRITE(nHDestino, SubStr(cBuffer,1,nPos-1))
				FWRITE(nHDestino, "<sheetData>")
				(::cAliasLin)->(DbSeek(Str(nCont,10)))	//Leitura das linhas
				While (::cAliasLin)->(!EOF()) .and. nCont==(::cAliasLin)->PLA
					nLinha := (::cAliasLin)->LIN
					cBuffer2	:= '<row r="'+cValToChar(nLinha)+'"'
					If !Empty((::cAliasLin)->OLEVEL)
						cBuffer2	+= ' outlineLevel="'+(::cAliasLin)->OLEVEL+'"'
					Endif
					If !Empty((::cAliasLin)->COLLAP)
						cBuffer2	+= ' collapsed="'+(::cAliasLin)->COLLAP+'"'
					Endif
					If !Empty((::cAliasLin)->CHIDDEN)
						cBuffer2	+= ' hidden="'+(::cAliasLin)->CHIDDEN+'"'
					Endif
					If !Empty((::cAliasLin)->CHEIGHT)
						cBuffer2	+= ' customHeight="'+(::cAliasLin)->CHEIGHT+'"'
						cBuffer2	+= ' ht="'+cValToChar((::cAliasLin)->HT)+'"'
					Endif
					cBuffer2	+= ">"
					FWRITE(nHDestino, cBuffer2)
					cBuffer2	:= ""
					(::cAliasCol)->(DbSetOrder(1))
					(::cAliasCol)->(DbSeek( Str(nCont,10)+Str(nLinha,10) ))	//Leitura das colunas
					While (::cAliasCol)->(!EOF()) .AND. nCont==(::cAliasCol)->PLA .AND. nLinha==(::cAliasCol)->LIN
						cBuffer2	:= '<c r="'+::Ref(nLinha,(::cAliasCol)->COL)+'"'
						If (::cAliasCol)->STY>=0
							cBuffer2	+= ' s="'+cValToChar((::cAliasCol)->STY)+'"'
						Endif
						If !Empty((::cAliasCol)->TIPO) .AND. (::cAliasCol)->TPVLR!="U" .AND. !( ((::cAliasCol)->TIPO =="d".AND.(::cAliasCol)->TPVLR=="N") .OR. (::cAliasCol)->TIPO =="n" )	//não incluir atributo "t" quando data serializada e numero
							cBuffer2	+= ' t="'+Alltrim((::cAliasCol)->TIPO)+'"'
						Endif
						cBuffer2	+= '>'
						nColuna	:= (::cAliasCol)->COL
						If !Empty((::cAliasCol)->FORMULA)
							cBuffer2	+= '<f>'+RTRIM((::cAliasCol)->FORMULA)+'</f>'
						Endif
						cBuffer2	+= '<v>'
						If (::cAliasCol)->TPVLR=="C"
							cBuffer2	+= Alltrim((::cAliasCol)->VLRTXT)
						ElseIf (::cAliasCol)->TPVLR=="U"
						Else
							cNumero		:= cValToChar((::cAliasCol)->VLRNUM)
							If (::cAliasCol)->VLRDEC<>0
								nPosPonto	:= At(".",cNumero)
								If nPosPonto==0
									cNumero	+= "."+Replicate("0",8)
								Else
									nQtdTmp	:= 8-Len(SubStr(cNumero,nPosPonto+1))
									If nQtdTmp>0
										cNumero	+= Replicate("0",nQtdTmp)
									Endif
								Endif
								cNumero	+= cValToChar((::cAliasCol)->VLRDEC)
							Endif
							cBuffer2	+= cNumero
						Endif
						cBuffer2	+= '</v>'
						cBuffer2	+= '</c>'
						FWRITE(nHDestino, cBuffer2)
						(::cAliasCol)->(DbSkip())
					EndDo
					cBuffer2	:= nil
					FWRITE(nHDestino, "</row>")
					(::cAliasLin)->(DbSkip())
				EndDo
				FWRITE(nHDestino, "</sheetData>")
				nBytesSalvo += FWRITE(nHDestino, SubStr(cBuffer,nPos+12))
				nBytesSalvo += 12
			Else
				nBytesSalvo := FWRITE(nHDestino, cBuffer,nBytesLer)
			Endif
			If nBytesSalvo < nBytesLer
				UserException("Erro de gravação do Destino. " + Str(nBytesLer,8,2) +;
				" bytes a SALVAR." + Str(nBytesSalvo,8,2) + " bytes gravados." + "Ferror = " + str(ferror(),4))
				EXIT
			Endif
			// Elimina do Total do Arquivo a quantidade de bytes copiados
			nBytesFalta -= nBytesLer
		EndDo
		FCLOSE(nHOrigem)
		FErase("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\worksheets\tmp"+::asheet[nCont][2])
		FCLOSE(nHDestino)

		If !Empty(::asheet[nCont][3])	//Se possuir comments
			::asheet[nCont][3]:Save2File(::asheet[nCont][4])
			AADD(::aFiles,::asheet[nCont][4])
		EndIf
		If !Empty(::asheet[nCont][5])	//Se possuir comments
			::asheet[nCont][5]:Save2File(::asheet[nCont][6])
			AADD(::aFiles,::asheet[nCont][6])
		EndIf

		For nCont2:=1 to Len(::aPlanilhas[nCont][5])
			::CriarFile("\"+::cNomeFile+"\xl\tables\"	,"table"+cValToChar(::aPlanilhas[nCont][5][nCont2]:nIdRelat)+".xml"		,::xls_table(nCont,nCont2)		,)
			AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\tables\table"+cValToChar(::aPlanilhas[nCont][5][nCont2]:nIdRelat)+".xml")
		Next
	Next

	If lServidor
		cArquivo	:= ::cLocalFile+'\'+::cNomeFile+'.xlsx'
		cArquivo2	:= ::cLocalFile+'\'+::cNomeFile2+'.xlsx'
		::cLocalFile:= ""
	Else
		cArquivo	:= '\tmpxls\'+::cTmpFile+'\'+::cNomeFile+'.xlsx'
		cArquivo2	:= '\tmpxls\'+::cTmpFile+'\'+::cNomeFile2+'.xlsx'
	Endif
	SplitPath(cArquivo,@cDrive,@cPath,@cNome,@cExtensao)
	cNome	:= SubStr(cArquivo,Rat("\",cArquivo)+1)	//Split não está respeitando o case original
	If !Empty(cPath)
		FWMakeDir(cPath,.F.)	//Cria a estrutura de pastas
	Endif

	If !FindFunction("FZIP")
		WaitRunSrv('"'+cAr7Zip+'" a -tzip "'+cRootPath+cArquivo+'" "'+cRootPath+'\tmpxls\'+::cTmpFile+'\'+::cNomeFile+'\*"',.T.,"C:\")
	Else
		If IsSrvUnix()	//Solução para servidor linux zipar arquivos com inicio "."
			WaitRunSrv('zip -r "'+cRootPath+replace(cArquivo,"\","/")+'" *',.T.,cRootPath+'/tmpxls/'+self:cTmpFile+'/'+self:cNomeFile+'/')
		Else
			StartJob("FZip",GetEnvServer(),.T.,cArquivo,::aFiles,"\tmpxls\"+::cTmpFile+'\'+::cNomeFile+'\')
		Endif
	Endif

	If !(::cNomeFile==::cNomeFile2)	//Ajusta Case
		FRename( cArquivo, cArquivo2, , .F. )
	Endif

	//Apaga arquivos temporarios listados
	For nCont:=1 to Len(::aFiles)
		If fErase(::aFiles[nCont],,.F.)<>0
			ConOut(::aFiles[nCont])
			ConOut("Ferror:"+cValToChar(ferror()))
		Endif
	Next

	DelPasta("\tmpxls\"+::cTmpFile+"\"+::cNomeFile)	//Apaga arquivos temporarios
	If substr(cArquivo2,1,8)<>"\tmpxls\"
		DelPasta("\tmpxls\"+::cTmpFile)
	Endif
	If !Empty(::cLocalFile)
		If GetRemoteType() == REMOTE_HTML
			CpyS2TW(cArquivo2, .T.)
		Else
			FWMakeDir(::cLocalFile,.F.)
			CpyS2T( cArquivo2,::cLocalFile,,.F.)
			cArquivo2	:= ::cLocalFile+'\'+::cNomeFile2+'.xlsx'
			::cArqGrv	:= cArquivo2
		Endif
		::lDelSrv	:= .F.
	Endif
Return cArquivo2

/*/{Protheus.doc} CpyPasta
Copia pasta para o servidor
@type function
@version 1.0
@author Saulo Gomes Martins
@since 17/12/2020
@param cCaminho, character, pasta no client
@param cCaminho2, character, pasta no servidor
/*/
Static Function CpyPasta(cCaminho,cCaminho2)
	Local nCont
	Local aFiles
	Local nHOrigem,nHDestino
	Local nBytesFalta,nBytesLer,nBytesLidos,cBuffer
	cCaminho	:= Replace(cCaminho,"/","\")
	cCaminho2	:= lower(cCaminho2)
	aFiles		:= Directory(cCaminho+"\*","HSD",,.F.)
	If Right(cCaminho2,1)=="\"
		cCaminho2	:= SubStr(cCaminho2,1,Len(cCaminho2)-1)
	Endif
	MakeDir(cCaminho2,,.F.)
	For nCont:=1 to Len(aFiles)
		If aFiles[nCont][1]=="." .or. aFiles[nCont][1]==".."
			Loop
		Endif
		If aFiles[nCont][5] $ "D"
			CpyPasta(cCaminho+"\"+aFiles[nCont][1],cCaminho2+"\"+aFiles[nCont][1])
		Else
			If !__COPYFILE(cCaminho+"\"+aFiles[nCont][1],cCaminho2+"\"+lower(aFiles[nCont][1]),,,.F.)
				nHOrigem	:= fOpen(cCaminho+"\"+aFiles[nCont][1])
				nHDestino	:= FCreate(cCaminho2+"\"+lower(aFiles[nCont][1]), , , .F.)
				nTamArquivo := Fseek(nHOrigem,0,2)
				Fseek(nHOrigem,0)
				nBytesFalta := nTamArquivo
				While nBytesFalta > 0
					nBytesLer := Min(nBytesFalta, 1024 )
					nBytesLidos := FREAD(nHOrigem, @cBuffer, nBytesLer )
					If nBytesLidos < nBytesLer
						UserException("Erro de Leitura da Origem. " + Str(nBytesLer,8,2) +;
						" bytes a LER." + Str(nBytesLidos,8,2) + " bytes Lidos." + "Ferror = " + str(ferror(),4))
						Exit
					Endif
					nBytesSalvo := FWRITE(nHDestino, cBuffer)
					If nBytesSalvo < nBytesLer
						UserException("Erro de gravação do Destino. " + Str(nBytesLer,8,2) +;
						" bytes a SALVAR." + Str(nBytesSalvo,8,2) + " bytes gravados." + "Ferror = " + str(ferror(),4))
						EXIT
					Endif
				EndDo
				FCLOSE(nHOrigem)
				FCLOSE(nHDestino)
				//UserException("Erro ao Copiar caminho "+cCaminho+"\"+aFiles[nCont][1]+" -> "+cCaminho2+"\"+aFiles[nCont][1])
			Endif
		Endif
	Next
Return

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
	Local aFiles	:= Directory(cCaminho+"\*","HSD",,.F.)
	For nCont:=1 to Len(aFiles)
		If aFiles[nCont][1]=="." .or. aFiles[nCont][1]==".."
			Loop
		Endif
		If aFiles[nCont][5] $ "D"
			DelPasta(cCaminho+"\"+aFiles[nCont][1])
		Else
//			ConOut("Deletando:"+cCaminho+"\"+aFiles[nCont][1])
			If fErase(cCaminho+"\"+aFiles[nCont][1],,.F.)<>0
				ConOut(cCaminho+"\"+aFiles[nCont][1])
				ConOut("Ferror:"+cValToChar(ferror()))
			Endif
		Endif
	Next
//	ConOut("Apagando pasta:"+cCaminho)
	If !DirRemove(cCaminho,,.F.)
		ConOut(cCaminho)
		ConOut("Ferror:"+cValToChar(ferror()))
	Endif
Return
//NÃO DOCUMENTAR
METHOD CriarFile(cLocal,cNome,cString) Class YExcel
	Local cDirServ	:= "\tmpxls\"+::cTmpFile
	Local lOk			:= .T.
	Local nFile
	If ValType(cString)!="C"
		return lOk
	Endif
	FWMakeDir(cDirServ+cLocal,.F.)
	//oFile	:= FWFileIOBase():New(cDirServ+cLocal+"\"+cNome)
	//oFile:SetCaseSensitive()
	If !File(cDirServ+cLocal+"\"+cNome,,.F.)
		nFile	:= FCreate(cDirServ+cLocal+"\"+cNome, , , .F.)
//		oFile:Create()
	Else
		fErase(cDirServ+cLocal+"\"+cNome,,.F.)
		nFile	:= FCreate(cDirServ+cLocal+"\"+cNome, , , .F.)
//		oFile:Create()
	Endif
	FClose(nFile)
//	oFile:Close()
	nFile	:= FOPEN(cDirServ+cLocal+"\"+cNome, FO_READWRITE,,.F.)
	cString	:= EncodeUTF8(cString)
	IF FWrite(nFile, cString, Len(cString)) < Len(cString)
		lOk	:= .F.
	Endif
	fClose(nFile)
Return lOk
//NÃO DOCUMENTAR, USADO NA GRAVAÇÃO DO SHEET
METHOD GravaFile(nFile,cString,cLocal,cArquivo) Class YExcel
Return GravaFile(nFile,cString,cLocal,cArquivo)

Static Function GravaFile(nFile,cString,cLocal,cArquivo)
	Local lOk			:= .T.
	If ValType(cString)=="C"
	Endif
	If !Empty(cArquivo)
		nFile	:= FOPEN(cLocal+"\"+cArquivo, FO_READWRITE,,.F.)
	Endif
	cString	:= EncodeUTF8(cString)
	FSeek(nFile, 0, FS_END)
	IF FWrite(nFile, cString, Len(cString)) < Len(cString)
		lOk	:= .F.
	Endif
Return lOk

/*/{Protheus.doc} YExcel::SetColLevel
Defini o nível das colunas informadas (agrupamento de colunas)
@type method
@version 1.0
@author Saulo Gomes Martins
@since 16/03/2021
@param nMin, numeric, Coluna Inicial
@param nMax, numeric, Coluna Final
@param nNivel, numeric, Nivel
@param lFechado, logical, Se esse nível está fechado
/*/
Method SetColLevel(nMin,nMax,nNivel,lFechado) Class YExcel
	Local nCont		:= nMin-1
	Local cPath
	Local cNivelAtu
	Local lsummaryRight	:= .T.		//Resumo abaixo
	Local csummaryRight	:= ::asheet[::nPlanilhaAt][1]:XPathGetAtt("/xmlns:worksheet/xmlns:sheetPr/xmlns:outlinePr","summaryRight")
	PARAMTYPE 0	VAR nMin			AS NUMERIC
	PARAMTYPE 1	VAR nMax			AS NUMERIC		OPTIONAL DEFAULT nMin
	PARAMTYPE 2	VAR nNivel			AS NUMERIC		OPTIONAL
	PARAMTYPE 3	VAR lFechado		AS LOGICAL		OPTIONAL DEFAULT .F.
	If !Empty(csummaryRight) .AND. csummaryRight=="0"
		lsummaryRight	:= .F.
	Endif
	If ValType(nNivel)!="N"
		lFechado	:= .F.
	Endif

	AjtColConf(self,nMin,nMax)
	If !lsummaryRight .AND. lFechado .AND. nCont>0
		cPath	:= ColNew(self,nCont)
		If lFechado
			If !Empty(::asheet[::nPlanilhaAt][1]:XPathGetAtt(cPath, "collapsed"))
				::asheet[::nPlanilhaAt][1]:XPathSetAtt(cPath, "collapsed"	, "1" )
			Else
				::asheet[::nPlanilhaAt][1]:XPathDelAtt(cPath, "collapsed")
			EndIf
		EndIf
	Endif

	For nCont:=nMin to nMax
		cPath	:= ColNew(self,nCont)
		
		cNivelAtu	:= ::asheet[::nPlanilhaAt][1]:XPathGetAtt(cPath, "outlineLevel")
		If !Empty(cNivelAtu)
			If nNivel>Val(cNivelAtu)
				::asheet[::nPlanilhaAt][1]:XPathSetAtt(cPath, "outlineLevel"	, cValToChar(nNivel) )
				If !lFechado
					::asheet[::nPlanilhaAt][1]:XPathDelAtt(cPath, "hidden")
				EndIf
			EndIf
		Else
			::asheet[::nPlanilhaAt][1]:XPathAddAtt(cPath, "outlineLevel"	, cValToChar(nNivel) )
			If !lFechado
				::asheet[::nPlanilhaAt][1]:XPathDelAtt(cPath, "hidden")
			EndIf
		EndIf

		If lFechado
			If !Empty(::asheet[::nPlanilhaAt][1]:XPathGetAtt(cPath, "hidden"))
				::asheet[::nPlanilhaAt][1]:XPathSetAtt(cPath, "hidden"	, "1" )
			Else
				::asheet[::nPlanilhaAt][1]:XPathAddAtt(cPath, "hidden"	, "1" )
			EndIf
		Endif
	Next
	If lsummaryRight .AND. lFechado
		cPath	:= ColNew(self,nCont)
		If lFechado
			If !Empty(::asheet[::nPlanilhaAt][1]:XPathGetAtt(cPath, "collapsed"))
				::asheet[::nPlanilhaAt][1]:XPathSetAtt(cPath, "collapsed"	, "1" )
			Else
				::asheet[::nPlanilhaAt][1]:XPathDelAtt(cPath, "collapsed")
			EndIf
		EndIf
	Endif
Return
/*/{Protheus.doc} ColNew
Cria uma nova definição de coluna ou usa a já existente
@type function
@version 1.0
@author Saulo Gomes Martins
@since 16/03/2021
@param oExcel, object, Objeto YExcel
@param nCont, numeric, Numero da coluna
@return characters, Path (caminho do xml)
/*/
Static Function ColNew(oExcel,nCont)
	Local cPath
	If !oExcel:asheet[oExcel:nPlanilhaAt][1]:XPathHasNode("/xmlns:worksheet/xmlns:cols/xmlns:col[@min='"+cValToChar(nCont)+"']")
		oExcel:asheet[oExcel:nPlanilhaAt][1]:XPathAddNode( "/xmlns:worksheet/xmlns:cols", "col", "" )
		cPath	:= "/xmlns:worksheet/xmlns:cols/xmlns:col[last()]"
	Else
		cPath	:= "/xmlns:worksheet/xmlns:cols/xmlns:col[@min='"+cValToChar(nCont)+"']"
	EndIf

	If !Empty(oExcel:asheet[oExcel:nPlanilhaAt][1]:XPathGetAtt(cPath, "min"))
		oExcel:asheet[oExcel:nPlanilhaAt][1]:XPathSetAtt(cPath, "min"		, cValToChar(nCont) )
	Else
		oExcel:asheet[oExcel:nPlanilhaAt][1]:XPathAddAtt(cPath, "min"		, cValToChar(nCont) )
	Endif
	
	If !Empty(oExcel:asheet[oExcel:nPlanilhaAt][1]:XPathGetAtt(cPath, "max"))
		oExcel:asheet[oExcel:nPlanilhaAt][1]:XPathSetAtt(cPath, "max"		, cValToChar(nCont) )
	Else
		oExcel:asheet[oExcel:nPlanilhaAt][1]:XPathAddAtt(cPath, "max"		, cValToChar(nCont) )
	EndIf
Return cPath

/*/{Protheus.doc} AddTamCol
Defini o tamanho de uma coluna ou varias colunas
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param nMin, numeric, Coluna inicial
@param nMax, numeric, Coluna final
@param nWidth, numeric, descricao
@param lbestFit, logical, melhor ajuste numerico
@param lcustomWidth, logical, tamanho customizado
@type method
/*/
Method AddTamCol(nMin,nMax,nWidth,lbestFit,lcustomWidth) Class YExcel
	Local nCont
	Local cPath
	PARAMTYPE 0	VAR nMin			AS NUMERIC
	PARAMTYPE 1	VAR nMax			AS NUMERIC		OPTIONAL DEFAULT nMin
	PARAMTYPE 2	VAR nWidth			AS NUMERIC		OPTIONAL DEFAULT 12
	PARAMTYPE 3	VAR lbestFit		AS LOGICAL		OPTIONAL DEFAULT .T.
	PARAMTYPE 4	VAR lcustomWidth	AS LOGICAL		OPTIONAL DEFAULT .T.
	nWidth := nWidth+0.7109375	//Microsoft excel soma esse valor na coluna
	
	AjtColConf(self,nMin,nMax)

	For nCont:=nMin to nMax
		cPath	:= ColNew(self,nCont)
		
		If !Empty(::asheet[::nPlanilhaAt][1]:XPathGetAtt(cPath, "width"))
			::asheet[::nPlanilhaAt][1]:XPathSetAtt(cPath, "width"	, cValToChar(nWidth) )
		Else
			::asheet[::nPlanilhaAt][1]:XPathAddAtt(cPath, "width"	, cValToChar(nWidth) )
		EndIf

		If lbestFit
			If !Empty(::asheet[::nPlanilhaAt][1]:XPathGetAtt(cPath, "bestFit"))
				::asheet[::nPlanilhaAt][1]:XPathSetAtt(cPath, "bestFit"	, "1" )
			Else
				::asheet[::nPlanilhaAt][1]:XPathAddAtt(cPath, "bestFit"	, "1" )
			EndIf
		Else
			::asheet[::nPlanilhaAt][1]:XPathDelAtt(cPath, "bestFit")
		Endif
		If lcustomWidth
			If !Empty(::asheet[::nPlanilhaAt][1]:XPathGetAtt(cPath, "customWidth"))
				::asheet[::nPlanilhaAt][1]:XPathSetAtt(cPath, "customWidth"	, "1" )
			Else
				::asheet[::nPlanilhaAt][1]:XPathAddAtt(cPath, "customWidth"	, "1" )
			EndIf
		Else
			::asheet[::nPlanilhaAt][1]:XPathDelAtt(cPath, "customWidth")
		Endif
	Next
Return

/*/{Protheus.doc} AjtColConf
Ajusta conflito de coluna entre min e max antes de incluir definição
@type function
@version 1.0
@author Saulo Gomes Martins
@since 16/03/2021
@param oExcel, object, Objeto yexcel
@param nMin, numeric, Coluna minima definida
@param nMax, numeric, Coluna maxima definida
/*/
Static Function AjtColConf(oExcel,nMin,nMax)
	Local nMinTmp
	Local nMaxTmp
	Local lConflito	:= .T.	//Verificar conflito de coluna
	Local aAtr
	Local aChildren
	Local nCont,nCont2,nCont3
	While lConflito
		lConflito	:= .F.
		aChildren := oExcel:asheet[oExcel:nPlanilhaAt][1]:XPathGetChildArray( "/xmlns:worksheet/xmlns:cols" )
		For nCont:=1 to Len(aChildren)
			nMinTmp	:= Val(oExcel:asheet[oExcel:nPlanilhaAt][1]:XPathGetAtt("/xmlns:worksheet/xmlns:cols/xmlns:col["+cValToChar(nCont)+"]","min"))
			nMaxTmp	:= Val(oExcel:asheet[oExcel:nPlanilhaAt][1]:XPathGetAtt("/xmlns:worksheet/xmlns:cols/xmlns:col["+cValToChar(nCont)+"]","max"))
			//Se tem conflito de intervalo vai dividir o intervalo
			If nMinTmp!=nMaxTmp .AND. ((nMin>=nMinTmp .AND. nMin<=nMaxTmp) .OR. (nMax>=nMinTmp .AND. nMax<=nMaxTmp))
				aAtr	:= oExcel:asheet[oExcel:nPlanilhaAt][1]:XPathGetAttArray("/xmlns:worksheet/xmlns:cols/xmlns:col["+cValToChar(nCont)+"]")
				For nCont2:=nMinTmp to nMaxTmp
					oExcel:asheet[oExcel:nPlanilhaAt][1]:XPathAddNode( "/xmlns:worksheet/xmlns:cols", "col", "" )
					For nCont3:=1 to Len(aAtr)
						If aAtr[nCont3][1]=="min"
							oExcel:asheet[oExcel:nPlanilhaAt][1]:XPathAddAtt("/xmlns:worksheet/xmlns:cols/xmlns:col[last()]", "min"	, cValToChar(nCont2) )
						ElseIf aAtr[nCont3][1]=="max"
							oExcel:asheet[oExcel:nPlanilhaAt][1]:XPathAddAtt("/xmlns:worksheet/xmlns:cols/xmlns:col[last()]", "max"	, cValToChar(nCont2) )
						Else
							oExcel:asheet[oExcel:nPlanilhaAt][1]:XPathAddAtt("/xmlns:worksheet/xmlns:cols/xmlns:col[last()]", aAtr[nCont3][1]	, aAtr[nCont3][2] )
						EndIf
					Next
				Next
				oExcel:asheet[oExcel:nPlanilhaAt][1]:XPathDelNode("/xmlns:worksheet/xmlns:cols/xmlns:col["+cValToChar(nCont)+"]")
				lConflito	:= .T.
				Exit
			Endif
		Next
	EndDo
Return

/*/{Protheus.doc} OpenRead
Abrir planilha e armazena conteudo para leitura
@author Saulo Gomes Martins
@since 03/03/2018
@version 1.0
@return logical, lRet Se conseguiu ler a planilha
@param cFile, characters, arquivo que será aberto
@param nPlanilha, numeric, numero(indexado em 1,2,3) da planilha a ser lida
@type method
/*/
METHOD OpenRead(cFile,nPlanilha) Class YExcel
	Local nRet
	Local cDrive, cDir, cNome, cExt
	Local nCont,nCont2,nCont3
	Local cTipo,cRef
	Local aChildren,aChildren2,aAtributos,oXml
	Local oXmlStyle
	Local cCamSrv	:= ""
	Local cCamLocal	:= ""
	Local cNomeNS	:= "ns"
	Local cNomeNS2	:= "ns"
	Local aStyles
	Local nPosR
	Local nPos
	Local cnumfmtid
	Local aAtrr
	Local oDataTime
	ConOut("[Warning] Method deprecated, see https://github.com/saulogm/advpl-excel/wiki/Exemplo-ler-planilha")
	PARAMTYPE 0	VAR cFile			AS CHARACTER
	PARAMTYPE 1	VAR nPlanilha		AS NUMERIC		OPTIONAL DEFAULT 1
	cFile	:= Alltrim(cFile)
	If !File(cFile,,.F.)
		ConOut("Arquivo nao encontrado!")
		Return .F.
	Endif
	If ValType(cRootPath)=="U"
		cRootPath	:= GetSrvProfString( "RootPath", "" )
	Endif
	::oCell := tHashMap():new()
	If !File("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\worksheets\sheet"+cValTochar(nPlanilha)+".xml",,.F.)
		SplitPath( cFile, @cDrive, @cDir, @cNome, @cExt)
		cNome	:= SubStr(cFile,Rat("\",cFile)+1)	//Split não está respeitando o case original
		FWMakeDir("\tmpxls\"+::cTmpFile+'\',.F.)
		FWMakeDir("\tmpxls\"+::cTmpFile+'\'+::cNomeFile+'\',.F.)
		If ":" $ UPPER(cFile)
			CpyT2S(cFile,"\tmpxls\"+::cTmpFile+'\',,.F.)
			cCamSrv	:= cRootPath+"\tmpxls\"+::cTmpFile+'\'+cNome
			cCamLocal:= "\tmpxls\"+::cTmpFile+'\'+cNome
		Else
			cCamSrv	:= cRootPath+cFile
			cCamLocal:= cFile
		Endif
		If !FindFunction("FZIP")
			WaitRunSrv('"'+cAr7Zip+'" x -tzip "'+cCamSrv+'" -o"'+cRootPath+'\tmpxls\'+::cTmpFile+'\'+::cNomeFile+'" * -r -y',.T.,"C:\")
			If !File("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\sharedStrings.xml",,.F.)
				nRet	:= -1
				ConOut("Arquivo nao descompactado!")
				Return .F.
			Else
				nRet	:= 0
			Endif
		Else
			nRet	:= StartJob("FUnZip",GetEnvServer(),.T.,cCamLocal,"\tmpxls\"+::cTmpFile+'\'+::cNomeFile+'\')
		Endif
		If nRet!=0
			ConOut(Ferror())
			ConOut("Arquivo nao descompactado!")
			Return .F.
		Endif
		oXml	:= TXmlManager():New()
		If oXML:ParseFile("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\sharedStrings.xml")
			oXML:XPathRegisterNs( "ns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main" )
			aChildren := oXML:XPathGetChildArray( "/ns:sst" )
			For nCont:=1 to Len(aChildren)
				::oString:Set(::nQtdString,oXML:XPathGetNodeValue("/ns:sst/ns:si["+cValToChar(nCont)+"]/ns:t"))
				::nQtdString++
			Next
		Endif
	Endif
	oXml	:= TXmlManager():New()
	oXML:ParseFile("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\worksheets\sheet"+cValTochar(nPlanilha)+".xml")
	aNs	:= oXML:XPathGetRootNsList()
	For nCont:=1 to Len(aNs)
		oXML:XPathRegisterNs( aNs[nCont][1], aNs[nCont][2] )
		If alltrim(lower(aNs[nCont][2]))==lower("http://schemas.openxmlformats.org/spreadsheetml/2006/main")
			cNomeNS	:= aNs[nCont][1]
		Endif
	Next
	oXmlStyle	:= TXmlManager():New()
	oXmlStyle:ParseFile("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\styles.xml")
	aNs	:= oXmlStyle:XPathGetRootNsList()
	For nCont:=1 to Len(aNs)
		oXmlStyle:XPathRegisterNs( aNs[nCont][1], aNs[nCont][2] )
		If alltrim(lower(aNs[nCont][2]))==lower("http://schemas.openxmlformats.org/spreadsheetml/2006/main")
			cNomeNS2	:= aNs[nCont][1]
		Endif
	Next
	aStyles	:= oXmlStyle:XPathGetChildArray("/"+cNomeNS2+":styleSheet/"+cNomeNS2+":cellXfs")

	aChildren := oXML:XPathGetChildArray("/"+cNomeNS+":worksheet/"+cNomeNS+":sheetData")
	::adimension	:= {{0,0},{999999,999999}}
	For nCont:=1 to Len(aChildren)	//Row
		aChildren2	:= oXML:XPathGetChildArray( aChildren[nCont][2] )
		For nCont2:=1 to Len(aChildren2)	//c
			cTipo		:= "N"
			aAtributos	:= oXML:XPathGetAttArray(aChildren2[nCont2][2])						//Atributos do elemento
			cRet		:= oXML:XPathGetNodeValue("/"+cNomeNS+":worksheet/"+cNomeNS+":sheetData/"+cNomeNS+":row["+cValToChar(nCont)+"]/"+cNomeNS+":c["+cValToChar(nCont2)+"]/"+cNomeNS+":v")
			nPosR		:= aScan(aAtributos,{|x| x[1]=="r"})
			If nPosR==0
				cRef		:= ::ref(nCont,nCont2)
				aPosicao	:= {nCont,nCont2}
			Else
				cRef		:= aAtributos[nPosR][2]
				aPosicao	:= ::LocRef(cRef)	//Retorna linha e coluna
			Endif
			If ::adimension[2][1]>aPosicao[1]	//Menor linha
				::adimension[2][1] := aPosicao[1]
			Endif
			If ::adimension[2][2]>aPosicao[2]	//Menor Coluna
				::adimension[2][2]	:= aPosicao[2]
			Endif
			If ::adimension[1][1]<aPosicao[1]	//Maior Linha
				::adimension[1][1]	:= aPosicao[1]
			Endif
			If ::adimension[1][2]<aPosicao[2]	//Maior Coluna
				::adimension[1][2]	:= aPosicao[2]
			Endif
			For nCont3:=1 to Len(aAtributos)
				If aAtributos[nCont3][1]=="t" .and. aAtributos[nCont3][2]=="str"
					cTipo	:= "C"
				ElseIf aAtributos[nCont3][1]=="t" .and. aAtributos[nCont3][2]=="inlineStrs"
					cTipo	:= "C"
					cRet		:= oXML:XPathGetNodeValue("/"+cNomeNS+":worksheet/"+cNomeNS+":sheetData/"+cNomeNS+":row["+cValToChar(nCont)+"]/"+cNomeNS+":c["+cValToChar(nCont2)+"]/"+cNomeNS+":is/"+cNomeNS+":t")
				ElseIf aAtributos[nCont3][1]=="t" .and. aAtributos[nCont3][2]=="s"
					cRet	:= ""
					cTipo	:= "C"
					::oString:Get(Val(oXML:XPathGetNodeValue("/"+cNomeNS+":worksheet/"+cNomeNS+":sheetData/"+cNomeNS+":row["+cValToChar(nCont)+"]/"+cNomeNS+":c["+cValToChar(nCont2)+"]/"+cNomeNS+":v")),@cRet)
				ElseIf aAtributos[nCont3][1]=="t" .and. aAtributos[nCont3][2]=="b"
					cTipo	:= "L"
					cRet	:= oXML:XPathGetNodeValue("/"+cNomeNS+":worksheet/"+cNomeNS+":sheetData/"+cNomeNS+":row["+cValToChar(nCont)+"]/"+cNomeNS+":c["+cValToChar(nCont2)+"]/"+cNomeNS+":v")=="1"
//			ElseIf aAtributos[nCont3][1]=="s" .and. aAtributos[nCont3][2]=="1"
//				cTipo	:= "D"
//				cRet	:= STOD("19000101")-2+Val(oXML:XPathGetNodeValue("/"+cNomeNS+":worksheet/"+cNomeNS+":sheetData/"+cNomeNS+":row["+cValToChar(nCont)+"]/"+cNomeNS+":c["+cValToChar(nCont2)+"]/"+cNomeNS+":v"))
				ElseIf aAtributos[nCont3][1]=="s"
					aAtrr	:= oXmlStyle:XPathGetAttArray(aStyles[Val(aAtributos[nCont3][2])+1][2])
					cnumfmtid	:= ""
					nPos	:= aScan(aAtrr,{|x| lower(x[1])=="numfmtid"})
					If nPos>0
						cnumfmtid	:= aAtrr[nPos][2]
						If "|"+cnumfmtid+"|" $ "|14|15|16|17|18|19|20|21|22|45|46|47|"
							cTipo		:= "D"
							oDataTime	:= yExcel_DateTime():New(,,oXML:XPathGetNodeValue("/"+cNomeNS+":worksheet/"+cNomeNS+":sheetData/"+cNomeNS+":row["+cValToChar(nCont)+"]/"+cNomeNS+":c["+cValToChar(nCont2)+"]/"+cNomeNS+":v"))
							cRet		:= oDataTime:GetDate()
							::oCell:Set(cRef+"_H",oDataTime:GetTime())
							FreeObj(oDataTime)
						Endif
					Endif
				Endif
			Next
			If cTipo=="N"
				::oCell:Set(cRef,Val(cRet))
			Else
				::oCell:Set(cRef,cRet)
			Endif
		Next
	Next
Return nRet==0

/*/{Protheus.doc} CellRead
Retorna o valor de uma celula, após o uso do método OpenRead()
@author Saulo Gomes Martins
@since 03/03/2018
@version 1.0
@return variadic, xValor Conteúdo da celula
@param nLinha, numeric, Linha da informação
@param nColuna, numeric, Coluna da informação
@param xDefault, variadic , Valor padrão caso não tenha a informação
@param lAchou, logical, passa por referencia se achou a informação da celula
@type method
/*/
Method CellRead(nLinha,nColuna,xDefault,lAchou,cOutro) Class YExcel
	Local cRef	:= ::Ref(nLinha,nColuna)
	Local xValor:= Nil
	ConOut("[Warning] Method deprecated, see https://github.com/saulogm/advpl-excel/wiki/Exemplo-ler-planilha")
	Default cOutro	:= ""
	lAchou	:= .T.
	If !::oCell:Get(cRef+cOutro,@xValor)
		xValor	:= xDefault
		lAchou	:= .F.
	Endif
Return xValor

/*/{Protheus.doc} CloseRead
Limpa a pasta temporaria
@author Saulo Gomes Martins
@since 03/03/2018
@version 1.0

@type method
/*/
METHOD CloseRead() Class YExcel
	ConOut("[Warning] Method deprecated, see https://github.com/saulogm/advpl-excel/wiki/Exemplo-ler-planilha")
	::oString:clean()
	::oCell:clean()
	::nQtdString := 0
	DelPasta("\tmpxls\"+::cTmpFile)
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
			cRet	+= "Z"
		Else
			cRet	+= ColunasIndex((nNum-(nNum % 26))/26)
			cRet	+= ColunasIndex(nNum % 26)
		Endif
	Else
		IF nNum % 26==0
			cRet	+= NumToString(((nNum-(nNum % 26))/26)-1)
			cRet	+= "Z"
		Else
			cRet	+= NumToString((nNum-(nNum % 26))/26)
			cRet	+= ColunasIndex(nNum % 26)
		Endif
	Endif
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
	Endif
Return nRet

Static Function ColunasIndex(xNum,nIdx)
	Local cRet		:= ""
	Default nIdx	:= 1
	If nIdx==1
		cRet	:= chr(xNum+64)
	Else
		cRet	:= asc(xNum)-64
	Endif
	// Local cRet		:= ""
	// Local nPos
	// Default nIdx	:= 1
	// nPos	:= aScan(aColIdx,{|x| x[nIdx]==xNum})
	// If nPos>0
	// 	If nIdx==1
	// 		cRet	:= aColIdx[nPos][2]
	// 	Else
	// 		cRet	:= aColIdx[nPos][1]
	// 	Endif
	// Endif
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
Class YExcelTag
	Data cNome
	Data cClassName
	Data oAtributos
	Data oIndice
	Data xValor
	Data oExcel			//Objeto referencia do yexcel
	Data xDados			//Outros dados
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
	Method PutTxml()
	Method LoadTagXml()
EndClass

Method New(cNome,xValor,oAtributo,oExcel) Class YExcelTag
	Local nCont
	PARAMTYPE 0	VAR cNome  AS CHARACTER
	PARAMTYPE 1	VAR xValor  AS ARRAY, CHARACTER, DATE, NUMERIC, LOGICAL, OBJECT		OPTIONAL DEFAULT Nil
	PARAMTYPE 2	VAR oAtributo  AS ARRAY,OBJECT		OPTIONAL DEFAULT tHashMap():new()
	::oExcel		:= oExcel
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
	Endif
	::cClassName	:= "YEXCELTAG"
	If ValType(::oExcel)=="O"
		AADD(::oExcel:aCleanObj,self)
	Endif
Return self

Method GetNome() Class YExcelTag
Return ::cNome

Method ClassName() Class YExcelTag
Return "YEXCELTAG"

Method SetValor(xValor,xIndice) Class YExcelTag
	If ValType(xIndice)=="U"
		::xValor	:= xValor
	ElseIf ValType(xIndice)=="N"
		::xValor[xIndice]	:= xValor
	ElseIf ValType(xIndice)=="C" .and. ValType(::xValor)=="A"
		::AddValor(xValor,xIndice)
	Else
		::xValor	:= xValor
	Endif
Return

Method GetValor(xIndice,xDefault) Class YExcelTag
	Local nPos
	If ValType(xIndice)=="U"
		xDefault	:=  ::xValor
	ElseIf ValType(xIndice)=="N"
		xDefault	:=  ::xValor[xIndice]
	ElseIf ValType(xIndice)=="C" .and. ValType(::xValor)=="A"
		If ::oIndice:Get(xIndice,@nPos)
			xDefault	:=  ::xValor[nPos]
		Endif
	Endif
Return xDefault

Method AddValor(xValor,xIndice) Class YExcelTag
	Local nPos
	If ValType(xIndice)=="C"
		If ::oIndice:Get(xIndice,@nPos)
			::xValor[nPos]	:= xValor
		Else
			AADD(::xValor,xValor)
			::oIndice:Set(xIndice,Len(::xValor))
		Endif
	ElseIf ValType(xIndice)=="N"
		::xValor[xIndice]	:= xValor
	Else
		AADD(::xValor,xValor)
	Endif
Return

Method AddAtributo(cAtributo,xValor) Class YExcelTag
	PARAMTYPE 0	VAR cAtributo  AS CHARACTER
	::oAtributos:Set(cAtributo,xValor)
Return

Method SetAtributo(cAtributo,xValor) Class YExcelTag
	PARAMTYPE 0	VAR cAtributo  AS CHARACTER
	If ValType(xValor)=="U"
		::oAtributos:Del(cAtributo)
	Else
		::oAtributos:Set(cAtributo,xValor)
	Endif
Return

Method GetAtributo(cAtributo,cDefault) Class YExcelTag
	Local xValor
	PARAMTYPE 0	VAR cAtributo  AS CHARACTER
	If ::oAtributos:Get(cAtributo,@xValor)
		Return xValor
	Endif
Return cDefault

Method GetTag(nFile,lFechaTag,lSoValor) Class YExcelTag
	Local cRet	:= ""
	Local nCont
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
				Endif
			Else
				GravaFile(nFile,'>')
				GravaFile(nFile,VarTipo(::xValor,nFile))
				If lFechaTag
					GravaFile(nFile,'</'+::cNome+'>')
				Endif
			Endif
		Endif
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
				Endif
			Else
				cRet	+= '>'
				cRet	+= VarTipo(::xValor)
				If lFechaTag
					cRet	+= '</'+::cNome+'>'
				Endif
			Endif
		Endif
	Endif
Return cRet

Method LoadTagXml(oXml,cCaminho) Class YExcelTag
	Local aChildren
	Local nCont
	Local nPos
	aChildren	:= oXml:XPathGetAttArray(cCaminho)
	For nCont:=1 to Len(aChildren)
		::SetAtributo(aChildren[nCont][1],aChildren[nCont][2])
	Next
	aChildren	:= oXml:XPathGetChildArray(cCaminho)
	For nCont:=1 to Len(aChildren)
		::AddValor(yExcelTag():New(aChildren[nCont][1]),{})
		nPos	:= Len(::xValor)
		::xValor[nPos]:LoadTagXml(oXml,aChildren[nCont][2])
	Next
	If Empty(aChildren)
		::SetValor(oXml:XPathGetNodeValue(cCaminho))
	Endif
Return self

Method PutTxml(oXml,cCaminho) Class YExcelTag
	Local aListAtt
	Local nCont
	oXml:XPathAddNode(cCaminho,::cNome,"")
	::oAtributos:List(@aListAtt)
	For nCont:=1 to Len(aListAtt)
		oXml:XPathAddAtt(cCaminho+"/xmlns:"+::cNome+"[last()]",aListAtt[nCont][1],Transform(aListAtt[nCont][2],""))
	Next
	If ValType(::xValor)!="U"
		VarTipo2(::xValor,oXml,cCaminho,::cNome)
	Endif
Return

Static Function VarTipo2(xValor,oXml,cCaminho,cNome)
	Local nCont,aList
	If ValType(xValor)=="A"
		For nCont:=1 to Len(xValor)
			VarTipo2(xValor[nCont],oXml,cCaminho,cNome)
		Next
	ElseIf ValType(xValor)=="O"
		If GetClassName(xValor)=="THASHMAP"
			xValor:List(@aList)
			aSort(aList,,,{|x,y| x[1]<y[1] })
			For nCont:=1 to Len(aList)
				VarTipo2(aList[nCont][2],oXml,cCaminho,cNome)
			Next
		Else
			xValor:PutTxml(oXml,cCaminho+"/xmlns:"+cNome+"[last()]")
		Endif
	Else
		oXml:XPathSetNode(cCaminho+"/xmlns:"+cNome+"[last()]",cNome,Transform(xValor,""))
	Endif
Return

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
		Endif
	Else
		cRet	+= Transform(xValor,"")
	Endif
Return cRet

Static Function Var2Chr(xValor)
	Local nCont,aList
	Local cRet	:= ""
	Local cTipo	:= ValType(xValor)
	If cTipo=="A"
		cRet	+= "{"
		For nCont:=1 to Len(xValor)
			If nCont>1
				cRet	+= ","
			Endif
			cRet	+= Var2Chr(xValor[nCont])
		Next
		cRet	+= "}"
	ElseIf cTipo=="O"
		If GetClassName(xValor)=="THASHMAP"
			xValor:List(@aList)
			aSort(aList,,,{|x,y| x[1]<y[1] })
			cRet	+= "["
			For nCont:=1 to Len(aList)
				If nCont>1
					cRet	+= ","
				Endif
				cRet	+= '"'+aList[nCont][1]+'":'+Var2Chr(aList[nCont][2])
			Next
			cRet	+= "]"
		ElseIf GetClassName(xValor)=="YEXCELTAG"
			cRet	+= xValor:GetTag()
		Endif
	ElseIf cTipo=="J"
		cRet	+= xValor:ToJson()
	ElseIf cTipo=="C"
		cRet	+= '"'+Transform(xValor,"")+'"'
	ElseIf cTipo=="L"
		cRet	+= If(xValor,".T.",".F.")
	ElseIf cTipo=="D"
		cRet	+= "STOD("+DTOS(xValor)+")"
	ElseIf cTipo=="U"
		cRet	+= "nil"
	Else
		cRet	+= cValToChar(xValor)
	Endif
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
Class YExcel_Table from yExcelTag
	Data oyExcel
	Data lAutoFilter
	Data aRef
	Data nPrimLinha
	Data nPrimColuna
	Data oColunas
	Data aColunas
	Data nLinha
	Data oTableColumns
	Data otableStyleInfo
	Data cNomeTabela
	Data nIdRelat
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

Method new(oyExcel,nLinha,nColuna,cNome) Class YExcel_Table
	_Super:New("table",{},,oyExcel)
	::oyExcel		:= oyExcel
	::aRef			:= {{nLinha,nColuna},{0,0}}
	::nPrimLinha	:= nLinha
	::nPrimColuna	:= nColuna
	::oColunas		:= tHashMap():New()
	::aColunas		:= {}
	::nLinha		:= 0
	::cNomeTabela	:= cNome
	::AddLine()
Return self

/*/{Protheus.doc} AddFilter
Adiciona filtro a tabela
@author Saulo Gomes Martins
@since 08/05/2017

@type method
/*/
Method AddFilter() Class YExcel_Table
	::lAutoFilter:= .T.
Return

/*/{Protheus.doc} Cell
Preenche informação da célula
@author Saulo Gomes Martins
@since 08/05/2017
@version undefined
@param cColuna, characters, Nome da coluna
@param xValor, variadic, conteudo da célula
@param cFormula, characters, (Opcional) Formula
@param nStyle, numeric, (Opcional) posição da formatação
@type method
/*/
METHOD Cell(cColuna,xValor,cFormula,nStyle) Class YExcel_Table
	Local aColuna,nColuna
	If ValType(cColuna)=="C"
		If !::oColunas:Get(cColuna,@aColuna)
			UserException("YExcel - Coluna informada não encontrado. Utilize o metodo AddColumn para adicionar a coluna:"+cValToChar(cColuna))
		Endif
		nColuna	:= aColuna[2]
		If Empty(nStyle)
			nStyle	:= aColuna[5]
		Endif
	Else
		nColuna	:= cColuna
	Endif
	::oyExcel:Cell(::nLinha,nColuna,xValor,cFormula,nStyle)
Return

/*/{Protheus.doc} AddLine
Adiciona uma nova linha
@author Saulo Gomes Martins
@since 08/05/2017
@param nQtd, numeric, Quantidade de linhas para avançar
@type method
/*/
Method AddLine(nQtd) Class YExcel_Table
	Default nQtd	:= 1
	::nLinha		+= nQtd
	::aRef[2][1]	:= ::nLinha
	If ::nLinha!=::nPrimLinha
		::oyExcel:AddRow(nQtd,::nLinha,::aRef[1][2],::aRef[2][2])
	Endif
return ::nLinha

/*/{Protheus.doc} AddColumn
Adiciona uma nova coluna a tabela
@author Saulo Gomes Martins
@since 08/05/2017
@version undefined
@param cNome, characters, descricao
@param nStyle, numeric, descricao
@type method
/*/
METHOD AddColumn(cNome,nStyle) Class YExcel_Table
	Local otableColumn
	Local nCont
//	Local nPosCol		:= aScan(self:GetValor(),{|x| x:GetNome()=="tableColumns"})
	::aRef[2][2]	+= 1
	::nLinha		:= ::nPrimLinha
	::oyExcel:AddCol(1,::aRef[2][2],::nPrimLinha,::aRef[2][1])

	nCont	:= Len(self:oTableColumns:GetValor())+1
	otableColumn	:= yExcelTag():New("tableColumn",{},,::oyExcel)
	otableColumn:SetAtributo("id",nCont)
	otableColumn:SetAtributo("name",cNome)
	self:oTableColumns:SetAtributo("count",nCont)
	self:oTableColumns:AddValor(otableColumn)
	::oColunas:Set(cNome,{::aRef[1][1],::aRef[1][2]+Len(::aColunas),otableColumn,nil,nStyle})
	AADD(::aColunas,cNome)
	::Cell(cNome,cNome)
Return


/*/{Protheus.doc} AddTotal
Adiciona um totalizador na coluna
@author Saulo Gomes Martins
@since 08/05/2017
@version undefined
@param cColuna, characters, Nome da coluna
@param xValor, variadic, Valor
@param cFunction, characters, Formula
@type method
@see https://support.office.com/pt-br/article/SUBTOTAL-Fun%C3%A7%C3%A3o-SUBTOTAL-7b027003-f060-4ade-9040-e478765b9939?NS=EXCEL&Version=16&SysLcid=1046&UiLcid=1046&AppVer=ZXL160&HelpId=xlmain11.chm60392&ui=pt-BR&rs=pt-BR&ad=BR
@obs PAG 2392
function-number				function-number					Function
(includes hidden values)	(excludes hidden values)
1							101							AVERAGE	MÉDIA
2							102							COUNT	CONTAR NUMERO
3							103							COUNTA	CONT.VALORES
4							104							MAX		MAX
5							105							MIN		MIN
6							106							PRODUCT	MULT
7							107							STDEV	DESVPAD
8							108							STDEVP	DESVPADP
9							109							SUM		SOMA
10							110							VAR		VAR
11							111							VARP	VARP
/*/
Method AddTotal(cColuna,xValor,cFunction,nStyle) Class YExcel_Table
	Local aColuna,otableColumn
	If ::oColunas:Get(cColuna,@aColuna)
		otableColumn	:= aColuna[3]
		aColuna[4]		:= xValor
		If ValType(nStyle)<>"U"
			aColuna[5]		:= nStyle
		Endif
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
				otableColumn:AddValor(yExcelTag():New("totalsRowFormula",cFunction,,::oyExcel),"totalsRowFormula")
			Endif
		Endif
	Endif
Return

/*/{Protheus.doc} AddTotais
Inclui a linha de totalizador
@author Saulo Gomes Martins
@since 08/05/2017

@type method
/*/
Method AddTotais() Class YExcel_Table
	Local nCont,xValor,cFormula
	Local aColuna,cRef
	
	::oyExcel:AddRow(1,::nLinha+1,::aRef[1][2],::aRef[2][2])

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
			Endif
		Endif
		If ValType(xValor)=="U" .and. ValType(cFormula)=="U"
			Loop
		Else
			::oyExcel:Cell(::aRef[2][1]+1,aColuna[2],xValor,cFormula,aColuna[5])
		Endif
	Next

Return

/*/{Protheus.doc} Finish
Finaliza a tabela criada
@author Saulo Gomes Martins
@since 03/05/2017
@version undefined
@type method

/*/
METHOD Finish() Class YExcel_Table
	Local nPosCol
	Local cRef
	cRef		:= ::oyExcel:Ref(::aRef[1][1],::aRef[1][2])+":"+::oyExcel:Ref(::aRef[2][1],::aRef[2][2])
	nPosCol		:= aScan(self:GetValor(),{|x| x:GetNome()=="autoFilter"})
	If ::lAutoFilter
		self:GetValor(nPosCol):SetAtributo("ref",cRef)
	Else
		aDel(self:GetValor(),nPosCol)
		aSize(self:GetValor(),Len(self:GetValor())-1)
	Endif
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
@type method
@OBS Annex G. (normative) Predefined SpreadsheetML Style Definitions
PAG 4426
	TableStyleMedium2	- AZUL|LINHA1-AZUL_CLARO|LINHA2-BRANCO|SEM BORDA
	TableStyleMedium9	- AZUL|LINHA1-AZUL_ESCURO|LINHA2-AZUL_CLARO|SEM BORDA
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
METHOD AddStyle(cNome,lLinhaTiras,lColTiras,lFormPrimCol,lFormUltCol) Class YExcel_Table
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
	Endif
	If lColTiras	//Colunas em tiras
		::otableStyleInfo:SetAtributo("showColumnStripes",1)
	Else
		::otableStyleInfo:SetAtributo("showColumnStripes",0)
	Endif
	If lFormPrimCol	//Exibir formato especial na primeira coluna da tabela
		::otableStyleInfo:SetAtributo("showFirstColumn",1)
	Else
		::otableStyleInfo:SetAtributo("showFirstColumn",0)
	Endif
	If lFormUltCol	//Exibir formato especial na ultima coluna da tabela
		::otableStyleInfo:SetAtributo("showLastColumn",1)
	Else
		::otableStyleInfo:SetAtributo("showLastColumn",0)
	Endif
Return

/*/{Protheus.doc} GetDateTime
Retorna objeto para manipulação de DateTime
@author Saulo Gomes Martins
@since 09/12/2019
@version 1.0
@param dData, date, Data para formatação
@param cTime, characters, Hora para formatação
@param nData, numeric, DataTime em formato numerico
@type method
/*/
METHOD GetDateTime(dData,cTime,nData) Class YExcel
	Local oDateTime	:= yExcel_DateTime():New(dData,cTime,nData)
	AADD(::aCleanObj,oDateTime)
Return oDateTime

/*/{Protheus.doc} yExcel_DateTime
Classe yExcel_DateTime para manipulação de DateTime
@author Saulo Gomes Martins
@since 10/12/2019
@version 1.0
@obs pag 4780
@type class
/*/
Class YExcel_DateTime
	Data dData
	Data cTime
	Data cNumero
	Data nNumero			//Numero com limite de 8 decimais
	Data nDecimal			//Numero com limite decimal acima de 8 posições
	Data cClassName			//Nome da Classe
	Data cName				//Nome da Classe
	Method New() CONSTRUCTOR
	METHOD ClassName()
	METHOD NumToDateTime()
	METHOD GetStrNumber()
	METHOD GetDate()
	METHOD GetTime()
	METHOD StrNumber()
EndClass

/*/{Protheus.doc} yExcel_DateTime:New
Construtor da classe yExcel_DateTime
@author Saulo Gomes Martins
@since 10/12/2019
@version 1.0
@return object, self objeto
@param dData, date, Data para iniciar o objeto
@param cTime, characters, Hora para iniciar o objeto
@param nData, numeric, (Opcional) Data e hora para iniciar o objeto
@type method
@obs enviar dData e cTime ou somente nData
/*/
Method New(dData,cTime,nData,nDec8,cDataUTC) Class YExcel_DateTime
	::dData		:= dData
	::cTime		:= cTime
	::nNumero	:= 0
	::nDecimal	:= 0
	::cClassName	:= "YEXCEL_DATETIME"
	::cName			:= "YEXCEL_DATETIME"
	If ValType(cDataUTC)=="C"
		::dData	:= STOD(Replace(SubStr(cDataUTC,1,10),"-",""))
		::cTime	:= SubStr(cDataUTC,12,8)
		::StrNumber()
	ElseIf ValType(::dData)=="D" .AND. ValType(cTime)=="C"
		::StrNumber()
	ElseIf ValType(nData)=="N" .OR. ValType(nData)=="C"
		::NumToDateTime(nData,nDec8)
	Endif
Return Self

/*/{Protheus.doc} yExcel_DateTime:NumToDateTime
Converte numero do excel em data e hora
@author Saulo Gomes Martins
@since 10/12/2019
@version 1.0
@param nData, numeric, numero da hora, aceita também string
@type method
/*/
Method NumToDateTime(nData,nDec8) Class YExcel_DateTime
	Local nInt
	Local nDec
	Local nHora
	Local nMinuto
	Local nSegundo
	Local f60			:= DEC_CREATE( "60" , 20, 19 )	
	Local f86400		:= DEC_CREATE( "86400" , 20, 19 )
	Local fNumTime
	Local fSeg,fMinuto,fHora
	If ValType(nData)=="N"
		nInt	:= Int(nData)
		nDec	:= nData-nInt
		fNumTime	:= DEC_CREATE(cValToChar(nDec), 20, 19 )
		::cNumero	:= cValToChar(nData)
		::nNumero	:= nInt+nDec
		::nDecimal	:= nDec8
	Else
		nPosPonto	:= At(".",nData)
		If nPosPonto==0
			nPosPonto	:= At(",",nData)
		Endif
		If nPosPonto==0
			nInt		:= Val(nData)
			nDec		:= 0
			::nNumero	:= nInt
		Else
			nInt		:= Val(SubStr(nData,1,nPosPonto-1))
			fNumTime	:= DEC_CREATE("0."+SubStr(nData,nPosPonto+1), 20, 19 )
			// nDec		:= Val("0."+SubStr(nData,nPosPonto+1))
			::nNumero	:= nInt+(Val("0."+SubStr(nData,nPosPonto+1,8)))
			::nDecimal	:= Val("0."+SubStr(nData,nPosPonto+8+1))*(10^8)
		Endif
		::cNumero	:= nData
	Endif
	If ValType(self:nDecimal)=="N"
		fNumTime	:= DEC_ADD(fNumTime,DEC_DIV(DEC_CREATE(cValToChar(self:nDecimal),20,19),DEC_CREATE("10000000000000000",20,19)))
	EndIf

	::dData	:= STOD("19000101")-2+nInt
	
	fSeg		:= DEC_MUL(fNumTime,f86400)
	fMinuto		:= DEC_DIV(fSeg,f60)
	fHora		:= DEC_DIV(fMinuto,f60)
	nHora		:= Int(DEC_TO_DBL(fHora))
	nMinuto		:= Int(DEC_TO_DBL(DEC_SUB(fMinuto,DEC_CREATE(cValToChar(nHora*60), 20, 19 ))))
	nSegundo	:= Round(DEC_TO_DBL(DEC_SUB(fSeg,DEC_CREATE(cValToChar((nMinuto*60)+(nHora*60*60)), 20, 19 ))),0)
	
	::cTime	:= ""
	// nHora	:= Int(nDec*86400/60/60)
	// nMinuto	:= Int(((nDec*86400/60/60)-nHora)*60)
	// nSegundo:= Round(((nDec*86400/60/60)-nHora-(nMinuto/60))*60*60,0)
	::cTime	+= StrZero(nHora,2)				//Hora
	::cTime	+= ":"+StrZero(nMinuto,2)		//Minuto
	::cTime	+= ":"+StrZero(nSegundo,2)		//Segundos
	//IntToHora(nDec*86400/60/60)
Return Self

/*/{Protheus.doc} yExcel_DateTime:GetStrNumber
Retorna o numero em formato string
@author Saulo Gomes Martins
@since 10/12/2019
@version 1.0

@type method
/*/
Method GetStrNumber() Class YExcel_DateTime
Return ::cNumero

/*/{Protheus.doc} yExcel_DateTime:GetDate
Retorna a data
@author Saulo Gomes Martins
@since 10/12/2019
@version 1.0

@type method
/*/
Method GetDate() Class YExcel_DateTime
Return ::dData

/*/{Protheus.doc} yExcel_DateTime:GetTime
Retorna a Hora no formato HH:MM:SS
@author Saulo Gomes Martins
@since 10/12/2019
@version 1.0

@type method
/*/
Method GetTime() Class YExcel_DateTime
Return ::cTime

/*/{Protheus.doc} yExcel_DateTime:StrNumber
Converte data e hora em string com numero representando data e hora do excel
@author Saulo Gomes Martins
@since 10/12/2019
@version 1.0

@type method
/*/
Method StrNumber() Class YExcel_DateTime
	Local aHora	:= SeparaHora(::cTime)
	Local nSeg	:= 0
	Local cNum
	Local f1SegDec	:= DEC_CREATE( "0.00001157407407407407" , 20, 19 )	// 1/86400
	Local fSeg,fMult,fMult2,fSub,f10,fMultFim
	nSeg		+= aHora[1]*3600
	nSeg		+= aHora[2]*60
	nSeg		+= aHora[3]
	nSeg		+= aHora[4]/1000
	fSeg		:= DEC_CREATE(cValToChar(nSeg),20, 19)
	fMult		:= DEC_MUL(fSeg,f1SegDec)
	fMult2		:= DEC_MUL(fSeg,f1SegDec)
	fMult2		:= DEC_RESCALE(fMult2, 8, 2)
	fSub		:= DEC_SUB(fMult,fMult2)
	f10			:= DEC_CREATE("10000000000000000",20, 19)
	fMultFim	:= DEC_MUL(f10,fSub)

	::nDecimal	:= Int(Val(cValToChar(DEC_RESCALE( fMultFim , 8, 2))))
	cNum		:= Replace(cValToChar(DEC_RESCALE(fMult, 8, 2)),"0.","")
	// nHora		+= (aHora[1]*100000000)				//Hora
	// nHora		+= (aHora[2]*100000000)/60			//Minuto
	// nHora		+= (aHora[3]*100000000)/60/60		//Segundo
	// nHora		+= (aHora[4]*100000000)/60/60/1000	//Milesimo
	// cNum		:= Replace(cValToChar(nHora/24),"0.","")
	// If (At(".",cNum)-1)>0
	// 	cNum:= SubStr(cNum,1,At(".",cNum)-1)
	// Else
	// 	cNum:= SubStr(cNum,1,Len(cNum))
	// Endif
	::cNumero	:= cValToChar(::dData-STOD("19000101")+2)+"."+cNum
	::nNumero	:= ::dData-STOD("19000101")+2+(&("0."+SubStr(cNum,1,8)))
	// If Len(cNum)>8
	// 	::nDecimal	:= &("0."+SubStr(cNum,9)+"*(10^8)")
	// Endif
Return ::cNumero

METHOD ClassName() Class YExcel_DateTime
Return "YEXCEL_DATETIME"

/*/{Protheus.doc} SeparaHora
Retorna Hora,Minuto,Segundo,Milesimo.
@author Saulo Gomes Martins
@since 09/12/2019
@version 1.0
@return Array, aHora 1-Hora|2-Munuto|3-Segundo|4-Milésimo de segundo
@param cHora, characters, Hora no Formato HH:MM:SS.MMMM
@type function
/*/
Static Function SeparaHora(cHora)
	Local nHoras	:= 0
	Local nMinutos	:= 0
	Local nSegundos	:= 0
	Local nMilesimo	:= 0	//Milésimo de segundo
	Local nPosSepara

	nPosSepara	:= At(":",cHora)
	If nPosSepara==0
		nHoras		:= Val(cHora)
	Else
		nHoras		:= Val(SubStr(cHora,1,nPosSepara-1))
		cHora		:= SubStr(cHora,nPosSepara+1)
		nPosSepara	:= At(":",cHora)
		If nPosSepara==0
			nMinutos	:= Val(cHora)
		Else
			nMinutos	:= Val(SubStr(cHora,1,nPosSepara-1))		///60
			cHora		:= SubStr(cHora,nPosSepara+1)
			nPosSepara	:= At(".",cHora)
			If nPosSepara==0
				nSegundos	:= Val(cHora)
			Else
				nSegundos	:= Val(SubStr(cHora,1,nPosSepara-1))	///60/60
				cHora		:= SubStr(cHora,nPosSepara+1)
				nMilesimo	:= Val(cHora)							///60/60/1000
			Endif
		Endif
	Endif
Return {nHoras,nMinutos,nSegundos,nMilesimo}

/*/{Protheus.doc} new_content_types
Criação do arquivo \[content_types].xml
@author Saulo Gomes Martins
@since 10/12/2019
@version 1.0
@param cXml, characters, xml para criação
@type method
/*/
Method new_content_types(cFile) Class YExcel
	Local nCont
	Local aNs
	Local cXml			:= ""
	::ocontent_types	:= TXMLManager():New()
	If Empty(cFile)	//Cria modelo em branco
		cXml	+= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
		cXml	+= '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
		cXml	+= '	<Default Extension="jpg" ContentType="image/jpeg"/>'
		cXml	+= '	<Default Extension="png" ContentType="image/png"/>'
		cXml	+= '	<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
		cXml	+= '	<Default Extension="xml" ContentType="application/xml"/>'
		cXml	+= '	<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>'
		cXml	+= '	<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
		cXml	+= '	<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
		cXml	+= '	<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
		cXml	+= '	<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
		cXml	+= '	<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
		cXml	+= '	<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
		cXml	+= '</Types>'
		::ocontent_types:Parse(cXml)
	Else
		::ocontent_types:ParseFile(cFile)
	Endif
	aNs	:= ::ocontent_types:XPathGetRootNsList()
	For nCont:=1 to Len(aNs)
		::ocontent_types:XPathRegisterNs( aNs[nCont][1], aNs[nCont][2] )
	Next
	AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\[Content_Types].xml")
Return

/*/{Protheus.doc} new_rels
Cria arquivo de relacionamento Relationship(rels)
@author Saulo Gomes Martins
@since 10/12/2019
@version 1.0
@return numeric, nPos Posição no array
@param cXml, characters, xml para criação
@param cCaminho, characters, caminho do arquivo
@type method
/*/
Method new_rels(cFile,cCaminho) Class YExcel
	Local nCont
	Local aNs
	Local oXML
	Local cXml			:= ""
	Default cCaminho	:= cFile
	oXML	:= TXMLManager():New()
	If Empty(cFile)	//Cria modelo em branco
		cXml	+= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
		cXml	+= '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
		cXml	+= '</Relationships>'
		oXML:Parse(cXml)
	Else
		oXml:ParseFile(cFile)
	Endif
	aNs	:= oXML:XPathGetRootNsList()
	For nCont:=1 to Len(aNs)
		oXML:XPathRegisterNs( aNs[nCont][1], aNs[nCont][2] )
	Next
	AADD(::aRels,{oXML,cCaminho,oXml:XPathChildCount("/xmlns:Relationships")})
Return Len(::aRels)

/*/{Protheus.doc} add_rels
Adiciona node no arquivo Relationship(rels)
@author Saulo Gomes Martins
@since 10/12/2019
@version 1.0
@return characters, cId rId criado
@param cCaminho, characters, caminho do arquivo de rel para gravar
@param cType, characters, atributo Type
@param cTarget, characters, atributo Target
@type method
/*/
Method add_rels(cCaminho,cType,cTarget) Class YExcel
	Local nPos
	Local cId
	If ValType(cCaminho)=="N"
		nPos	:= cCaminho
	ElseIf ValType(cCaminho)=="C"
		If SubStr(cCaminho,1,1)!="\"
			cCaminho	:= "\"+cCaminho
		Endif
		nPos	:= aScan(::aRels,{|x| x[2]==cCaminho })
	Endif
	If nPos==0
		nPos	:= ::new_rels(,cCaminho)
	Endif
	::aRels[nPos][3]++
	cId	:= "rId"+cValToChar(::aRels[nPos][3])
	::aRels[nPos][1]:XPathAddNode( "/xmlns:Relationships", "Relationship", "" )
	::aRels[nPos][1]:XPathAddAtt( "/xmlns:Relationships/xmlns:Relationship[last()]", "Type"		, cType )
	::aRels[nPos][1]:XPathAddAtt( "/xmlns:Relationships/xmlns:Relationship[last()]", "Target"	, cTarget )
	::aRels[nPos][1]:XPathAddAtt( "/xmlns:Relationships/xmlns:Relationship[last()]", "Id"		, cId )
Return cId
/*/{Protheus.doc} YExcel::Get_rels
Retorna atributos do Relationships relacionado ao ID
@type method
@version 1.0
@author Saulo Gomes Martins
@since 17/03/2021
@param cCaminho, character, Caminho do rel
@param cId, character, Id relacionado
@param cAtrRet, character, nome do atributo
@return character, conteudo do atributo
/*/
Method Get_rels(cCaminho,cId,cAtrRet) Class YExcel
	Local oTmp := TXmlManager():New()
	Local aNs
	Local xRet
	Local nCont
	oTmp:ParseFile(cCaminho)
	aNs	:= oTmp:XPathGetRootNsList()
	For nCont:=1 to Len(aNs)
		oTmp:XPathRegisterNs( aNs[nCont][1], aNs[nCont][2] )
	Next
	If oTmp:XPathHasNode("/xmlns:Relationships/xmlns:Relationship[@Id='"+cId+"']")
		xRet	:= oTmp:XPathGetAtt("/xmlns:Relationships/xmlns:Relationship[@Id='"+cId+"']",cAtrRet)
	Endif
	FreeObj(oTmp)
Return xRet
/*/{Protheus.doc} YExcel::FindRels
Busca o relationships com filtros
@type method
@version 1.0
@author Saulo Gomes Martins
@since 18/03/2021
@param cCaminho, character, Caminho do rels
@param cAtrRet, character, Atributo para retorno
@param cId, character, Id para filtro
@param cType, character, Type para filtro
@param cTarget, character, Target para filtro
@return character, Atributo
/*/
Method FindRels(cCaminho,cAtrRet,cId,cType,cTarget) Class YExcel
	Local nPos
	Local cRet
	Local cFiltro	:= ""
	If ValType(cCaminho)=="N"
		nPos	:= cCaminho
	ElseIf ValType(cCaminho)=="C"
		If SubStr(cCaminho,1,1)!="\"
			cCaminho	:= "\"+cCaminho
		Endif
		nPos	:= aScan(::aRels,{|x| x[2]==cCaminho })
	Endif
	If nPos>0
		If !Empty(cId)
			cFiltro	+="@Id='"+cId+"'"
		EndIf
		If !Empty(cType)
			If !Empty(cFiltro)
				cFiltro	+= " and "
			Endif
			cFiltro	+="@Type='"+cType+"'"
		EndIf
		If !Empty(cTarget)
			If !Empty(cFiltro)
				cFiltro	+= " and "
			Endif
			cFiltro	+="@Target='"+cTarget+"'"
		EndIf
		If !Empty(cFiltro)
			cFiltro	:= "["+cFiltro+"]"
		Endif
		If ::aRels[nPos][1]:XPathHasNode("/xmlns:Relationships/xmlns:Relationship"+cFiltro)
			cRet	:= ::aRels[nPos][1]:XPathGetAtt("/xmlns:Relationships/xmlns:Relationship"+cFiltro,cAtrRet)
		EndIf
	Endif
Return cRet

/*/{Protheus.doc} new_app
Cria arquivo \docprops\app.xml
@author Saulo Gomes Martins
@since 10/12/2019
@version 1.0
@param cXml, characters, xml para leitura
@type method
/*/
Method new_app(cFile) Class YExcel
	Local nCont
	Local aNs
	Local cXml			:= ""
	::oapp	:= TXMLManager():New()
	If Empty(cFile)	//Cria modelo em branco
		cXml	+= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
		cXml	+= '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
		cXml	+= '	<Application>Microsoft Excel</Application>'
		cXml	+= '	<DocSecurity>0</DocSecurity>'
		cXml	+= '	<ScaleCrop>false</ScaleCrop>'
		cXml	+= '	<HeadingPairs>'
		cXml	+= '		<vt:vector size="2" baseType="variant">'
		cXml	+= '			<vt:variant>'
		cXml	+= '				<vt:lpstr>Planilhas</vt:lpstr>'
		cXml	+= '			</vt:variant>'
		cXml	+= '			<vt:variant>'
		cXml	+= '				<vt:i4>1</vt:i4>'
		cXml	+= '			</vt:variant>'
		cXml	+= '		</vt:vector>'
		cXml	+= '	</HeadingPairs>'
		cXml	+= '	<TitlesOfParts>'
		cXml	+= '		<vt:vector size="1" baseType="lpstr">'
		cXml	+= '			<vt:lpstr>Plan1</vt:lpstr>'
		cXml	+= '		</vt:vector>'
		cXml	+= '	</TitlesOfParts>'
		cXml	+= '	<Company>Microsoft</Company>'
		cXml	+= '	<LinksUpToDate>false</LinksUpToDate>'
		cXml	+= '	<SharedDoc>false</SharedDoc>'
		cXml	+= '	<HyperlinksChanged>false</HyperlinksChanged>'
		cXml	+= '	<AppVersion>16.0300</AppVersion>'
		cXml	+= '</Properties>'
		::oapp:Parse(cXml)
	Else
		::oapp:ParseFile(cFile)
	Endif
	aNs	:= ::oapp:XPathGetRootNsList()
	For nCont:=1 to Len(aNs)
		::oapp:XPathRegisterNs( aNs[nCont][1], aNs[nCont][2] )
	Next
	AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\docprops\app.xml")
Return

/*/{Protheus.doc} new_core
Cria arquivo \docprops\core.xml
@author Saulo Gomes Martins
@since 10/12/2019
@version 1.0
@param cXml, characters, xml para leitura
@type method
/*/
Method new_core(cFile) Class YExcel
	Local nCont
	Local aNs
	Local aRet
	Local cXml			:= ""
	::ocore	:= TXMLManager():New()
	If Empty(cXml)	//Cria modelo em branco
		cXml	+= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
		cXml	+= '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
		cXml	+= '	<dc:creator>Totvs - Protheus</dc:creator>'
		cXml	+= '	<cp:lastModifiedBy>Totvs - Protheus</cp:lastModifiedBy>'
		aRet	:= LocalToUTC(DTOS(Date()),Time())
		cXml	+= '	<dcterms:created xsi:type="dcterms:W3CDTF">'+SUBSTR(aRet[1],1,4)+"-"+SUBSTR(aRet[1],5,2)+"-"+SUBSTR(aRet[1],7,2)+'T'+aRet[2]+'Z</dcterms:created>'
		cXml	+= '	<dcterms:modified xsi:type="dcterms:W3CDTF">'+SUBSTR(aRet[1],1,4)+"-"+SUBSTR(aRet[1],5,2)+"-"+SUBSTR(aRet[1],7,2)+'T'+aRet[2]+'Z</dcterms:modified>'
		cXml	+= '</cp:coreProperties>'
		aRet	:= nil
		::ocore:Parse(cXml)
	Else
		::ocore:ParseFile(cFile)
	Endif
	aNs	:= ::ocore:XPathGetRootNsList()
	For nCont:=1 to Len(aNs)
		::ocore:XPathRegisterNs( aNs[nCont][1], aNs[nCont][2] )
	Next
	AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\docprops\core.xml")
Return

/*/{Protheus.doc} new_workbook
Cria arquivo \xl\workbook.xml
@author Saulo Gomes Martins
@since 10/12/2019
@version 1.0
@param cXml, characters, xml para leitura
@type method
/*/
Method new_workbook(cFile) Class YExcel
	Local nCont
	Local aNs
	Default cXml			:= ""
	::oworkbook	:= TXMLManager():New()
	If Empty(cFile)	//Cria modelo em branco
		cXml	+= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
		cXml	+= '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
		cXml	+= '	<fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="17927"/>'
		cXml	+= '	<workbookPr defaultThemeVersion="124226"/>'
		cXml	+= '	<bookViews>'
		cXml	+= '		<workbookView xWindow="240" yWindow="135" windowWidth="20115" windowHeight="8250"/>'
		cXml	+= '	</bookViews>'
		cXml	+= '	<sheets>'
		cXml	+= '	</sheets>'
		cXml	+= '	<definedNames/>'
		cXml	+= '</workbook>'
		::oworkbook:Parse(cXml)
	Else
		::oworkbook:ParseFile(cFile)
	Endif
	aNs	:= ::oworkbook:XPathGetRootNsList()
	For nCont:=1 to Len(aNs)
		::oworkbook:XPathRegisterNs( aNs[nCont][1], aNs[nCont][2] )
	Next
	AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\workbook.xml")
Return

/*/{Protheus.doc} new_draw
Cria arquivo \xl\drawings\drawingX.xml
@author Saulo Gomes Martins
@since 10/12/2019
@version 1.0
@return numeric, nPos posição do draw no array
@param cXml, characters, xml para leitura
@param cCaminho, characters, caminho para gravar
@type method
/*/
Method new_draw(cFile,cCaminho) Class YExcel
	Local nCont
	Local aNs
	Local oXML
	Default cXml			:= ""
	oXML	:= TXMLManager():New()
	If Empty(cFile)	//Cria modelo em branco
		cXml	+= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
		cXml	+= '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
		cXml	+= '</xdr:wsDr>'
		oXML:Parse(cXml)
	else
		oXML:ParseFile(cFile)
	Endif
	aNs	:= oXML:XPathGetRootNsList()
	For nCont:=1 to Len(aNs)
		oXML:XPathRegisterNs( aNs[nCont][1], aNs[nCont][2] )
	Next
	AADD(::aDraw,{oXML,cCaminho,oXml:XPathChildCount("/xdr:wsDr")})
Return Len(::aDraw)


/*/{Protheus.doc} new_comment
Cria arquivo \xl\comment.xml
@author Saulo Gomes Martins
@since 10/12/2019
@version 1.0
@param cXml, characters, xml para leitura
@type method
/*/
Method new_comment(cFile) Class YExcel
	Local nCont
	Local aNs
	Local oXml
	Default cXml			:= ""
	oXml	:= TXMLManager():New()
	If Empty(cFile)	//Cria modelo em branco
		cXml	+= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
		cXml	+= '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
		cXml	+= '<authors/>'
		cXml	+= '<commentList/>'
		cXml	+= '</comments>'
		oXml:Parse(cXml)
	Else
		oXml:ParseFile(cFile)
	Endif
	aNs	:= oXml:XPathGetRootNsList()
	For nCont:=1 to Len(aNs)
		oXml:XPathRegisterNs( aNs[nCont][1], aNs[nCont][2] )
	Next
Return oXml

/*/{Protheus.doc} ajustNS
Ajuste para criar node com namespace diferente
@author Saulo Gomes Martins
@since 10/12/2019
@version 1.0
@param oXml, object, Objeto TXMLManager
@param cText1, characters, texto com node errado
@param cText2, characters, texto com node para correção
@type function
/*/
Static Function ajustNS(oXml,cText1,cText2)
	Local aNs,nCont
	oXml:Parse(Replace(oXml:Save2String(),cText1,cText2))
	aNs	:= oXML:XPathGetRootNsList()
	For nCont:=1 to Len(aNs)
		oXML:XPathRegisterNs( aNs[nCont][1], aNs[nCont][2] )
	Next
Return

/*/{Protheus.doc} new_styles
Cria arquivo \xl\styles.xml
@author Saulo Gomes Martins
@since 10/12/2019
@version 1.0

@type method
/*/
Method new_styles(cFile) Class YExcel
	Local nCont
	Local aNs
	Default cXml			:= ""
	::oStyle	:= TXMLManager():New()
	If Empty(cFile)	//Cria modelo em branco
		cXml	+= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
		cXml	+= '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
		cXml	+= '<numFmts count="0"/>'
		cXml	+= '<fonts count="0" x14ac:knownFonts="1"/>'
		cXml	+= '<fills count="0"/>'
		cXml	+= '<borders count="0"/>'
		cXml	+= '<cellStyleXfs count="1">'
		cXml	+= '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'
		cXml	+= '</cellStyleXfs>'
		cXml	+= '<cellXfs count="0"/>
		cXml	+= '<cellStyles count="1">
		cXml	+= '<cellStyle name="Normal" xfId="0" builtinId="0"/>
		cXml	+= '</cellStyles>
		cXml	+= '<dxfs count="0"/>'
		cXml	+= '</styleSheet>'
		::oStyle:Parse(cXml)
	Else
		::oStyle:ParseFile(cFile)
	Endif
	aNs	:= ::oStyle:XPathGetRootNsList()
	For nCont:=1 to Len(aNs)
		::oStyle:XPathRegisterNs( aNs[nCont][1], aNs[nCont][2] )
	Next
	AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\styles.xml")
Return
/*/{Protheus.doc} SheetTmp
Cria um sheet temporario vazio
@type function
@version 1.0
@author Saulo Gomes Martins
@since 17/03/2021
@return object, TXmlManager tmp
/*/
Static Function SheetTmp()
	Local oXml
	Local cXml	:= ""
	Local nCont
	Local aNs

	cXml	+= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
	cXml	+= '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
	cXml	+= '</worksheet>'
	oXml:= TXmlManager():New()
	oXml:Parse(cXml)
	aNs	:= oXML:XPathGetRootNsList()
	For nCont:=1 to Len(aNs)
		oXML:XPathRegisterNs( aNs[nCont][1], aNs[nCont][2] )
	Next
Return oXml
/*/{Protheus.doc} xls_sheet
Cria arquivo \xl\worksheets\sheetX.xml
@author Saulo Gomes Martins
@since 10/12/2019
@version 1.0

@type method
/*/
Method xls_sheet(cFile,cCaminho) Class YExcel
	Local nCont
	Local aNs
	Local oXML
	Local cXml			:= ""
	oXML	:= TXMLManager():New()
	If Empty(cFile)	//Cria modelo em branco
		cXml	+= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
		cXml	+= '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">'
		cXml	+= '<sheetPr/>'
		cXml	+= '<dimension ref="A1"/>'
		cXml	+= '<sheetViews>'
			cXml	+= '<sheetView tabSelected="0" workbookViewId="0">'
			cXml	+= '<selection sqref="A1"/>'
			cXml	+= '</sheetView>'
		cXml	+= '</sheetViews>'
		cXml	+= '<sheetFormatPr defaultRowHeight="15"/>'
		cXml	+= '<cols/>'
		cXml	+= '<sheetData/>'
		cXml	+= '<autoFilter/>'
		cXml	+= '<mergeCells/>'
		cXml	+= '</worksheet>'
		oXML:Parse(cXml)
	Else
		oXML:ParseFile(cFile)
	Endif

	aNs	:= oXML:XPathGetRootNsList()
	For nCont:=1 to Len(aNs)
		oXML:XPathRegisterNs( aNs[nCont][1], aNs[nCont][2] )
	Next
	AADD(::asheet,{oXML,cCaminho,/*oxml Comment*/,/*arquivo Comment*/,/*vmlDrawing*/,/*caminho vmlDraw*/})
Return Len(::asheet)


Method new_vmlDrawing(cFile) Class YExcel
	Local nCont
	Local aNs
	Local oXML
	Local cXml			:= ""
	oXML	:= TXMLManager():New()
	If Empty(cFile)	//Cria modelo em branco
		cXml	+= '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">'
		// cXml	+= '	<o:shapelayout v:ext="edit">'
		// cXml	+= '		<o:idmap v:ext="edit" data="1"/>'
		// cXml	+= '	</o:shapelayout>'
		// cXml	+= '	<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">'
		// cXml	+= '		<v:stroke joinstyle="miter"/>'
		// cXml	+= '		<v:path gradientshapeok="t" o:connecttype="rect"/>'
		// cXml	+= '	</v:shapetype>'
		// cXml	+= '	<v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;'
		// cXml	+= '  margin-left:59.25pt;margin-top:1.5pt;width:96pt;height:64.5pt;z-index:1;'
		// cXml	+= '  visibility:hidden" fillcolor="#ffffc0" o:insetmode="auto">'
		// cXml	+= '		<v:fill color2="#ffffc0"/>'
		// cXml	+= '		<v:shadow on="t" color="black" obscured="t"/>'
		// cXml	+= '		<v:path o:connecttype="none"/>'
		// cXml	+= '		<v:textbox style="mso-direction-alt:auto">'
		// cXml	+= '			<div style="text-align:left"/>'
		// cXml	+= '		</v:textbox>'
		// cXml	+= '		<x:ClientData ObjectType="Note">'
		// cXml	+= '			<x:MoveWithCells/>'
		// cXml	+= '			<x:SizeWithCells/>'
		// cXml	+= '			<x:AutoFill>False</x:AutoFill>'
		// // cXml	+= '			<x:Anchor>'
		// // cXml	+= '				1, 15, 0, 2, 3, 15, 4, 8</x:Anchor>'
		// cXml	+= '			<x:Row>3</x:Row>'
		// cXml	+= '			<x:Column>0</x:Column>'
		// cXml	+= '		</x:ClientData>'
		// cXml	+= '	</v:shape>'
		cXml	+= '</xml>'
 		oXML:Parse(cXml)
	Else
		oXML:ParseFile(cFile)
	Endif
	aNs	:= oXML:XPathGetRootNsList()
	For nCont:=1 to Len(aNs)
		oXML:XPathRegisterNs( aNs[nCont][1], aNs[nCont][2] )
	Next
Return oXML

Static Function AjustXML(oXml2)
	Local oXML	:= TXMLManager():New()
	Local aNs
	Local nCont
	oXML:Parse(oXml2:Save2String())
	aNs	:= oXML:XPathGetRootNsList()
	For nCont:=1 to Len(aNs)
		oXML:XPathRegisterNs( aNs[nCont][1], aNs[nCont][2] )
	Next
	FreeObj(oXml2)
Return oXML
/*/{Protheus.doc} aScanOrdem
Organiza de acordo com ordem enviada no array
@type function
@version 1.0
@author Saulo Gomes Martins
@since 17/03/2021
@param aArray, array, Array com ordens {"primeiro","segundo"}
@param cValor, character, valor para analisar ordem
@return numeric, Posição da ordem
/*/
Static Function aScanOrdem(aArray,cValor)
	Local nPos := aScan(aArray,cValor)
	If nPos==0
		nPos	:= 99999
	EndIf
Return nPos
/*/{Protheus.doc} Xml2Xml
Analisa dois objetos de XML para replicar os dados
@type function
@version 1.0
@author Saulo Gomes Martins
@since 17/03/2021
@param oXml, object, xml para receber os dados
@param oXml2, object, xml para enviar os dados
@param cPath, character, Caminho do nó que será enviado
@param cFiltro, character, Envia apenas tags do filtro
@param cNaoFazer, character, Não envia determinada Tag
@param lAdd, logical, Incluir direto a tag ou analisa se já existe
@param cPath2, character, caminho relativo ao xml2
@param aOrdem, array, Ordem para organizar as tags
/*/
Static Function Xml2Xml(oXml,oXml2,cPath,cFiltro,cNaoFazer,lAdd,cPath2,aOrdem)
	Local aChildren,aAtrr
	Local nCont,nCont2
	Local cPos
	Local cValor
	Local nPosOrdem
	Default lAdd	:= .T.
	Default cPath2	:= cPath
	If !Empty(aOrdem) 
		nPosOrdem	:= aScan(aOrdem,{|x| x[1]==Replace(Replace(cPath,"[1]",""),"[last()]","")})
	Endif
	aChildren	:=  oXML2:XPathGetChildArray(cPath2)
	If nPosOrdem>0
		//Se foi enviado a ordem para essa tag, reorganiza aChildren
		aSort(aChildren,,,{|x,y| (aScanOrdem(aOrdem[nPosOrdem][2],x[1])*1000)+Val(SubStr(x[2],rat("[",x[2])+1,rat("]",x[2])-rat("[",x[2])-1))<(aScanOrdem(aOrdem[nPosOrdem][2],y[1])*1000)+Val(SubStr(y[2],rat("[",y[2])+1,rat("]",y[2])-rat("[",y[2])-1)) })
	EndIf
	For nCont:=1 to Len(aChildren)
		If !Empty(cFiltro) .and. cFiltro!=aChildren[nCont][1]
			Loop
		EndIf
		If !Empty(cNaoFazer) .AND. cNaoFazer==aChildren[nCont][1]
			Loop
		EndIf
		If lAdd
			oXml:XPathAddNode(cPath,aChildren[nCont][1],"")
			cPos	:= "last()"
		ElseIf !oXml:XPathHasNode(cPath+"/xmlns:"+aChildren[nCont][1])
			oXml:XPathAddNode(cPath,aChildren[nCont][1],"")
			cPos	:= "last()"
		Else
			cPos	:= "1"
		EndIf
		//Atributos
		aAtrr	:= oXML2:XPathGetAttArray(aChildren[nCont][2])
		For nCont2:=1 to Len(aAtrr)
			If (aChildren[nCont][1]=="drawing".OR.aChildren[nCont][1]=="tablePart").AND.aAtrr[nCont2][1]=="id"
				SetAtrr(oXml,cPath+"/xmlns:"+aChildren[nCont][1]+"["+cPos+"]","r:"+aAtrr[nCont2][1],aAtrr[nCont2][2])
			Else
				SetAtrr(oXml,cPath+"/xmlns:"+aChildren[nCont][1]+"["+cPos+"]",aAtrr[nCont2][1],aAtrr[nCont2][2])
			EndIf
		Next
		//Replica filhos
		If !Empty(oXML2:XPathGetChildArray(aChildren[nCont][2]))
			Xml2Xml(oXml,oXml2,cPath+"/xmlns:"+aChildren[nCont][1]+"["+cPos+"]",,,,aChildren[nCont][2],aOrdem)
		Else
			cValor	:= EncodeUTF8(Replace(oXML2:XPathGetNodeValue(aChildren[nCont][2]),"&","&amp;"))
			oXml:XPathSetNode(cPath+"/xmlns:"+aChildren[nCont][1]+"["+cPos+"]",aChildren[nCont][1],cValor)
		EndIf
	Next
Return

/*/{Protheus.doc} xls_table
Cria arquivo \xl\tables\tableX.xml
@author Saulo Gomes Martins
@since 10/12/2019
@version 1.0

@type method
/*/
Method xls_table(nCont,nCont2) Class YExcel
	Local cRet	:= ""
	cRet	+= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
	cRet	+= ::aPlanilhas[nCont][5][nCont2]:GetTag()
Return cRet

/*/{Protheus.doc} xls_sharedStrings
Cria arquivo /xl/sharedStrings.xml
@author Saulo Gomes Martins
@since 10/12/2019
@version 1.0
@param nFile, numeric, header do arquivo
@type method
/*/
Method xls_sharedStrings(nFile) Class YExcel
	Local nTam
	Local cRet	:= ""
	Local cTexto
	nTam	:= ::nQtdString
	cRet	+= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
	cRet	+= '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="'+cValToChar(nTam)+'" uniqueCount="'+cValToChar(nTam)+'">'
	FWRITE(nFile,cRet)
	cRet	:= ""

	(::cAliasStr)->(DbSetOrder(2))	//POS
	(::cAliasStr)->(DbGoTop())
	While (::cAliasStr)->(!EOF())
		cRet	+= '<si>'
		cTexto	:= EncodeUTF8(&((::cAliasStr)->VLRMEMO))
		If Valtype(cTexto)!="C"
			cTexto	:= &((::cAliasStr)->VLRMEMO)
			cTexto	:= Replace(cTexto,chr(129),"")
			cTexto	:= Replace(cTexto,chr(141),"")
			cTexto	:= Replace(cTexto,chr(143),"")
			cTexto	:= Replace(cTexto,chr(144),"")
			cTexto	:= Replace(cTexto,chr(157),"")
			cTexto	:= EncodeUTF8(cTexto)
			If Valtype(cTexto)!="C"
				cTexto	:= ""
			Endif
		Endif
		cRet	+= '<t xml:space="preserve"><![CDATA['+cTexto+']]></t>'
		cRet	+= '</si>'
		FWRITE(nFile,cRet)
		cRet	:= ""
		(::cAliasStr)->(DbSkip())
	EndDo
	cRet	+= '</sst>'
	FWRITE(nFile,cRet)
	cRet	:= ""
Return cRet
/*/{Protheus.doc} YExcel::Read_sharedStrings
Ler strings compartilhada para gravar no banco
@type method
@version 1.0
@author Saulo Gomes Martins
@since 17/03/2021
@param cFile, character, Nome do arquivo
/*/
Method Read_sharedStrings(cFile) Class YExcel
	Local oTmp
	Local aNs
	Local nCont
	oTmp := TXmlManager():New()
	oTmp:ParseFile(cFile)
	aNs	:= oTmp:XPathGetRootNsList()
	For nCont:=1 to Len(aNs)
		oTmp:XPathRegisterNs( aNs[nCont][1], aNs[nCont][2] )
	Next
	For nCont:=1 to Val(oTmp:XPathGetAtt("/xmlns:sst","count"))
		::SetStrComp(oTmp:XPathGetNodeValue("/xmlns:sst/xmlns:si["+cValToChar(nCont)+"]/xmlns:t"))
	Next
	AADD(::aFiles,cFile)
	FreeObj(oTmp)
Return
/*/{Protheus.doc} FWMakeDir
Cria o caminho completo, a função padrão não mantem o case
@type function
@version 1.0
@author Saulo Gomes Martins
@since 17/03/2021
@param cCaminho, character, Caminho
@param lShowMsg, logical, não usado
/*/
Static Function FWMakeDir(cCaminho,lShowMsg)
	Local aPastas
	Local nCont
	Local cCamiAtu	:= ""
	cCaminho	:= Replace(cCaminho,"/","\")
	aPastas		:= StrToKArr(cCaminho,"\")
	For nCont:=1 to Len(aPastas)
		If !(":" $ aPastas[nCont])
			cCamiAtu	+= "\"
		Endif
		cCamiAtu	+= aPastas[nCont]
		If ":" $ aPastas[nCont]
			Loop
		Endif
		MakeDir(cCamiAtu,,.F.)
	Next
Return
/*/{Protheus.doc} SetAtrr
Altera, inclui ou exclui um atributo
@type function
@version 1.0
@author Saulo Gomes Martins
@since 17/03/2021
@param oXml, object, Objeto xml
@param cPath, character, caminho
@param cAtrr, character, Atributo
@param cValAtrr, character, Valor
/*/
Static Function SetAtrr(oXml,cPath,cAtrr,cValAtrr)
	If Empty(cValAtrr)
		oXml:XPathDelAtt(cPath,cAtrr)
	ElseIf !Empty(oXml:XPathGetAtt(cPath,cAtrr))
		oXml:XPathSetAtt(cPath,cAtrr,cValAtrr)
	Else
		oXml:XPathAddAtt(cPath,cAtrr,cValAtrr)
	Endif
Return

/*/{Protheus.doc} yxlsthem
Cria thema do YExcel
@author Saulo Gomes Martins
@since 10/12/2019
@version 1.0

@type function
/*/
user function yxlsthe2()
	Local cRet := ""

	cRet += PlainH_1()
	cRet += PlainH_2()
	cRet += PlainH_3()
	cRet += PlainH_4()
	cRet += PlainH_5()
	cRet += PlainH_6()
	cRet += PlainH_7()
	cRet += PlainH_8()
	cRet += PlainH_9()

Return(cRet)

Static Function PlainH_1()
	Local cRet := ""

	cRet += '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CHR(13)+CHR(10)
	cRet += '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Tema do Office">' + CHR(13)+CHR(10)
	cRet += "	<a:themeElements>" + CHR(13)+CHR(10)
	cRet += '		<a:clrScheme name="Escritório">' + CHR(13)+CHR(10)
	cRet += "			<a:dk1>" + CHR(13)+CHR(10)
	cRet += '				<a:sysClr val="windowText" lastClr="000000"/>' + CHR(13)+CHR(10)
	cRet += "			</a:dk1>" + CHR(13)+CHR(10)
	cRet += "			<a:lt1>" + CHR(13)+CHR(10)
	cRet += '				<a:sysClr val="window" lastClr="FFFFFF"/>' + CHR(13)+CHR(10)
	cRet += "			</a:lt1>" + CHR(13)+CHR(10)
	cRet += "			<a:dk2>" + CHR(13)+CHR(10)
	cRet += '				<a:srgbClr val="1F497D"/>' + CHR(13)+CHR(10)
	cRet += "			</a:dk2>" + CHR(13)+CHR(10)
	cRet += "			<a:lt2>" + CHR(13)+CHR(10)
	cRet += '				<a:srgbClr val="EEECE1"/>' + CHR(13)+CHR(10)
	cRet += "			</a:lt2>" + CHR(13)+CHR(10)
	cRet += "			<a:accent1>" + CHR(13)+CHR(10)
	cRet += '				<a:srgbClr val="4F81BD"/>' + CHR(13)+CHR(10)
	cRet += "			</a:accent1>" + CHR(13)+CHR(10)
	cRet += "			<a:accent2>" + CHR(13)+CHR(10)
	cRet += '				<a:srgbClr val="C0504D"/>' + CHR(13)+CHR(10)
	cRet += "			</a:accent2>" + CHR(13)+CHR(10)
	cRet += "			<a:accent3>" + CHR(13)+CHR(10)
	cRet += '				<a:srgbClr val="9BBB59"/>' + CHR(13)+CHR(10)
	cRet += "			</a:accent3>" + CHR(13)+CHR(10)
	cRet += "			<a:accent4>" + CHR(13)+CHR(10)
	cRet += '				<a:srgbClr val="8064A2"/>' + CHR(13)+CHR(10)
	cRet += "			</a:accent4>" + CHR(13)+CHR(10)
	cRet += "			<a:accent5>" + CHR(13)+CHR(10)
	cRet += '				<a:srgbClr val="4BACC6"/>' + CHR(13)+CHR(10)
	cRet += "			</a:accent5>" + CHR(13)+CHR(10)
Return(cRet)

Static Function PlainH_2()
	Local cRet := ""

	cRet += "			<a:accent6>" + CHR(13)+CHR(10)
	cRet += '				<a:srgbClr val="F79646"/>' + CHR(13)+CHR(10)
	cRet += "			</a:accent6>" + CHR(13)+CHR(10)
	cRet += "			<a:hlink>" + CHR(13)+CHR(10)
	cRet += '				<a:srgbClr val="0000FF"/>' + CHR(13)+CHR(10)
	cRet += "			</a:hlink>" + CHR(13)+CHR(10)
	cRet += "			<a:folHlink>" + CHR(13)+CHR(10)
	cRet += '				<a:srgbClr val="800080"/>' + CHR(13)+CHR(10)
	cRet += "			</a:folHlink>" + CHR(13)+CHR(10)
	cRet += "		</a:clrScheme>" + CHR(13)+CHR(10)
	cRet += '		<a:fontScheme name="Escritório">' + CHR(13)+CHR(10)
	cRet += "			<a:majorFont>" + CHR(13)+CHR(10)
	cRet += '				<a:latin typeface="Cambria"/>' + CHR(13)+CHR(10)
	cRet += '				<a:ea typeface=""/>' + CHR(13)+CHR(10)
	cRet += '				<a:cs typeface=""/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Arab" typeface="Times New Roman"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Hebr" typeface="Times New Roman"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Thai" typeface="Tahoma"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Ethi" typeface="Nyala"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Beng" typeface="Vrinda"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Gujr" typeface="Shruti"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Khmr" typeface="MoolBoran"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Knda" typeface="Tunga"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Guru" typeface="Raavi"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Cans" typeface="Euphemia"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Cher" typeface="Plantagenet Cherokee"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Tibt" typeface="Microsoft Himalaya"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Thaa" typeface="MV Boli"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Deva" typeface="Mangal"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Telu" typeface="Gautami"/>' + CHR(13)+CHR(10)
Return(cRet)

Static Function PlainH_3()
	Local cRet := ""

	cRet += '				<a:font script="Taml" typeface="Latha"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Syrc" typeface="Estrangelo Edessa"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Orya" typeface="Kalinga"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Mlym" typeface="Kartika"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Laoo" typeface="DokChampa"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Sinh" typeface="Iskoola Pota"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Mong" typeface="Mongolian Baiti"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Viet" typeface="Times New Roman"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Uigh" typeface="Microsoft Uighur"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Geor" typeface="Sylfaen"/>' + CHR(13)+CHR(10)
	cRet += "			</a:majorFont>" + CHR(13)+CHR(10)
	cRet += "			<a:minorFont>" + CHR(13)+CHR(10)
	cRet += '				<a:latin typeface="Calibri"/>' + CHR(13)+CHR(10)
	cRet += '				<a:ea typeface=""/>' + CHR(13)+CHR(10)
	cRet += '				<a:cs typeface=""/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Arab" typeface="Arial"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Hebr" typeface="Arial"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Thai" typeface="Tahoma"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Ethi" typeface="Nyala"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Beng" typeface="Vrinda"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Gujr" typeface="Shruti"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Khmr" typeface="DaunPenh"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Knda" typeface="Tunga"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Guru" typeface="Raavi"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Cans" typeface="Euphemia"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Cher" typeface="Plantagenet Cherokee"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Tibt" typeface="Microsoft Himalaya"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Thaa" typeface="MV Boli"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Deva" typeface="Mangal"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Telu" typeface="Gautami"/>' + CHR(13)+CHR(10)
Return(cRet)

Static Function PlainH_4()
	Local cRet := ""

	cRet += '				<a:font script="Taml" typeface="Latha"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Syrc" typeface="Estrangelo Edessa"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Orya" typeface="Kalinga"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Mlym" typeface="Kartika"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Laoo" typeface="DokChampa"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Sinh" typeface="Iskoola Pota"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Mong" typeface="Mongolian Baiti"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Viet" typeface="Arial"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Uigh" typeface="Microsoft Uighur"/>' + CHR(13)+CHR(10)
	cRet += '				<a:font script="Geor" typeface="Sylfaen"/>' + CHR(13)+CHR(10)
	cRet += "			</a:minorFont>" + CHR(13)+CHR(10)
	cRet += "		</a:fontScheme>" + CHR(13)+CHR(10)
	cRet += '		<a:fmtScheme name="Escritório">' + CHR(13)+CHR(10)
	cRet += "			<a:fillStyleLst>" + CHR(13)+CHR(10)
	cRet += "				<a:solidFill>" + CHR(13)+CHR(10)
	cRet += '					<a:schemeClr val="phClr"/>' + CHR(13)+CHR(10)
	cRet += "				</a:solidFill>" + CHR(13)+CHR(10)
	cRet += '				<a:gradFill rotWithShape="1">' + CHR(13)+CHR(10)
	cRet += "					<a:gsLst>" + CHR(13)+CHR(10)
	cRet += '						<a:gs pos="0">' + CHR(13)+CHR(10)
	cRet += '							<a:schemeClr val="phClr">' + CHR(13)+CHR(10)
	cRet += '								<a:tint val="50000"/>' + CHR(13)+CHR(10)
	cRet += '								<a:satMod val="300000"/>' + CHR(13)+CHR(10)
	cRet += "							</a:schemeClr>" + CHR(13)+CHR(10)
	cRet += "						</a:gs>" + CHR(13)+CHR(10)
	cRet += '						<a:gs pos="35000">' + CHR(13)+CHR(10)
	cRet += '							<a:schemeClr val="phClr">' + CHR(13)+CHR(10)
	cRet += '								<a:tint val="37000"/>' + CHR(13)+CHR(10)
	cRet += '								<a:satMod val="300000"/>' + CHR(13)+CHR(10)
	cRet += "							</a:schemeClr>" + CHR(13)+CHR(10)
	cRet += "						</a:gs>" + CHR(13)+CHR(10)
Return(cRet)

Static Function PlainH_5()
	Local cRet := ""

	cRet += '						<a:gs pos="100000">' + CHR(13)+CHR(10)
	cRet += '							<a:schemeClr val="phClr">' + CHR(13)+CHR(10)
	cRet += '								<a:tint val="15000"/>' + CHR(13)+CHR(10)
	cRet += '								<a:satMod val="350000"/>' + CHR(13)+CHR(10)
	cRet += "							</a:schemeClr>" + CHR(13)+CHR(10)
	cRet += "						</a:gs>" + CHR(13)+CHR(10)
	cRet += "					</a:gsLst>" + CHR(13)+CHR(10)
	cRet += '					<a:lin ang="16200000" scaled="1"/>' + CHR(13)+CHR(10)
	cRet += "				</a:gradFill>" + CHR(13)+CHR(10)
	cRet += '				<a:gradFill rotWithShape="1">' + CHR(13)+CHR(10)
	cRet += "					<a:gsLst>" + CHR(13)+CHR(10)
	cRet += '						<a:gs pos="0">' + CHR(13)+CHR(10)
	cRet += '							<a:schemeClr val="phClr">' + CHR(13)+CHR(10)
	cRet += '								<a:shade val="51000"/>' + CHR(13)+CHR(10)
	cRet += '								<a:satMod val="130000"/>' + CHR(13)+CHR(10)
	cRet += "							</a:schemeClr>" + CHR(13)+CHR(10)
	cRet += "						</a:gs>" + CHR(13)+CHR(10)
	cRet += '						<a:gs pos="80000">' + CHR(13)+CHR(10)
	cRet += '							<a:schemeClr val="phClr">' + CHR(13)+CHR(10)
	cRet += '								<a:shade val="93000"/>' + CHR(13)+CHR(10)
	cRet += '								<a:satMod val="130000"/>' + CHR(13)+CHR(10)
	cRet += "							</a:schemeClr>" + CHR(13)+CHR(10)
	cRet += "						</a:gs>" + CHR(13)+CHR(10)
	cRet += '						<a:gs pos="100000">' + CHR(13)+CHR(10)
	cRet += '							<a:schemeClr val="phClr">' + CHR(13)+CHR(10)
	cRet += '								<a:shade val="94000"/>' + CHR(13)+CHR(10)
	cRet += '								<a:satMod val="135000"/>' + CHR(13)+CHR(10)
	cRet += "							</a:schemeClr>" + CHR(13)+CHR(10)
	cRet += "						</a:gs>" + CHR(13)+CHR(10)
	cRet += "					</a:gsLst>" + CHR(13)+CHR(10)
	cRet += '					<a:lin ang="16200000" scaled="0"/>' + CHR(13)+CHR(10)
Return(cRet)

Static Function PlainH_6()
	Local cRet := ""

	cRet += "				</a:gradFill>" + CHR(13)+CHR(10)
	cRet += "			</a:fillStyleLst>" + CHR(13)+CHR(10)
	cRet += "			<a:lnStyleLst>" + CHR(13)+CHR(10)
	cRet += '				<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">' + CHR(13)+CHR(10)
	cRet += "					<a:solidFill>" + CHR(13)+CHR(10)
	cRet += '						<a:schemeClr val="phClr">' + CHR(13)+CHR(10)
	cRet += '							<a:shade val="95000"/>' + CHR(13)+CHR(10)
	cRet += '							<a:satMod val="105000"/>' + CHR(13)+CHR(10)
	cRet += "						</a:schemeClr>" + CHR(13)+CHR(10)
	cRet += "					</a:solidFill>" + CHR(13)+CHR(10)
	cRet += '					<a:prstDash val="solid"/>' + CHR(13)+CHR(10)
	cRet += "				</a:ln>" + CHR(13)+CHR(10)
	cRet += '				<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">' + CHR(13)+CHR(10)
	cRet += "					<a:solidFill>" + CHR(13)+CHR(10)
	cRet += '						<a:schemeClr val="phClr"/>' + CHR(13)+CHR(10)
	cRet += "					</a:solidFill>" + CHR(13)+CHR(10)
	cRet += '					<a:prstDash val="solid"/>' + CHR(13)+CHR(10)
	cRet += "				</a:ln>" + CHR(13)+CHR(10)
	cRet += '				<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">' + CHR(13)+CHR(10)
	cRet += "					<a:solidFill>" + CHR(13)+CHR(10)
	cRet += '						<a:schemeClr val="phClr"/>' + CHR(13)+CHR(10)
	cRet += "					</a:solidFill>" + CHR(13)+CHR(10)
	cRet += '					<a:prstDash val="solid"/>' + CHR(13)+CHR(10)
	cRet += "				</a:ln>" + CHR(13)+CHR(10)
	cRet += "			</a:lnStyleLst>" + CHR(13)+CHR(10)
	cRet += "			<a:effectStyleLst>" + CHR(13)+CHR(10)
	cRet += "				<a:effectStyle>" + CHR(13)+CHR(10)
	cRet += "					<a:effectLst>" + CHR(13)+CHR(10)
	cRet += '						<a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">' + CHR(13)+CHR(10)
	cRet += '							<a:srgbClr val="000000">' + CHR(13)+CHR(10)
	cRet += '								<a:alpha val="38000"/>' + CHR(13)+CHR(10)
Return(cRet)

Static Function PlainH_7()
	Local cRet := ""

	cRet += "							</a:srgbClr>" + CHR(13)+CHR(10)
	cRet += "						</a:outerShdw>" + CHR(13)+CHR(10)
	cRet += "					</a:effectLst>" + CHR(13)+CHR(10)
	cRet += "				</a:effectStyle>" + CHR(13)+CHR(10)
	cRet += "				<a:effectStyle>" + CHR(13)+CHR(10)
	cRet += "					<a:effectLst>" + CHR(13)+CHR(10)
	cRet += '						<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">' + CHR(13)+CHR(10)
	cRet += '							<a:srgbClr val="000000">' + CHR(13)+CHR(10)
	cRet += '								<a:alpha val="35000"/>' + CHR(13)+CHR(10)
	cRet += "							</a:srgbClr>" + CHR(13)+CHR(10)
	cRet += "						</a:outerShdw>" + CHR(13)+CHR(10)
	cRet += "					</a:effectLst>" + CHR(13)+CHR(10)
	cRet += "				</a:effectStyle>" + CHR(13)+CHR(10)
	cRet += "				<a:effectStyle>" + CHR(13)+CHR(10)
	cRet += "					<a:effectLst>" + CHR(13)+CHR(10)
	cRet += '						<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">' + CHR(13)+CHR(10)
	cRet += '							<a:srgbClr val="000000">' + CHR(13)+CHR(10)
	cRet += '								<a:alpha val="35000"/>' + CHR(13)+CHR(10)
	cRet += "							</a:srgbClr>" + CHR(13)+CHR(10)
	cRet += "						</a:outerShdw>" + CHR(13)+CHR(10)
	cRet += "					</a:effectLst>" + CHR(13)+CHR(10)
	cRet += "					<a:scene3d>" + CHR(13)+CHR(10)
	cRet += '						<a:camera prst="orthographicFront">' + CHR(13)+CHR(10)
	cRet += '							<a:rot lat="0" lon="0" rev="0"/>' + CHR(13)+CHR(10)
	cRet += "						</a:camera>" + CHR(13)+CHR(10)
	cRet += '						<a:lightRig rig="threePt" dir="t">' + CHR(13)+CHR(10)
	cRet += '							<a:rot lat="0" lon="0" rev="1200000"/>' + CHR(13)+CHR(10)
	cRet += "						</a:lightRig>" + CHR(13)+CHR(10)
	cRet += "					</a:scene3d>" + CHR(13)+CHR(10)
	cRet += "					<a:sp3d>" + CHR(13)+CHR(10)
	cRet += '						<a:bevelT w="63500" h="25400"/>' + CHR(13)+CHR(10)
Return(cRet)

Static Function PlainH_8()
	Local cRet := ""

	cRet += "					</a:sp3d>" + CHR(13)+CHR(10)
	cRet += "				</a:effectStyle>" + CHR(13)+CHR(10)
	cRet += "			</a:effectStyleLst>" + CHR(13)+CHR(10)
	cRet += "			<a:bgFillStyleLst>" + CHR(13)+CHR(10)
	cRet += "				<a:solidFill>" + CHR(13)+CHR(10)
	cRet += '					<a:schemeClr val="phClr"/>' + CHR(13)+CHR(10)
	cRet += "				</a:solidFill>" + CHR(13)+CHR(10)
	cRet += '				<a:gradFill rotWithShape="1">' + CHR(13)+CHR(10)
	cRet += "					<a:gsLst>" + CHR(13)+CHR(10)
	cRet += '						<a:gs pos="0">' + CHR(13)+CHR(10)
	cRet += '							<a:schemeClr val="phClr">' + CHR(13)+CHR(10)
	cRet += '								<a:tint val="40000"/>' + CHR(13)+CHR(10)
	cRet += '								<a:satMod val="350000"/>' + CHR(13)+CHR(10)
	cRet += "							</a:schemeClr>" + CHR(13)+CHR(10)
	cRet += "						</a:gs>" + CHR(13)+CHR(10)
	cRet += '						<a:gs pos="40000">' + CHR(13)+CHR(10)
	cRet += '							<a:schemeClr val="phClr">' + CHR(13)+CHR(10)
	cRet += '								<a:tint val="45000"/>' + CHR(13)+CHR(10)
	cRet += '								<a:shade val="99000"/>' + CHR(13)+CHR(10)
	cRet += '								<a:satMod val="350000"/>' + CHR(13)+CHR(10)
	cRet += "							</a:schemeClr>" + CHR(13)+CHR(10)
	cRet += "						</a:gs>" + CHR(13)+CHR(10)
	cRet += '						<a:gs pos="100000">' + CHR(13)+CHR(10)
	cRet += '							<a:schemeClr val="phClr">' + CHR(13)+CHR(10)
	cRet += '								<a:shade val="20000"/>' + CHR(13)+CHR(10)
	cRet += '								<a:satMod val="255000"/>' + CHR(13)+CHR(10)
	cRet += "							</a:schemeClr>" + CHR(13)+CHR(10)
	cRet += "						</a:gs>" + CHR(13)+CHR(10)
	cRet += "					</a:gsLst>" + CHR(13)+CHR(10)
	cRet += '					<a:path path="circle">' + CHR(13)+CHR(10)
	cRet += '						<a:fillToRect l="50000" t="-80000" r="50000" b="180000"/>' + CHR(13)+CHR(10)
Return(cRet)

Static Function PlainH_9()
	Local cRet := ""

	cRet += "					</a:path>" + CHR(13)+CHR(10)
	cRet += "				</a:gradFill>" + CHR(13)+CHR(10)
	cRet += '				<a:gradFill rotWithShape="1">' + CHR(13)+CHR(10)
	cRet += "					<a:gsLst>" + CHR(13)+CHR(10)
	cRet += '						<a:gs pos="0">' + CHR(13)+CHR(10)
	cRet += '							<a:schemeClr val="phClr">' + CHR(13)+CHR(10)
	cRet += '								<a:tint val="80000"/>' + CHR(13)+CHR(10)
	cRet += '								<a:satMod val="300000"/>' + CHR(13)+CHR(10)
	cRet += "							</a:schemeClr>" + CHR(13)+CHR(10)
	cRet += "						</a:gs>" + CHR(13)+CHR(10)
	cRet += '						<a:gs pos="100000">' + CHR(13)+CHR(10)
	cRet += '							<a:schemeClr val="phClr">' + CHR(13)+CHR(10)
	cRet += '								<a:shade val="30000"/>' + CHR(13)+CHR(10)
	cRet += '								<a:satMod val="200000"/>' + CHR(13)+CHR(10)
	cRet += "							</a:schemeClr>" + CHR(13)+CHR(10)
	cRet += "						</a:gs>" + CHR(13)+CHR(10)
	cRet += "					</a:gsLst>" + CHR(13)+CHR(10)
	cRet += '					<a:path path="circle">' + CHR(13)+CHR(10)
	cRet += '						<a:fillToRect l="50000" t="50000" r="50000" b="50000"/>' + CHR(13)+CHR(10)
	cRet += "					</a:path>" + CHR(13)+CHR(10)
	cRet += "				</a:gradFill>" + CHR(13)+CHR(10)
	cRet += "			</a:bgFillStyleLst>" + CHR(13)+CHR(10)
	cRet += "		</a:fmtScheme>" + CHR(13)+CHR(10)
	cRet += "	</a:themeElements>" + CHR(13)+CHR(10)
	cRet += "	<a:objectDefaults/>" + CHR(13)+CHR(10)
	cRet += "	<a:extraClrSchemeLst/>" + CHR(13)+CHR(10)
	cRet += "</a:theme>" + CHR(13)+CHR(10)
Return(cRet)

// Static Function LimpStrC()
// 	If !DbSqlExec(::cAliasLin,"INSERT INTO "+::cAliasLin+" (PLA,LIN) SELECT DISTINCT C.PLA,C.LIN FROM "+::cAliasCol+" C LEFT JOIN "+::cAliasLin+" L ON C.PLA=L.PLA AND C.LIN=L.LIN WHERE L.LIN IS NULL",::cDriver)
// 		UserException("YExcel - Erro ao incluir linhas. "+TCSqlError())
// 	Endif
// SELECT ROW_NUMBER() OVER (ORDER BY STR.ROWID)-1 SEQUENCIA,* 
// FROM STRSC049180 STR
// LEFT JOIN COLSC049160 COL
// ON COL.TIPO='s' AND COL.VLRNUM=STR.POS
// WHERE COL.VLRNUM IS NULL
// ORDER BY 1
// ;
// Return
//EM DESENVOLVIMENTO
/*
formula=
 expression ;
expression=
 "(", expression, ")" |
 constant |
 prefix-operator, expression |
 expression, infix-operator, expression |
 expression, postfix-operator |
 cell-reference |
 function-call |
 name ;
*/
Class YExcelVar
	Data cTipo				//T=TEXTO;N=NUMERO;L=LOGICO;R=REFERENCIA;F=FUNÇÃO;E=EXPRESSÃO
	Data xValor				//VALOR
	Data cFuncao			//Nome da função
	Data cPrefixOper		//Operador inicio
	Data cPosfixOper		//Operador final
	Data cFormula
	Method New()
	Method GetValue()
	Method SetValue()
	Method ADDValue()
	Method GetLen()
	Method SetType()
	Method SetFomula()
	Method GetPre()
	Method SetPre()
	Method SetPos()
	Method SetFuncao()
EndClass

Method New(cTipo) Class YExcelVar
	::xValor		:= {}
	::cTipo			:= cTipo
Return Self

Method GetValue() Class YExcelVar
Return ::xValor

Method GetLen() Class YExcelVar
Return Len(::xValor)

Method SetValue(xValor) Class YExcelVar
	::xValor	:= xValor
Return self

Method ADDValue(xValorPar) Class YExcelVar
	// Local oValor	:= yExcelVar():New()
	AADD(::xValor,xValorPar)
	// oValor:SetValue(xValor)
Return xValorPar

Method SetType(cTipo) Class YExcelVar
	::cTipo	:= cTipo
Return self

Method SetFomula(cFormula) Class YExcelVar
	::cFormula	:= cFormula
Return self

Method GetPre(cPrefixOper) Class YExcelVar
Return ::cPrefixOper

Method SetPre(cPrefixOper) Class YExcelVar
	::cPrefixOper	:= cPrefixOper
	cPrefixOper		:= ""
Return self

Method SetPos(cPosfixOper) Class YExcelVar
	::cPosfixOper	:= cPosfixOper
	cPosfixOper		:= ""
Return self

Method SetFuncao(cFuncao) Class YExcelVar
	::cFuncao	:= cFuncao
Return self

Static aAriOpe	:= {"-","%","^","*","/","+"}
Static cTexOpe	:= "&"
Static aLogOpe	:= {"=","<",">"}//{"=","<>","<","<=",">",">="}

Static aModicadores := {"(",")","=","+","%","^","*","/","+","&","<",">",","}

User Function ytstfor2()
	Local cFormula	:= '"Valor:"&A1+SUM(C3:C4)+123.132*'
	Local oRet
	oParseFor	:= yExcelfunction():New()
	oRet	:= oParseFor:Parse(cFormula)
	VarInfo("oRet",oRet,,.F.)
Return

Class YExcelfunction
	Data aDados
	Method New()
	Method Parse()
EndClass

Method New() Class YExcelfunction
Return self

Method Parse(cFormula,nQtdLido) Class YExcelfunction
	Local cTexto	:= ""
	Local cTexto2	:= ""
	Local nTam		:= Len(cFormula)
	Local cTipo		:= ""
	Local cNumero	:= ""
	Local cValor
	Local lRef
	Local oRet		:= yExcelVar():New()
	Local nQtdSoma	:= 0
	Local cPre		:= ""
	Local cPos
	Local nCont
	Local oTmp

	nQtdLido		:= 0
	oRet:SetValue(Array(0))
	For nCont:=1 to Len(cFormula)
		nQtdSoma	:= 0
		cNumero		:= ""
		cValor		:= ""
		cTexto		:= SubStr(cFormula,nCont,1)
		cTexto2		:= SubStr(cFormula,nCont+1,1)	//Proximo letra
		cTipo		:= tpExpressao(cTexto)
		If cTexto=='"'	//Texto
			cValor		:= ""
			nCont++
			While nCont<=nTam
				cTexto		:= SubStr(cFormula,nCont,1)
				If cTexto=='"'
					Exit
				Endif
				cValor		+= cTexto
				nCont++
			EndDo
			oTmp:=yExcelVar():New()
			oTmp:SetValue(cValor):SetType("T"):SetPre(@cPre):SetPos(@cPos)
			oRet:ADDValue(oTmp)
		ElseIf aScan(aAriOpe,cTexto)>0	//Operadores simples
			cPre	+= cTexto
		ElseIf aScan(aLogOpe,cTexto)>0	//Operadores logicos
				cPre	+= cTexto
		ElseIf cTexto=="&"				//Operadores texto
			cPre	+= cTexto
		ElseIf cTexto=="("
			oTmp	:= ::Parse(SubStr(cFormula,nCont+1),@nQtdSoma)
			oTmp:SetFuncao("("):SetType("E"):SetPre(@cPre):SetPos(@cPos)
			oRet:ADDValue(oTmp)
			nCont	+= nQtdSoma+2			
		ElseIf cTexto==")"	//Sair da função
			Exit
		ElseIf cTexto==","
			Loop
		ElseIf cTipo=="N"
			cNumero	:= cTexto
			While (cTexto2=="." .OR. tpExpressao(cTexto2)=="N") .AND. nCont<=nTam
				cNumero	+= cTexto2
				nCont++
				cTexto2	:= SubStr(cFormula,nCont+1,1)
			EndDo
			oTmp:=yExcelVar():New()
			oTmp:SetValue(Val(cNumero)):SetType("N"):SetPre(@cPre):SetPos(@cPos)
			oRet:ADDValue(oTmp)
		ElseIf cTipo=="C"
			lRef		:= .F.
			cValor		:= ""
			While aScan(aModicadores,cTexto)==0 .AND. nCont<=nTam
				If tpExpressao(cTexto)=="N" .OR. cTexto==":"
					lRef		:= .T.
				Endif
				cValor		+= cTexto
				nCont++
				cTexto		:= SubStr(cFormula,nCont,1)
			EndDo
			If lRef
				oTmp:=yExcelVar():New()
				oTmp:SetValue(cValor):SetType("R"):SetPre(@cPre):SetPos(@cPos)
				oRet:ADDValue(oTmp)
			ElseIf cTexto=="("		//Formula
				// oRet:SetFuncao(cValor)
				oTmp	:= ::Parse(SubStr(cFormula,nCont+1),@nQtdSoma)
				oTmp:SetFuncao(cValor):SetType("F"):SetPre(@cPre):SetPos(@cPos)
				oRet:ADDValue(oTmp)
				nCont	+= nQtdSoma+2
			ElseIf cValor=="TRUE".or.cValor=="FALSE"
				oTmp:=yExcelVar():New()
				oTmp:SetValue(cValor=="TRUE"):SetType("L"):SetPre(@cPre):SetPos(@cPos)
				oRet:ADDValue(oTmp)
			Else
				oTmp:=yExcelVar():New()
				oTmp:SetValue(cValor):SetType("R"):SetPre(@cPre):SetPos(@cPos)
				oRet:ADDValue(oTmp)
			Endif
			nCont--
		Endif
	Next
	// If !Empty(cPos) .AND. oRet:GetLen()>0
	// 	oRet:GetValue()[oRet:GetLen()]:SetPos(@cPos)
	// Endif
	nQtdLido	:= nCont-1
	oRet:SetFomula(SubStr(cFormula,1,nQtdLido))
Return oRet

Static Function tpExpressao(cExpression)
	Local cTipo :=""
	Local nAscii	:= asc(cExpression)
	If nAscii>=48 .and. nAscii<=57	//Numeros
		cTipo	:= "N"
	ElseIf nAscii>=65 .and. nAscii<=90	//Letra Maisculas
		cTipo	:= "C"
	ElseIf nAscii>=97 .and. nAscii<=122	//Letra Minusculas
		cTipo	:= "C"
	Endif
Return cTipo

