#include "Totvs.ch"
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
@obs As linhas e colunas devem ser informadas sempre de forma sequencial crescente
@obs Até a versão 7.00.131227 deve ser intalado no servidor o compactador 7zip
	Adicionar no appserver.ini
	[GENERAL]
	LOCAL7ZIP=C:\Program Files\7-Zip\7z.exe
@OBS
RECURSOS DISPONIVEIS
* Definir células String,Numérica,data,DateTime,Logica,formula
* Adicionar novas planilhas(Nome,Cor)
* Cor de preenchimento(simples,efeito de preenchimento)
* Alinhamento(Horizontal,Vertical,Reduzir para Caber,Quebra Texto,Angulo de Rotação)
* Formato da célula
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
* Cria nome para referência de célula ou intervalo
* Agrupamento de linha
* Imagens
* Exibir/Oculta linhas de Grade
* Definir linha para repetir na impressão

* Leitura simples dos dados
@type class
/*/
//Dummy Function
User Function YExcel()
Return .T.

CLASS YExcel
	Data oString			//String compartilhadas
	Data nQtdString			//Quantidade de string conmpartilhadas
	Data adimension			//Dimensão da planilha
	Data cClassName			//Nome da Classe
	Data cName				//Nome da Classe
	Data osheetData			//Objeto com dados das linhas
	Data cTmpFile			//Arquivo temporario criado no servidor
	Data cNomeFile			//Nome do arquivo para gerar
	Data nFileTmpRow		//nHeader do Arquivo temporario de linhas
	Data cPlanilhaAt
	Data aPlanilhas
	Data osheetViews
	Data oCols
	Data oAutoFilter
	Data oMergeCells
	Data aConditionalFormatting
	Data lRowDef
	Data aSpanRow
	Data nTamLinha
	Data nColunaAtual
	Data nPriodFormCond
	Data otableParts
	Data atable
	Data aFiles
	Data nIdRelat
	Data nCont
	Data osheetPr
	Data oCell
	Data nNumFmtId
	Data cPagOrientation
	Data adrawing		//arquivo drawing de cada sheet
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
	METHOD showGridLines()	//Exibir ou ocultar linhas de grade
	METHOD SetPrintTitles()	//Definir linha para repetir na impressão
	METHOD SetPagOrientation()	//Definir orientação da pagina na impressão

	METHOD GetDateTime()

	//Leitura de planilha
	METHOD OpenRead()
	METHOD CellRead()
	METHOD CloseRead()

	//Interno
	METHOD CriarFile()		//Cria arquivos temporarios
	METHOD GravaFile()		//Grava em arquivos temporarios
	METHOD GravaRow()		//Grava temporario de linhas
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
	METHOD Border()			//Cria Borda com todas opções
	Method AddFmtNum()		//Cria formato para numeros

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

	//Tabela
	METHOD AddTabela()

	//Inicializar TXmlManager
	METHOD new_content_types()
	METHOD new_rels()
	METHOD add_rels()
	METHOD new_app()
	METHOD new_core()
	METHOD new_workbook()
	METHOD new_draw()
	METHOD xls_sheet()
	METHOD xls_table()
	METHOD xls_sharedStrings()
	METHOD new_styles()
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
@param [cEscopo], characters, Planilha de escopo
@type function
/*/
METHOD AddNome(cNome,nLinha,nColuna,nLinha2,nColuna2,cRefPar,cPlanilha,cEscopo) CLASS YExcel
	Local cRef			:= ""
	Local nPos			:= 0
	PARAMTYPE 0	VAR cNome			AS CHARACTER
	PARAMTYPE 1	VAR nLinha			AS NUMERIC			OPTIONAL
	PARAMTYPE 2	VAR nColuna	  		AS NUMERIC			OPTIONAL
	PARAMTYPE 3	VAR nLinha2	  		AS NUMERIC			OPTIONAL
	PARAMTYPE 4	VAR nColuna2  		AS NUMERIC			OPTIONAL
	PARAMTYPE 5	VAR cRefPar	 	 	AS CHARACTER		OPTIONAL
	PARAMTYPE 6	VAR cPlanilha	  	AS CHARACTER		OPTIONAL DEFAULT ::cPlanilhaAt
	PARAMTYPE 7	VAR cEscopo	  		AS CHARACTER		OPTIONAL

	If ValType(cRefPar)=="U"
		If !Empty(cPlanilha)
			cRef	:= "'"+cPlanilha+"'!"
		EndIf
		cRef	+= ::Ref(nLinha,nColuna,.T.,.T.)
		If Valtype(nLinha2)<>"U" .OR. Valtype(nColuna2)<>"U"
			cRef	+= ":"+::Ref(nLinha2,nColuna2,.T.,.T.)
		Endif
	Else
		cRef	:= cRefPar
	EndIf
	If ValType(cEscopo)=="C"
		nPos	:= aScan(::aPlanilhas,{|x| x[2]==cEscopo })
	EndIf
	::oworkbook:XPathAddNode( "/xmlns:workbook/xmlns:definedNames"						, "definedName"			, cRef )
	::oworkbook:XPathAddAtt( "/xmlns:workbook/xmlns:definedNames/xmlns:definedName[last()]", "name"				, cNome)
	If nPos>0
		::oworkbook:XPathAddAtt( "/xmlns:workbook/xmlns:definedNames/xmlns:definedName[last()]", "localSheetId"		, cValToChar(nPos-1))
	EndIf
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
@type function
@obs pag 1566
/*/
METHOD SetPrintTitles(nLinha,nLinha2,cRefPar,cPlanilha) CLASS YExcel
	Default nLinha2	:= nLinha
	::AddNome("_xlnm.Print_Titles",nLinha,,nLinha2,,cRefPar,cPlanilha,::cPlanilhaAt)
Return

/*/{Protheus.doc} SetPagOrientation
Informa a orientação do papel na impressão
@author Saulo Gomes Martins
@since 12/12/2019
@version 1.0
@param cOrientation, characters, descricao
@type function
@obs pag 1667
/*/
METHOD SetPagOrientation(cOrientation) CLASS YExcel
	Default cOrientation := "default"
	If lower(cOrientation)+"|" $ "default|landscape|portrait|"
		::cPagOrientation	:= cOrientation
	EndIf
Return ::cPagOrientation

METHOD ClassName() CLASS YExcel
Return "YEXCEL"

/*/{Protheus.doc} New
Construtor da classe
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param cNomeFile, characters, Nome do arquivo para gerar
@type function
/*/
METHOD New(cNomeFile) CLASS YExcel
	Local nPos
	If ValType(cAr7Zip)=="U"
		cAr7Zip := GetPvProfString("GENERAL", "LOCAL7ZIP" , "C:\Program Files\7-Zip\7z.exe" , GetAdv97() )
	Endif
	PARAMTYPE 0	VAR cNomeFile  AS CHARACTER 		OPTIONAL DEFAULT lower(CriaTrab(,.F.))
	::cClassName	:= "YEXCEL"
	::cName			:= "YEXCEL"
	::oString		:= tHashMap():new()
	::oCell			:= tHashMap():new()	//Usado no leitura simples
	::nQtdString	:= 0
	::nNumFmtId		:= 167
	::aPlanilhas	:= {}
	::cTmpFile		:= lower(CriaTrab(,.F.))
	::cNomeFile		:= lower(cNomeFile)
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
	::new_app()
	::new_core()
	::new_workbook()
	::new_content_types()
	::new_styles()
	::AddFont(11,"FF000000","Calibri","2")
	::Borda()	//Sem borda
	::CorPreenc(,,"none")
	::CorPreenc(,,"gray125")
	::AddStyles(0/*numFmtId*/,/*fontId*/,/*fillId*/,/*borderId*/,/*xfId*/,/*aValores*/,/*aOutrosAtributos*/)	//Sem Formatação
	::AddStyles(14/*numFmtId*/,/*fontId*/,/*fillId*/,/*borderId*/,/*xfId*/,/*aValores*/,{{"applyNumberFormat","0"}}/*aOutrosAtributos*/)	//Formato Data padrão
	::AddStyles(166/*numFmtId*/,/*fontId*/,/*fillId*/,/*borderId*/,/*xfId*/,/*aValores*/,{{"applyNumberFormat","0"}}/*aOutrosAtributos*/)	//Formato Data time padrão

	nPos	:= ::new_rels(,"\_rels\.rels")
	::add_rels(nPos,"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument","xl/workbook.xml")
	::add_rels(nPos,"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties","docProps/core.xml")
	::add_rels(nPos,"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties","docProps/app.xml")

	//Defini formato Moeda padrão brasileiro
	::oStyle:XPathAddNode( "/xmlns:styleSheet/xmlns:numFmts", "numFmt", "" )
	::oStyle:XPathAddAtt( "/xmlns:styleSheet/xmlns:numFmts/xmlns:numFmt[last()]", "formatCode"	, '_-"R$"\ * #,##0.00_-;\-"R$"\ * #,##0.00_-;_-"R$"\ * "-"??_-;_-@_-' )
	::oStyle:XPathAddAtt( "/xmlns:styleSheet/xmlns:numFmts/xmlns:numFmt[last()]", "numFmtId"	, "44" )
	::oStyle:XPathSetAtt("/xmlns:styleSheet/xmlns:numFmts","count",cValToChar(Val(::oStyle:XPathGetAtt("/xmlns:styleSheet/xmlns:numFmts","count"))+1))

	//Defini formato 166
	::oStyle:XPathAddNode( "/xmlns:styleSheet/xmlns:numFmts", "numFmt", "" )
	::oStyle:XPathAddAtt( "/xmlns:styleSheet/xmlns:numFmts/xmlns:numFmt[last()]", "formatCode"	, "dd/mm/yyyy\ hh:mm;@" )
	::oStyle:XPathAddAtt( "/xmlns:styleSheet/xmlns:numFmts/xmlns:numFmt[last()]", "numFmtId"	, "166" )
	::oStyle:XPathSetAtt("/xmlns:styleSheet/xmlns:numFmts","count",cValToChar(Val(::oStyle:XPathGetAtt("/xmlns:styleSheet/xmlns:numFmts","count"))+1))


	AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\sharedStrings.xml")
	AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\theme\theme1.xml")

	nPos	:= ::new_rels(,"\xl\_rels\workbook.xml.rels")
	::add_rels(nPos,"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme","theme/theme1.xml")
	::add_rels(nPos,"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles","styles.xml")
	::add_rels(nPos,"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings","sharedStrings.xml")
Return self


/*/{Protheus.doc} ADDImg
Adiciona imagem para ser usado
@author Saulo Gomes Martins
@since 06/01/2019
@version 1.0
@return nID, ID da imagem
@param cImg, characters, Localização da imagem
@type function
/*/
METHOD ADDImg(cImg) CLASS YExcel
	Local cDrive, cDir, cNome, cExt
	Local cDirImg	:= "\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\media\"
	PARAMTYPE 0	VAR cImg		AS CHARACTER

	If !File(cImg,,.F.)
		UserException("YExcel - Imagem não encontrada ("+cImg+")")
	EndIf

	::nIDMedia++
	FWMakeDir(cDirImg,.F.)
	SplitPath( cImg, @cDrive, @cDir, @cNome, @cExt)
	cNome	:= SubStr(cImg,Rat("\",cImg)+1)
	If ":" $ UPPER(cImg)
		CpyT2S(cImg,cDirImg,,.F.)
		FRename(cDirImg+cNome,cDirImg+"image"+cValToChar(::nIDMedia)+cExt,,.F.)
	Else
		__COPYFILE(cImg,cDirImg+"image"+cValToChar(::nIDMedia)+cExt,,,.F.)
	EndIf
	AADD(::aFiles,cDirImg+"image"+cValToChar(::nIDMedia)+cExt)
	AADD(::aImagens,{::nIDMedia,"image"+cValToChar(::nIDMedia)+cExt})

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
@param [cUnidade], characters, Unidade da dimensão da imagem. padrão em pixel
@param nRot, numeric, rotação da imagem
@type function
@OBS pag 3166
/*/
METHOD Img(nID,nLinha,nColuna,nX,nY,cUnidade,nRot,nQtdPlan) CLASS YExcel
	Local nPos
	Local cCellType
	Local cID
	Local cIdDraw
	Default nQtdPlan	:= Len(::aPlanilhas)
	PARAMTYPE 0	VAR nID			AS NUMERIC
	PARAMTYPE 1	VAR nLinha		AS NUMERIC
	PARAMTYPE 2	VAR nColuna		AS NUMERIC
	PARAMTYPE 3	VAR nY			AS NUMERIC
	PARAMTYPE 4	VAR nX			AS NUMERIC
	PARAMTYPE 5	VAR cUnidade	AS CHARACTER	OPTIONAL DEFAULT "px"
	PARAMTYPE 6	VAR nRot		AS NUMERIC		OPTIONAL DEFAULT 0

	If aScan(::aImagens,{|x| x[1]==nID })==0
		UserException("YExcel - Imagem não cadastrada, usar metodo ADDImg. ID("+cValToChar(nID)+")")
	EndIf

	cUnidade	:= lower(cUnidade)
	//Converte para  EMUs (English Metric Units)
	If cUnidade=="px"
		nX	:= nX*36000*0.2645833333
		nY	:= nY*36000*0.2645833333
	ElseIf cUnidade=="cm"
		nX	:= nX*36000
		nY	:= nY*36000
	EndIf
	Default cCellType	:= "oneCellAnchor"
	//absolute	- Não mover ou redimensionar com linhas / colunas subjacentes
	//oneCell	- Mova-se com células, mas não redimensione
	//twoCell	- Mover e redimensionar com células âncoras

	If Empty(::adrawing)
		::nIdRelat++
		nPos	:= ::nIdRelat
		cID		:= ::add_rels("\xl\worksheets\_rels\sheet"+cValToChar(nQtdPlan)+".xml.rels","http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing","../drawings/drawing"+cValToChar(nPos)+".xml")

		::aPlanilhas[nQtdPlan][3]	:= ::new_draw(,"\xl\drawings\drawing"+cValToChar(nPos)+".xml")

		::odrawing:SetAtributo("r:id",cID)
		::odrawing:xDados	:= nPos
		AADD(::adrawing,nPos)		//Cria o arquivo \xl\drawings\drawing1
		AADD(::aworkdrawing,nPos)	//Cria o arquivo
		//Adiciona um nova drawing no content_types
		::ocontent_types:XPathAddNode( "/xmlns:Types", "Override", "" )
		::ocontent_types:XPathAddAtt( "/xmlns:Types/xmlns:Override[last()]", "PartName"	, "/xl/drawings/drawing"+cValToChar(Len(::aworkdrawing))+".xml" )
		::ocontent_types:XPathAddAtt( "/xmlns:Types/xmlns:Override[last()]", "ContentType", "application/vnd.openxmlformats-officedocument.drawing+xml" )
	EndIf
	nPos	:= ::aPlanilhas[nQtdPlan][3]
	cIdDraw	:= ::add_rels("\xl\drawings\_rels\drawing"+cValToChar(::odrawing:xDados)+".xml.rels","http://schemas.openxmlformats.org/officeDocument/2006/relationships/image","../media/"+::aImagens[nID][2])
	::aDraw[nPos][1]:XPathAddNode( "/xdr:wsDr", cCellType, "" )
	::aDraw[nPos][1]:XPathAddAtt( "/xdr:wsDr/xdr:"+cCellType+"[last()]", "editAs"	, "oneCell" )

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
	::aDraw[nPos][1]:XPathAddAtt(	"/xdr:wsDr/xdr:"+cCellType+"[last()]/xdr:pic/xdr:nvPicPr/xdr:cNvPr","id", cValToChar(Len(::aImgdraw)+1) )
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
	PARAMTYPE 0	VAR cFile			AS CHARACTER
	PARAMTYPE 1	VAR nPlanilha	  	AS NUMERIC		OPTIONAL DEFAULT 1
	cFile	:= Alltrim(cFile)
	If !File(cFile,,.F.)
		ConOut("Arquivo nao encontrado!")
		Return .F.
	EndIf
	If ValType(cRootPath)=="U"
		cRootPath	:= GetSrvProfString( "RootPath", "" )
	EndIf
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
		EndIf
		If !FindFunction("FZIP")
			WaitRunSrv('"'+cAr7Zip+'" x -tzip "'+cCamSrv+'" -o"'+cRootPath+'\tmpxls\'+::cTmpFile+'\'+::cNomeFile+'" * -r -y',.T.,"C:\")
			If !File("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\sharedStrings.xml",,.F.)
				nRet	:= -1
				ConOut("Arquivo nao descompactado!")
				Return .F.
			Else
				nRet	:= 0
			EndIf
		Else
			nRet	:= FUnZip(cCamLocal,"\tmpxls\"+::cTmpFile+'\'+::cNomeFile+'\')
		EndIf
		If nRet!=0
			ConOut(Ferror())
			ConOut("Arquivo nao descompactado!")
			Return .F.
		EndIf
		oXml	:= TXmlManager():New()
		If oXML:ParseFile("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\sharedStrings.xml")
			oXML:XPathRegisterNs( "ns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main" )
			aChildren := oXML:XPathGetChildArray( "/ns:sst" )
			For nCont:=1 to Len(aChildren)
				::oString:Set(::nQtdString,oXML:XPathGetNodeValue("/ns:sst/ns:si["+cValToChar(nCont)+"]/ns:t"))
				::nQtdString++
			Next
		EndIf
	EndIf
	oXml	:= TXmlManager():New()
	oXML:ParseFile("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\worksheets\sheet"+cValTochar(nPlanilha)+".xml")
	aNs	:= oXML:XPathGetRootNsList()
	For nCont:=1 to Len(aNs)
		oXML:XPathRegisterNs( aNs[nCont][1], aNs[nCont][2] )
		If alltrim(lower(aNs[nCont][2]))==lower("http://schemas.openxmlformats.org/spreadsheetml/2006/main")
			cNomeNS	:= aNs[nCont][1]
		EndIf
	Next
	oXmlStyle	:= TXmlManager():New()
	oXmlStyle:ParseFile("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\styles.xml")
	aNs	:= oXmlStyle:XPathGetRootNsList()
	For nCont:=1 to Len(aNs)
		oXmlStyle:XPathRegisterNs( aNs[nCont][1], aNs[nCont][2] )
		If alltrim(lower(aNs[nCont][2]))==lower("http://schemas.openxmlformats.org/spreadsheetml/2006/main")
			cNomeNS2	:= aNs[nCont][1]
		EndIf
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
    		EndIf
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
//    			ElseIf aAtributos[nCont3][1]=="s" .and. aAtributos[nCont3][2]=="1"
//    				cTipo	:= "D"
//    				cRet	:= STOD("19000101")-2+Val(oXML:XPathGetNodeValue("/"+cNomeNS+":worksheet/"+cNomeNS+":sheetData/"+cNomeNS+":row["+cValToChar(nCont)+"]/"+cNomeNS+":c["+cValToChar(nCont2)+"]/"+cNomeNS+":v"))
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
    					EndIf
    				EndIf
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
Method CellRead(nLinha,nColuna,xDefault,lAchou,cOutro) Class YExcel
	Local cRef	:= ::Ref(nLinha,nColuna)
	Local xValor:= Nil
	Default cOutro	:= ""
	lAchou	:= .T.
	If !::oCell:Get(cRef+cOutro,@xValor)
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
	Local cID
	Local nFile,oSelection
	Local nQtdPlanilhas	:= Len(::aPlanilhas)
	Local nCont
	Local oCorPlan
	Private oSelf	:= Self
	PARAMTYPE 0	VAR cNome		  	AS CHARACTER		OPTIONAL DEFAULT "Planilha"+cValToChar(nQtdPlanilhas+1)
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
	EndIf
	cNome	:= EncodeUTF8(cNome)
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
		::xls_sheet(nFile)
		fClose(nFile)
		fErase("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\worksheets\tmprow.xml",,.F.)

		If ::nIdRelat>0
			For nCont:=1 to Len(::atable)
				::nCont	:= nCont
				::CriarFile("\"+::cNomeFile+"\xl\tables\"	,"table"+cValToChar(::atable[nCont]:nIdRelat)+".xml"		,::xls_table()		,)
				AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\tables\table"+cValToChar(::atable[nCont]:nIdRelat)+".xml")
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
	::osheetPr:AddValor(yExcelTag():New("pageSetUpPr",,{{"fitToPage","1"}}))	//Flag indicating whether the Fit to Page print option is enabled. pag 1675
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
	::odrawing		:= yExcelTag():New("drawing",)
	::adrawing		:= {}
	::aImgdraw		:= {}
	::nRowoutlineLevel	:= nil
	::lRowcollapsed		:= .F.
	::lRowHidden		:= .F.

	//Cria arquivo temporario de gravação das linhas
	::CriarFile("\"+::cNomeFile+"\xl\worksheets"	,"tmprow.xml"			,""			,)
	GravaFile(@::nFileTmpRow,"","\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\worksheets","tmprow.xml")
	//Cria nova planilha
	nQtdPlanilhas++
	//Adiciona dentro do workbooks o relacionamento na planilha
	cID	:= ::add_rels("\xl\_rels\workbook.xml.rels","http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet","worksheets/sheet"+cValToChar(nQtdPlanilhas)+".xml")

	AADD(::aPlanilhas,{cID,cNome,/*draw*/})

	::oworkbook:XPathAddNode( "/xmlns:workbook/xmlns:sheets", "sheet", "" )
	::oworkbook:XPathAddAtt( "/xmlns:workbook/xmlns:sheets/xmlns:sheet[last()]", "name"		, cNome)
	::oworkbook:XPathAddAtt( "/xmlns:workbook/xmlns:sheets/xmlns:sheet[last()]", "sheetId"	, cValToChar(nQtdPlanilhas))
	::oworkbook:XPathAddAtt( "/xmlns:workbook/xmlns:sheets/xmlns:sheet[last()]", "r:id"		, cID)

	//Adiciona um nova Planilha no content_types
	::ocontent_types:XPathAddNode( "/xmlns:Types", "Override", "" )
	::ocontent_types:XPathAddAtt( "/xmlns:Types/xmlns:Override[last()]", "PartName"	, "/xl/worksheets/sheet"+cValToChar(nQtdPlanilhas)+".xml" )
	::ocontent_types:XPathAddAtt( "/xmlns:Types/xmlns:Override[last()]", "ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" )
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
		UserException("YExcel - As linhas devem ser informadas sempre de forma sequencial crescente. Linha informada:"+cValToChar(nLinha)+" | Linha Atual:"+cValToChar(::adimension[2][1])+".")
	EndIf
	If ::adimension[1][1]==nLinha .AND. nColuna==::nColunaAtual
		UserException("YExcel - Não é possivel redefinir o valor da celula gravada.")
	ElseIf ::adimension[1][1]==nLinha .AND. nColuna<=::nColunaAtual
		UserException("YExcel - As colunas devem ser informadas sempre de forma sequencial crescente. Coluna informada:"+cValToChar(nColuna)+" | Coluna Atual:"+cValToChar(::nColunaAtual)+".")
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

/*/{Protheus.doc} showGridLines
Se vai exibir ou ocultar linhas de grade na planilha
@author Saulo Gomes Martins
@since 11/12/2019
@version 1.0
@param lView, logical, Se falso oculta linhas de grade
@type function
@obs pag 1709
/*/
METHOD showGridLines(lView) CLASS YExcel
	::osheetViews:GetValor():SetAtributo("showGridLines",If(lView,"1","0"))
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
@obs pag 1601 - 18.3.1.2
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

Return {nTamanho,cCorRGB,cNome,cfamily,cScheme,lNegrito,lItalico,lSublinhado,lTachado}

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
Return {cFgCor,cBgCor,cType}

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
Return {cTipo,cCor,cModelo}

/*/{Protheus.doc} ADDdxf
Cria estilo para formatação condicional
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param [aFont], array, objeto criado pelo metodo :Font() com fonte
@param [aCorPreenc], array, objeto com cor criado pelo metodo :Preench() de preenchimento
@param [aBorda], object, objeto criado pelo metodo :ObjBorda() com borda
@return posição do estilo
@type function
/*/
METHOD ADDdxf(aFont,aCorPreenc,aBorda) CLASS YExcel
	Local nTamdxfs

	::oStyle:XPathAddNode( "xmlns:styleSheet/xmlns:dxfs","dxf", "" )
	nTamdxfs	:= Val(::oStyle:XPathGetAtt("xmlns:styleSheet/xmlns:dxfs","count"))+1
	::oStyle:XPathSetAtt("xmlns:styleSheet/xmlns:dxfs","count",cValToChar(nTamdxfs))

	//Font
	If ValType(aFont)=="A"
		::AddFont(aFont[1],aFont[2],aFont[3],aFont[4],aFont[5],aFont[6],aFont[7],aFont[8],aFont[9],"xmlns:styleSheet/xmlns:dxfs/xmlns:dxf[last()]")
	EndIf
	//Preenchimento
	If ValType(aCorPreenc)=="A"
		::CorPreenc(aCorPreenc[1],aCorPreenc[2],aCorPreenc[3],"xmlns:styleSheet/xmlns:dxfs/xmlns:dxf[last()]")
	EndIf
	//Borda
	If ValType(aBorda)=="A"
		::Borda(aBorda[1],aBorda[2],aBorda[3],"xmlns:styleSheet/xmlns:dxfs/xmlns:dxf[last()]")
	EndIf
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
METHOD AddFont(nTamanho,cCorRGB,cNome,cfamily,cScheme,lNegrito,lItalico,lSublinhado,lTachado,cLocal) CLASS YExcel
	Local nTamFonts := 0
	PARAMTYPE 0	VAR nTamanho		AS NUMERIC				OPTIONAL DEFAULT 11
	PARAMTYPE 1	VAR cCorRGB			AS CHARACTER,NUMERIC	OPTIONAL DEFAULT "FF000000"
	PARAMTYPE 2	VAR cNome	  		AS CHARACTER			OPTIONAL DEFAULT "Calibri"
	PARAMTYPE 3	VAR cfamily	  		AS CHARACTER			OPTIONAL
	PARAMTYPE 4	VAR cScheme	  		AS CHARACTER			OPTIONAL
	PARAMTYPE 5	VAR lNegrito	  	AS LOGICAL				OPTIONAL DEFAULT .F.
	PARAMTYPE 6	VAR lItalico	  	AS LOGICAL				OPTIONAL DEFAULT .F.
	PARAMTYPE 7	VAR lSublinhado	  	AS LOGICAL				OPTIONAL DEFAULT .F.
	PARAMTYPE 8	VAR lTachado	  	AS LOGICAL				OPTIONAL DEFAULT .F.
	PARAMTYPE 9	VAR cLocal	  		AS CHARACTER			OPTIONAL DEFAULT "/xmlns:styleSheet/xmlns:fonts"	///xmlns:styleSheet/xmlns:dxfs/xmlns:dxf[last()]

	If ValType(cCorRGB)=="C" .and. Len(cCorRGB)==6
		cCorRGB	:= "FF"+cCorRGB
	EndIf
	If cLocal=="/xmlns:styleSheet/xmlns:fonts"
		nTamFonts	:= Val(::oStyle:XPathGetAtt(cLocal,"count"))+1
		::oStyle:XPathSetAtt(cLocal,"count",cValToChar(nTamFonts))
	EndIf
	::oStyle:XPathAddNode( cLocal, "font", "" )

	If lNegrito
		::oStyle:XPathAddNode( cLocal+"/xmlns:font[last()]", "b", "" )
	EndIf
	If lItalico
		::oStyle:XPathAddNode( cLocal+"/xmlns:font[last()]", "i", "" )
	EndIf
	If lTachado
		::oStyle:XPathAddNode( cLocal+"/xmlns:font[last()]", "strike", "" )
	EndIf
	If lSublinhado
		::oStyle:XPathAddNode( cLocal+"/xmlns:font[last()]", "u", "" )
	EndIf

	::oStyle:XPathAddNode( cLocal+"/xmlns:font[last()]", "sz", "" )
	::oStyle:XPathAddAtt( cLocal+"/xmlns:font[last()]/xmlns:sz", "val"	, cValToChar(nTamanho) )

	::oStyle:XPathAddNode( cLocal+"/xmlns:font[last()]", "color", "" )
	If ValType(cCorRGB)=="N"
		::oStyle:XPathAddAtt( cLocal+"/xmlns:font[last()]/xmlns:color", "indexed"	, cValToChar(cCorRGB) )
	Else
		If ValType(cCorRGB)=="C" .and. Len(cCorRGB)==6
			cCorRGB	:= "FF"+cCorRGB
		EndIf
		::oStyle:XPathAddAtt( cLocal+"/xmlns:font[last()]/xmlns:color", "rgb"	, cCorRGB )
	EndIf

	::oStyle:XPathAddNode( cLocal+"/xmlns:font[last()]", "name", "" )
	::oStyle:XPathAddAtt( cLocal+"/xmlns:font[last()]/xmlns:name", "val"	, cNome )

	If !Empty(cfamily)
		::oStyle:XPathAddNode( cLocal+"/xmlns:font[last()]", "family", "" )
		::oStyle:XPathAddAtt( cLocal+"/xmlns:font[last()]/xmlns:family", "val"	, cfamily )
	EndIf
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
	EndIf
return nTamFonts-1

/*/{Protheus.doc} CorPreenc
Adiciona cor de preenchimento para ser usado no estilo das células
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param [cBgCor], characters, Cor em Alpha+RGB do preenchimento
@param [cFgCor], characters, Cor em Aplha+RGB do fundo
@param [cType], characters, tipo de preenchimento(padrão solid)
@type function

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
METHOD CorPreenc(cFgCor,cBgCor,cType,cLocal) CLASS YExcel
	Local nPos
	Default cType	:= "solid"
	PARAMTYPE 3	VAR cLocal	  		AS CHARACTER			OPTIONAL DEFAULT "/xmlns:styleSheet/xmlns:fills"	///xmlns:styleSheet/xmlns:dxfs/xmlns:dxf[last()]

	::oStyle:XPathAddNode( cLocal, "fill", "" )
	nPos	:= Val(::oStyle:XPathGetAtt(cLocal,"count"))+1
	If cLocal=="/xmlns:styleSheet/xmlns:fills"
		::oStyle:XPathSetAtt(cLocal,"count",cValToChar(nPos))
	EndIf

	::oStyle:XPathAddNode( cLocal+"/xmlns:fill[last()]", "patternFill", "" )
	::oStyle:XPathAddAtt( cLocal+"/xmlns:fill[last()]/xmlns:patternFill", "patternType"	, cType )
	If cType != "none"
		::oStyle:XPathAddNode( cLocal+"/xmlns:fill[last()]/xmlns:patternFill", "fgColor", "" )
		If ValType(cFgCor)=="C"
			If Len(cFgCor)==6
				cFgCor	:= "FF"+cFgCor
			EndIf
			::oStyle:XPathAddAtt( cLocal+"/xmlns:fill[last()]/xmlns:patternFill/xmlns:fgColor", "rgb"	, cFgCor )
		Elseif ValType(cFgCor)=="N"
			::oStyle:XPathAddAtt( cLocal+"/xmlns:fill[last()]/xmlns:patternFill/xmlns:fgColor", "indexed"	, cValToChar(cFgCor) )
		Else
			::oStyle:XPathAddAtt( cLocal+"/xmlns:fill[last()]/xmlns:patternFill/xmlns:fgColor", "indexed"	, "64" )	//indexed="64" System Foreground n/a
		EndiF
		::oStyle:XPathAddNode( cLocal+"/xmlns:fill[last()]/xmlns:patternFill", "bgColor", "" )
		If ValType(cBgCor)=="C"
			If Len(cBgCor)==6
				cBgCor	:= "FF"+cBgCor
			EndIf
			::oStyle:XPathAddAtt( cLocal+"/xmlns:fill[last()]/xmlns:patternFill/xmlns:bgColor", "rgb"	, cBgCor )
		Elseif ValType(cBgCor)=="N"
			::oStyle:XPathAddAtt( cLocal+"/xmlns:fill[last()]/xmlns:patternFill/xmlns:bgColor", "indexed"	, cValToChar(cBgCor) )
		Else
			::oStyle:XPathAddAtt( cLocal+"/xmlns:fill[last()]/xmlns:patternFill/xmlns:bgColor", "indexed"	, "65" )	//indexed="65" System Background n/a 	pag:1775
		EndIf
	EndIf

Return nPos-1

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
METHOD EfeitoPreenc(nAngulo,aCores,ctype,nleft,nright,ntop,nbottom,cLocal) CLASS YExcel
	Local nPos
	PARAMTYPE 7	VAR cLocal	  		AS CHARACTER			OPTIONAL DEFAULT "/xmlns:styleSheet/xmlns:fills"	///xmlns:styleSheet/xmlns:dxfs/xmlns:dxf[last()]
	::oStyle:XPathAddNode( cLocal, "fill", "" )
	nPos	:= Val(::oStyle:XPathGetAtt(cLocal,"count"))+1
	If cLocal=="/xmlns:styleSheet/xmlns:fills"
		::oStyle:XPathSetAtt(cLocal,"count",cValToChar(nPos))
	EndIf
	::gradientFill(nAngulo,aCores,ctype,nleft,nright,ntop,nbottom,cLocal+"/xmlns:fill[last()]")
Return nPos-1

METHOD gradientFill(nAngulo,aCores,ctype,nleft,nright,ntop,nbottom,cLocal) CLASS YExcel	//Pag 1779
	Local nCont
	PARAMTYPE 0	VAR nAngulo 		AS NUMERIC		OPTIONAL
	PARAMTYPE 1	VAR aCores	 		AS ARRAY
	PARAMTYPE 2	VAR ctype	 		AS CHARACTER	OPTIONAL
	PARAMTYPE 3	VAR nleft	 		AS NUMERIC		OPTIONAL
	PARAMTYPE 4	VAR nright	 		AS NUMERIC		OPTIONAL
	PARAMTYPE 5	VAR ntop	 		AS NUMERIC		OPTIONAL
	PARAMTYPE 6	VAR nbottom	 		AS NUMERIC		OPTIONAL
	PARAMTYPE 7	VAR cLocal	  		AS CHARACTER			OPTIONAL DEFAULT "/xmlns:styleSheet/xmlns:fills/xmlns:fill[last()]"	///xmlns:styleSheet/xmlns:dxfs/xmlns:dxf[last()]

	::oStyle:XPathAddNode( cLocal, "gradientFill", "" )

	If ValType(ctype)!="U" .and. !(ctype $ "path|linear")
		UserException("YExcel - Tipo invalido para efeito de preenchimento.(path|linear)")
	EndIf

	If ValType(ctype)!="U" .and. ctype=="path"
		::oStyle:XPathAddAtt( cLocal+"/xmlns:gradientFill[last()]", "type"	, ctype )
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
		::oStyle:XPathAddAtt( cLocal+"/xmlns:gradientFill[last()]", "left"		, cValToChar(nleft) )
		::oStyle:XPathAddAtt( cLocal+"/xmlns:gradientFill[last()]", "right"		, cValToChar(nright) )
		::oStyle:XPathAddAtt( cLocal+"/xmlns:gradientFill[last()]", "top"		, cValToChar(ntop) )
		::oStyle:XPathAddAtt( cLocal+"/xmlns:gradientFill[last()]", "bottom"	, cValToChar(nbottom) )
	Else
		Default nAngulo	:= 90
		::oStyle:XPathAddAtt( cLocal+"/xmlns:gradientFill[last()]", "degree"	, cValToChar(nAngulo) )
	EndIf
	For nCont:=1 to Len(aCores)
		If !(aCores[nCont][2]>=0 .and. aCores[nCont][2]<=1)
			UserException("YExcel - Definição de cor varia de 0 a 1. Valor informado:"+cValToChar(aCores[nCont][2]))
		EndIf
		If Len(aCores[nCont][1])==6
			aCores[nCont][1]	:= "FF"+aCores[nCont][1]
		EndIf
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

@type function
@Obs pode juntar os tipo. Exemplo "ED"-Esquerda e direita

/*/
METHOD Borda(cTipo,cCor,cModelo,cLocal) CLASS YExcel
	Local nPos
	Local cLeft,cRight,cTop,cBottom,cDiagonal
	PARAMTYPE 0	VAR cTipo	  		AS CHARACTER			OPTIONAL DEFAULT ""
	PARAMTYPE 1	VAR cCor	  		AS CHARACTER			OPTIONAL DEFAULT "FF000000"
	PARAMTYPE 2	VAR cModelo	  		AS CHARACTER			OPTIONAL DEFAULT "thin"
	PARAMTYPE 3	VAR cLocal	  		AS CHARACTER			OPTIONAL DEFAULT "/xmlns:styleSheet/xmlns:borders"	///xmlns:styleSheet/xmlns:dxfs/xmlns:dxf[last()]
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
		nPos	:= ::Border(cModelo,cModelo,cModelo,cModelo,,cCor,cCor,cCor,cCor,,cLocal)
	Else
		nPos	:= ::Border(cLeft,cRight,cTop,cBottom,cDiagonal,cCor,cCor,cCor,cCor,cCor,cLocal)
	EndIf
Return nPos-1

METHOD Border(cleft,cright,ctop,cbottom,cdiagonal,cCorleft,cCorright,cCortop,cCorbottom,cCordiagonal,cLocal) CLASS YExcel
	Local nPos
	PARAMTYPE 10	VAR cLocal	  		AS CHARACTER			OPTIONAL DEFAULT "/xmlns:styleSheet/xmlns:borders"	///xmlns:styleSheet/xmlns:dxfs/xmlns:dxf[last()]

	::oStyle:XPathAddNode( cLocal, "border", "" )
	nPos	:= Val(::oStyle:XPathGetAtt(cLocal,"count"))+1
	If cLocal=="/xmlns:styleSheet/xmlns:borders"
		::oStyle:XPathSetAtt(cLocal,"count",cValToChar(nPos))
	EndIf

	::oStyle:XPathAddNode( cLocal+"/xmlns:border[last()]", "left", "" )
	If ValType(cleft)<>"U"
		::oStyle:XPathAddAtt( cLocal+"/xmlns:border[last()]/xmlns:left", "style"	, cleft )
		If ValType(cCorleft)<>"U"
			::oStyle:XPathAddNode( cLocal+"/xmlns:border[last()]/xmlns:left", "color", "" )
			If ValType(cCorleft)=="C"
				::oStyle:XPathAddAtt( cLocal+"/xmlns:border[last()]/xmlns:left/xmlns:color", "rgb"	,cCorleft )
			ElseIf ValType(cCorleft)=="N"
				::oStyle:XPathAddAtt( cLocal+"/xmlns:border[last()]/xmlns:left/xmlns:color", "indexed"	,cValToChar(cCorleft) )
			EndIf
		EndIf
	EndIf

	::oStyle:XPathAddNode( cLocal+"/xmlns:border[last()]", "right", "" )
	If ValType(cright)<>"U"
		::oStyle:XPathAddAtt( cLocal+"/xmlns:border[last()]/xmlns:right", "style"	, cright )
		If ValType(cCorright)<>"U"
			::oStyle:XPathAddNode( cLocal+"/xmlns:border[last()]/xmlns:right", "color", "" )
			If ValType(cCorright)=="C"
				::oStyle:XPathAddAtt( cLocal+"/xmlns:border[last()]/xmlns:right/xmlns:color", "rgb"	,cCorright )
			ElseIf ValType(cCorright)=="N"
				::oStyle:XPathAddAtt( cLocal+"/xmlns:border[last()]/xmlns:right/xmlns:color", "indexed"	,cValToChar(cCorright) )
			EndIf
		EndIf
	EndIf

	::oStyle:XPathAddNode( cLocal+"/xmlns:border[last()]", "top", "" )
	If ValType(ctop)<>"U"
		::oStyle:XPathAddAtt( cLocal+"/xmlns:border[last()]/xmlns:top", "style"	, ctop )
		If ValType(cCortop)<>"U"
			::oStyle:XPathAddNode( cLocal+"/xmlns:border[last()]/xmlns:top", "color", "" )
			If ValType(cCortop)=="C"
				::oStyle:XPathAddAtt( cLocal+"/xmlns:border[last()]/xmlns:top/xmlns:color", "rgb"	,cCortop )
			ElseIf ValType(cCortop)=="N"
				::oStyle:XPathAddAtt( cLocal+"/xmlns:border[last()]/xmlns:top/xmlns:color", "indexed"	,cValToChar(cCortop) )
			EndIf
		EndIf
	EndIf

	::oStyle:XPathAddNode( cLocal+"/xmlns:border[last()]", "bottom", "" )
	If ValType(cbottom)<>"U"
		::oStyle:XPathAddAtt( cLocal+"/xmlns:border[last()]/xmlns:bottom", "style"	, cbottom )
		If ValType(cCorbottom)<>"U"
			::oStyle:XPathAddNode( cLocal+"/xmlns:border[last()]/xmlns:bottom", "color", "" )
			If ValType(cCorbottom)=="C"
				::oStyle:XPathAddAtt( cLocal+"/xmlns:border[last()]/xmlns:bottom/xmlns:color", "rgb"	,cCorbottom )
			ElseIf ValType(cCorbottom)=="N"
				::oStyle:XPathAddAtt( cLocal+"/xmlns:border[last()]/xmlns:bottom/xmlns:color", "indexed"	,cValToChar(cCorbottom) )
			EndIf
		EndIf
	EndIf

	::oStyle:XPathAddNode( cLocal+"/xmlns:border[last()]", "diagonal", "" )
	If ValType(cdiagonal)<>"U"
		::oStyle:XPathAddAtt( cLocal+"/xmlns:border[last()]/xmlns:diagonal", "style"	, cdiagonal )
		If ValType(cCordiagonal)<>"U"
			::oStyle:XPathAddNode( cLocal+"/xmlns:border[last()]/xmlns:diagonal", "color", "" )
			If ValType(cCordiagonal)=="C"
				::oStyle:XPathAddAtt( cLocal+"/xmlns:border[last()]/xmlns:diagonal/xmlns:color", "rgb"	,cCordiagonal )
			ElseIf ValType(cCordiagonal)=="N"
				::oStyle:XPathAddAtt( cLocal+"/xmlns:border[last()]/xmlns:diagonal/xmlns:color", "indexed"	,cValToChar(cCordiagonal) )
			EndIf
		EndIf
	EndIf
Return nPos

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
	Local cformatCode
	Local cDecimal
	Local cNumero	:= ""
	Local cNegINIAli:= ""
	Local cNegFIMAli:= ""
	Local nPosCor
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
		cPrefixo	:= '"'+cPrefixo+'"'
		cNumero		:= cPrefixo+cNumero
	EndIf
	If !Empty(cSufixo)
		cSufixo		:= '"'+cSufixo+'"'
		cNumero		:= cNumero+cSufixo
	EndIf
	If !Empty(cNegINI)
		cNegINIAli	:= "_"+cNegINI
	EndIf
	If !Empty(cNegFIM)
		cNegFIMAli	:= "_"+cNegFIM
	EndIf
	If !Empty(cValorZero)
		cValorZero	:= '"'+cValorZero+'"'
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


	If Empty(nNumFmtId)
		nNumFmtId	:= ::nNumFmtId++
	EndIf
	If !::oStyle:XPathHasNode( "/xmlns:styleSheet/xmlns:numFmts/numFmt[@numFmtId='"+cValToChar(nNumFmtId)+"']")	//Se não existe o ID
		::oStyle:XPathAddNode( "/xmlns:styleSheet/xmlns:numFmts", "numFmt", "" )
		::oStyle:XPathAddAtt( "/xmlns:styleSheet/xmlns:numFmts/xmlns:numFmt[last()]", "numFmtId"	, cValToChar(nNumFmtId) )
		::oStyle:XPathSetAtt("/xmlns:styleSheet/xmlns:numFmts","count",cValToChar(Val(::oStyle:XPathGetAtt("/xmlns:styleSheet/xmlns:numFmts","count"))+1))
	EndIf
	::oStyle:XPathAddAtt( "/xmlns:styleSheet/xmlns:numFmts/xmlns:numFmt[@numFmtId='"+cValToChar(nNumFmtId)+"']", "formatCode"	, cformatCode )

Return nNumFmtId	//Não retorna a posição, mas o atributo numFmtId

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
	Local cLocal	:= "/xmlns:styleSheet/xmlns:cellXfs"
	Local aListAtt
	Local nCont,nCont2
	Local nPos
	PARAMTYPE 0	VAR numFmtId			AS NUMERIC 		OPTIONAL
	PARAMTYPE 1	VAR fontId				AS NUMERIC 		OPTIONAL
	PARAMTYPE 2	VAR fillId				AS NUMERIC 		OPTIONAL
	PARAMTYPE 3	VAR borderId			AS NUMERIC 		OPTIONAL
	PARAMTYPE 4	VAR xfId  				AS NUMERIC 		OPTIONAL DEFAULT 0
	PARAMTYPE 5	VAR aValores  			AS ARRAY 		OPTIONAL DEFAULT {}
	PARAMTYPE 6	VAR aOutrosAtributos	AS ARRAY 		OPTIONAL DEFAULT {}
	If ValType(fontId)=="N" .AND. (fontId+1)>Val(::oStyle:XPathGetAtt("/xmlns:styleSheet/xmlns:fonts","count"))
		UserException("YExcel - Fonte informada("+cValToChar(fontId)+") não definido. Utilize o indice informado pelo metodo :AddFont()")
	ElseIf ValType(fillId)=="N" .AND. (fillId+1)>Val(::oStyle:XPathGetAtt("/xmlns:styleSheet/xmlns:fills","count"))
		UserException("YExcel - Cor Preenchimento informado("+cValToChar(fillId)+") não definido. Utilize o indice informado pelo metodo :CorPreenc()")
	ElseIf ValType(borderId)=="N" .AND. (borderId+1)>Val(::oStyle:XPathGetAtt("/xmlns:styleSheet/xmlns:borders","count"))
		UserException("YExcel - Borda informada("+cValToChar(borderId)+") não definido. Utilize o indice informado pelo metodo :Borda()")
	EndIf

	::oStyle:XPathAddNode( cLocal, "xf", "" )
	nPos	:= Val(::oStyle:XPathGetAtt(cLocal,"count"))+1
	::oStyle:XPathSetAtt(cLocal,"count",cValToChar(nPos))

	If ValType(numFmtId)<>"U"
		::oStyle:XPathAddAtt( cLocal+"/xmlns:xf[last()]", "numFmtId"			, cValToChar(numFmtId) )
		::oStyle:XPathAddAtt( cLocal+"/xmlns:xf[last()]", "applyNumberFormat"	, "1" )
	Else
		::oStyle:XPathAddAtt( cLocal+"/xmlns:xf[last()]", "numFmtId"			, "0" )
	EndIf

	If ValType(fontId)<>"U"
		::oStyle:XPathAddAtt( cLocal+"/xmlns:xf[last()]", "fontId"				, cValToChar(fontId) )
		::oStyle:XPathAddAtt( cLocal+"/xmlns:xf[last()]", "applyFont"			, "1" )
	Else
		::oStyle:XPathAddAtt( cLocal+"/xmlns:xf[last()]", "fontId"				, "0" )
	EndIf

	If ValType(fillId)<>"U"
		::oStyle:XPathAddAtt( cLocal+"/xmlns:xf[last()]", "fillId"				, cValToChar(fillId) )
		::oStyle:XPathAddAtt( cLocal+"/xmlns:xf[last()]", "applyFill"			, "1" )
	Endif
	If ValType(borderId)<>"U"
		::oStyle:XPathAddAtt( cLocal+"/xmlns:xf[last()]", "borderId"			, cValToChar(borderId) )
		::oStyle:XPathAddAtt( cLocal+"/xmlns:xf[last()]", "applyBorder"			, "1" )
	Else
		::oStyle:XPathAddAtt( cLocal+"/xmlns:xf[last()]", "borderId"			, "0" )
	EndIf

	::oStyle:XPathAddAtt( cLocal+"/xmlns:xf[last()]", "xfId"					, cValToChar(xfId) )

	For nCont:=1 to Len(aValores)
		::oStyle:XPathAddNode( cLocal+"/xmlns:xf[last()]", aValores[nCont]:GetNome(), "" )
		If aValores[nCont]:GetNome()=="alignment"
			If aScan(self:oStyle:XPathGetAttArray(cLocal+"/xmlns:xf[last()]"),{|x| x[1]=="applyAlignment"})>0
				::oStyle:XPathSetAtt( cLocal+"/xmlns:xf[last()]", "applyAlignment"		, "1" )
			Else
				::oStyle:XPathAddAtt( cLocal+"/xmlns:xf[last()]", "applyAlignment"		, "1" )
			EndIf
		EndIf
		aValores[nCont]:oAtributos:List(@aListAtt)
		For nCont2:=1 to Len(aListAtt)
			If aScan(self:oStyle:XPathGetAttArray(cLocal+"/xmlns:xf[last()]/xmlns:"+aValores[nCont]:GetNome()),{|x| x[1]==aListAtt[nCont2][1] })>0
				::oStyle:XPathSetAtt( cLocal+"/xmlns:xf[last()]/xmlns:"+aValores[nCont]:GetNome(), aListAtt[nCont2][1]			, cValToChar(aListAtt[nCont2][2]) )
			Else
				::oStyle:XPathAddAtt( cLocal+"/xmlns:xf[last()]/xmlns:"+aValores[nCont]:GetNome(), aListAtt[nCont2][1]			, cValToChar(aListAtt[nCont2][2]) )
			EndIf
		Next
	Next

	For nCont:=1 to Len(aOutrosAtributos)
		::oStyle:XPathSetAtt( cLocal+"/xmlns:xf[last()]", aOutrosAtributos[nCont][1]	, cValToChar(aOutrosAtributos[nCont][2]) )
	Next
return nPos-1


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

/*METHOD Addhyperlink(nLinha,nColuna,cLocation,cId,ctooltip,cDisplay) CLASS YExcel
	PARAMTYPE 0	VAR cRef			AS CHARACTER 		OPTIONAL
	PARAMTYPE 1	VAR cLocation		AS CHARACTER 		OPTIONAL
	PARAMTYPE 2	VAR cId				AS CHARACTER 		OPTIONAL
	PARAMTYPE 3	VAR ctooltip		AS CHARACTER 		OPTIONAL
	PARAMTYPE 4	VAR cDisplay  		AS CHARACTER 		OPTIONAL
Return*/

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
	Local cRet		:= ""
	Default llinha	:= .F.
	Default lColuna	:= .F.
	If llinha
		cLinha	:= "$"
	EndIf
	If lColuna
		cColuna	:= "$"
	EndIf
	If ValType(nColuna)!="U"
		cRet	+= cColuna+NumToString(nColuna)
	EndIf
	If ValType(nLinha)!="U"
		cRet	+= cLinha+cValToChar(nLinha)
	EndIf
Return cRet


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
Return {Val(cLinha),If(!Empty(cColuna),::StringToNum(cColuna),0)}


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
METHOD AddTabela(cNome,nLinha,nColuna,nQtdPlan) CLASS YExcel
	Local nPos
	Local oTable
	Local cID
	Default nQtdPlan	:= Len(::aPlanilhas)
	PARAMTYPE 0	VAR cNome  AS CHARACTER 		OPTIONAL DEFAULT lower(CriaTrab(,.F.))
	PARAMTYPE 1	VAR nLinha  AS NUMERIC 			OPTIONAL DEFAULT ::adimension[2][1]
	PARAMTYPE 2	VAR nColuna  AS NUMERIC
	::nIdRelat++
	nPos	:= ::nIdRelat

	oTable	:= yExcel_Table():New(self,nLinha,nColuna,cNome) //yExcelTag():New("table",{},)
	oTable:nIdRelat	:= nPos
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

	cID		:= ::add_rels("\xl\worksheets\_rels\sheet"+cValToChar(nQtdPlan)+".xml.rels","http://schemas.openxmlformats.org/officeDocument/2006/relationships/table","../tables/table"+cValToChar(oTable:nIdRelat)+".xml")
	::otableParts:AddValor(yExcelTag():New("tablePart",nil,{{"r:id",cID}}))
	::otableParts:SetAtributo("count",Len(::atable)+1)

	//Adiciona um nova Tabela
	::ocontent_types:XPathAddNode( "/xmlns:Types", "Override", "" )
	::ocontent_types:XPathAddAtt( "/xmlns:Types/xmlns:Override[last()]", "PartName"	, "/xl/tables/table"+cValToChar(oTable:nIdRelat)+".xml" )
	::ocontent_types:XPathAddAtt( "/xmlns:Types/xmlns:Override[last()]", "ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml" )
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
	Local cPath
	Local cDrive,cNome,cExtensao
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

	FWMakeDir("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\")
	FWMakeDir("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\docprops")
	FWMakeDir("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl")
	::ocontent_types:Save2File("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\[content_types].xml")
	::oapp:Save2File("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\docprops\app.xml")
	::ocore:Save2File("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\docprops\core.xml")
	::oworkbook:Save2File("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\workbook.xml")
	::oStyle:Save2File("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\styles.xml")

	For nCont:=1 to Len(::aRels)
		If !Empty(::aRels[nCont][3])
			FWMakeDir("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+SubStr(::aRels[nCont][2],1,rAt("\",::aRels[nCont][2])-1),.F.)
			::aRels[nCont][1]:Save2File("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+::aRels[nCont][2])
			AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+::aRels[nCont][2])
		EndIf
	Next
	For nCont:=1 to Len(::aDraw)
		If !Empty(::aDraw[nCont][3])
			FWMakeDir("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+SubStr(::aDraw[nCont][2],1,rAt("\",::aDraw[nCont][2])-1),.F.)
			::aDraw[nCont][1]:Save2File("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+::aDraw[nCont][2])
			AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+::aDraw[nCont][2])
		EndIf
	Next

	::CriarFile("\"+::cNomeFile+"\xl"				,"sharedStrings.xml"	,""						,)
	GravaFile(@nFile,"","\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl","sharedStrings.xml")
	::xls_sharedStrings(nFile)
	fClose(nFile)
	nFile	:= nil

	::CriarFile("\"+::cNomeFile+"\xl\theme"			,"theme1.xml"			,u_yxlsthem()			,)

	nQtdPlanilhas	:= Len(::aPlanilhas)
	::CriarFile("\"+::cNomeFile+"\xl\worksheets"	,"sheet"+cValToChar(nQtdPlanilhas)+".xml"			,""			,)
	GravaFile(@nFile,"","\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\worksheets","sheet"+cValToChar(nQtdPlanilhas)+".xml")
	::xls_sheet(nFile)
	fClose(nFile)
	fErase("\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\worksheets\tmprow.xml",,.F.)

	If ::nIdRelat>0
		For nCont:=1 to Len(::atable)
			::nCont	:= nCont
			::CriarFile("\"+::cNomeFile+"\xl\tables\"	,"table"+cValToChar(::atable[nCont]:nIdRelat)+".xml"		,::xls_table()		,)
			AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\tables\table"+cValToChar(::atable[nCont]:nIdRelat)+".xml")
		Next
	EndIf

	If lServidor
		cArquivo	:= cLocal+'\'+::cNomeFile+'.xlsx'
		cLocal		:= ""
	Else
		cArquivo	:= '\tmpxls\'+::cTmpFile+'\'+::cNomeFile+'.xlsx'
	EndIf
	SplitPath(cArquivo,@cDrive,@cPath,@cNome,@cExtensao)
	cNome	:= SubStr(cArquivo,Rat("\",cArquivo)+1)	//Split não está respeitando o case original
	If !Empty(cPath)
		FWMakeDir(cPath,.F.)	//Cria a estrutura de pastas
	EndIF

	If !FindFunction("FZIP")
		WaitRunSrv('"'+cAr7Zip+'" a -tzip "'+cRootPath+cArquivo+'" "'+cRootPath+'\tmpxls\'+::cTmpFile+'\'+::cNomeFile+'\*"',.T.,"C:\")
	Else
		fZip(cArquivo,::aFiles,"\tmpxls\"+::cTmpFile+'\'+::cNomeFile+'\')
	EndIf

//	For nCont:=1 to Len(::aFiles)
//		If fErase(::aFiles[nCont],,.F.)<>0
//			ConOut(::aFiles[nCont])
//			ConOut("Ferror:"+cValToChar(ferror()))
//		EndIf
//	Next

	DelPasta("\tmpxls\"+::cTmpFile+"\"+::cNomeFile)	//Apaga arquivos temporarios
	If substr(cArquivo,1,8)<>"\tmpxls\"
		DelPasta("\tmpxls\"+::cTmpFile)
	EndIf
	If !Empty(cLocal)
		If GetRemoteType() == REMOTE_HTML
			CpyS2TW(cArquivo, .T.)
		Else
			FWMakeDir(cLocal,.F.)
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
	Local aFiles	:= Directory(cCaminho+"\*","HSD",,.F.)
	For nCont:=1 to Len(aFiles)
		If aFiles[nCont][1]=="." .or. aFiles[nCont][1]==".."
			Loop
		EndIf
		If aFiles[nCont][5] $ "D"
			DelPasta(cCaminho+"\"+aFiles[nCont][1])
		Else
//			ConOut("Deletando:"+cCaminho+"\"+aFiles[nCont][1])
			If fErase(cCaminho+"\"+aFiles[nCont][1],,.F.)<>0
				ConOut(cCaminho+"\"+aFiles[nCont][1])
				ConOut("Ferror:"+cValToChar(ferror()))
			EndIf
		EndIf
	Next
//	ConOut("Apagando pasta:"+cCaminho)
	If !DirRemove(cCaminho,,.F.)
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
	//oFile	:= FWFileIOBase():New(cDirServ+cLocal+"\"+cNome)
	//oFile:SetCaseSensitive()
	If !File(cDirServ+cLocal+"\"+cNome,,.F.)
		nFile	:= FCreate(cDirServ+cLocal+"\"+cNome, , , .F.)
//		oFile:Create()
	Else
		fErase(cDirServ+cLocal+"\"+cNome,,.F.)
		nFile	:= FCreate(cDirServ+cLocal+"\"+cNome, , , .F.)
//		oFile:Create()
	EndIf
	FClose(nFile)
//	oFile:Close()
	nFile	:= FOPEN(cDirServ+cLocal+"\"+cNome, FO_READWRITE,,.F.)
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
		nFile	:= FOPEN(cLocal+"\"+cArquivo, FO_READWRITE,,.F.)
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
	f	:= replace(f,"<","&lt;")
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
		::SetAtributo("s","1")		//Adiciona o estilo padrão de data
		//::SetAtributo("t","d")	//Adiciona o estilo padrão de data
		//::SetV(SUBSTR(DTOS(v),1,4)+"-"+SUBSTR(DTOS(v),5,2)+"-"+SUBSTR(DTOS(v),7,2))
		If !Empty(v)
			::SetV(v-STOD("19000101")+2)
		Else
			::SetV(" ")
		EndIf
	ElseIf cTipo=="O" .and. GetClassName(v)=="YEXCEL_DATETIME"
		::SetAtributo("s","2")			//Adiciona o estilo padrão de data time
		::SetV(v:GetStrNumber())
	Else
		::SetV(v)
	EndIf
	If ValType(nStyle)=="N"
		If nStyle+1>Val(::oyExcel:oStyle:XPathGetAtt("/xmlns:styleSheet/xmlns:cellXfs","count"))
			UserException("YExcel - Estilo informado("+cValToChar(nStyle)+") não definido. Utilize o indice informado pelo metodo AddStyles")
		Else
			::SetAtributo("s",nStyle)
		EndIf
	EndIf
Return self

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
Class yExcelTag
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
	Local otableColumn
//	Local nPosCol		:= aScan(self:GetValor(),{|x| x:GetNome()=="tableColumns"})
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
	Local nPosCol
	Local cRef
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

/*/{Protheus.doc} GetDateTime
Retorna objeto para manipulação de DateTime
@author Saulo Gomes Martins
@since 09/12/2019
@version 1.0
@param dData, date, Data para formatação
@param cTime, characters, Hora para formatação
@param nData, numeric, DataTime em formato numerico
@type function
/*/
METHOD GetDateTime(dData,cTime,nData) Class yExcel
Return yExcel_DateTime():New(dData,cTime,nData)

/*/{Protheus.doc} yExcel_DateTime
Classe yExcel_DateTime para manipulação de DateTime
@author Saulo Gomes Martins
@since 10/12/2019
@version 1.0

@type class
/*/
Class yExcel_DateTime
	Data dData
	Data cTime
	Data cNumero
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
@return self, objeto
@param dData, date, Data para iniciar o objeto
@param cTime, characters, Hora para iniciar o objeto
@param [nData], numeric, Data e hora para iniciar o objeto
@type function
@obs enviar dData e cTime ou somente nData
/*/
Method New(dData,cTime,nData) class yExcel_DateTime
	::dData	:= dData
	::cTime	:= cTime
	::cClassName	:= "YEXCEL_DATETIME"
	::cName			:= "YEXCEL_DATETIME"
	If ValType(::dData)=="D" .AND. ValType(cTime)=="C"
		::StrNumber()
	ElseIf ValType(nData)=="N" .OR. ValType(nData)=="C"
		::NumToDateTime(nData)
	EndIf
Return Self

/*/{Protheus.doc} yExcel_DateTime:NumToDateTime
Converte numero do excel em data e hora
@author Saulo Gomes Martins
@since 10/12/2019
@version 1.0
@param nData, numeric, numero da hora, aceita também string
@type function
/*/
Method NumToDateTime(nData) Class yExcel_DateTime
	Local nInt
	Local nDec
	Local nHora
	Local nMinuto
	Local nSegundo
	If ValType(nData)=="N"
		nInt	:= Int(nData)
		nDec	:= nData-nInt
		::cNumero	:= cValToChar(nData)
	Else
		nPosPonto	:= At(".",nData)
		If nPosPonto==0
			nPosPonto	:= At(",",nData)
		EndIf
		If nPosPonto==0
			nInt		:= Val(nData)
			nDec		:= 0
		Else
			nInt	:= Val(SubStr(nData,1,nPosPonto-1))
			nDec	:= Val("0."+SubStr(nData,nPosPonto+1))
		EndIf
		::cNumero	:= nData
	EndIf
	::dData	:= STOD("19000101")-2+nInt
	::cTime	:= ""
	nHora	:= Int(nDec*86400/60/60)
	nMinuto	:= Int(((nDec*86400/60/60)-nHora)*60)
	nSegundo:= Round(((nDec*86400/60/60)-nHora-(nMinuto/60))*60*60,0)
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

@type function
/*/
Method GetStrNumber() Class yExcel_DateTime
Return ::cNumero

/*/{Protheus.doc} yExcel_DateTime:GetDate
Retorna a data
@author Saulo Gomes Martins
@since 10/12/2019
@version 1.0

@type function
/*/
Method GetDate() Class yExcel_DateTime
Return ::dData

/*/{Protheus.doc} yExcel_DateTime:GetTime
Retorna a Hora no formato HH:MM:SS
@author Saulo Gomes Martins
@since 10/12/2019
@version 1.0

@type function
/*/
Method GetTime() Class yExcel_DateTime
Return ::cTime

/*/{Protheus.doc} yExcel_DateTime:StrNumber
Converte data e hora em string com numero representando data e hora do excel
@author Saulo Gomes Martins
@since 10/12/2019
@version 1.0

@type function
/*/
Method StrNumber() Class yExcel_DateTime
	Local nHora	:= 0
	Local aHora	:= SeparaHora(::cTime)
	nHora		+= (aHora[1]*100000000)				//Hora
	nHora		+= (aHora[2]*100000000)/60			//Minuto
	nHora		+= (aHora[3]*100000000)/60/60		//Segundo
	nHora		+= (aHora[4]*100000000)/60/60/1000	//Milesimo
	cNum:= Replace(cValToChar(nHora/24),"0.","")
	If (At(".",cNum)-1)>0
		cNum:= SubStr(cNum,1,At(".",cNum)-1)
	Else
		cNum:= SubStr(cNum,1,Len(cNum))
	EndIf
	::cNumero	:= cValToChar(::dData-STOD("19000101")+2)+"."+cNum
Return ::cNumero

METHOD ClassName() CLASS yExcel_DateTime
Return "YEXCEL_DATETIME"

/*/{Protheus.doc} SeparaHora
Retorna Hora,Minuto,Segundo,Milesimo.
@author Saulo Gomes Martins
@since 09/12/2019
@version 1.0
@return aHora, 1-Hora|2-Munuto|3-Segundo|4-Milésimo de segundo
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
			EndIf
		EndIf
	EndIf
Return {nHoras,nMinutos,nSegundos,nMilesimo}

/*/{Protheus.doc} new_content_types
Criação do arquivo \[content_types].xml
@author Saulo Gomes Martins
@since 10/12/2019
@version 1.0
@param cXml, characters, xml para criação
@type function
/*/
Method new_content_types(cXml) class YExcel
	Local nCont
	Local aNs
	Default cXml			:= ""
	::ocontent_types	:= TXMLManager():New()
	If Empty(cXml)	//Cria modelo em branco
		cXml	+= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
		cXml	+= '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
		cXml	+= '	<Default Extension="jpg" ContentType="image/jpeg"/>'
		cXml	+= '	<Default Extension="png" ContentType="image/png"/>'
		cXml	+= '	<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
		cXml	+= '	<Default Extension="xml" ContentType="application/xml"/>'
		cXml	+= '	<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
		cXml	+= '	<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
		cXml	+= '	<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
		cXml	+= '	<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
		cXml	+= '	<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
		cXml	+= '	<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
		cXml	+= '</Types>'
	EndIf
	::ocontent_types:Parse(cXml)
	aNs	:= ::ocontent_types:XPathGetRootNsList()
	For nCont:=1 to Len(aNs)
		::ocontent_types:XPathRegisterNs( aNs[nCont][1], aNs[nCont][2] )
	Next
	AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\[content_types].xml")
Return

/*/{Protheus.doc} new_rels
Cria arquivo de relacionamento Relationship(rels)
@author Saulo Gomes Martins
@since 10/12/2019
@version 1.0
@return nPos, Posição no array
@param cXml, characters, xml para criação
@param cCaminho, characters, caminho do arquivo
@type function
/*/
Method new_rels(cXml,cCaminho) class YExcel
	Local nCont
	Local aNs
	Local oXML
	Default cXml			:= ""
	oXML	:= TXMLManager():New()
	If Empty(cXml)	//Cria modelo em branco
		cXml	+= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
		cXml	+= '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
		cXml	+= '</Relationships>'
	EndIf
	oXML:Parse(cXml)
	aNs	:= oXML:XPathGetRootNsList()
	For nCont:=1 to Len(aNs)
		oXML:XPathRegisterNs( aNs[nCont][1], aNs[nCont][2] )
	Next
	AADD(::aRels,{oXML,cCaminho,0})
Return Len(::aRels)

/*/{Protheus.doc} add_rels
Adiciona node no arquivo Relationship(rels)
@author Saulo Gomes Martins
@since 10/12/2019
@version 1.0
@return cId, rId criado
@param cCaminho, characters, caminho do arquivo de rel para gravar
@param cType, characters, atributo Type
@param cTarget, characters, atributo Target
@type function
/*/
Method add_rels(cCaminho,cType,cTarget) class YExcel
	Local nPos
	Local cId
	If ValType(cCaminho)=="N"
		nPos	:= cCaminho
	ElseIf ValType(cCaminho)=="C"
		If SubStr(cCaminho,1,1)!="\"
			cCaminho	:= "\"+cCaminho
		EndIf
		nPos	:= aScan(::aRels,{|x| x[2]==cCaminho })
	EndIf
	If nPos==0
		nPos	:= ::new_rels(,cCaminho)
	EndIf
	::aRels[nPos][3]++
	cId	:= "rId"+cValToChar(::aRels[nPos][3])
	::aRels[nPos][1]:XPathAddNode( "/xmlns:Relationships", "Relationship", "" )
	::aRels[nPos][1]:XPathAddAtt( "/xmlns:Relationships/xmlns:Relationship[last()]", "Type"		, cType )
	::aRels[nPos][1]:XPathAddAtt( "/xmlns:Relationships/xmlns:Relationship[last()]", "Target"	, cTarget )
	::aRels[nPos][1]:XPathAddAtt( "/xmlns:Relationships/xmlns:Relationship[last()]", "Id"		, cId )
Return cId

/*/{Protheus.doc} new_app
Cria arquivo \docprops\app.xml
@author Saulo Gomes Martins
@since 10/12/2019
@version 1.0
@param cXml, characters, xml para leitura
@type function
/*/
Method new_app(cXml) class YExcel
	Local nCont
	Local aNs
	Default cXml			:= ""
	::oapp	:= TXMLManager():New()
	If Empty(cXml)	//Cria modelo em branco
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
	EndIf
	::oapp:Parse(cXml)
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
@type function
/*/
Method new_core(cXml) class YExcel
	Local nCont
	Local aNs
	Local aRet
	Default cXml			:= ""
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
	EndIf
	::ocore:Parse(cXml)
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
@type function
/*/
Method new_workbook(cXml) class YExcel
	Local nCont
	Local aNs
	Default cXml			:= ""
	::oworkbook	:= TXMLManager():New()
	If Empty(cXml)	//Cria modelo em branco
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
	EndIf
	::oworkbook:Parse(cXml)
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
@return nPos, posição do draw no array
@param cXml, characters, xml para leitura
@param cCaminho, characters, caminho para gravar
@type function
/*/
Method new_draw(cXml,cCaminho) class YExcel
	Local nCont
	Local aNs
	Local oXML
	Default cXml			:= ""
	oXML	:= TXMLManager():New()
	If Empty(cXml)	//Cria modelo em branco
		cXml	+= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
		cXml	+= '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
		cXml	+= '</xdr:wsDr>'
	EndIf
	oXML:Parse(cXml)
	aNs	:= oXML:XPathGetRootNsList()
	For nCont:=1 to Len(aNs)
		oXML:XPathRegisterNs( aNs[nCont][1], aNs[nCont][2] )
	Next
	AADD(::aDraw,{oXML,cCaminho,0})
Return Len(::aDraw)

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

@type function
/*/
Method new_styles(cXml) class YExcel
	Local nCont
	Local aNs
	Default cXml			:= ""
	::oStyle	:= TXMLManager():New()
	If Empty(cXml)	//Cria modelo em branco
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
	EndIf
	::oStyle:Parse(cXml)
	aNs	:= ::oStyle:XPathGetRootNsList()
	For nCont:=1 to Len(aNs)
		::oStyle:XPathRegisterNs( aNs[nCont][1], aNs[nCont][2] )
	Next
	AADD(::aFiles,"\tmpxls\"+::cTmpFile+"\"+::cNomeFile+"\xl\styles.xml")
Return

/*/{Protheus.doc} xls_sheet
Cria arquivo \xl\worksheets\sheetX.xml
@author Saulo Gomes Martins
@since 10/12/2019
@version 1.0

@type function
/*/
Method xls_sheet(nFile) class YExcel
	Local nCont
	Local cRet	:= ""
	Local nTamArquivo,nBytesFalta,cBuffer,nBytesLer,nBytesLidos,nBytesSalvo
	Local cFunName	:= FUNNAME()
	Default cFunName:= ""
	If Type("cUserName")=="C"
		If !Empty(cFunName)
			cFunName	+= "|"
		EndIf
		cFunName	+= cUserName
	EndIf
	cRet	+= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
	cRet	+= '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
	cRet	+= ::osheetPr:GetTag()
	If ::aDimension[1][1]>0
		cRet	+= '<dimension ref="'+::NumToString(::aDimension[2][2])+cValToChar(::aDimension[2][1])+":"+::NumToString(::aDimension[1][2])+cValToChar(::aDimension[1][1])+'"/>'
	Else
		cRet	+= '<dimension ref="A1"/>'
	EndIf
	cRet	+= ::osheetViews:GetTag()
	cRet	+= '<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>'
	cRet	+= ::oCols:GetTag()
	::GravaFile(nFile,cRet)
	cRet:=""
	FWRITE(nFile,"<sheetData>")
	nTamArquivo := Fseek(::nFileTmpRow,0,2)
	nBytesFalta := nTamArquivo
	Fseek(::nFileTmpRow,0)
	cBuffer	:= Space(2048)
	While nBytesFalta > 0
		nBytesLer 	:= Min(nBytesFalta, 2048 )
		nBytesLidos := FREAD(::nFileTmpRow, @cBuffer, nBytesLer )
		nBytesSalvo := FWRITE(nFile, cBuffer,nBytesLer)
		nBytesFalta -= nBytesLer
	EndDo
	FCLOSE(::nFileTmpRow)
	FWRITE(nFile,"</sheetData>")

	cRet	+= If(ValType(::oAutoFilter)=="O",::oAutoFilter:GetTag(),"")
	cRet	+= If(!Empty(::oMergeCells:GetValor()),::oMergeCells:GetTag(),"")
	If !Empty(::aConditionalFormatting)
		For nCont:=1 to Len(::aConditionalFormatting)
			cRet	+= ::aConditionalFormatting[nCont]:GetTag()
		Next
	EndIf
	cRet	+= '<pageMargins left="0.511811024" right="0.511811024" top="0.78740157499999996" bottom="0.78740157499999996" header="0.31496062000000002" footer="0.31496062000000002"/>'
	cRet	+= '<pageSetup paperSize="9" fitToWidth="1" fitToHeight="0" orientation="'+::cPagOrientation+'" />
	cRet	+= '<headerFooter>'
//	cRet	+= '<oddHeader>&amp;R'+cFunName
//	If !Empty(cFunName)
//		cRet	+= CRLF
//	EndIf
//	cRet	+= DTOC(date())+" "+SUBSTR(TIME(),1,5)+'</oddHeader>'
	cRet	+= '<oddFooter>&amp;LTOTVS&amp;RPág &amp;P/&amp;N'+CRLF
	If !Empty(cFunName)
		cRet	+= cFunName+CRLF
	EndIf
	cRet	+= DTOC(date())+" "+SUBSTR(TIME(),1,5)
	cRet	+= '</oddFooter>'
	cRet	+= '</headerFooter>'
	cRet	+= if(!Empty(::atable),::otableParts:GetTag(),"")
	cRet	+= if(!Empty(::adrawing),::odrawing:GetTag(),"")
	cRet	+= '</worksheet>'
	::GravaFile(nFile,cRet)
	cRet:=""
Return

/*/{Protheus.doc} xls_table
Cria arquivo \xl\tables\tableX.xml
@author Saulo Gomes Martins
@since 10/12/2019
@version 1.0

@type function
/*/
Method xls_table() class YExcel
	Local cRet	:= ""
	cRet	+= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
	cRet	+= ::atable[::nCont]:GetTag()
Return cRet

/*/{Protheus.doc} xls_sharedStrings
Cria arquivo /xl/sharedStrings.xml
@author Saulo Gomes Martins
@since 10/12/2019
@version 1.0
@param nFile, numeric, header do arquivo
@type function
/*/
Method xls_sharedStrings(nFile) class YExcel
	Local nCont
	Local aString
	Local cRet	:= ""
	::oString:list(@aString)
	aSort(aString,,,{|x,y| x[2]<y[2] })
	cRet	+= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
	cRet	+= '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="'+cValToChar(Len(aString))+'" uniqueCount="'+cValToChar(Len(aString))+'">'
	FWRITE(nFile,cRet)
	cRet	:= ""
	For nCont:=1 to Len(aString)
		cRet	+= '<si>'
		cRet	+= '<t><![CDATA['+EncodeUTF8(aString[nCont][1])+']]></t>'
		cRet	+= '</si>'
		FWRITE(nFile,cRet)
		cRet	:= ""
	Next
	cRet	+= '</sst>'
	FWRITE(nFile,cRet)
	cRet	:= ""
Return cRet