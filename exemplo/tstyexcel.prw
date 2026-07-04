#include 'protheus.ch'
#include 'parmtype.ch'

/*/{Protheus.doc} tstyexcel
Teste da classe YExcel
@author Saulo Gomes Martins
@since 08/05/2017
@vesion 2.0

@type function
@obs para LEITURA/EDIÇÃO ver função YTstRW no fim do fonte
/*/
user function tstyexcel()
	Local oExcel
	Local oAlinhamento,oQuebraTxt,o45Graus,oAliRecuoE,oAliRecuoD
	Local nCont,nCont2
	Local nStyle1
	Local nPosCor,nPosCorI,nPosCorP,nPosCor3,nPosCorEfe,nPosCorEf2
	Local nPosBorda,nBordaAll
	Local nPosfont1,nPosfont2,nPosfont3,nPosfont4,nPosFont5
	Local nFmtnum3
	Local oPosStyle,oPos3dec,oPosMoeda,oPosM45g,oPosMoeda2,oPosQuebra,oPosBorverm,oPosFonts,oPosCab,oPosEfe,oPosEfe2
	Local oEstilo1,oRegest1
	Local nIdImg
	Local oFont
	Local oCorPre,oCorPre2,oCorPre3
	Local oBorda
	Local nPosVerm,nPosVerd,nPosAmar
	Local oTabela
	Local nTotalvenda,nVenda
	Local lSubTotal	:= .F.
	Local oStyleLink
	Local oRecuoE,oRecuoD
	Local oRegraLinha
	Local jCab
	Local aOnlyFieds
	Local lSx3
	Local lExibirCab
	Local lCombo
	Local cAlias
	Local nStart

	// RpcSetEnv("T1","M SP 01")
	//RpcSetEnv("01","010101")
	RpcSetEnv("99","01")
	nStart := Seconds()
	oExcel	:= YExcel():new("TstYExcel",,"A")
	// oExcel	:= YExcel():new(,"C:\temp\novo.xlsx")
	// oExcel	:= YExcel():new(,)

	//Definição de Cor Transparecia+RGB
	nPosCor			:= oExcel:CorPreenc("FF0000FF")	//Cor de Fundo Azul
	nPosCorI		:= oExcel:CorPreenc("FFB8CCE4")	//Cor de Fundo Azul impa
	nPosCorP		:= oExcel:CorPreenc("FFDCE6F1")	//Cor de Fundo Azul par
	nPosCor3		:= oExcel:CorPreenc("FF4F81BD")	//Cor de Fundo Azul Escuro
					//EfeitoPreenc(nAngulo,aCores,ctype,nleft,nright,ntop,nbottom)
	nPosCorEfe		:= oExcel:EfeitoPreenc(90,{{"FFFFFF",0},{"0072C4",1}})							//Efeito linear
	nPosCorEf2		:= oExcel:EfeitoPreenc(,{{"FFFFFF",0},{"0072C4",1}},"path",0.2,0.8,0.2,0.8)		//Efeito do centro

						//cHorizontal,cVertical,lReduzCaber,lQuebraTexto,ntextRotation,nRecuo
	oAlinhamento	:= oExcel:Alinhamento("center","center")
	oQuebraTxt		:= oExcel:Alinhamento("center","center",,.T.)
	o45Graus		:= oExcel:Alinhamento(,,,,45)
	oAliRecuoE		:= oExcel:Alinhamento("left",,,,,2)			//Recuo esquerdo
	oAliRecuoD		:= oExcel:Alinhamento("right",,,,,2)		//Recuo direito
	oSemBloq		:= oExcel:Cellprotection(.F.)		//Célula não Bloqueada
	oOculForm		:= oExcel:Cellprotection(,.T.)		//Célula Oculta formulas
						//cTipo,cCor,cModelo
	nPosBorda		:= oExcel:Borda("ALL","FFFF0000","thick")
	nBordaAll		:= oExcel:Borda("ALL")

						//nTamanho,cCorRGB,cNome,cfamily,cScheme,lNegrito,lItalico,lSublinhado,lTachado
	nPosFont1		:= oExcel:AddFont(20,"FFFFFFFF","Arial","2")
	nPosFont2		:= oExcel:AddFont(20,56,"Calibri","2",,.T.,.T.,.T.,.T.)
	nPosFont3		:= oExcel:AddFont(11,"FFFFFFFF","Calibri","2")
	nPosFont4		:= oExcel:AddFont(11,"FFFF0000","Calibri","2")
	nPosFont5		:= oExcel:AddFont(11,"FF0000FF","Calibri","2",,,,.T.)

	nFmtNum3		:= oExcel:AddFmtNum(3/*nDecimal*/,.T./*lMilhar*/,/*cPrefixo*/,/*cSufixo*/,"("/*cNegINI*/,")"/*cNegFim*/,/*cValorZero*/,/*cCor*/,"Red"/*cCorNeg*/,/*nNumFmtId*/)

	//oExcel:NewStyle(id Estilo para herdar):SetnumFmt(id FmtNum):Setfont(id da fonte):Setfill(id fundo):Setborder(id borda):SetaValores(Array alinhamentos)
	oPosStyle	:= oExcel:NewStyle():Setfont(nPosFont1):Setfill(nPosCor):Setborder():SetaValores({oAlinhamento})
	oPos3Dec	:= oExcel:NewStyle():SetnumFmt(nFmtNum3)
	oPosMoeda	:= oExcel:NewStyle():SetnumFmt(44)
	oPosM45G	:= oExcel:NewStyle(oPosMoeda):SetaValores({o45Graus})
	oPosMoeda2	:= oExcel:NewStyle(oPosMoeda):Setborder(nBordaAll)
	oPosQuebra	:= oExcel:NewStyle():SetaValores({oQuebraTxt})
	oPosBorVerm	:= oExcel:NewStyle():SetaValores({oQuebraTxt}):Setborder(nPosBorda)
	oPosFonts	:= oExcel:NewStyle():Setfont(nPosFont2)
	oStyleLink	:= oExcel:NewStyle():Setfont(nPosFont5)
	
	oPosCAB		:= oExcel:NewStyle():SetFont(nPosFont3):Setfill(nPosCor3)

	oPosEfe		:= oExcel:NewStyle():Setfill(nPosCorEfe)
	oPosEfe2	:= oExcel:NewStyle():Setfill(nPosCorEf2)

	oEstilo1	:= oExcel:NewStyle():Setborder(nBordaAll)	//Estilo com bordas

	oRecuoE		:= oExcel:NewStyle():SetaValores({oAliRecuoE})
	oRecuoD		:= oExcel:NewStyle():SetaValores({oAliRecuoD})

	oSubTotal	:= oExcel:NewStyle(oEstilo1):Setfill(nPosCor3):SetFont(nPosFont3)
	
	oProtecao1	:= oExcel:NewStyle():SetaValores({oSemBloq})
	oProtecao2	:= oExcel:NewStyle():SetaValores({oOculForm})

	//REGRAS DE ESTILO - EXEMPLOS DE REGRAS POSSIVEIS
	// oStyRule := oExcel:NewStyRules()
	// bBloco	:= {|nLinha,nColuna,oExcel| Logico }
	// oStyRule:AddStyle(bBloco,Estilo)		//Cria regra para definir o estilo
	// oStyRule:AddnumFmt(bBloco,idFmtNum)	//Cria regra para definir formato
	// oStyRule:AddFont(bBloco,idFont)		//Cria regra para definir a fonte
	// oStyRule:Addfill(bBloco,idFill)		//Cria regra para definir o preenchimento de fundo
	// oStyRule:Addborder(bBloco,idBorder)	//Cria regra para definir a borda
	// oStyRule:AddValores(bBloco,aValores)	//Cria regra para alinamentos

	//Criar Agrupamento de estilo
	oRegEst1 := oExcel:NewStyRules()
	oRegEst1:AddStyle({|nLin,nCol,oObjExcel| lSubTotal    }, oExcel:NewStyle(oEstilo1):Setfill(nPosCor3):SetFont(nPosFont3) )	//Linhas SubTotais
	oRegEst1:AddStyle({|nLin,nCol,oObjExcel| nLin % 2 ==1 }, oExcel:NewStyle(oEstilo1):Setfill(nPosCorI) )	//Linhas ímpar aplica estilo herdados com borda e fundo azul
	oRegEst1:AddStyle({|nLin,nCol,oObjExcel| nLin % 2 ==0 }, oExcel:NewStyle(oEstilo1):Setfill(nPosCorP) )	//Linhas pares aplica estilo herdados com borda e fundo azul2
	oRegEst1:AddnumFmt({|nLin,nCol,oObjExcel| nCol==3	  }, 44 )											//Se coluna B irá definir formato da célula moeda
	oRegEst1:AddnumFmt({|nLin,nCol,oObjExcel| nCol==4	  }, nFmtNum3 )										//Se coluna C irá definir formato da célula numero com 3 casas decimais
	oRegEst1:AddFont({|nLin,nCol,oObjExcel| nCol==4	 	 }, nPosFont4 )										//Se coluna C irá definir fonte vermelha


	//===================PRIMEIRA PLANILHA====================================
	oExcel:ADDPlan(/*cNome*/)		//Adiciona uma planilha em branco
	oExcel:SetdefaultRowHeight(12.75)	//Defini altura padrão das linhas não informadas
	// //Defini o tamanho das colunas
	//Primeira_coluna,Ultima_coluna,Largura,AjusteNumero,customWidth
	oExcel:AddTamCol(1,1,15.00)
	oExcel:AddTamCol(2,2,20.00)
	oExcel:AddTamCol(3,5,15.00)
	oExcel:AddTamCol(5,6,20.00)
	oExcel:ColHidden(6)

	//Cadastra imagem
	If File("\Star_Wars_Logo.png")
		nIDImg		:= oExcel:ADDImg("\Star_Wars_Logo.png")	//Imagem no Protheus_data
				//nID,nLinha,nColuna,nX,nY,cUnidade,nRot
		oExcel:Img(nIDImg,7,7,200,121,/*"px"*/,)	//Usa imagem cadastrada
	EndIf
	//Para alterações deve primeiro posicionar na Celula pelo Pos(linha,coluna) ou PosR(referencia)
	//For nCont:=1 to 2	//Até linha 2
	//	For nCont2:=1 to 6	//Até Coluna 6
	//		oExcel:Pos(nCont,nCont2):SetValue("TESTE EXCEL"):SetStyle(oPosStyle)
	//	Next
	//Next
	oExcel:Pos(1,1):SetValue("TESTE EXCEL"):SetStyle(oPosStyle)
	oExcel:mergeCells(1,1,2,6)											//Mescla as células A1:F2
	//Textos
	oExcel:Pos(3,1):SetValue("Olá Mundo!"):SetStyle(oPosBorVerm)						//Texto simples
	oExcel:Pos(3,2):SetValue("Texto grande para quebra em linhas"):SetStyle(oPosQuebra)	//Texto grande
	oExcel:Pos(3,3):SetValue("Negrito,Italico,Sublinhado,Tachado"):SetStyle(oPosFonts)	//Formatando letra
	oExcel:SetRowH(30.75,3)	//Defini o tamanho da linha 3
	//Numeros
	oExcel:Pos(5,1):SetValue(100):SetStyle(oPos3Dec)				//Numero
	oExcel:Pos(5,2):SetValue(-100.2):SetStyle(oPos3Dec)				//Numero negativo
	oExcel:Pos(5,3):SetValue(1000):SetStyle(oPosMoeda)				//Campo Numerico formato moeda
	oExcel:AddNome("VALOR1",5,1)									//Defini nome da referencia de célula
	oExcel:AddNome("VALORES",5,1,5,3)								//Defini o nome do intervalo
	oExcel:Pos(6,1):SetValue(2,"1+1")								//Formula simples
	oExcel:Pos(6,2):SetValue(102,"VALOR1+A6")						//Formula com células
	oExcel:Pos(6,3):SetValue(999.8,"SUM(VALORES)")					//Formula com funções
	oExcel:Pos(6,4):SetValue(1099.8,oExcel:Ref(5,1)+"+"+oExcel:Ref(6,3))	//Usando metodo Ref para localizar posição da celula
	//Datas
	oExcel:Pos(8,1):SetValue(date())								//Data
	oExcel:Pos(8,2):SetDateTime(date(),time())						//Date time
	oExcel:Pos(8,3):SetValue(date()):SetStyle(oExcel:NewStyle():SetnumFmt(oExcel:AddFmt("[$-pt-BR]mmm-aaa;@")))	//Data formato mes-ano
	oExcel:Pos(8,4):SetDateTime(CTOD(""),"00:00:01"):SetStyle(oExcel:NewStyle():SetnumFmt(oExcel:AddFmt("hh:mm:ss;@")))								//Date time
	oExcel:Pos(8,5):SetValue(oExcel:GetDateTime(date(),time())+10)	//DateTime + 10 dias

	//Logicos
	oExcel:Pos(10,1):SetValue(.T.):SetStyle(oPosEfe)				//C5	Campo Logico
	oExcel:Pos(10,2):SetValue(.F.):SetStyle(oPosEfe2)				//C6	Campo Logico falso

	oExcel:Pos(12,1):SetValue("FORMATAÇÃO CONDICIONAL")
	oExcel:mergeCells(12,1,12,3)
	oExcel:SetRowH(20,13,18)	//Altera a linha 13-18
	oExcel:Pos(13,1):SetValue(-10)
	oExcel:Pos(14,1):SetValue(0)
	oExcel:Pos(15,1):SetValue(5)
	oExcel:Pos(16,1):SetValue(10)
	oExcel:Pos(17,1):SetValue(20)
	oExcel:Pos(18,1):SetValue(25)

	//Adiciona Link
	oExcel:PosR("A20"):SetValue("Link Planilha Teste"):Addhyperlink("Teste!A1","Ir para teste"):SetStyle(oStyleLink)

	oExcel:PosR("B21"):AddComment("se não tem nada de bom a dizer não diga nada","Desconhecido")
	oExcel:PosR("B21"):AddComment()	//Deleta o comentario
	oExcel:PosR("B22"):AddComment("Que a Força esteja com você!","Mestre Jedi")

	oExcel:Pos(24,1):SetValue("Texto com recuo esquerdo"):SetStyle(oRecuoE)
	oExcel:Pos(24,5):SetValue("Texto com recuo direito"):SetStyle(oRecuoD)
	oExcel:mergeCells(24,5,24,6)

	//Proteção em planilha
	oExcel:Pos(25,1):SetValue("Senha 123"):SetStyle(oProtecao1)	//Permitir editar essa célula
	oExcel:Pos(25,2):SetValue(30,"15+15"):SetStyle(oProtecao2)	//Ocultar formula
	oExcel:SetsheetProtection("123")
	
	nStyle1	:= oExcel:GetStyle(5,3)		//Pega estilo da primeira celula

	//FORMATAÇÃO CONDICIONAL
	//Cria objetos para ser usado na formatação
	oFont	:= oExcel:Font(12,"FFFFFF","Calibri","2",,.T.,.F.,.F.,.F.)	//Cor Branca Negrito
	oCorPre	:= oExcel:Preenc("FF0000")									//Fundo Vermelho
	oCorPre2:= oExcel:Preenc("00FF00")									//Fundo Verde
	oCorPre3:= oExcel:Preenc("FFFF00")									//Fundo Amarelo
	oBorda	:= oExcel:ObjBorda("ALL","000000")							//Borda Preta
	//Cria o Estilo			oFont,oCorPreenc,oBorda
	nPosVerm	:= oExcel:ADDdxf(oFont,oCorPre,oBorda)
	nPosVerd	:= oExcel:ADDdxf(oFont,oCorPre2,oBorda)
	nPosAmar	:= oExcel:ADDdxf(,oCorPre3,oBorda)
	//OBS: Os estilos são criados para worksheet do arquivo, podendo ser usado em outras planilhas(abas)

	//Cria as regras	cRefDe,cRefAte,nEstilo,operator,xFormula
	oExcel:FormatCond(oExcel:Ref(13,1),oExcel:Ref(18,1),nPosVerm,"<",0)				//Numero negativo em vermelho
	oExcel:FormatCond(oExcel:Ref(13,1),oExcel:Ref(18,1),nPosVerd,"between",{10,20})	//Entre 10 e 20
	oExcel:FormatCond(oExcel:Ref(13,1),oExcel:Ref(18,1),nPosAmar,"=","0")			//igual a zero


	oExcel:SetHeader("&A","A&KFF0000B&K0070C0C&K01+000D&K07+037E","&D"+CHR(10)+"&T")		//Configura Cabeçalho
	oExcel:SetFooter("&18A&36B","&BTeste Excel","Pág &P/&N")		//Configura Rodapé

	//===================SEGUNDA PLANILHA====================================
	oExcel:ADDPlan("Teste" ,"00AA00")		//Adiciona nova planilha
	oExcel:SetPlanAt(2)			//Apenas teste o metodo
	oExcel:SetPlanAt("Teste")	//Apenas teste o metodo
	oExcel:SetPagOrientation("portrait")	//default|landscape|portrait
	//oExcel:SetPrintArea(1,1,5,20)	//Define área de impressão
	oExcel:AddTamCol(1,2,12.00)
	oExcel:AddTamCol(3,3,20.00)
	oExcel:AddTamCol(4,4,12.00)
	oExcel:AddTamCol(5,6,18.00)
	oExcel:SetsumRight(.F.)				//Defini que o agrupamento vai ser na esquerda
	oExcel:SetColLevel(4,5,1,.T.)		//Agrupa coluna 4 e 5 fechado
	If File("\Star_Wars_Logo.png")
		oExcel:Img(nIDImg,2,6,121,200,/*"px"*/,270)	//Usa imagem com rotação de 270
	EndIf

	oExcel:Pos(1,1):SetValue("Linha"):SetStyle(oPosCAB)
	oExcel:Pos(1,2):SetValue("Filial"):SetStyle(oPosCAB)
	oExcel:Pos(1,3):SetValue("Venda"):SetStyle(oPosCAB)
	oExcel:Pos(1,4):SetValue("Numero"):SetStyle(oPosCAB)
	oExcel:Pos(1,5):SetValue("Data Venda"):SetStyle(oPosCAB)
	nCont2	:= 1
	For nCont:=2 to 110
		oExcel:NivelLinha(2,,If(nCont2==1,.F.,.T.))	//NivelLinha(nNivel,lFechado,lOculto)	PROXIMAS LINHAS A SER CRIADO COM NÍVEL 2
		oExcel:Pos(nCont,1):SetValue(nCont):SetStyle(oRegEst1)
		oExcel:Pos(nCont,2):SetValue("Filial "+cValToChar(nCont2)):SetStyle(oRegEst1)
		oExcel:Pos(nCont,3):SetValue(Randomize(1,100)):SetStyle(oRegEst1)
		oExcel:Pos(nCont,4):SetValue(Randomize(1,100)*(1+(Randomize(0,200)/100))):SetStyle(oRegEst1)
		oExcel:Pos(nCont,5):SetValue(date()+nCont):SetStyle(oRegEst1)
		If nCont % 10 ==0
			lSubTotal	:= .T.
			oExcel:AddNome("VENDA"+cValToChar(nCont2),nCont-8,3,nCont,3)
			nCont++
			oExcel:NivelLinha(nil,If(nCont2==1,.F.,.T.))
			oExcel:Pos(nCont,1):SetValue("Sub Total Filial"):SetStyle(oRegEst1)
			oExcel:Pos(nCont,2):SetValue(cValToChar(nCont2)):SetStyle(oRegEst1)
			oExcel:Pos(nCont,3):SetValue(0,"SUBTOTAL(9,VENDA"+cValToChar(nCont2)+")"):SetStyle(oRegEst1)
			oExcel:Pos(nCont,4):SetValue(""):SetStyle(oRegEst1)
			oExcel:Pos(nCont,5):SetValue(""):SetStyle(oRegEst1)
			nCont2++
		EndIf
		lSubTotal	:= .F.
	Next
	oExcel:NivelLinha()
	lSubTotal	:= .T.
	oExcel:Pos(nCont,1):SetValue("Total Geral"):SetStyle(oRegEst1)
	oExcel:Pos(nCont,2):SetValue(""):SetStyle(oRegEst1)
	oExcel:Pos(nCont,3):SetValue(0,'SUMIF(A2:'+oExcel:Ref(nCont-1,1)+',"Sub Total Filial",C2:'+oExcel:Ref(nCont-1,3)+')'):SetStyle(oRegEst1)
	oExcel:Pos(nCont,4):SetValue(""):SetStyle(oRegEst1)
	oExcel:Pos(nCont,5):SetValue(""):SetStyle(oRegEst1)

	oExcel:AutoFilter(1,1,nCont,4)	//Auto filtro
	oExcel:AddPane(1,1)	//Congela primeira linha e primeira coluna


	//===================TERCEIRA PLANILHA====================================
	//TESTE COM FORMATAR COMO TABELA
	oExcel:ADDPlan("Tabela","0000FF")		//Adiciona nova planilha

	oExcel:AddTamCol(1,2,12.00)
	oExcel:AddTamCol(3,3,20.00)
	oExcel:AddTamCol(4,4,12.00)
	oExcel:AddTamCol(5,6,18.00)
	oExcel:SetPrintTitles(1,1)				//Linha de/ate que irá repetir na impressão de paginas
	oExcel:showGridLines(.F.)				//Oculta linhas de grade
	//oExcel:Cell(1,1,"teste",,)
	oTabela	:= oExcel:AddTabela("Tabela1",1,1)	//Cria uma tabela de estilos
	oTabela:AddStyle("TableStyleMedium15"/*cNome*/,.T./*lLinhaTiras*/,/*lColTiras*/,/*lFormPrimCol*/,/*lFormUltCol*/)	//Cria os estilos,Cab:Preto|Linha:Cinza,Branco
	oTabela:AddFilter()				//Adiciona filtros a tabela
	oTabela:AddColumn("Linha")		//Adiciona coluna Linha
	oTabela:AddColumn("Filial")		//Adiciona coluna Filial
	oTabela:AddColumn("Venda")		//Adiciona coluna Venda
	oTabela:AddColumn("Data Venda")	//Adiciona coluna Data Venda

	nTotalVenda	:= 0	//Valor Total da venda
	nCont2		:= 1	//Variavel de controle
	For nCont:=2 to 100
		oTabela:AddLine()				//Adiciona nova linha
		//Preenche as células
		oTabela:Cell("Linha",nCont,,)
		oTabela:Cell("Filial","Filial "+cValToChar(nCont2),,)
		nVenda		:= Randomize(1,100)
		nTotalVenda	+= nVenda
		oTabela:Cell("Venda",nVenda,,)
		oTabela:Cell(4,date()+nCont,,)
		If nCont % 10 ==0
			nCont2++
		EndIf
	Next
	oTabela:AddTotal("Linha","TOTAL","")							//Preenche texto TOTAL na linha totalizadora da coluna Linha
	oTabela:AddTotal("Filial",99,"SUBTOTAL(103,Tabela1[Filial])")	//Usa função COUNTA(Contar Valores)
	oTabela:AddTotal("Venda",nTotalVenda,"SUM")		//Usa função SUM(Somar) para totalizar a coluna venda
	oTabela:AddTotais()	//Adiciona linha de totais
	oTabela:Finish()	//Fecha a edição da tabela

	// EXEMPLO PREENCHER COM ALIAS
	oExcel:ADDPlan("SB1","E26B0A")		//Adiciona nova planilha
	SB1->(DBSetOrder(4))	//B1_FILIAL, B1_GRUPO, B1_COD
	//SB1->(DBSetFilter({|| B1_GRUPO<>'    '},"B1_GRUPO<>'    '"))
	SB1->(DBSetFilter({|| B1_TIPO='KT'},"B1_TIPO='KT'"))
	SB1->(DbGoTop())
	oExcel:Pos(1,1):SetValue("Cadastro de produtos")

	//Definir detalhe dos campos manualmente
	jCab	:= jSonObject():New()
		//NewFldTab(jCab,cCampo,cDescricao,nTamanho,cPicture,cCombo,xStyle,nOrdem,cTipo,cDados,lNewCampo,lHidden,nTamDados)
	oExcel:NewFldTab(jCab,"B1_PRV1")
	jCab["B1_PRV1"]["style"]	:= oPosMoeda	//informar campo moeda
	oExcel:NewFldTab(jCab,"B1_GRUPO")
	jCab["B1_GRUPO"]["ordem"]	:= 1			//Alterar ordem para 1
	oExcel:NewFldTab(jCab,"NEWCAMPO","Registro",10,"9",,,,"N","Recno()",.T.)

	lSx3		:= .T.	//Buscar definições de campo da SX3
	lExibirCab	:= .T.	//Exibir Cabeçalho da tabela
	lCombo		:= .T.	//Traduzir campos combobox para a descrição
	aOnlyFieds	:= {"B1_COD","B1_DESC","B1_TIPO","B1_GRUPO","B1_UM","B1_LOCPAD","B1_PICM","B1_PRV1","B1_RASTRO","B1_UREV"}	//Campos para exibir
	//Criar SubTotal
		//DefSubTotal(cChave,lSubTotal,lTotalGeral,lAgrupar,nNivel)
	oExcel:DefSubTotal("B1_GRUPO",.T.,.T.,.T.,2)
		//AddSubTotal(cCampo,cFuncao,bFormula,bAdvpl,bAdvplExib,xValorIni)
	oExcel:AddSubTotal("B1_GRUPO","last",/*bFormula*/,/*bAdvpl*/,/*bAdvplExib*/,/*xValorIni*/)	//Ultimo conteudo
	oExcel:AddSubTotal("B1_PRV1","9",/*bFormula*/,/*bAdvpl*/,/*bAdvplExib*/,/*xValorIni*/)	//Somar
	oExcel:AddSubTotal("B1_PICM","1",/*bFormula*/,/*bAdvpl*/,/*bAdvplExib*/,/*xValorIni*/)	//Média
	oExcel:AddSubTotal("B1_LOCPAD","formula",{|cLinIni,cLinFim,cColuna,cChave,lTotalGeral| "SUBTOTAL(3,"+cColuna+cLinIni+":"+cColuna+cLinFim+")" }/*bFormula*/,{|aValAtu,xValor| aValAtu[1]+=1 }/*bAdvpl*/,{|aValAtu,xValor| cValToChar(aValAtu[1]) }/*bAdvplExib*/,0/*xValorIni*/)	//Formula
	oExcel:AddSubTotal("B1_COD","formula",/*bFormula*/,{|aValAtu,xValor,cChave| aValAtu[1]:=cChave }/*bAdvpl*/,{|aValAtu| If(oExcel:cSubTotalTp=="T","Total Geral","Total do Grupo: "+aValAtu[1]) }/*bAdvplExib*/,""/*xValorIni*/)	//Exibir texto nos totalizadores

	//Regra para formatação de linha	
	oRegraLinha	:= oExcel:NewRuleLine({|| oExcel:cSubTotalTp })	//Tipo de linha a ser impressa
	oRegraLinha:AddRegra("S",oSubTotal,.F.)						//Igual a S-SubTotal
	oRegraLinha:AddRegra("T",oSubTotal,.F.)						//Igual a T-TotalGeral
	//		Alias2Tab(cAlias,oStyle,lSx3,jCab,lExibirCab,lCombo,aOnlyFieds,aRegraStyle,oStyleLinha,lFiltro)
	oExcel:Alias2Tab("SB1",oPosCAB,lSx3,jCab,lExibirCab,,aOnlyFieds,oRegraLinha,oEstilo1)
	
	
	// EXEMPLO PREENCHER COM QUERY E FORMATO TABELA DO EXCEL
	cAlias := MpSysOpenQuery("SELECT A1_CGC,A1_COD,A1_LOJA,A1_PESSOA,A1_NOME,A1_DESC,A1_LC,A1_MSALDO,A1_MCOMPRA,A1_ULTCOM,A1_RECCOFI FROM "+RetSqlName("SA1")+" WHERE A1_COD<='000100' AND D_E_L_E_T_=' '")
	oExcel:ADDPlan("SA1","1F497D")			//Adiciona nova planilha
	oTabela	:= oExcel:AddTabela("Tabela2")	//Cria uma tabela de estilos
	oTabela:AddStyle("TableStyleMedium9"/*cNome*/,.T./*lLinhaTiras*/,/*lColTiras*/,/*lFormPrimCol*/,/*lFormUltCol*/)
	jCab	:= jSonObject():New()	//Definir do cabeçalho
		//NewFldTab(jCab,cCampo,cDescricao,nTamanho,cPicture,cCombo,xStyle,nOrdem,cTipo,cDados,lNewCampo)
	oExcel:NewFldTab(jCab,"A1_COD","Código do Cliente",10,,,)
	oExcel:NewFldTab(jCab,"A1_PESSOA",,,,"F=Física;J=Jurídica",)
	oExcel:NewFldTab(jCab,"A1_DESC",,,"@E 999,999,999.99",,)
	oExcel:NewFldTab(jCab,"A1_MSALDO",,,,,oPosMoeda)
	
	//Regra para formatação de linha
	oRegraLinha	:= oExcel:NewRuleLine({|| oExcel:cCampo+"|"+(cAlias)->A1_RECCOFI })	//Formata apenas 1 coluna da linha
				//AddRegra(cRegra,oStyle,lPrincipal)
	oRegraLinha:AddRegra("A1_RECCOFI|S",oPosEfe,.F.)
	//		Alias2Tab(cAlias,oStyle,lSx3,jCab,lExibirCab,lCombo,aOnlyFieds,aRegraStyle,oStyleLinha)
	oTabela:Alias2Tab(cAlias,,.T.,jCab,,,,oRegraLinha)
	oTabela:AddTotal(1,0,"COUNTA")	//Usa função COUNTA(Contar Valores)
	oTabela:AddTotais()	//Adiciona linha de totais
	oTabela:Finish()	//Fecha a edição da tabela
	(cAlias)->(DbCloseArea())

	oExcel:Save(GetTempPath())
	oExcel:OpenApp()
	oExcel:Close()
	conout( "Tempo: " + LTrim( Str( Seconds()-nStart ) ) + " segundos" )
return

/*/{Protheus.doc} YxlsRead
Testa leitura simples do xlsx
@author Saulo Gomes Martins
@since 03/03/2018
@version 1.0

@type function
/*/
User Function YxlsRead()
	//DEPRECATED, ver função U_YTstRW
	Local oExcel	:= YExcel():new("TesteXlsx")	//Cria teste
	Local cTexto	:= "Texto teste"
	Local nNumero	:= 123.09
	Local lLogico	:= .T.
	Local dData		:= date()
	Local oDateTime := oExcel:GetDateTime(date(),time())	//Formatando DateTime
	Local nColuna,nLinha
	oExcel:ADDPlan()
	oExcel:Cell(1,1,cTexto,,)
	oExcel:Cell(2,1,nNumero,,)
	oExcel:Cell(3,1,lLogico,,)
	oExcel:Cell(4,1,dData,,)
	oExcel:Cell(5,1,oDateTime)
	oExcel:ADDPlan()
	oExcel:Cell(1,1,"OK",,)
	cArquivo	:= oExcel:Gravar(GetTempPath(),.F.)	//Não abrir arquivo
	ConOut(cArquivo)
	oExcel	:= YExcel():new()
	oExcel:OpenRead(cArquivo)
	For nLinha	:= 1 to oExcel:adimension[1][1]
		For nColuna	:= 1 to oExcel:adimension[1][2]
			ConOut("Tipo:"+ValType(oExcel:CellRead(nLinha,nColuna)))
			ConOut(oExcel:CellRead(nLinha,nColuna))
		Next
		If nLinha==5
			oDateTime	:= oExcel:GetDateTime(,,oExcel:CellRead(nLinha,1))
			ConOut("Formato data")
			ConOut(oDateTime:GetDate())
			ConOut(oDateTime:GetTime())
			ConOut(oDateTime:GetStrNumber())
		EndIf
	Next
	ConOut("Ler planilha 2")
	oExcel:OpenRead(cArquivo,2)
	For nLinha	:= 1 to oExcel:adimension[1][1]
		For nColuna	:= 1 to oExcel:adimension[1][2]
			ConOut("Tipo:"+ValType(oExcel:CellRead(nLinha,nColuna)))
			ConOut(oExcel:CellRead(nLinha,nColuna))
		Next
	Next
	oExcel:CloseRead()
	FreeObj(oDateTime)
	//DEPRECATED
Return
/*/{Protheus.doc} YTstRW
Teste leitura e escrita
@type function
@version 1.0
@author Saulo Gomes Martins
@since 30/03/2021
/*/
User Function YTstRW()
	Local aTamLin
	Local nCont
	Local nCont2
	//Local xValor
	Local oExcel	:= YExcel():new(,GetTempPath()+"TstYExcel.xlsx")
	Local nStart := Seconds()
	Local nElapsed
	ConOut(TIME())
	oExcel:Pos(1,1):SetValue(oExcel:GetValue(1,1)+" - editado")
	aTamLin	:= oExcel:LinTam()
	For nCont:=aTamLin[1] to 100//aTamLin[2]
		ConOut("Linha:"+cValToChar(nCont))
		/*If nCont==1
			ConOut(oExcel:GetString(nCont,1,"inlineStr"))
			ConOut(oExcel:GetString(nCont,2,"inlineStr"))
			ConOut(oExcel:GetString(nCont,3,"inlineStr"))
			ConOut(oExcel:GetString(nCont,4,"inlineStr"))
		Else
			ConOut(oExcel:GetNumber(nCont,1))
			ConOut(oExcel:GetString(nCont,2,"inlineStr"))
			ConOut(oExcel:GetNumber(nCont,3))
			//aData	:= oExcel:GetDtTime(nCont,4)
			//ConOut(aData[1])
			//ConOut(aData[2])
			ConOut(oExcel:GetDate(nCont,4))
		Endif*/
		aTamCol	:= oExcel:ColTam(nCont)
		If aTamCol[1]>0
			For nCont2:=aTamCol[1] to aTamCol[2]
				xValor	:= oExcel:GetValue(nCont,nCont2)
				If ValType(xValor)=="O"
					VarInfo(oExcel:Ref(nCont,nCont2),xValor,,.F.)
				Else
					ConOut(oExcel:Ref(nCont,nCont2)+"["+cValToChar(xValor)+"]")
				EndIf
			Next
		EndIf
		oExcel:Pos(nCont,5):SetValue("Editado")
	Next
	ConOut(TIME())
	nElapsed	:= Seconds() - nStart
	conout( "Tempo: " + LTrim( Str( nElapsed ) ) + " segundos" )
	oExcel:Save("c:\temp")
	oExcel:OpenApp()
	oExcel:Close()
Return

User Function ytstOlaMundo
	Local oExcel := YExcel():new()
	oExcel:ADDPlan() //Cria Planinha
	oExcel:Pos(1,1):SetValue("Olá Mundo") //Escreve
	oExcel:Save()    //Salvar
	oExcel:OpenApp() //Abrir Excel
	oExcel:Close()   //Fechar e limpar objeto
Return
/*/{Protheus.doc} yTst2xl4
Teste leitura
@type function
@version 1.0
@author Saulo Gomes Martins
@since 30/03/2021
/*/
User Function yTst2xl4()
	Local aTamLin
	Local nContP,nContL,nContC
	Local xValor
	Local oExcel	:= YExcel():new(,"D:\temp\Saulo.xlsx")

	For nContP:=1 to oExcel:LenPlanAt()	//Ler as Planilhas
		oExcel:SetPlanAt(nContP)		//Informa qual a planilha atual
		ConOut("Planilha:"+oExcel:GetPlanAt("2"))	//Nome da Planilha
		aTamLin	:= oExcel:LinTam() 		//Linha inicio e fim da linha
		For nContL:=aTamLin[1] to aTamLin[2]
			aTamCol	:= oExcel:ColTam(nContL) //Coluna inicio e fim
			If aTamCol[1]>0	//Se a linha tem algum valor
				For nContC:=aTamCol[1] to aTamCol[2]
					xValor	:= oExcel:GetValue(nContL,nContC)	//Conteúdo 
					If ValType(xValor)=="O"
						ConOut(oExcel:Ref(nContL,nContC)+"["+cValToChar(xValor:GetDate())+"]["+cValToChar(xValor:GetTime())+"]")
						//VarInfo(oExcel:Ref(nContL,nContC),xValor,,.F.)
					Else
						ConOut(oExcel:Ref(nContL,nContC)+"["+cValToChar(xValor)+"]")
					EndIf
				Next
			EndIf
		Next
	Next
	oExcel:Close()
Return


User Function tst2Excel()
	Local nStart, nElapsed
	RpcSetEnv("99","01")
	conout( "YExcel")
	nStart := Seconds()
	//tstFwXlsx()
	//TesteyExcel()
	//tstBulk()
	YTstLeitura()
	//tstBulk2()
	//U_ytst6()
	nElapsed	:= Seconds() - nStart
	ntam		:= Directory("c:\temp\Pasta1.xlsx","HSD")[1][2]
	conout( "Tempo: " + LTrim( Str( nElapsed ) ) + " segundos" )
	conout( "Tamanho Arquivo: " + LTrim( Str( Round(ntam/1024/1024,3) ) ) + " MB" )

Return

Static Function fazernada(cValor)
Return

Static Function TesteyExcel()
	Local nCont,nCont2
	Local oExcel		:= YExcel():new("Pasta1",,"A")
	Local nPosBorda2	:= oExcel:Borda("ALL")
	Local nPosBordas	:= oExcel:AddStyles(/*numFmtId*/,/*fontId*/,/*fillId*/,nPosBorda2/*borderId*/,/*xfId*/,)
	oExcel:ADDPlan(/*cNome*/)		//Adiciona uma planilha em branco
	// oExcel:SetDefRow(.F.,{1,4})	//Definir a coluna inicial e final da linha, importante para performace da classe
	oExcel:Cell(1,1,"Linha",,nPosBordas)
	oExcel:Cell(1,2,"Filial",,nPosBordas)
	oExcel:Cell(1,3,"Venda",,nPosBordas)
	oExcel:Cell(1,4,"Data Venda",,nPosBordas)
	nCont2	:= 1
	cSubTotais	:= ""
	For nCont:=2 to 10000
		oExcel:Pos(nCont,1):SetValue(nCont)
		oExcel:Pos(nCont,2):SetValue("Filial"+cValtoChar(nCont))
		oExcel:Pos(nCont,3):SetValue(Randomize(1,100))
		oExcel:Pos(nCont,4):SetValue(date()+nCont)
	Next
	oExcel:Gravar("c:\temp",.F.,.T.)
Return

Static Function tstBulk()
	Local nCont,nCont2
	Local oExcel		:= YExcel():new("Pasta1",,"B")
	Local nPosBorda2	:= oExcel:Borda("ALL")
	//Local nPosCorI		:= oExcel:CorPreenc("FFB8CCE4")	//Cor de Fundo Azul impa
	Local nPosBordas	:= oExcel:NewStyle():Setborder(nPosBorda2)
	//Local oRegEst1 		:= oExcel:NewStyRules()
	//oRegEst1:AddStyle({|nLin,nCol,oObjExcel| nLin % 2 ==1 }, oExcel:NewStyle(nPosBordas):Setfill(nPosCorI) )	//Linhas ímpar aplica estilo herdados com borda e fundo azul
	//Local oPosFill		:= oExcel:NewStyle():Setfill(oExcel:CorPreenc("FF0000"))
	Local aCampos		:= {}
	Local aRegraStyle
	Local nElapsed,nStart
	oExcel:ADDPlan(/*cNome*/)		//Adiciona uma planilha em branco
	// oExcel:SetDefRow(.F.,{1,4})	//Definir a coluna inicial e final da linha, importante para performace da classe
	oExcel:Cell(1,1,"Linha",,nPosBordas)
	oExcel:Cell(1,2,"Filial",,nPosBordas)
	oExcel:Cell(1,3,"Venda",,nPosBordas)
	oExcel:Cell(1,4,"Data Venda",,nPosBordas)
	//BulkNewField(nColuna,cTipo,nTamanho,cCombo,oStyle,lFormula,lDatetime,cCampo,cDados)
	AADD(aCampos,oExcel:BulkNewField(1,"N"))
	AADD(aCampos,oExcel:BulkNewField(2,"C",12))
	AADD(aCampos,oExcel:BulkNewField(3,"N"))
	//AADD(aCampos,oExcel:BulkNewField(3,"C","1=Sim;2=Não"))
	AADD(aCampos,oExcel:BulkNewField(4,"D"))

	//aRegraStyle	:= oExcel:NewRuleLine({|| If(nCont%10==0,"vermelho","") },{"vermelho",oPosFill,.F.}):GetArray()
	//oExcel:SetsumBelow(.T.)				//Defini que o agrupamento de linhas vai ser em baixo
	//oExcel:SetsumRight(.F.)				//Defini que o agrupamento vai ser na esquerda
	//oExcel:SetColLevel(4,5,1,.T.)		//Agrupa coluna 4 e 5 fechado

	oExcel:DefBulkLine(aCampos,aRegraStyle,.F.)
	cSubTotais	:= ""
	oExcel:nLinha	:= 2
	nStart	:= Seconds()
	nCont2	:= 1
	For nCont:=2 to 100000
		//If nCont2==1
		//	oExcel:NivelLinha(nil,.T.,.F.)
		//ElseIf nCont2==2
		//	oExcel:NivelLinha(1,.T.,.T.)
		//ElseIf nCont2>11
		//	nCont2 := 0
		//EndIf
		//nCont2++
		oExcel:SetValueBulk(nCont)
		oExcel:SetValueBulk("Filial"+cValtoChar(nCont))
		oExcel:SetValueBulk(Randomize(1,100))
		//oExcel:SetValueBulk(cValToChar(Randomize(1,3)))
		oExcel:SetValueBulk(date()+nCont)
		oExcel:SetBulkLine()
	Next
	oExcel:NivelLinha()
	oExcel:FlushBulk()
	nElapsed	:= Seconds() - nStart
	conout( "Linhas Tempo: " + LTrim( Str( nElapsed ) ) + " segundos" )
	nStart	:= Seconds()
	oExcel:Gravar("c:\temp",.F.,.T.)
	nElapsed	:= Seconds() - nStart
	conout( "Gravar: " + LTrim( Str( nElapsed ) ) + " segundos" )
Return

Static Function tstBulk2()
	Local nCont,nCont2
	Local oExcel		:= YExcel():new("Pasta1",,"A")
	Local aCampos		:= {}
	oExcel:ADDPlan(/*cNome*/)		//Adiciona uma planilha em branco
	// oExcel:SetDefRow(.F.,{1,4})	//Definir a coluna inicial e final da linha, importante para performace da classe
	//1-Numero da coluna,2-Tipo de conteudo(C,N,L,D),3-Logico se vai ter Formula,4-Logico se é datetime,5-Objeto Estilo
	For nCont:=1 to 16385	//Limite maximo de colunas do excel
		AADD(aCampos,oExcel:BulkNewField(nCont,"C"))
	Next

	oExcel:DefBulkLine(aCampos)
	nCont2	:= 1
	cSubTotais	:= ""
	For nCont:=2 to 16385
		oExcel:SetValueBulk("abcdefghijklmnopqrstuvxz abcdefghijklmnopqrstuvxz abcdefghijklmnopqrstuvxz "+cValtoChar(nCont))
	Next
	oExcel:SetBulkLine()
	oExcel:FlushBulk()
	oExcel:Gravar("c:\temp",.F.,.T.)
Return


Static Function tstFwXlsx()
	Local nCont,nCont2
	Local oExcel		:= FwMsExcelXlsx():New()
	Local nElapsed,nStart
	oExcel:SetWriteinFile(.T.)
	////Habilita dados de processamento no BD e com Bulk ocorrendo com até 200 mil registros.
	//lWriteDb := oExcel:SetWriteinDb(.T., 200000)
	//If !lWriteDb
	//	conout("Não foi possível habilitar o recurso de dados em disco, o processamento consumirá a memória do servidor.")
	//EndIf
	oExcel:AddworkSheet("WorkSheet1")
	oExcel:AddTable ("WorkSheet1","Table1")
	oExcel:AddColumn("WorkSheet1","Table1","Linha",1,2,.F., )
	oExcel:AddColumn("WorkSheet1","Table1","Filial",1,1,.F., )
	oExcel:AddColumn("WorkSheet1","Table1","Venda",1,3,.F., )
	oExcel:AddColumn("WorkSheet1","Table1","Data Venda",1,4,.F., )

	nStart	:= Seconds()
	nCont2	:= 1
	For nCont:=2 to 100000
		oExcel:AddRow("WorkSheet1","Table1",{nCont,"Filial"+cValtoChar(nCont),Randomize(1,100),date()+nCont})
	Next
	nElapsed	:= Seconds() - nStart
	conout( "Linhas Tempo: " + LTrim( Str( nElapsed ) ) + " segundos" )
	nStart	:= Seconds()
	oExcel:Activate()
	oExcel:GetXMLFile("c:\temp\fwTESTE.xlsx")
	oExcel:DeActivate()
	nElapsed	:= Seconds() - nStart
	conout( "Gravar: " + LTrim( Str( nElapsed ) ) + " segundos" )
Return

Static Function YTstLeitura()
	Local aTamLin
	Local nCont//,nCont2
	//Local xValor
	Local oExcel
	Local nStart := Seconds()
	Local nElapsed
	
	oExcel	:= YExcel():new(,"c:\temp\Pasta1.xlsx")
	nElapsed	:= Seconds() - nStart
	conout( "Carregar xlsx: " + LTrim( Str( nElapsed ) ) + " segundos" )

	nStart := Seconds()
	aTamLin	:= oExcel:LinTam()
	For nCont:=aTamLin[1] to aTamLin[2]
		//fazernada(oExcel:Pos(nCont,1):GetValue())
		//fazernada(oExcel:Pos(nCont,2):GetValue())
		//fazernada(oExcel:Pos(nCont,3):GetValue())
		//fazernada(oExcel:Pos(nCont,4):GetValue())

		ConOut(oExcel:Pos(nCont,1):GetNumber())
		ConOut(oExcel:Pos(nCont,2):GetString())
		ConOut(oExcel:Pos(nCont,3):GetNumber())
		ConOut(oExcel:Pos(nCont,4):GetDate())
		//If nCont>2
		//Exit
		//EndIf
	Next
	ConOut(TIME())
	nElapsed	:= Seconds() - nStart
	conout( "Leitura: " + LTrim( Str( nElapsed ) ) + " segundos" )
	oExcel:Close()
Return

User Function TstMerge()
	Local oExcel		:= YExcel():new("Pasta1")
	Local nPosBorda2	:= oExcel:Borda("ALL")
	Local nPosBordas	:= oExcel:AddStyles(/*numFmtId*/,/*fontId*/,/*fillId*/,nPosBorda2/*borderId*/,/*xfId*/,)
	Local nCont
	Local nStart		:= Seconds()
	ConOut("=================="+TIME())
	oExcel:ADDPlan(/*cNome*/)		//Adiciona uma planilha em branco
	// oExcel:SetDefRow(.F.,{1,4})	//Definir a coluna inicial e final da linha, importante para performace da classe
	oExcel:Cell(1,1,"Linha",,nPosBordas)
	oExcel:Cell(1,2,"Filial",,nPosBordas)
	oExcel:Cell(1,3,"Venda",,nPosBordas)
	oExcel:Cell(1,4,"Data Venda",,nPosBordas)
	nCont2	:= 1
	cSubTotais	:= ""
	For nCont:=2 to 10
		oExcel:Pos(nCont,1):SetValue(nCont)
		oExcel:Pos(nCont,2):SetValue("Filial"+cValtoChar(nCont))
		oExcel:Pos(nCont,3):SetValue(Randomize(1,100))
		oExcel:Pos(nCont,4):SetValue(date()+nCont)
		//oExcel:mergeCells(nCont,4,nCont,6)
	Next

	//Cria objetos para ser usado na formatação
	oFont	:= oExcel:Font(12,"FFFFFF","Calibri","2",,.T.,.F.,.F.,.F.)	//Cor Branca Negrito
	oCorPre	:= oExcel:Preenc("FF0000")									//Fundo Vermelho
	oCorPre2:= oExcel:Preenc("00FF00")									//Fundo Verde
	oCorPre3:= oExcel:Preenc("FFFF00")									//Fundo Amarelo
	oBorda	:= oExcel:ObjBorda("ALL","000000")							//Borda Preta
	//Cria o Estilo			oFont,oCorPreenc,oBorda
	nPosVerm	:= oExcel:ADDdxf(oFont,oCorPre,oBorda)
	nPosVerd	:= oExcel:ADDdxf(oFont,oCorPre2,oBorda)
	nPosAmar	:= oExcel:ADDdxf(,oCorPre3,oBorda)

	oExcel:FormatCond(oExcel:Ref(2,3),oExcel:Ref(nCont-1,3),nPosAmar,"=","0")			//igual a zero

	oExcel:Gravar("c:\temp",.T.,.T.)
	ConOut("=================="+TIME()+"|FIM")
	ConOut(Seconds() - nStart)
Return

User Function TstEdit()
	oExcel := YExcel():new(,"C:\Temp\TESTE.xlsx")
	oExcel:Pos(10,3):SetValue("TESTE")
	oExcel:Pos(11,3):SetValue("12/08/2022")
	oExcel:Save(GetTempPath())
	oExcel:OpenApp() //Abrir Excel
	oExcel:Close() //Fechar e limpar objeto
Return

User Function tstMemoria()
	Local oExcel := YExcel():new(,,"M")  //Inicia em memoria
	Local nCont
	oExcel:ADDPlan() //Cria Planinha
	oExcel:AddTamCol(1,2,15.00)
	oExcel:Pos(1,1):SetValue("Olá Mundo") //Altera linha 1 coluna 1
	oExcel:Pos(2,1):SetValue("Segunda linha") //Altera linha 2 coluna 1
	oExcel:AddTamCol(2,2,25.00)
	oExcel:AddTamCol(2,3,35.00)
	For nCont:=1 to 30
		oExcel:Pos(Randomize(1,30),Randomize(1,30)):SetValue(nCont) //Volta para linha 1 coluna 1
	Next
	oExcel:AddPane(2,3)
	oExcel:Save()    //Salvar
	oExcel:OpenApp() //Abrir Excel
	oExcel:Close()   //Fechar e limpar objeto
Return


	//ElseIf ::lMemoria
	//	If nLinha<::nLinha .And. ;//Linha para posicionamento é menor que atual
	//		!::aPlanilhas[nPlanilha][7]:Get(nLinha,@cPathLinha)	//E linha não existe
	//			UserException("YExcel - gravação em arquivo deve ser sequencial. Linha Atual "+cValToChar(::nLinha)+" Linha enviada "+cValToChar(nLinha))
	//	ElseIf nLinha<>::nLinha
	//		If ::aPlanilhas[nPlanilha][7]:Get("C|"+cValToChar(nLinha),@nCQtd)

	//			//cRefUlt	:= ::asheet[::nPlanilhaAt][1]:XPathGetAtt(cPathLinha,"c","")
	//			//If ::aPlanilhas[nPlanilha][7]:Get(::cRef,@cPathColuna)
	//			//If nColuna < 
	//		EndIf
	//	Endif

User Function tst3Excel()
	Local oExcel	:= YExcel():new("NomeArquivo")
	oExcel:ADDPlan()
	oExcel:Pos(1,1):SetValue("Olá Mundo")
	oExcel:ColHidden(3)	//Ocultar a coluna C
	oExcel:Save()    //Salvar
	oExcel:OpenApp() //Abrir Excel
	oExcel:Close()   //Fechar e limpar objeto
Return
