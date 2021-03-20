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
	Local oAlinhamento,oQuebraTxt,o45Graus
	Local nCont,nCont2
	Local oDateTime
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
	oExcel	:= YExcel():new("TstYExcel")
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

						//cHorizontal,cVertical,lReduzCaber,lQuebraTexto,ntextRotation
	oAlinhamento	:= oExcel:Alinhamento("center","center")
	oQuebraTxt		:= oExcel:Alinhamento("center","center",,.T.)
	o45Graus		:= oExcel:Alinhamento(,,,,45)
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
	oPosStyle	:= oExcel:NewStyle():Setfont(nPosFont1):Setfill(nPosCor):Setborder(nPosBorda):SetaValores({oAlinhamento})
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
	
	// //Defini o tamanho das colunas
	//Primeira_coluna,Ultima_coluna,Largura,AjusteNumero,customWidth
	oExcel:AddTamCol(1,1,15.00)
	oExcel:AddTamCol(2,2,20.00)
	oExcel:AddTamCol(3,5,15.00)
	oExcel:AddTamCol(5,6,20.00)

	//Cadastra imagem
	If File("\Star_Wars_Logo.png")
		nIDImg		:= oExcel:ADDImg("\Star_Wars_Logo.png")	//Imagem no Protheus_data
				//nID,nLinha,nColuna,nX,nY,cUnidade,nRot
		oExcel:Img(nIDImg,7,7,200,121,/*"px"*/,)	//Usa imagem cadastrada
	EndIf
	//Para alterações deve primeiro posicionar na Celula pelo Pos(linha,coluna) ou PosR(referencia)
	oExcel:Pos(1,1):SetValue("TESTE EXCEL"):SetStyle(oPosStyle)
	oExcel:mergeCells(1,1,2,6)											//Mescla as células A1:F2
	//Textos
	oExcel:Pos(3,1):SetValue("Olá Mundo!"):SetStyle(oPosBorVerm)						//Texto simples
	oExcel:Pos(3,2):SetValue("Texto grande para quebra em linhas"):SetStyle(oPosQuebra)	//Texto grande
	oExcel:Pos(3,3):SetValue("Negrito,Italico,Sublinhado,Tachado"):SetStyle(oPosFonts)	//Formatando letra
	oExcel:SetRowH(60.75,3)	//Defini o tamanho da linha 3
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
	oDateTime	:= oExcel:GetDateTime(date(),time())				//Formatando DateTime
	oExcel:Pos(8,2):SetValue(oDateTime)								//Date time
	oExcel:Pos(8,3):SetValue(date()):SetStyle(oExcel:NewStyle():SetnumFmt(oExcel:AddFmt("[$-pt-BR]mmm-aaa;@")))	//Data formato mes-ano
	oDateTime	:= oExcel:GetDateTime(CTOD(""),"00:00:01")			//Formatando DateTime
	oExcel:Pos(8,4):SetValue(oDateTime):SetStyle(oExcel:NewStyle():SetnumFmt(oExcel:AddFmt("hh:mm:ss;@")))								//Date time

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
	oExcel:PosR("A4"):SetValue("Link Planilha Teste"):Addhyperlink("Teste!A1","Ir para teste"):SetStyle(oStyleLink)

	oExcel:PosR("E8"):AddComment("se não tem nada de bom a dizer não diga nada","Desconhecido")
	oExcel:PosR("E8"):AddComment()	//Deleta o comentario
	oExcel:PosR("E7"):AddComment("Que a Força esteja com você!","Mestre Jedi")

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
	oExcel:AddTamCol(1,2,12.00)
	oExcel:AddTamCol(3,3,20.00)
	oExcel:AddTamCol(4,4,12.00)
	oExcel:AddTamCol(5,6,18.00)
	oExcel:SetsumRight(.F.)			//Defini que o agrupamento vai ser na esquerda
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
		oExcel:Pos(nCont,1):SetValue(nCont)
		oExcel:Pos(nCont,2):SetValue("Filial "+cValToChar(nCont2))
		oExcel:Pos(nCont,3):SetValue(Randomize(1,100))
		oExcel:Pos(nCont,4):SetValue(Randomize(1,100)*(1+(Randomize(0,200)/100)))
		oExcel:Pos(nCont,5):SetValue(date()+nCont)
		oExcel:SetStyle(oRegEst1,nCont,1,nCont,5)	//Defini estilo da linha
		If nCont % 10 ==0
			lSubTotal	:= .T.
			oExcel:AddNome("VENDA"+cValToChar(nCont2),nCont-8,3,nCont,3)
			nCont++
			oExcel:NivelLinha(nil,If(nCont2==1,.F.,.T.))
			oExcel:Pos(nCont,1):SetValue("Sub Total Filial")
			oExcel:Pos(nCont,2):SetValue(cValToChar(nCont2))
			oExcel:Pos(nCont,3):SetValue(0,"SUBTOTAL(9,VENDA"+cValToChar(nCont2)+")")
			oExcel:Pos(nCont,4):SetValue("")
			oExcel:Pos(nCont,5):SetValue("")
			oExcel:SetStyle(oRegEst1,nCont,1,nCont,5)	//Defini estilo da linha
			nCont2++
		EndIf
		lSubTotal	:= .F.
	Next
	oExcel:NivelLinha()
	lSubTotal	:= .T.
	oExcel:Pos(nCont,1):SetValue("Total Geral")
	oExcel:Pos(nCont,2):SetValue("")
	oExcel:Pos(nCont,3):SetValue(0,'SUMIF(A2:'+oExcel:Ref(nCont-1,1)+',"Sub Total Filial",C2:'+oExcel:Ref(nCont-1,3)+')')
	oExcel:Pos(nCont,4):SetValue("")
	oExcel:Pos(nCont,5):SetValue("")
	oExcel:SetStyle(oRegEst1,nCont,1,nCont,5)	//Defini estilo da linha
	oExcel:SetRowLevel(2,nCont-1,1,.F.)			//Defini da llinha 2 até linha anterior como nível 1

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
	//oTabela:AddStyle("TableStyleMedium2"/*cNome*/,.T./*lLinhaTiras*/,/*lColTiras*/,.T./*lFormPrimCol*/,/*lFormUltCol*/)	//Cria os estilos,Cab:Azul|Linha:Azul,Branco
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

	oExcel:Save(GetTempPath())
	oExcel:OpenApp()
	oExcel:Close()
return

/*/{Protheus.doc} YxlsRead
Testa leitura simples do xlsx
@author Saulo Gomes Martins
@since 03/03/2018
@version 1.0

@type function
/*/
User Function YxlsRead()
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
Return

User Function YTstRW()
	Local aTamLin
	Local nCont,nCont2
	Local xValor
	Local oExcel	:= YExcel():new(,GetTempPath()+"TstYExcel.xlsx")
	oExcel:Pos(1,1):SetValue(oExcel:GetValue(1,1)+" - editado")
	oExcel:SetFooter("&18A&36B","&BTeste Excel","Pag &P/&N")		//Configura Rodapé
	aTamLin	:= oExcel:LinTam()
	For nCont:=aTamLin[1] to aTamLin[2]
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
	Next
	//TODO VERIFICAR VALORES GRAVADOS, POIS PARECE ESTÁ INCOERENTE
	
	// VarInfo("valor",oExcel:GetValue(8,1))
	// VarInfo("valor",oExcel:GetValue(8,2))
	// VarInfo("valor",oExcel:GetValue(8,3))
	// VarInfo("valor",oExcel:GetValue(8,4))
	oExcel:Save("c:\temp")
	oExcel:OpenApp()
	oExcel:Close()
Return

//Teste de bordas
User Function yTst2xl2()
	Local oExcel 	:= YExcel():new()
	Local nBorda1	:= oExcel:Borda("ALL","FFFF0000","thick")
	Local nBorda2	:= oExcel:Borda("LR","FF000000","thick")
	Local nBorda3	:= oExcel:Borda("ALL","FF000000","mediumDashDot")
	Local nBorda4	:= oExcel:Borda("ALL","FF000000","slantDashDot")
	Local nBorda5	:= oExcel:Borda("ALL","FF000000","double")
	Local nBorda6	:= oExcel:Borda("ALL","FF000000","dashed")
	Local nSty1		:= oExcel:AddStyles(/*numFmtId*/,/*fontId*/,/*fillId*/,nBorda1/*borderId*/,/*xfId*/,)
	Local nSty2		:= oExcel:AddStyles(/*numFmtId*/,/*fontId*/,/*fillId*/,nBorda2/*borderId*/,/*xfId*/,)
	Local nSty3		:= oExcel:AddStyles(/*numFmtId*/,/*fontId*/,/*fillId*/,nBorda3/*borderId*/,/*xfId*/,)
	Local nSty4		:= oExcel:AddStyles(/*numFmtId*/,/*fontId*/,/*fillId*/,nBorda4/*borderId*/,/*xfId*/,)
	Local nSty5		:= oExcel:AddStyles(/*numFmtId*/,/*fontId*/,/*fillId*/,nBorda5/*borderId*/,/*xfId*/,)
	Local nSty6		:= oExcel:AddStyles(/*numFmtId*/,/*fontId*/,/*fillId*/,nBorda6/*borderId*/,/*xfId*/,)
	oExcel:ADDPlan()
	oExcel:AddTamCol(1,1,30.00)
	oExcel:AddTamCol(3,3,30.00)
	oExcel:Pos(1,1):SetValue(GetDateTime(date(),time())):SetStyle(nSty1)
	oExcel:Pos(1,3):SetValue("Borda esquerda e direita"):SetStyle(nSty2)
	oExcel:Pos(3,1):SetValue("borda média traço-ponto"):SetStyle(nSty3)
	oExcel:Pos(3,3):SetValue("borda traço-ponto inclinado"):SetStyle(nSty4)
	oExcel:Pos(5,1):SetValue("borda de linha dupla"):SetStyle(nSty5)
	oExcel:Pos(5,3):SetValue("borda tracejado"):SetStyle(nSty6)
	oExcel:Save(GetTempPath())
	oExcel:OpenApp()
	oExcel:Close()
Return
