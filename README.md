# advpl-excel
Gerar Excel formato xlsx com menor consumo de memoria e mais otimizado possível.

# Recursos disponíveis
* Definir células String,Numerica,data,Logica,formula
* Adicionar novas planilhas
* Cor de preenchimento
* Alinhamento
* Formato da celula
* Mesclar células
* Auto Filtro
* Congelar Paineis(colunas e linhas)
* Definir tamanho da coluna
* Definir tamanho da linha
* Letra: Fonte,Tamanho,Cor,Negrito,Italico,Sublinhado,Tachado
* Bordas: (Left,Right,Top,Bottom),Cor,Estilo
* Formatação condicional:(operador,formula)(font,fundo,bordas)
* Formatar como tabela

# Testes
![Exemplo1](https://github.com/saulogm/advpl-excel/raw/master/exemplo/excel1.png)

![Exemplo2](https://github.com/saulogm/advpl-excel/raw/master/exemplo/excel2.png)

![Exemplo2](https://github.com/saulogm/advpl-excel/raw/master/exemplo/excel3.png)

Exemplo de uso:
```advpl
#include 'protheus.ch'
#include 'parmtype.ch'

/*/{Protheus.doc} tstyexcel
Teste da classe YExcel
@author Saulo Gomes Martins
@since 08/05/2017

@type function
/*/
user function tstyexcel()
	Local oExcel	:= YExcel():new()
	Local nPosCor,nPosFont,nPosStyle,nPosMoeda,nPosQuebra
	ConOut(Time())
	oExcel:new("Pasta2")
	oExcel:ADDPlan(/*cNome*/)		//Adiciona uma planilha em branco
	//Definição de Cor Transparecia+RGB
	nPosCor			:= oExcel:CorPreenc("FF0000FF")	//Cor de Fundo Azul

						//cHorizontal,cVertical,lReduzCaber,lQuebraTexto,ntextRotation
	oAlinhamento	:= oExcel:Alinhamento("center","center")
	oQuebraTxt		:= oExcel:Alinhamento(,,,.T.)
	o45Graus		:= oExcel:Alinhamento(,,,,45)
						//cTipo,cCor,cModelo
	nPosBorda		:= oExcel:Borda("ALL","FFFF0000","thick")
	nPosBorda2		:= oExcel:Borda("ALL")
	
						//nTamanho,cCorRGB,cNome,cfamily,cScheme,lNegrito,lItalico,lSublinhado,lTachado
	nPosFont		:= oExcel:AddFont(20,"FFFFFFFF","Calibri","2")
	nPosFont2		:= oExcel:AddFont(20,56,"Calibri","2",,.T.,.T.,.T.,.T.)
	
	nPosStyle	:= oExcel:AddStyles(/*numFmtId*/,nPosFont/*fontId*/,nPosCor/*fillId*/,/*borderId*/,/*xfId*/,{oAlinhamento})
	nPosMoeda	:= oExcel:AddStyles(44/*numFmtId*/,/*fontId*/,/*fillId*/,/*borderId*/,/*xfId*/,{o45Graus})
	nPosMoeda2	:= oExcel:AddStyles(44/*numFmtId*/,/*fontId*/,/*fillId*/,nPosBorda2/*borderId*/,/*xfId*/)
	nPosQuebra	:= oExcel:AddStyles(/*numFmtId*/,/*fontId*/,/*fillId*/,/*borderId*/,/*xfId*/,{oQuebraTxt})
	nPosBorVerm	:= oExcel:AddStyles(/*numFmtId*/,/*fontId*/,/*fillId*/,nPosBorda/*borderId*/,/*xfId*/,{oQuebraTxt})
	nPosFonts	:= oExcel:AddStyles(/*numFmtId*/,nPosFont2/*fontId*/,/*fillId*/,/*borderId*/,/*xfId*/,)
	nPosBordas	:= oExcel:AddStyles(/*numFmtId*/,/*fontId*/,/*fillId*/,nPosBorda2/*borderId*/,/*xfId*/,)
	nPosBorDt	:= oExcel:AddStyles(14/*numFmtId*/,/*fontId*/,/*fillId*/,nPosBorda2/*borderId*/,/*xfId*/,)	//borda com data
	
	oExcel:Cell(1,1,"TESTE EXCEL",,nPosStyle)
	oExcel:mergeCells(1,1,2,6)						//Mescla as células A1:B2
	oExcel:Cell(3,1,100)							//A3	Numero
	oExcel:Cell(3,2,2,"1+1")						//B3	Formula simples
	oExcel:Cell(4,1,102,"A3+B3")					//A4	Formula com células
	oExcel:Cell(4,2,202,"SUM(A3:A4)")				//B4	Formula com funções
	oExcel:Cell(5,1,"Olá Mundo!",,nPosBorVerm)		//A5	Texto simples
	oExcel:Cell(5,2,date())							//B5	Data
	oExcel:Cell(5,3,.T.)							//C5	Campo Logico
	oExcel:Cell(5,4,1000,,nPosMoeda)				//D5	Campo Numerico formato moeda
	oExcel:nTamLinha	:= 30.75					//Defini o tamanho das proximas linha criadas
	oExcel:Cell(6,3,.F.)							//C6	Campo Logico falso
	oExcel:Cell(6,5,"Texto grande para quebra em linhas",,nPosQuebra)	//E6	Texto grande	
	oExcel:Cell(6,6,0,oExcel:Ref(3,1)+"+"+oExcel:Ref(3,2),)				//F6	Usando metodo RefSTR para localizar posição da celula
	oExcel:Cell(6,7,"Negrito,Italico,Sublinhado,Tachado",,nPosFonts)	//G6	Texto grande	
	oExcel:nTamLinha	:= nil
	
	oExcel:Cell(7,1,"FORMATAÇÃO CONDICIONAL")
	oExcel:mergeCells(7,1,7,3)
	oExcel:Cell(8,1,-10)
	oExcel:Cell(9,1,0)
	oExcel:Cell(10,1,5)
	oExcel:Cell(11,1,10)
	oExcel:Cell(12,1,20)
	oExcel:Cell(13,1,25)
	
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
	oExcel:FormatCond(oExcel:Ref(8,1),oExcel:Ref(13,1),nPosVerm,"<",0)				//Numero negativo em vermelho
	oExcel:FormatCond(oExcel:Ref(8,1),oExcel:Ref(13,1),nPosVerd,"between",{10,20})	//Entre 10 e 20
	oExcel:FormatCond(oExcel:Ref(8,1),oExcel:Ref(13,1),nPosAmar,"=","0")			//igual a zero
	
	//Defini o tamanho das colunas
	//Primeira_coluna,Ultima_coluna,Largura,AjusteNumero,customWidth
	oExcel:AddTamCol(1,2,12.00)
	oExcel:AddTamCol(3,3,20.00)
	oExcel:AddTamCol(4,6,12.00)
	
	//Teste de 50mil células - 20 segundos
	oExcel:ADDPlan("Teste")		//Adiciona nova planilha
	oExcel:SetDefRow(.T.,{1,4})	//Definir a coluna inicial e final da linha, importante para performace da classe
	oExcel:Cell(1,1,"Linha",,nPosBordas)
	oExcel:Cell(1,2,"Filial",,nPosBordas)
	oExcel:Cell(1,3,"Venda",,nPosBordas)
	oExcel:Cell(1,4,"Data Venda",,nPosBordas)
//	For nCont2:=5 to 50
//		oExcel:Cell(1,nCont2,"Coluna "+cValToChar(nCont2))
//	Next
	nCont2	:= 1
	For nCont:=2 to 100
		oExcel:Cell(nCont,1,nCont,,nPosBordas)
		oExcel:Cell(nCont,2,"Filial "+cValToChar(nCont2),,nPosBordas)
		oExcel:Cell(nCont,3,Randomize(1,100),,nPosMoeda2)
		oExcel:Cell(nCont,4,date()+nCont,,nPosBorDt)
		If nCont % 10 ==0
			nCont2++
		EndIf
//		For nCont2:=5 to 50
//			oExcel:Cell(nCont,nCont2,nCont2)
//		Next
	Next
	oExcel:Cell(nCont,1,99,"COUNT(A2:A"+cValToChar(nCont-1)+")")
	oExcel:Cell(nCont,3,0,"SUM(C2:C"+cValToChar(nCont-1)+")",nPosMoeda2)
	
	oExcel:AutoFilter(1,1,nCont,4)	//Auto filtro
	oExcel:AddPane(1,1)	//Congela primeira linha e primeira coluna

	//TESTE COM FORMATAR COMO TABELA
	oExcel:ADDPlan("Tabela")		//Adiciona nova planilha
	oExcel:SetDefRow(.T.,{1,4})		//Definir a coluna inicial e final da linha, importante para performace da classe
	oTabela	:= oExcel:AddTabela("Tabela1",1,1)	//Cria uma tabela de estilos
	oTabela:AddStyle("TableStyleMedium15"/*cNome*/,.T./*lLinhaTiras*/,/*lColTiras*/,/*lFormPrimCol*/,/*lFormUltCol*/)	//Cria os estilos,Cab:Preto|Linha:Cinza,Branco
	//oTabela:AddStyle("TableStyleMedium2"/*cNome*/,.T./*lLinhaTiras*/,/*lColTiras*/,.T./*lFormPrimCol*/,/*lFormUltCol*/)	//Cria os estilos,Cab:Azul|Linha:Azul,Branco
	oTabela:AddFilter()				//Adiciona filtros a tabela
	oTabela:AddColumn("Linha")		//Adiciona coluna Linha
	oTabela:AddColumn("Filial")		//Adiciona coluna Filial
	oTabela:AddColumn("Venda")		//Adiciona coluna Venda
	oTabela:AddColumn("Data Venda")	//Adiciona coluna Data Venda
	
	oTabela:AddLine()				//Cria a primeira linha
	//Preenche com o cabeçario na primeira linnha
	oTabela:Cell("Linha","Linha",,)
	oTabela:Cell("Filial","Filial",,,)
	oTabela:Cell("Venda","Venda",,,)
	oTabela:Cell("Data Venda","Data Venda",,)
	
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

	oExcel:Gravar("c:\temp",.T.,.T.)
return
```
# Instalação
1. Instalar o 7-Zip (http://www.7-zip.org/)
2. Configurar o appserver.ini com o caminho do 7-Zip.
```
[GENERAL]
LOCAL7ZIP=C:\Program Files\7-Zip\7z.exe
```
3. Aplicar patch yexcel.ptm

### Dúvidas
- Email: saulomax@gmail.com
- GitHub: https://github.com/saulogm
