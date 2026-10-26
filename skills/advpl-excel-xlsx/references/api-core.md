# API core — custom.excel.Xlsx

As assinaturas abaixo vêm de `src/custom.excel.tlpp`. Quase todo método que "age na célula" retorna `self`, o que permite encadear: `oExcel:Pos(1,1):SetValue("x"):SetStyle(oSty)`.

## Sumário
1. Ciclo de vida
2. Planilhas
3. Posicionamento e referências
4. Valores, fórmulas e datas
5. Leitura
6. Mesclar, dimensões, visibilidade
7. Congelar, filtro, grade
8. Nomes, hyperlink, comentário
9. Agrupamento (outline)
10. Imagens
11. Inserir linhas/colunas

---

## 1. Ciclo de vida

| Método | Descrição |
|---|---|
| `New(cNomeFile, cFileOpen, cTipo)` | `cNomeFile`: nome do arquivo sem extensão (se omitido, é aleatório). `cFileOpen`: abre um xlsx existente (chama `Open`). `cTipo`: `"A"` (padrão, arquivo sequencial), `"M"` (memória), `"B"` (banco temporário/bulk) |
| `Open(cFile, cSenha)` → lógico | Abre um xlsx, inclusive criptografado (MS-OFFCRYPTO Agile). Reinicia o objeto, então uma instância pode abrir vários arquivos em sequência |
| `GetError()` → caractere | Mensagem da última falha do `Open` |
| `Save(cLocal)` → caractere | Grava e retorna o caminho do arquivo. `cLocal` pode ser uma pasta do client (`GetTempPath()`, `"c:\temp"`) ou do servidor (`"\temp"`) |
| `OpenApp()` | Abre o arquivo salvo no aplicativo padrão (ShellExecute). Só com SmartClient |
| `Destroy()` / `Close()` | Libera os objetos internos e remove os temporários |
| `Gravar(cLocal, lAbrir, lDelSrv)` | v1: Save + OpenApp opcional + Close numa chamada. Não chame `Close` de novo depois |

```tlpp
oExcel := custom.excel.Xlsx():New(cNomeFile = "Vendas", cTipo = "A")
oExcel := custom.excel.Xlsx():New(cFileOpen = GetTempPath() + "Vendas.xlsx")	//abrir para ler/editar
oExcel := custom.excel.Xlsx():New()	//depois oExcel:Open(cArq, cSenha)
```

## 2. Planilhas

| Método | Descrição |
|---|---|
| `ADDPlan(cNome, cCor)` → id | Cria uma planilha e passa a editá-la. `cCor` é a cor da aba em RGB (`"00AA00"`). Pode ser chamado sem nome |
| `SetPlanAt(xPlan)` → lógico | Seleciona a planilha por índice ou nome |
| `GetPlanAt(cRet)` | `"1"` (padrão) retorna o índice; `"2"` retorna o nome |
| `LenPlanAt()` | Quantidade de planilhas |
| `SetPlanName(cNome)` | Renomeia a planilha atual |

Os métodos de coluna, painel, impressão e grade valem para a planilha atual, por isso devem ser chamados **depois** do `ADDPlan`.

## 3. Posicionamento e referências

| Método | Descrição |
|---|---|
| `Pos(nLinha, nColuna, nPlanilha)` → self | Posiciona a célula. Índices começam em 1 |
| `PosR(cRef)` → self | Posiciona por referência: `PosR("A20")` |
| `Ref(nLinha, nColuna, lLinhaAbs, lColunaAbs)` → caractere | `Ref(3,3)` = `"C3"`; `Ref(3,3,.T.,.T.)` = `"$C$3"` |
| `LocRef(cRef)` → `{nLin,nCol}` | `LocRef("C5")` = `{5,3}` |
| `NumToString(nCol)` / `StringToNum(cCol)` | Converte entre número e letra da coluna (3 ↔ "C") |

Os dados públicos `oExcel:nLinha` e `oExcel:nColuna` guardam a posição atual. `oExcel:nLinha := 2` pode ser usado antes de um bulk.

## 4. Valores, fórmulas e datas

**`SetValue(xValor, cFormula)`**: aceita caractere, numérico, lógico, data, objeto `DateTime` ou `""` (célula vazia, útil para aplicar estilo).

```tlpp
oExcel:Pos(2,1):SetValue("Texto")
oExcel:Pos(2,2):SetValue(1234.56)
oExcel:Pos(2,3):SetValue(.T.)
oExcel:Pos(2,4):SetValue(Date())
oExcel:Pos(2,5):SetValue(oExcel:GetDateTime(dData = Date(), cTime = Time()))	//data e hora
oExcel:Pos(2,6):SetValue(oExcel:GetDateTime(dData = Date(), cTime = Time()) + 10)	//DateTime aceita + - * /
oExcel:Pos(4,1):SetValue(, "SUM(A2:A3)")	//inglês, vírgula, sem "="; o valor é calculado no Save
oExcel:Pos(5,1):SetValue(, oExcel:Ref(2,1) + "+" + oExcel:Ref(3,1))
oExcel:Pos(6,1):SetValue(xValor = 600, cFormula = "SUM(VALORES)")	//xValor informado é mantido; VALORES criado antes com AddNome
```

- **Fórmulas sem valor são calculadas pela classe** no `Save`/`Save2Html`, nos modos A, M e B (motor AdvPL com as funções do Excel). O que não for calculável fica sem valor e o workbook é marcado para o Excel recalcular ao abrir. `SetCalcFormula(.F.)` desliga o cálculo, `AddFuncFormula(cNome, bBloco, nMin, nMax, cFlags)` registra uma função própria e `SetCalcLocal("en-US")` muda a localidade de TEXT/VALUE/DATEVALUE (padrão pt-BR). Veja [formulas.md](formulas.md).
- `SetDateTime(dDate, cTime)` → self: equivale a `SetValue(GetDateTime(...))`.
- `GetDateTime(dData, cTime, nData)` → um objeto `custom.excel.DateTime`, com `GetDate()`, `GetTime()`, `GetStrNumber()` e operadores `+ - * /` e de comparação.
- Uma célula com data e hora precisa de um formato de exibição. Crie um estilo com `AddFmt("dd/mm/yyyy hh:mm:ss;@")` (ver estilos.md) ou use o id nativo 22.
- `Cell(nLinha, nColuna, xValor, cFormula, xStyle)` é da API v1. Em código novo use `Pos():SetValue():SetStyle()`.

## 5. Leitura (tipos "M" e "B", ou arquivo aberto)

| Método | Retorno |
|---|---|
| `GetValue(nLinha, nColuna, xDefault, @lAchou)` | Valor da célula. Linha e coluna padrão são as da posição atual. Data e hora vêm como objeto `DateTime` |
| `Pos(l,c):GetString()` / `GetNumber()` / `GetDate()` / `GetBoolean()` / `GetDtTime()` | Valor convertido para o tipo |
| `GetFormula()` | Fórmula da célula posicionada |
| `LinTam(nColuna, nColuna2)` | `{primeiraLinha, ultimaLinha}` |
| `ColTam(nLinha, nLinha2)` | `{primeiraColuna, ultimaColuna}`. Se `[1] == 0`, a linha está vazia |
| `GetStyle(nLinha, nColuna)` | Estilo da célula |

```tlpp
oExcel := custom.excel.Xlsx():New()
If !oExcel:Open(cArq, cSenha)
	ConOut(oExcel:GetError())
Else
	For nPlan := 1 To oExcel:LenPlanAt()
		oExcel:SetPlanAt(nPlan)
		aLinhas := oExcel:LinTam()
		For nLin := aLinhas[1] To aLinhas[2]
			aCols := oExcel:ColTam(nLin)
			If aCols[1] == 0
				Loop
			EndIf
			For nCol := aCols[1] To aCols[2]
				xValor := oExcel:GetValue(nLin, nCol)
				If ValType(xValor) == "O"	//DateTime
					dData := xValor:GetDate()
				EndIf
			Next
			FwFreeArray(aCols)
		Next
		FwFreeArray(aLinhas)
	Next
EndIf
oExcel:Destroy()
FreeObj(oExcel)
```

Para editar um arquivo aberto, basta posicionar, fazer `SetValue` e chamar `Save` novamente.

## 6. Mesclar, dimensões, visibilidade

| Método | Descrição |
|---|---|
| `mergeCells(nLinha, nColuna, nLinha2, nColuna2)` | Os 4 parâmetros são obrigatórios. O valor e o estilo vão na célula superior esquerda (escreva-a antes) |
| `AddTamCol(nMin, nMax, nWidth, lbestFit, lcustomWidth)` | Largura das colunas `nMin..nMax` (padrão 12). Chame depois do ADDPlan e antes de escrever |
| `ColHidden(nColuna, lHidden)` | Oculta a coluna (padrão `.T.`) |
| `SetRowH(nHeight, nLinha, nLinha2)` | Altura das linhas. Sem linha informada, vale para as linhas criadas depois. No modo `"A"`, chamar antes de gravar as linhas do intervalo (linha já gravada gera erro) |
| `SetdefaultRowHeight(nHeight)` | Altura padrão (15) |

## 7. Congelar, filtro, grade

| Método | Descrição |
|---|---|
| `AddPane(nySplit, nxSplit)` | Congela painéis. `AddPane(1,0)` congela a linha 1; `AddPane(1,1)` congela a linha 1 e a coluna A |
| `AutoFilter(nLinha, nColuna, nLinha2, nColuna2)` | Autofiltro. **Um por planilha** |
| `showGridLines(lView)` | `.F.` oculta as linhas de grade |

## 8. Nomes, hyperlink, comentário

| Método | Descrição |
|---|---|
| `AddNome(cNome, nLinha, nColuna, nLinha2, nColuna2, cRefPar, cPlanilha, cEscopo)` | Nome definido para uma célula ou intervalo. Crie antes da fórmula que o usa |
| `Addhyperlink(cLocation, ctooltip, cDisplay, cRef)` → self | Link interno na célula posicionada (`"Plan2!A1"`). `cLocation` vazio remove o link. Não muda a aparência, então aplique uma fonte azul sublinhada |
| `AddComment(cText, cAutor, nleft, ntop, nWidth, nHeight)` → self | Comentário na célula posicionada. Sem texto, remove |

```tlpp
oExcel:PosR("A20"):SetValue("Ir para Teste"):SetStyle(oStyLink):Addhyperlink(cLocation = "Teste!A1", ctooltip = "Ir para teste")
oExcel:PosR("B21"):SetValue("Obs"):AddComment(cText = "Valor revisado", cAutor = "Financeiro")
```

## 9. Agrupamento (outline)

| Método | Descrição |
|---|---|
| `NivelLinha(nNivel, lFechado, lOculto)` | Nível (1-7) **das linhas escritas a seguir**. `NivelLinha()` sem argumentos reseta. `NivelLinha(nil, lFechado)` marca a linha de resumo |
| `SetRowLevel(nLinha, nLinha2, nNivel, lFechado)` | Nível para um intervalo de linhas |
| `SetsumBelow(l)` | Resumo abaixo (`.T.`) ou acima |
| `SetsumRight(l)` | Resumo à direita. **Chame antes** de `SetColLevel` |
| `SetColLevel(nMin, nMax, nNivel, lFechado)` | Agrupa colunas |

## 10. Imagens

```tlpp
If File("\logo.png")	//Protheus_data ou caminho local
	nIdImg := oExcel:ADDImg("\logo.png")	//registra uma vez (gera erro se o arquivo não existir)
	oExcel:Img(nID = nIdImg, nLinha = 1, nColuna = 1, nX = 200, nY = 80)	//nX/nY obrigatórios (px)
	oExcel:Img(nID = nIdImg, nLinha = 10, nColuna = 5, nX = 100, nY = 40, nRot = 270)	//reutiliza, com rotação
EndIf
```

Assinatura: `Img(nID, nLinha, nColuna, nX, nY, cUnidade = "px", nRot, cCellType)`. A mesma imagem pode ser usada em outras planilhas.

## 11. Inserir linhas/colunas (tipo "M"/arquivo aberto)

`InsertRowEmpty(nLinha, nLinha2)`, `InsertCellEmpty(nLinha, nColuna, nLinha2, nColuna2)`, `AddRow(nQtd, nLinhaIni, nColunaIni, nColunaFim)` e `AddCol(nQtd, nColunaIni, nLinhaIni, nLinhaFim)`.
