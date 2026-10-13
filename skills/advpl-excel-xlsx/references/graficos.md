# Gráficos — namespace `custom.excel.chart`

## Sumário
1. Fluxo básico
2. NewChart — tipos
3. Métodos comuns (Chart e Serie)
4. Métodos por tipo
5. Combinação (ChartCombo)
6. Exemplos por tipo

---

## 1. Fluxo básico

1. Escreva os dados na planilha, em ordem sequencial se o tipo for "A".
2. Crie o gráfico com `oGraf := oExcel:NewChart("tipo")` ou `custom.excel.chart.ChartX():New(oExcel)`.
3. Configure título, legenda e opções.
4. Adicione as séries com `AddSerie(...)`, passando a referência absoluta `"Plan!$B$2:$B$7"`.
5. Ancore com `Insert(nLinha, nColuna, nLinha2, nColuna2)`, **uma única vez** e antes do `Save`.

O gráfico fica na planilha **atual**. Gráficos e séries são registrados na instância e liberados no `oExcel:Destroy()`, então não use `FreeObj` neles.

A referência da série inclui o nome da planilha. Se o nome tiver espaço, use aspas simples: `"'Minha Plan'!$B$2:$B$7"`. Os arrays opcionais `aVal`/`aCat` são gravados como cache, e o Excel os mostra antes de recalcular (útil para LibreOffice e para visualizações prévias).

## 2. `NewChart(cTipo = "col")`: atalhos

| cTipo | Classe |
|---|---|
| `col` (colunas), `bar` (barras horizontais) | `ChartBar` |
| `line` | `ChartLine` |
| `pie` | `ChartPie` |
| `scatter` | `ChartScatter` |
| `area` | `ChartArea` |
| `stock` / `acoes` | `ChartStock` |
| `radar` | `ChartRadar` |
| `doughnut` / `rosca` | `ChartDoughnut` |
| `bubble` / `bolhas` | `ChartBubble` |
| `ofpie` | `ChartOfPie` |
| `surface` / `superficie` | `ChartSurface` |
| `combo` | `ChartCombo` |

Também é possível instanciar a classe direto: `custom.excel.chart.ChartScatter():New(oExcel)`.

## 3. Métodos comuns

**Chart:**
- `AddSerie(cNome, cRefVal, cRefCat, aVal, aCat, cCor)` retorna um objeto `Serie`.
- `SetTitulo(cTitulo)`.
- `SetLegenda(lExibir, cPosicao)`: `cPosicao` é `"r"` (padrão), `"l"`, `"t"` ou `"b"`.
- `SetGrade(lExibir)`.
- `Set3D(lAtivo, nRotX = 15, nRotY = 20, nPerspectiva = 30)`: só para Bar, Line, Pie, Area e Surface. Nos demais gera UserException.
- `Insert(nLinha = 1, nColuna = 1, nLinha2 = +15, nColuna2 = +8)`: índices a partir de 1.

**Serie:**
- `SetCor("FF4F81BD")`.
- `SetMarcador(l)`.
- `SetSuavizar(l)`.
- `SetExplosao(n%)` para pizza.
- `SetXY(cRefX, cRefY, aX, aY)` para dispersão e bolhas.
- `SetTamanho(cRefTam, aTam)` para bolhas.

## 4. Métodos por tipo

| Classe | Métodos |
|---|---|
| ChartBar | `SetDirecao("col"\|"bar")`, `SetAgrupamento("clustered"\|"stacked"\|"percentStacked")`, `SetGapWidth(0-500, padrão 150)` |
| ChartLine | `SetAgrupamento("standard"\|"stacked"\|"percentStacked")`, `SetMarcador(l)`, `SetSuavizar(l)` (ignorados em 3D) |
| ChartPie | `SetRotulos(lExibir, lValor, lPercentual, lCategoria)`, `SetAnguloIni(n)` |
| ChartScatter | `SetEstilo("marker"\|"lineMarker"\|"smoothMarker"\|"line"\|"smooth")` |
| ChartArea | `SetAgrupamento(...)` |
| ChartStock | Não tem métodos próprios. Adicione as séries na ordem Abertura (opcional), Máxima, Mínima, Fechamento |
| ChartRadar | `SetEstilo("standard"\|"marker"\|"filled")` |
| ChartDoughnut | `SetFuro(10-90, padrão 50)` + os métodos de Pie |
| ChartOfPie | `SetTipo("pie"\|"bar")`, `SetDivisao(cTipo = "auto"\|"pos"\|"val"\|"percent", nValor)` |
| ChartBubble | `SetEscala(1-300)`, `SetTamRepr("area"\|"w")` |
| ChartSurface | `SetWireframe(l)` |
| ChartCombo | `AddChart(oChart, lEixoSecundario)` |

## 5. Combinação: `ChartCombo`

Os gráficos membros são instanciados direto (`ChartBar():New(oExcel)`, `ChartLine():New(oExcel)`) e recebem as próprias séries. Depois entram no combo com `AddChart`. **Só `ChartBar`, `ChartLine` e `ChartArea` podem ser membros, e sem 3D.** Qualquer outro tipo gera UserException no `AddChart`. Crie os membros com `New(oExcel)`, nunca com `New(Nil)`: o parâmetro é obrigatório e é ele que registra o membro para limpeza. `AddChart` retorna o próprio membro. A legenda e o título ficam no combo.

```tlpp
oCombo := oExcel:NewChart("combo")
oCombo:SetTitulo("Vendas x Margem")
oCombo:SetLegenda(lExibir = .T., cPosicao = "b")
oCol := custom.excel.chart.ChartBar():New(oExcel)
oCol:AddSerie(cNome = "Vendas", cRefVal = "Dados!$B$2:$B$7", cRefCat = "Dados!$A$2:$A$7", cCor = "FF4F81BD")
oLin := custom.excel.chart.ChartLine():New(oExcel)
oLin:AddSerie(cNome = "Margem", cRefVal = "Dados!$C$2:$C$7", cRefCat = "Dados!$A$2:$A$7", cCor = "FFC0504D")
oCombo:AddChart(oCol)
oCombo:AddChart(oChart = oLin, lEixoSecundario = .T.)	//eixo de valores à direita
oCombo:Insert(nLinha = 9, nColuna = 1, nLinha2 = 24, nColuna2 = 9)
```

## 6. Exemplos por tipo

```tlpp
//Dados: A1:D7 (Mês, Sudeste, Sul, Nordeste), escritos linha a linha com Pos():SetValue()
cPlan := "Dados"

//Colunas agrupadas
oGraf := oExcel:NewChart("col")
oGraf:SetTitulo("Vendas por Região")
oGraf:AddSerie(cNome = "Sudeste", cRefVal = cPlan+"!$B$2:$B$7", cRefCat = cPlan+"!$A$2:$A$7", aVal = aSudeste, aCat = aMeses)
oGraf:AddSerie(cNome = "Sul", cRefVal = cPlan+"!$C$2:$C$7", cRefCat = cPlan+"!$A$2:$A$7")
oGraf:Insert(nLinha = 1, nColuna = 9, nLinha2 = 16, nColuna2 = 17)

//Barras horizontais empilhadas, legenda no rodapé, cores fixas
oGraf := oExcel:NewChart("bar")
oGraf:SetAgrupamento("stacked")
oGraf:SetLegenda(lExibir = .T., cPosicao = "b")
oGraf:AddSerie(cNome = "Sudeste", cRefVal = cPlan+"!$B$2:$B$7", cRefCat = cPlan+"!$A$2:$A$7", cCor = "FF4F81BD")
oGraf:Insert(nLinha = 18, nColuna = 9, nLinha2 = 33, nColuna2 = 17)

//Linha com marcadores e uma série suavizada
oGraf := oExcel:NewChart("line")
oGraf:SetMarcador(.T.)
oSerie := oGraf:AddSerie(cNome = "Nordeste", cRefVal = cPlan+"!$D$2:$D$7", cRefCat = cPlan+"!$A$2:$A$7")
oSerie:SetSuavizar(.T.)
oGraf:Insert(nLinha = 1, nColuna = 19, nLinha2 = 16, nColuna2 = 27)

//Pizza com percentual (3D opcional)
oGraf := oExcel:NewChart("pie")
oGraf:Set3D(lAtivo = .T., nRotX = 30)
oGraf:SetRotulos(lExibir = .T., lPercentual = .T., lCategoria = .F.)
oGraf:AddSerie(cNome = "Sudeste", cRefVal = cPlan+"!$B$2:$B$7", cRefCat = cPlan+"!$A$2:$A$7")
oGraf:Insert(nLinha = 18, nColuna = 19, nLinha2 = 33, nColuna2 = 27)

//Dispersão: a série usa SetXY
oGraf := custom.excel.chart.ChartScatter():New(oExcel)
oGraf:SetEstilo("marker")
oGraf:SetLegenda(lExibir = .F.)
oSerie := oGraf:AddSerie(cNome = "Sorvetes")
oSerie:SetXY(cRefX = cPlan+"!$F$2:$F$11", cRefY = cPlan+"!$G$2:$G$11", aX = aTemp, aY = aSorvete)
oGraf:Insert(nLinha = 35, nColuna = 9, nLinha2 = 50, nColuna2 = 17)

//Bolhas: SetXY + SetTamanho
oGraf := oExcel:NewChart("bolhas")
oSerie := oGraf:AddSerie(cNome = "Indicador")
oSerie:SetXY(cRefX = cPlan+"!$L$2:$L$7", cRefY = cPlan+"!$M$2:$M$7")
oSerie:SetTamanho(cRefTam = cPlan+"!$N$2:$N$7")
oGraf:Insert(nLinha = 60, nColuna = 1, nLinha2 = 75, nColuna2 = 9)

//Rosca
oGraf := oExcel:NewChart("rosca")
oGraf:SetFuro(60)
oGraf:SetRotulos(lExibir = .T., lPercentual = .T.)

//Pizza de pizza: as 2 últimas fatias vão para o secundário
oGraf := oExcel:NewChart("ofpie")
oGraf:SetTipo("pie")
oGraf:SetDivisao(cTipo = "pos", nValor = 2)

//Ações (OHLC)
oGraf := oExcel:NewChart("stock")
oGraf:AddSerie(cNome = "Abertura", cRefVal = cPlan+"!$G$2:$G$7", cRefCat = cPlan+"!$F$2:$F$7")
oGraf:AddSerie(cNome = "Máxima", cRefVal = cPlan+"!$H$2:$H$7", cRefCat = cPlan+"!$F$2:$F$7")
oGraf:AddSerie(cNome = "Mínima", cRefVal = cPlan+"!$I$2:$I$7", cRefCat = cPlan+"!$F$2:$F$7")
oGraf:AddSerie(cNome = "Fechamento", cRefVal = cPlan+"!$J$2:$J$7", cRefCat = cPlan+"!$F$2:$F$7")

//Área empilhada, radar preenchido, superfície 3D
oExcel:NewChart("area"):SetAgrupamento("stacked")
oExcel:NewChart("radar"):SetEstilo("filled")
oExcel:NewChart("superficie"):Set3D(.T.)
```

Os três últimos exemplos só mostram a configuração. Na prática, guarde o objeto em uma variável para chamar `AddSerie` e `Insert`.

Exemplo completo no repositório: `exemplo/custom.excel.grafico.teste.tlpp` (funções `tstGrafico`, `tstGrafico2`, `tstGraficoComb`). Elas aceitam uma instância externa (`oExcelPar`) para adicionar a planilha de gráficos em um arquivo já aberto.
