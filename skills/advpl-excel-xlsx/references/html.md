# Html — `Save2Html` e classe `custom.excel.Html`

## Sumário
1. Quando usar e ciclo de vida
2. Save2Html
3. NewHtml — métodos do objeto Html
4. Inputs por coluna
5. Blocos de código (planilha, linha, coluna, célula)
6. JavaScript no html e envio por formulário
7. Fórmulas no navegador
8. O que é convertido e limitações
9. Exemplo completo

---

## 1. Quando usar e ciclo de vida

Use quando o usuário pedir o relatório/planilha "em html", "para abrir no navegador", "com campos editáveis", "para enviar por e-mail como página" ou "tela web a partir do Excel". O mesmo objeto `custom.excel.Xlsx` gera o xlsx (`Save`) e o html (`Save2Html`). O html é **um único arquivo** com CSS e JavaScript embutidos; só o Chart.js vem de URL e só quando há gráfico.

```
New → (estilos) → ADDPlan → escrever → [oHtml := NewHtml() + configuração] → Save / Save2Html (qualquer ordem) → [OpenApp()] → Destroy() → FreeObj(oExcel)
```

- Tipos `"M"` e `"B"`: `Save` e `Save2Html` em qualquer ordem, sempre antes do `Destroy()`.
- Tipo `"A"`: o primeiro `Save`/`Save2Html` **encerra a gravação** (depois disso `Pos`/`ADDPlan`/bulk geram exceção).
- `Save2Html` **depois** do `Save` funciona sempre, sem `NewHtml()`: as linhas (tipo `"A"`) e as imagens ficam preservadas após o `Save` até o `Destroy` (`oExcel:lHtml`, padrão `.T.`). O `NewHtml()` serve só para configurar o html.
- Para não preservar (só xlsx, economizando disco), defina `oExcel:lHtml := .F.` antes do `Save`; nesse caso o `Save2Html` depois do `Save` não funciona no tipo `"A"` nem mostra imagens.
- Os temporários de `\tmpxls\` só são removidos no `Destroy()`/`Close()` — sempre chame o `Destroy()`.
- `OpenApp()` abre o último arquivo gerado (xlsx ou html). Não usar em job/REST.
- O objeto `oHtml` é liberado pelo `Destroy()` do Excel: **não** chamar `FreeObj(oHtml)`.

## 2. Save2Html

`oExcel:Save2Html(cLocal, oHtml)` → caminho do `.html` (mesmo nome do xlsx).

| Parâmetro | Descrição |
|---|---|
| `cLocal` | Pasta no client (`GetTempPath()`) ou no servidor iniciando com `\` (fica no servidor). Padrão `GetTempPath()`, ou `\` sem interface |
| `oHtml` | Configuração criada com `NewHtml()`. Opcional |

## 3. NewHtml — métodos do objeto Html

`oHtml := oExcel:NewHtml()`. Todos os métodos retornam `self`.

| Método | Uso |
|---|---|
| `SetTitulo(cTitulo)` | `<title>` |
| `SetColInput(xColuna, cTipo, nLinIni, nLinFim, nPlanilha, cNome, lSomenteLeitura, aOpcoes)` | Coluna como input (seção 4) |
| `SetBlocoPlanilha(bBloco)` / `SetBlocoLinha(bBloco)` / `SetBlocoColuna(xColuna, bBloco, nPlanilha)` / `SetBlocoCelula(bBloco)` | Blocos de código (seção 5) |
| `AddCss(cCss)` / `AddScript(cJs)` / `AddHead(cHtml)` | Conteúdo adicional |
| `SetAction(cUrl, cMetodo, cTxtBotao)` | `<form>` com botão de envio |
| `SetChartJs(cUrl)` | URL do Chart.js (padrão CDN jsdelivr 4.4.1). Sem internet: caminho acessível ao navegador |
| `SetFormulaJs(cUrl)` | URL do formula.js (padrão CDN jsdelivr `@formulajs/formulajs@4.6.1`), carregado só quando há fórmula. Sem internet: caminho acessível ao navegador; `""` não carrega (seção 7) |
| `SetAbas(lAbas)` | `.T.` abas por planilha (padrão), `.F.` todas em sequência |
| `SetPlanilhas(aPlanilhas)` | Somente as planilhas informadas (número ou nome) |
| `SetGrade(lGrade)` | Força exibir/ocultar linhas de grade |
| `SetLarguraAuto(lAuto)` | Largura das colunas pelo conteúdo |

`xColuna` (em `SetColInput` e `SetBlocoColuna`) aceita número, letra (`"B"`) ou nome da coluna/campo de uma **tabela formatada** (`AddTabela`/`Alias2Tab`). `nPlanilha` padrão = planilha atual do Excel no momento da chamada (`0` = todas): configure depois de escrever na planilha certa ou informe `nPlanilha`.

## 4. Inputs por coluna

```tlpp
oHtml:SetColInput(xColuna = "B", cTipo = "number", nLinIni = 2, cNome = "qtd")
oHtml:SetColInput(xColuna = "Status", cTipo = "select", aOpcoes = {{"A","Aprovado"},{"R","Reprovado"}})	//coluna de tabela formatada
oHtml:SetColInput(xColuna = 6, cTipo = "checkbox", nLinIni = 2, lSomenteLeitura = .T.)
```

- `cTipo`: `auto` (pelo tipo do valor: N→number, D→date, T→datetime, L→checkbox, demais→text), `text`, `number`, `date`, `datetime`, `checkbox`, `select`, `textarea`.
- Linhas: padrão a partir da linha 1; em coluna de tabela formatada, só as linhas de dados (sem cabeçalho e total). Informe `nLinIni` para não transformar o cabeçalho em input.
- `name`: referência da célula (`B3`) ou `cNome_linha` (`qtd_3`).

## 5. Blocos de código

Recebem um Json alterável (textos em CP-1252; `antes`, `depois` e `html` são HTML sem escape).

| Bloco | Chamada | Campos do Json |
|---|---|---|
| Planilha | `Eval(b, jPlan, nPlanilha, oExcel)` | `titulo`, `antes`, `depois`, `class`, `attr` (Json), `ignorar` |
| Linha | `Eval(b, jLin, nLinha, nPlanilha, oExcel)` | `class`, `style`, `attr` (Json), `antes`/`depois` (HTML fora do `<tr>`), `ignorar` |
| Célula / Coluna | `Eval(b, jCel, nLinha, nColuna, nPlanilha, oExcel)` | `valor`, `tipo` (C,N,L,D,T,U), `texto`, `html`, `formula`, `estilo`, `ref`, `tag`, `class`, `style`, `attr` (Json), `input` (Json), `antes`/`depois` (dentro da célula), `ignorar` |

- O bloco de coluna executa **depois** do bloco de célula.
- `valor` de data e hora (`tipo` `T`) é o serial numérico do Excel; data (`D`) vem como Date.
- Criar input em uma célula: `jCel["input"] := JsonObject():New()` com `tipo`, `nome`, `somenteLeitura`, `opcoes`. Remover: `jCel["input"] := nil`.
- Blocos são executados para todas as células: mantenha-os leves em grande volume.

```tlpp
oHtml:SetBlocoLinha({|jLin,nLinha,nPlanilha,oExcel| If(nLinha>1 .and. nLinha%2==0, jLin["class"] := "par", nil) })
oHtml:SetBlocoColuna(xColuna = "D", bBloco = {|jCel| If(jCel["tipo"]=="N" .and. jCel["valor"]<0, jCel["style"] += ";color:red", nil) })
oHtml:SetBlocoCelula({|jCel| If(jCel["tipo"]!="U", jCel["attr"]["data-ref"] := jCel["ref"], nil) })
oHtml:AddCss("tr.par td{background:#f7f7f7}")
```

## 6. JavaScript no html e envio por formulário

`window.xlsxHtml`: `getDados()` (inputs em `{"Planilha":[{lin,col,ref,nome,valor}]}`), `valor(td)`, `valores("Plan!A1:B5")`, `celulas("A1:B5", secao)`, `mostra(nomeOuNumero)`, `colNum`, `colLetra`. Evento `xlsxhtml:pronto` no `document` (`e.detail` = `xlsxHtml`).

`SetAction(cUrl)`: as planilhas ficam em um `<form>`; cada input é enviado pelo nome e o campo oculto `xh_dados` leva `JSON.stringify(getDados())`.

## 7. Fórmulas no navegador

Cada célula com fórmula leva o texto em `data-f` e o valor calculado pelo AdvPL. No navegador:

- Ao alterar um input, só as fórmulas que dependem da célula são recalculadas (grafo de dependências, inclusive entre planilhas, nomes definidos e `Tab[Coluna]`), e o resultado é exibido com o formato numérico do estilo.
- Fórmulas sem valor gravado são calculadas ao abrir a página.
- Motor próprio (mesma gramática do Excel) com as funções mais usadas: agregação, `SUBTOTAL`/`AGGREGATE`, `SUMIF(S)`/`COUNTIF(S)`, lógicas, `LET`/`LAMBDA`/`MAP`/`REDUCE`, texto (`TEXT` na localidade do `SetCalcLocal`), datas, `VLOOKUP`/`XLOOKUP`/`INDEX`/`MATCH`/`OFFSET`/`INDIRECT`, matrizes dinâmicas (só a célula âncora). As demais funções do Excel vêm do **formula.js** (`SetFormulaJs`).
- Função que não existe em nenhum dos dois (ou formula.js não carregado): a célula mantém o valor calculado pelo AdvPL.
- Célula com fórmula configurada como input: o input prevalece (não recalcula).

| `window.xlsxHtml` | Uso |
|---|---|
| `recalcula(ref)` | Recalcula as dependentes de `ref` (ex.: `"Plan!B2:B9"`, após alterar células por script); sem argumento recalcula todas |
| `calcula(cFormula, plan)` | Avalia uma fórmula avulsa na planilha (número ou nome; padrão a visível) |
| `valorFormula(ref)` | Valor atual da célula (calculado) |
| `formata(valor, codigo)` | Formata como o formato numérico do Excel (`"#,##0.00"`) |
| `addFuncao(nome, fn, flags)` | Registra/substitui função: `fn(args, contexto)` recebe os argumentos avaliados; `flags` `"E"` aplica elemento a elemento, `"V"` volátil, `"L"` recebe os nós sem avaliar |

Evento `xlsxhtml:recalculo` no `document` (`e.detail.celulas` = `<td>` recalculados).

```js
document.addEventListener('xlsxhtml:recalculo',function(e){console.log(e.detail.celulas.length+' células');});
xlsxHtml.addFuncao('DOBRO',function(a){return a[0]*2;},'E');
```

## 8. O que é convertido e limitações

| Excel | Html |
|---|---|
| Valores e formato numérico do estilo | Texto formatado em pt-BR (decimal `,`, milhar `.`), valor bruto em `data-v` |
| Fórmulas | Valor calculado pelo AdvPL (ou o `xValor` informado) e recálculo no navegador ao alterar inputs (seção 7) |
| Estilos, mesclagem, larguras, colunas/linhas ocultas, altura | CSS `.sN` (degradê linear e `path`, recuo esquerdo/direito, texto girado), `colspan`/`rowspan`, `<colgroup>` com largura fixa da tabela e altura padrão das linhas da planilha |
| Texto maior que a célula | Ocupa as vizinhas vazias como no Excel (mesclada não transborda; alinhado à direita corta à esquerda) |
| Cor da aba (`tabColor`) | Aba do html na mesma cor |
| Grade, `AddPane`, `AutoFilter` | Grade, linhas/colunas fixas, filtro/ordenação no cabeçalho (estilo Excel: lista de valores com pesquisa, carga progressiva e "Selecionar Tudo", árvore ano/mês/dia nas datas; Filtros de Texto, Número e Data conforme o tipo da coluna) |
| `NivelLinha`, `SetColLevel` | Barra lateral (linhas) e superior (colunas) com +/- por grupo, mais os botões de nível. Respeita grupos recolhidos e `summaryBelow`/`summaryRight` |
| Tabela formatada | Cores aproximadas de `TableStyleLight/Medium/Dark N`, listras, bordas, total e filtro |
| Formatação condicional | `cellIs`, duplicados/exclusivos, `FormatCondTempo`, `DataBar` (degradê), `IconSet`. A prioridade vale entre todas as regras da planilha. **Não** avalia `FormatCond(...,"FORMULA",...)` |
| Comentário / hyperlink interno | Dica + marcador / link para a célula de destino |
| Imagens | `<img>` base64 na célula âncora, com rotação (90/270 graus: a área informada é a da imagem já girada, como no Excel) |
| Gráficos | Chart.js na área do `Insert`, com aparência próxima do Excel (cores do tema, grade, escala a partir de zero, rótulos de dados, colunas/pizza 3D, ações em candlestick, bolhas, pizza de pizza, combinado com eixo secundário). Usa os arrays da série ou lê as células das referências. Superfície não suportada; pizza de barras é desenhada como pizza de pizza |

## 9. Exemplo completo

```tlpp
#include "tlpp-core.th"
#include "protheus.ch"
Namespace custom.relatorio

User Function InvHtml() As Character
	Local oExcel As Object
	Local oHtml As Object
	Local oStyCab As Object
	Local oErro As Object
	Local nCont As Numeric
	Local cHtml := "" As Character

	oExcel	:= custom.excel.Xlsx():New(cNomeFile = "Inventario", cTipo = "M")
	Try
		oStyCab	:= oExcel:NewStyle():Setfont(oExcel:AddFont(nTamanho = 11, xCorRGB = "FFFFFFFF", lNegrito = .T.)):Setfill(oExcel:CorPreenc("FF4F81BD"))
		oExcel:ADDPlan(cNome = "Estoque")
		oExcel:Pos(1,1):SetValue("Produto"):SetStyle(oStyCab)
		oExcel:Pos(1,2):SetValue("Saldo"):SetStyle(oStyCab)
		oExcel:Pos(1,3):SetValue("Contagem"):SetStyle(oStyCab)
		For nCont:=1 to 10
			oExcel:Pos(nCont+1,1):SetValue("Produto "+cValToChar(nCont))
			oExcel:Pos(nCont+1,2):SetValue(nCont*10)
			oExcel:Pos(nCont+1,3):SetValue(0)
		Next
		oExcel:AutoFilter(nLinha = 1, nColuna = 1, nLinha2 = 11, nColuna2 = 3)
		oExcel:AddPane(nySplit = 1)

		oHtml	:= oExcel:NewHtml()
		oHtml:SetTitulo("Inventário")
		oHtml:SetColInput(xColuna = "C", cTipo = "number", nLinIni = 2, cNome = "contagem")
		oHtml:SetAction(cUrl = "https://meuservidor/rest/inventario")

		oExcel:Save(GetTempPath())
		cHtml	:= oExcel:Save2Html(GetTempPath(), oHtml)
		oExcel:OpenApp()
	Catch oErro
		ConOut("Erro: "+oErro:Description)
	EndTry
	oExcel:Destroy()
	FreeObj(oExcel)
Return cHtml
```

Exemplos com os tipos A, M e B: `exemplo/custom.excel.html.teste.tlpp` (`u_tstHtml`, `u_tstHtmlA`, `u_tstHtmlB`).
