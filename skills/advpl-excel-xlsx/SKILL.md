---
name: advpl-excel-xlsx
description: Gera, lê e edita arquivos Excel (.xlsx) em AdvPL/TLPP/Protheus usando a biblioteca advpl-excel — classe custom.excel.Xlsx():New() (antiga YExcel():New()). Cobre planilhas, estilos, fórmulas, tabelas formatadas, preenchimento por alias/query (Alias2Tab), bulk de grande volume, formatação condicional, imagens, gráficos (custom.excel.chart), proteção/senha, leitura de xlsx e gravação dos mesmos dados em html (Save2Html, com colunas como input e blocos de código). Use SEMPRE que o usuário pedir para gerar/exportar relatório ou planilha Excel/xlsx em Protheus, ADVPL ou TLPP, mencionar custom.excel.Xlsx, YExcel, YExcel_Style, Save2Html, "exportar para excel", "planilha com gráfico", "xlsx com senha", "ler planilha xlsx no Protheus", "relatório em html a partir do excel", "tabela html com campos editáveis", ou editar um fonte que já usa YExcel (nesse caso migrar para custom.excel.Xlsx). Não use para FWMsExcel/FWMsExcelXlsx/FWPrinterXlsx, a menos que o usuário peça para migrar para a advpl-excel.
---

# advpl-excel: gerar Excel com `custom.excel.Xlsx`

A biblioteca **advpl-excel** (TLPP, namespace `custom.excel`) grava o XML do .xlsx direto, sem precisar de Excel no servidor. A classe principal é `custom.excel.Xlsx`. `YExcel` (arquivo `YEXCEL.prw`) é só uma subclasse de compatibilidade (`Class YExcel From custom.excel.Xlsx`): tem os mesmos métodos e parâmetros, e existe apenas para não quebrar código legado.

## 1. Decisão inicial

| Situação | O que fazer |
|---|---|
| Código novo | Usar `custom.excel.Xlsx():New(...)` e as classes `custom.excel.*` / `custom.excel.chart.*` |
| Fonte existente com `YExcel()`, `YExcel_Style`, `YExcel_Table`, `YExcel_DateTime`… | **Migrar para o padrão novo** no trecho em que for mexer, seguindo [references/migracao-yexcel.md](references/migracao-yexcel.md) |
| Fonte `.prw` (AdvPL) | Usar o nome completo da classe (`custom.excel.Xlsx():New(...)`), mas **argumentos posicionais**, porque argumento nomeado (`cNome = "x"`) só existe em TLPP |
| Fonte `.tlpp` | Preferir argumentos nomeados (`oExcel:ADDPlan(cNome = "Dados")`). Eles documentam a chamada e evitam erro de posição em métodos com muitos parâmetros |

Antes de alterar um fonte existente, veja se o usuário fez alterações manuais (git diff/status). Se algo parecer errado, pergunte antes de "corrigir".

Se o projeto tiver `src/custom.excel.tlpp`, confira a assinatura de métodos pouco comuns com `Grep "^Method NomeMetodo\("` (use `-i`). O código é a fonte da verdade: se a wiki divergir dele, siga o código.

## 2. Escolha do `cTipo` no `New`

`custom.excel.Xlsx():New(cNomeFile, cFileOpen, cTipo)`

| cTipo | Quando usar | Restrição |
|---|---|---|
| `"A"` (padrão) | Relatórios comuns, escrita direta em arquivo, menor uso de memória | **A escrita precisa ser sequencial**: linhas crescentes e, dentro da linha, colunas crescentes. Voltar gera `UserException`. Não permite ler de volta (`GetValue`, `LinTam`, `ColTam`) |
| `"M"` | Escrever em ordem aleatória, ler ou alterar células já escritas, editar arquivo aberto | Tudo fica em memória, então não serve para volumes enormes |
| `"B"` | Grande volume (100 mil linhas ou mais) e bulk (`SetValueBulk`) | Usa tabelas temporárias (SQLite_TMP/TMPDB) |

`New(cFileOpen = "c:\temp\arq.xlsx")` abre um arquivo existente. Se o arquivo tiver senha, use `New()` seguido de `Open(cArq, cSenha)`, porque `New(cFileOpen)` gera exceção em arquivo criptografado.

## 3. Template base (TLPP)

```tlpp
#include "tlpp-core.th"
#include "protheus.ch"
Namespace custom.<modulo>

/*/{Protheus.doc} RelExcel
Gera relatório em Excel ...
@type function
@author ...
@since dd/mm/aaaa
@version 1.0
@return character, caminho do arquivo gerado
/*/
User Function RelExcel() As Character
	Local oExcel As Object
	Local oStyCab As Object
	Local oStyMoeda As Object
	Local nFonteCab As Numeric
	Local nCorCab As Numeric
	Local nBorda As Numeric
	Local nLin As Numeric
	Local cArquivo As Character
	Local cAlias As Character

	oExcel		:= custom.excel.Xlsx():New(cNomeFile = "RelExcel", cTipo = "A")

	//Estilos (criar antes de escrever as células)
	nFonteCab	:= oExcel:AddFont(nTamanho = 11, xCorRGB = "FFFFFFFF", cNome = "Calibri", cfamily = "2", lNegrito = .T.)
	nCorCab		:= oExcel:CorPreenc("FF4F81BD")
	nBorda		:= oExcel:Borda("ALL")
	oStyCab		:= oExcel:NewStyle():Setfont(nFonteCab):Setfill(nCorCab):Setborder(nBorda):SetaValores({oExcel:Alinhamento(cHorizontal = "center", cVertical = "center")})
	oStyMoeda	:= oExcel:NewStyle():SetnumFmt(44):Setborder(nBorda)	//44 = moeda R$

	oExcel:ADDPlan(cNome = "Dados")
	oExcel:AddTamCol(nMin = 1, nMax = 1, nWidth = 12)
	oExcel:AddTamCol(nMin = 2, nMax = 2, nWidth = 40)
	oExcel:AddTamCol(nMin = 3, nMax = 3, nWidth = 16)
	oExcel:AddPane(nySplit = 1, nxSplit = 0)	//Congela a linha 1

	//Cabeçalho (linha 1, colunas em ordem crescente)
	oExcel:Pos(1,1):SetValue("Código"):SetStyle(oStyCab)
	oExcel:Pos(1,2):SetValue("Descrição"):SetStyle(oStyCab)
	oExcel:Pos(1,3):SetValue("Valor"):SetStyle(oStyCab)

	cAlias	:= MpSysOpenQuery("SELECT ... FROM " + RetSqlName("SB1") + " WHERE D_E_L_E_T_ = ' '")
	nLin	:= 1
	While (cAlias)->(!Eof())
		nLin++
		oExcel:Pos(nLin,1):SetValue(AllTrim((cAlias)->B1_COD))
		oExcel:Pos(nLin,2):SetValue(AllTrim((cAlias)->B1_DESC))
		oExcel:Pos(nLin,3):SetValue((cAlias)->B1_PRV1):SetStyle(oStyMoeda)
		(cAlias)->(DbSkip())
	EndDo
	(cAlias)->(DbCloseArea())

	oExcel:AutoFilter(nLinha = 1, nColuna = 1, nLinha2 = Max(nLin,2), nColuna2 = 3)

	cArquivo	:= oExcel:Save(GetTempPath())	//Retorna o caminho do arquivo gerado
	oExcel:OpenApp()						//Só com interface (SmartClient); nunca em job/schedule
	oExcel:Destroy()						//Libera objetos internos e temporários do servidor
	FreeObj(oExcel)
Return cArquivo
```

Um relatório com muitas colunas vindas de uma query costuma ficar mais simples com `Alias2Tab` (a coluna, o título e a máscara vêm da SX3). Veja [references/tabelas-bulk-alias.md](references/tabelas-bulk-alias.md).

## 4. Regras que evitam erro

Cada regra existe porque o erro correspondente acontece na prática:

- **No tipo "A", escreva em ordem.** Linha 1 → 2 → 3, e dentro da linha, coluna 1 → 2 → 3. Com duas tabelas lado a lado, preencha linha por linha e não tabela por tabela (é o que fazem os exemplos de gráfico). Se a lógica precisar voltar, use `cTipo = "M"`.
- **Fórmula:** nomes das funções em inglês, `,` como separador e sem `=` inicial. O valor é calculado pela própria classe no `Save`/`Save2Html` (motor AdvPL com as funções do Excel): `SetValue(, "SUM(A2:A3)")`. Se `xValor` for informado, ele é mantido sem recalcular. O que não puder ser calculado (funções do ambiente do Excel, como CUBE*/RTD/WEBSERVICE) fica para o Excel recalcular ao abrir. Detalhes, `AddFuncFormula`, `SetCalcLocal` e a lista de funções em [references/formulas.md](references/formulas.md). Para montar referências use `oExcel:Ref(nLin,nCol)` (`Ref(3,3,.T.,.T.)` retorna `$C$3`).
- **Nome definido** (`AddNome`) deve existir antes da fórmula que o usa.
- **Estilos são criados uma vez e reutilizados.** Cada `NewStyle`/`AddFont`/`CorPreenc` dentro de um laço infla o styles.xml. Para variar por linha ou coluna use herança (`NewStyle(oPai)`) ou regras (`NewStyRules`).
- **Nomes reais dos parâmetros:** `AddFont(..., xCorRGB = ...)` (mas `Font(..., cCorRGB = ...)` na formatação condicional), `AddStyles(..., nborderId = ...)`, `Cell(..., xStyle = ...)` e `Borda(cTipo = "ALL"|"T"|"B"|"L"|"R"|combinação)`.
- **Parâmetros que confundem:**
  - `CorPreenc(cFgCor, cBgCor)` recebe a cor de frente primeiro. `Preenc(cBgCor, cFgCor)`, usado na formatação condicional, recebe o fundo primeiro.
  - As letras de borda também mudam: `Borda` usa T/B/L/R e `ObjBorda` usa C/B/E/D.
- **Cores:** ARGB (`"FF4F81BD"`) ou RGB (`"4F81BD"`).
- **Um único `AutoFilter` por planilha.** Uma tabela formatada usa `oTab:AddFilter()` em vez dele.
- **Tabela formatada:** `oTab:Finish()` é obrigatório. `AddTotal` vem antes de `AddTotais()`. Use `AddColumn(cNome, nStyle, cNameField)` para dar um estilo padrão e um apelido à coluna. Depois de um `oTab:Alias2Tab`, `Cell` e `AddTotal` aceitam o nome do campo. O estilo de cabeçalho de um campo vai em `NewFldTab(..., xStyleCab = oSty)` (ou `jCab[campo]["stylecab"]`).
- **Bulk:** `FlushBulk()` antes do `Save`.
- **Subtotais:** `DefSubTotal`/`AddSubTotal` antes do `Alias2Tab`.
- **Agrupamento:** `SetsumRight` antes de `SetColLevel`. `NivelLinha(n)` vale para as linhas seguintes até `NivelLinha()` sem argumentos.
- **Combo:** só `ChartBar`, `ChartLine` e `ChartArea` sem 3D podem ser membros, e eles são criados com `New(oExcel)` (`New(Nil)` falha).
- **Gráfico:** cada gráfico recebe `Insert(...)` uma única vez, antes do `Save`. `Set3D` só existe para barras, linhas, pizza, área e superfície; nos outros gera exceção.
- **Html (`Save2Html`):** o html mostra o valor das fórmulas calculado pelo motor AdvPL (ou o `xValor` informado). No tipo "A", o primeiro `Save`/`Save2Html` encerra a gravação. As linhas e imagens ficam preservadas após o `Save` até o `Destroy`, então `Save2Html` pode vir depois do `Save` sem `NewHtml()` (o `NewHtml()` serve só para configurar o html). Não chame `FreeObj` no objeto do `NewHtml()` (o `Destroy` libera). Detalhes em [references/html.md](references/html.md).
- **Datas:** `SetValue(date())` grava uma data. Para data e hora use `SetValue(oExcel:GetDateTime(dData = dDate, cTime = cHora))` ou `SetDateTime(dDate, cTime)`. Na leitura, uma célula de data e hora retorna um objeto `DateTime` (`:GetDate()`, `:GetTime()`).

## 5. Ciclo de vida e liberação de memória

```
New → (estilos) → ADDPlan → escrever → Save(cLocal) → [OpenApp()] → Destroy() → FreeObj(oExcel)
```

Para gerar também o html: `oHtml := oExcel:NewHtml()` (opcional, para configurar), depois `oExcel:Save2Html(cLocal, oHtml)` antes ou depois do `Save`, sempre antes do `Destroy()`. `OpenApp()` abre o último arquivo gerado.

- `Save(cLocal)` retorna o caminho do arquivo. Se `cLocal` for uma pasta local (ex.: `GetTempPath()`), o arquivo é copiado para o client. Um caminho do servidor (ex.: `"\temp"`) mantém o arquivo no servidor.
- `OpenApp()` faz ShellExecute no client. Não o use em job, schedule, REST ou `RpcSetEnv` sem interface.
- `Destroy()` (equivale a `Close()`) remove as pastas e tabelas temporárias e libera tudo o que a instância criou e registrou em `aCleanObj`: estilos (`NewStyle`), regras (`NewStyRules`, `NewRuleLine`), tabelas (`AddTabela`), gráficos e séries (`NewChart`/`AddSerie` e classes `ChartX():New(oExcel)`), campos de bulk, `jCab` passado ao `NewFldTab` e objetos `DateTime`. **Não chame `FreeObj` nesses objetos depois do `Destroy`**, porque seria liberar duas vezes. Depois do `Destroy` chame apenas `FreeObj(oExcel)`.
- Arrays locais criados por você (`aCampos`, `aOnlyFieds`, arrays de dados) devem ser liberados com `FwFreeArray()`. Objetos ou JSON criados por você e não entregues à lib devem ser liberados com `FreeObj()`. Faça isso em todos os fluxos de saída, inclusive no `Catch`/tratamento de erro, e sem liberar variável não inicializada.
- Havendo tratamento de erro, garanta o `Destroy()` também no caminho de erro, senão ficam temporários em `\tmpxls\` no servidor:

```tlpp
Try
	...
	cArquivo := oExcel:Save(GetTempPath())
Catch oErro
	ConOut("Erro ao gerar Excel: " + oErro:Description)
EndTry
If oExcel <> Nil
	oExcel:Destroy()
	FreeObj(oExcel)
EndIf
```

- `Gravar(cLocal, lAbrir, lDelSrv)` é da API v1 e faz Save + OpenApp opcional + Close numa chamada. Em código novo prefira a sequência explícita.

## 6. Encoding dos fontes gerados

Os fontes AdvPL/TLPP são **CP-1252 sem BOM, com CRLF e indentação por TAB**, e os comentários ficam em PT-BR (ProtheusDOC). Um fonte **novo** pode ser escrito e depois convertido para CP-1252 via PowerShell. Um fonte **existente com acentos** não deve ser editado com Edit/Write (a gravação em UTF-8 corrompe o arquivo inteiro): edite via PowerShell, lendo e gravando em CP-1252:

```powershell
$enc = [Text.Encoding]::GetEncoding(1252)
$t = [IO.File]::ReadAllText($p, $enc)          #confira que a leitura retornou conteúdo
#... $t = $t.Replace($ancoraUnica, $novo)       #valide que a âncora ocorre exatamente 1 vez
[IO.File]::WriteAllBytes($p, $enc.GetBytes($t))
```

Monte as strings entre parênteses ou em variáveis: em PowerShell a `,` tem precedência sobre o `+`. Depois de gravar, valide que não há BOM, nem LF sem CR, nem a sequência `EF BF BD`. Se o arquivo for corrompido, restaure as linhas afetadas a partir do `git show HEAD:<arquivo>`.

## 7. Onde consultar cada assunto

| Preciso de… | Leia |
|---|---|
| New/Open/Save, planilhas, Pos/PosR/Ref, SetValue, fórmulas, datas, leitura (GetValue/LinTam/ColTam), merge, larguras e alturas, congelar, filtro, nomes, hyperlink, comentário, agrupamento, imagem | [references/api-core.md](references/api-core.md) |
| Fonte, cor, borda, alinhamento, formato numérico, herança de estilo, regras de estilo, formatação condicional (dxf, DataBar, IconSet, duplicados, datas) | [references/estilos.md](references/estilos.md) |
| Tabela formatada do Excel, preencher a partir de alias ou query (Alias2Tab, NewFldTab, subtotais, NewRuleLine), bulk de grande volume | [references/tabelas-bulk-alias.md](references/tabelas-bulk-alias.md) |
| Gráficos (colunas, barras, linha, pizza, dispersão, área, ações, radar, rosca, bolhas, pizza de pizza, superfície, combinação) | [references/graficos.md](references/graficos.md) |
| Senha de abertura (SetPassword/Open), proteção de planilha, área de impressão, cabeçalho/rodapé, orientação | [references/protecao-impressao.md](references/protecao-impressao.md) |
| Cálculo automático de fórmulas: funções suportadas, SetCalcFormula, AddFuncFormula, SetCalcLocal, limitações | [references/formulas.md](references/formulas.md) |
| Gravar em html (Save2Html, NewHtml): inputs por coluna, blocos de código por planilha/linha/coluna/célula, CSS/JS, formulário, limitações | [references/html.md](references/html.md) |
| Converter código YExcel/v1 (Cell, AddStyles, Gravar, OpenRead/CellRead) para o padrão novo | [references/migracao-yexcel.md](references/migracao-yexcel.md) |
| Receitas completas prontas (relatório SX3 com subtotal, tabela, gráfico, senha, leitura) | [references/exemplos.md](references/exemplos.md) |

## 8. Requisitos do ambiente (vale mencionar ao usuário quando relevante)

- A lib precisa estar compilada no RPO: `custom.excel.tlpp`, `custom.excel.chart.tlpp`, `custom.excel.crypto.tlpp`, `custom.excel.html.tlpp` (Save2Html), os fontes do cálculo de fórmulas `custom.excel.calc*.tlpp` e, para legado, `YEXCEL.prw`.
- Gráficos no html usam o Chart.js de `cdn.jsdelivr.net`: sem internet no navegador, informe outro caminho com `oHtml:SetChartJs(cUrl)`.
- Os temporários ficam em `\tmpxls\` no RootPath. A compactação usa FZip. Para compatibilidade com LibreOffice em servidor Windows, é possível instalar o 7-Zip e configurar `[GENERAL] LOCAL7ZIP` no appserver.ini.
- `Alias2Tab` com `lSx3 = .T.` precisa de ambiente aberto (dicionário SX3 disponível).
