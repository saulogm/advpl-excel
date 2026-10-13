# Estilos e formatação condicional

## Sumário
1. Modelo de estilos
2. Componentes (fonte, preenchimento, borda, formato, alinhamento, proteção)
3. NewStyle e herança
4. Regras de estilo (NewStyRules)
5. Formatação condicional

---

## 1. Modelo de estilos

Os **componentes** são criados uma vez e cada um devolve um id numérico ou um objeto:
- `AddFont` → id
- `CorPreenc` / `EfeitoPreenc` → id
- `Borda` → id
- `AddFmtNum` / `AddFmt` → id de numFmt
- `Alinhamento` / `Cellprotection` → objeto

Depois eles são combinados em um `Style` com `NewStyle()`, que é aplicado com `SetStyle(oStyle)`.

Crie os estilos **antes** de escrever as células e reutilize-os. Criar um estilo por célula ou dentro de um laço infla o arquivo e deixa a geração lenta.

## 2. Componentes

### Fonte: `AddFont(nTamanho, xCorRGB, cNome, cfamily, cScheme, lNegrito, lItalico, lSublinhado, lTachado)` → id
- Padrões: tamanho 11, cor `"FF000000"`, fonte Calibri, flags `.F.`.
- `xCorRGB` aceita ARGB/RGB em texto ou um índice numérico de cor (ex.: `56`). **O nome do parâmetro é `xCorRGB`**. Já o `Font()` da formatação condicional usa `cCorRGB`.
- `cfamily`: `"0"` não aplicável, `"1"` Roman, `"2"` Swiss, `"3"` Modern, `"4"` Script, `"5"` Decorative.
```tlpp
nFonteBranca := oExcel:AddFont(nTamanho = 11, xCorRGB = "FFFFFFFF", cNome = "Calibri", cfamily = "2", lNegrito = .T.)
nFonteLink   := oExcel:AddFont(nTamanho = 11, xCorRGB = "FF0000FF", cNome = "Calibri", cfamily = "2", lSublinhado = .T.)
```

### Preenchimento
- `CorPreenc(cFgCor, cBgCor, cType = "solid")` → id. Para cor sólida basta o primeiro parâmetro: `CorPreenc("FF4F81BD")`. `cType` também aceita padrões como `gray125`, `lightGray`, `darkHorizontal`, `lightGrid`…
- `EfeitoPreenc(nAngulo = 90, aCores, cType = "linear"|"path", nleft, nright, ntop, nbottom)` → id de gradiente.
```tlpp
nGradLin := oExcel:EfeitoPreenc(nAngulo = 90, aCores = {{"FFFFFF",0},{"0072C4",1}})
nGradCen := oExcel:EfeitoPreenc(aCores = {{"FFFFFF",0},{"0072C4",1}}, ctype = "path", nleft = 0.2, nright = 0.8, ntop = 0.2, nbottom = 0.8)
```

### Borda
- `Borda(cTipo, cCor = "FF000000", cModelo = "thin")` → id. `cTipo`: `"ALL"`, `"T"`, `"B"`, `"L"`, `"R"` ou uma combinação (`"LTB"`).
- Modelos: `thin`, `medium`, `thick`, `dashed`, `dotted`, `double`, `hair`, `dashDot`, `dashDotDot`, `mediumDashed`, `mediumDashDot`, `mediumDashDotDot`, `slantDashDot`, `none`.
- `Border(cleft, cright, ctop, cbottom, cdiagonal, cCorleft, cCorright, cCortop, cCorbottom, cCordiagonal)` define o modelo de cada lado separadamente.
```tlpp
nBordaAll  := oExcel:Borda("ALL")
nBordaVerm := oExcel:Borda(cTipo = "ALL", cCor = "FFFF0000", cModelo = "thick")
```

### Formato numérico
- `AddFmtNum(nDecimal = 0, lMilhar = .F., cPrefixo, cSufixo, cNegINI = "-", cNegFim, cValorZero, cCor, cCorNeg, nNumFmtId)` → id. As cores podem ser Black, Blue, Cyan, Green, Magenta, Red, White, Yellow ou um índice de 1 a 56.
- `AddFmt(cformatCode, nNumFmtId)` → id, para um código de formato livre.
- `Masc2Style(cPicture, oStyle)` cria um estilo a partir de uma máscara Protheus (`"@E 999,999.99"`).
```tlpp
nFmt3Dec := oExcel:AddFmtNum(nDecimal = 3, lMilhar = .T., cNegINI = "(", cNegFim = ")", cCorNeg = "Red")
nFmtMes  := oExcel:AddFmt("[$-pt-BR]mmm-aaa;@")
nFmtHora := oExcel:AddFmt("hh:mm:ss;@")
nFmtDtHr := oExcel:AddFmt("dd/mm/yyyy hh:mm:ss;@")
```

**Ids nativos úteis** (usados direto em `SetnumFmt`):

| Id | Formato |
|---|---|
| 0 | Geral |
| 1 | `0` |
| 2 | `0.00` |
| 3 | `#,##0` |
| 4 | `#,##0.00` |
| 9 | `0%` |
| 10 | `0.00%` |
| 11 | Científico |
| 14 | Data |
| 20 | `h:mm` |
| 21 | `h:mm:ss` |
| 22 | Data e hora |
| **44** | **Moeda R$ (contábil)**, redefinido pela lib |
| 46 | `[h]:mm:ss` |
| 49 | `@` texto |

### Alinhamento: `Alinhamento(cHorizontal, cVertical, lReduzCaber, lQuebraTexto, ntextRotation, nRecuo)` → objeto
- `cHorizontal`: `left`, `center`, `right`, `justify`, `fill`, `general`, `distributed`, `centerContinuous`.
- `cVertical`: `top`, `center`, `bottom`, `justify`, `distributed`.
```tlpp
oCentro := oExcel:Alinhamento(cHorizontal = "center", cVertical = "center")
oQuebra := oExcel:Alinhamento(cHorizontal = "center", cVertical = "center", lQuebraTexto = .T.)
o45     := oExcel:Alinhamento(ntextRotation = 45)
oRecuo  := oExcel:Alinhamento(cHorizontal = "left", nRecuo = 2)
```

### Proteção: `Cellprotection(llocked, lhidden)` → objeto
Só tem efeito junto com `SetsheetProtection` (ver protecao-impressao.md). `llocked = .F.` deixa a célula editável e `lhidden = .T.` oculta a fórmula.

## 3. NewStyle e herança

`NewStyle(oStyleClone)` → um objeto `custom.excel.Style`. Os setters são encadeáveis: `SetnumFmt(id)`, `Setfont(id)`, `Setfill(id)`, `Setborder(id)`, `SetaValores({objAlinhamento, objProtecao})`, `SetAtt(...)`.

Quando você passa um estilo pai, tudo o que não for definido no filho é **herdado** dele, e se o pai mudar depois, a mudança chega ao filho.

```tlpp
oStyBase  := oExcel:NewStyle():Setborder(nBordaAll)
oStyCab   := oExcel:NewStyle(oStyBase):Setfont(nFonteBranca):Setfill(oExcel:CorPreenc("FF4F81BD")):SetaValores({oCentro})
oStyMoeda := oExcel:NewStyle(oStyBase):SetnumFmt(44)
oStyTotal := oExcel:NewStyle(oStyMoeda):Setfont(oExcel:AddFont(lNegrito = .T.))
oStyEdit  := oExcel:NewStyle():SetaValores({oExcel:Cellprotection(llocked = .F.)})

oExcel:Pos(1,1):SetValue("Total"):SetStyle(oStyCab)
oExcel:SetStyle(oStyBase, 2, 1, 10, 5)	//aplica em um intervalo: SetStyle(xStyle, nLinha, nColuna, nLinha2, nColuna2)
```

`SetStyle(xStyle, ...)` aceita um objeto Style, um objeto StyleRules ou um id numérico (v1, criado por `AddStyles`).

## 4. Regras de estilo: `NewStyRules()`

As regras decidem o estilo por célula no momento do `SetStyle`. O bloco recebe `{|nLin, nCol, oExcel| lógico}` e pode ler variáveis externas que estejam no escopo (ex.: um `lSubTotal` controlado pelo laço).

| Método | Efeito |
|---|---|
| `AddStyle(bRule, xStyle)` | Estilo completo |
| `AddnumFmt(bRule, nFmtNum)` | Formato numérico |
| `AddFont(bRule, fontId)` | Fonte |
| `Addfill(bRule, fillId)` | Preenchimento |
| `Addborder(bRule, nborderId)` | Borda |
| `AddValores(bRule, aValores)` | Alinhamento ou proteção |

```tlpp
oZebra := oExcel:NewStyRules()
oZebra:AddStyle(bRule = {|nLin,nCol,oObj| lSubTotal }, xStyle = oStyTotal)
oZebra:AddStyle(bRule = {|nLin,nCol,oObj| nLin % 2 == 1 }, xStyle = oExcel:NewStyle(oStyBase):Setfill(oExcel:CorPreenc("FFB8CCE4")))
oZebra:AddStyle(bRule = {|nLin,nCol,oObj| nLin % 2 == 0 }, xStyle = oExcel:NewStyle(oStyBase):Setfill(oExcel:CorPreenc("FFDCE6F1")))
oZebra:AddnumFmt(bRule = {|nLin,nCol,oObj| nCol == 3 }, nFmtNum = 44)
...
oExcel:Pos(nLin,nCol):SetValue(xValor):SetStyle(oZebra)
```

## 5. Formatação condicional

Os estilos da formatação condicional são **dxf**, que pertencem à pasta de trabalho e podem ser reutilizados em qualquer planilha. Para criá-los, combine:
- `Font(nTamanho, cCorRGB, cNome, cfamily, cScheme, lNegrito, lItalico, lSublinhado, lTachado)`, que tem a mesma ordem do AddFont;
- `Preenc(cBgCor, cFgCor, cType)`: **o fundo vem primeiro**, ao contrário do CorPreenc;
- `ObjBorda(cTipo, cCor, cModelo)`: `cTipo` usa `C` (cima), `B` (baixo), `E` (esquerda), `D` (direita) ou `ALL`.

Esses objetos são passados para `ADDdxf(aFont, aCorPreenc, aBorda)`, que devolve o id do dxf.

```tlpp
nDxfVerm := oExcel:ADDdxf(aFont = oExcel:Font(nTamanho = 11, cCorRGB = "FFFFFF", lNegrito = .T.), aCorPreenc = oExcel:Preenc("FF0000"), aBorda = oExcel:ObjBorda(cTipo = "ALL", cCor = "000000"))
nDxfVerd := oExcel:ADDdxf(aCorPreenc = oExcel:Preenc("00B050"))
nDxfAmar := oExcel:ADDdxf(aCorPreenc = oExcel:Preenc("FFFF00"))
```

| Método | Uso |
|---|---|
| `FormatCond(cRefDe, cRefAte, nEstilo, cOperator, xFormula)` | Operadores: `=`, `!=`, `>`, `>=`, `<`, `<=`, `$` (contém), `!$` (não contém), `between` (`xFormula = {de, ate}`), `FORMULA` |
| `AddFormatCond(cRefDe, cRefAte, nEstilo, cType, xFormula, cOperator, nPrioridade)` | Forma genérica |
| `DataBar(cRefDe, cRefAte, cCor = "FF638EC6", lMostrarValor, cTipoMin, xValMin, cTipoMax, xValMax)` | Barra de dados. Tipos: `min`, `max`, `num`, `percent`, `percentile`, `formula` |
| `IconSet(cRefDe, cRefAte, cIconSet = "3TrafficLights1", lReverso, lMostrarValor)` | Ícones: `3Arrows`, `3TrafficLights1`, `4Rating`, `5Quarters`… |
| `FormatCondDup(cRefDe, cRefAte, nEstilo, lUnico)` | Duplicados (`.F.`, padrão) ou únicos (`.T.`) |
| `FormatCondTempo(cRefDe, cRefAte, nEstilo, cPeriodo)` | `today`, `yesterday`, `tomorrow`, `last7Days`, `thisWeek`, `lastWeek`, `nextWeek`, `thisMonth`, `lastMonth`, `nextMonth`. Um período inválido gera UserException |

```tlpp
oExcel:FormatCond(cRefDe = oExcel:Ref(2,3), cRefAte = oExcel:Ref(nUltLin,3), nEstilo = nDxfVerm, cOperator = "<", xFormula = 0)
oExcel:FormatCond(cRefDe = "C2", cRefAte = "C50", nEstilo = nDxfAmar, cOperator = "between", xFormula = {10,20})
oExcel:DataBar(cRefDe = "D2", cRefAte = "D50")
oExcel:IconSet(cRefDe = "E2", cRefAte = "E50", cIconSet = "3Arrows")
oExcel:FormatCondDup(cRefDe = "A2", cRefAte = "A50", nEstilo = nDxfVerm)
oExcel:FormatCondTempo(cRefDe = "F2", cRefAte = "F50", nEstilo = nDxfVerd, cPeriodo = "thisWeek")
```
