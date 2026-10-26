# Cálculo de fórmulas (motor AdvPL)

A classe `custom.excel.Xlsx` calcula as fórmulas gravadas sem valor. Basta informar a fórmula; o `<v>` (valor em cache exibido por visualizadores, `GetValue` e `Save2Html`) é preenchido no `Save`/`Save2Html`.

```tlpp
oExcel:Pos(2,1):SetValue(10)
oExcel:Pos(3,1):SetValue(20)
oExcel:Pos(4,1):SetValue(, "SUM(A2:A3)")	//sem valor: o motor calcula 30
oExcel:Pos(5,1):SetValue(, "VLOOKUP(""X"",Plan2!A:B,2,FALSE)")
```

## Como funciona

- Funciona nos três modos de gravação: **A** (arquivo), **M** (memória) e **B** (banco).
- Fórmulas em inglês, `,` como separador e sem `=`, como já era exigido. Prefixos `_xlfn.`/`_xlws.` são aceitos.
- Referências: `A1`, `$A$1`, intervalos, linhas/colunas inteiras (`A:A`, `1:1`), outras planilhas (`'Plan 1'!A1`), nomes definidos (`AddNome`), referências estruturadas de tabela (`Tab[Col]`, `Tab[[#Totals],[Col]]`, `Tab[@Col]`), matrizes constantes (`{1,2;3,4}`).
- Uma fórmula pode depender de outra fórmula: a ordem de cálculo é resolvida pelas dependências. Ciclos não são calculados.
- Quando `xValor` é informado junto com a fórmula, o valor enviado é mantido (nada é recalculado).
- Se alguma fórmula não puder ser calculada (função do ambiente do Excel, função desconhecida, ciclo), a célula fica sem valor e o workbook recebe `<calcPr fullCalcOnLoad="1"/>`: o Excel recalcula tudo ao abrir.
- Fórmulas iguais deslocadas (a mesma fórmula linha a linha) reaproveitam a árvore já analisada, o que mantém o bulk rápido (≈ 40 mil fórmulas em ~24 s no teste de desempenho).

## Métodos

| Método | Uso |
|---|---|
| `SetCalcFormula(lCalc)` | Liga/desliga o cálculo automático (padrão `.T.`). |
| `AddFuncFormula(cNome, bBloco, nMin, nMax, cFlags)` | Registra ou substitui uma função. Bloco `{\|oMotor,aArgs\| xRetorno}`: `aArgs` traz os argumentos avaliados (referências como `custom.excel.calc.Ref`; converta com `oMotor:Num/Txt/Log/Esc/ParaMat`). `cFlags`: `E` expande matrizes elemento a elemento, `V` volátil, `X` não calculável. |
| `SetCalcLocal(cLocal)` | Localidade usada em funções de texto/data (`TEXT`, `DATEVALUE`, `VALUE`...): `"pt-BR"` (padrão) ou `"en-US"`. |
| `Formula(cFormula, lPtBr, lL1C1)` | Prepara uma fórmula em PT-BR e/ou L1C1, convertida para inglês/A1 no `SetValue`/`Cell` (veja abaixo). |

## Fórmula em português e L1C1

`oExcel:Formula(cFormula, lPtBr, lL1C1)` devolve um objeto `custom.excel.Formula` que é aceito no lugar do texto da fórmula em `SetValue(xValor, oFml)` e `Cell(nLin, nCol, xValor, oFml)`. A fórmula é analisada uma única vez; no `SetValue` ela é convertida para inglês/A1 relativa à célula posicionada, então o mesmo objeto pode ser gravado em quantas células forem necessárias.

```tlpp
oFml := oExcel:Formula('SE(LC[-1]>0;LC[-2]/LC[-1];0)')		//PT-BR + L1C1 (padrão)
For nLinha := 2 To 1000
	oExcel:Cell(nLinha,5,,oFml)						//IF(D2>0,C2/D2,0), IF(D3>0,C3/D3,0)...
Next
FreeObj(oFml)
oExcel:Pos(1001,5):SetValue(,oExcel:Formula('SOMA(L2C:L[-1]C)'))	//SUM(E$2:E1000)
oExcel:Pos(1002,5):SetValue(,oExcel:Formula('SUM(R2C:R[-1]C)',.F.))	//só L1C1, nomes em inglês
oExcel:Pos(1003,5):SetValue(,oExcel:Formula('SOMA(E2:E1000)',,.F.))	//só PT-BR, referências A1
```

| Parâmetro | Padrão | `.T.` | `.F.` |
|---|---|---|---|
| `lPtBr` | `.T.` | Fórmula em português: traduz nomes de funções, `VERDADEIRO`/`FALSO`, erros e separadores | Fórmula em inglês (`,` separa argumentos, `.` decimal) |
| `lL1C1` | `.T.` | Referências L1C1 | Referências A1 (sem alteração) |

Com os dois `.F.` a fórmula é gravada como foi escrita.

**PT-BR (`lPtBr`)**

- Nomes das funções do Excel em português para todas as funções do motor (`SOMA`→`SUM`, `SE`→`IF`, `PROCV`→`VLOOKUP`, `CONT.SE`→`COUNTIF`, `ÍNDICE`→`INDEX`, `CORRESP`→`MATCH`, `ARRED`→`ROUND`, `HOJE`→`TODAY`, `EXT.TEXTO`→`MID`...). Acentos são opcionais (`MÉDIA` ou `MEDIA`) e maiúsculas/minúsculas são indiferentes. Nome sem tradução (inglês, função própria do `AddFuncFormula`, nome definido) é mantido.
- `VERDADEIRO`/`FALSO` → `TRUE`/`FALSE`; erros `#N/D`, `#VALOR!`, `#NOME?`, `#NÚM!`, `#NULO!`, `#DESPEJAR!` → `#N/A`, `#VALUE!`, `#NAME?`, `#NUM!`, `#NULL!`, `#SPILL!`.
- Separadores: `;` entre argumentos vira `,`; `,` decimal vira `.`; na matriz constante `\` separa colunas e `;` separa linhas (`{1\2;3\4}` → `{1,2;3,4}`).
- Textos entre aspas, nomes de planilha entre apóstrofos e referências estruturadas (`Tab[Coluna]`) são copiados sem alteração; dentro de colchetes use a sintaxe em inglês (`Tab[[#Totals],[Valor]]`).

**L1C1 (`lL1C1`)**

| L1C1 | Significado | Na célula E5 |
|---|---|---|
| `L2C3` | Linha 2, coluna 3 (absoluta) | `$C$2` |
| `L[-1]C[2]` | 1 linha acima, 2 colunas à direita | `G4` |
| `LC[-1]` | Mesma linha, coluna anterior | `D5` |
| `L2C` | Linha 2 absoluta, mesma coluna | `E$2` |
| `L2:L4` / `L` | Linhas inteiras / linha atual | `$2:$4` / `5:5` |
| `C3` / `C[1]` / `C1:C3` | Colunas inteiras | `$C:$C` / `F:F` / `$A:$C` |

- `R` é aceito no lugar de `L` (`R1C1`). Referências a outra planilha: `Plan2!L1C1`, `'Plan 1'!LC[-1]`.
- Uma referência que resulte fora da planilha (linha ou coluna menor que 1) gera exceção no `SetValue` informando a fórmula e a célula.

## Limitações

- Números em AdvPL têm cerca de 15 dígitos significativos; valores abaixo de ~1E-15 viram 0 nas operações aritméticas.
- Funções de matriz dinâmica (FILTER, SORT, UNIQUE, SEQUENCE...) gravam apenas a célula âncora; o derramamento (spill) é refeito pelo Excel.
- `TEXT` e as conversões de texto seguem a localidade definida em `SetCalcLocal`.
- Referências estruturadas só funcionam para tabelas criadas na mesma execução (não em arquivos abertos com `Open`).
- Funções voláteis (NOW, TODAY, RAND...) são calculadas no servidor e recalculadas pelo Excel ao abrir.
- As funções marcadas abaixo como "não calculada no servidor" dependem do ambiente do Excel (cubos OLAP, RTD, web, Copilot, Python, previsão ETS...).

## Cobertura

Total: 539 funções (504 calculadas, 35 dependentes do ambiente do Excel)

### Banco de dados, web e ambiente do Excel

| Função | Situação |
|---|---|
| ANCHORARRAY | Calculada |
| CALL | Não calculada no servidor (o Excel calcula ao abrir) |
| COPILOT | Não calculada no servidor (o Excel calcula ao abrir) |
| CUBEKPIMEMBER | Não calculada no servidor (o Excel calcula ao abrir) |
| CUBEMEMBER | Não calculada no servidor (o Excel calcula ao abrir) |
| CUBEMEMBERPROPERTY | Não calculada no servidor (o Excel calcula ao abrir) |
| CUBERANKEDMEMBER | Não calculada no servidor (o Excel calcula ao abrir) |
| CUBESET | Não calculada no servidor (o Excel calcula ao abrir) |
| CUBESETCOUNT | Não calculada no servidor (o Excel calcula ao abrir) |
| CUBEVALUE | Não calculada no servidor (o Excel calcula ao abrir) |
| DATESTRING | Não calculada no servidor (o Excel calcula ao abrir) |
| DAVERAGE | Calculada |
| DCOUNT | Calculada |
| DCOUNTA | Calculada |
| DGET | Calculada |
| DMAX | Calculada |
| DMIN | Calculada |
| DPRODUCT | Calculada |
| DSTDEV | Calculada |
| DSTDEVP | Calculada |
| DSUM | Calculada |
| DVAR | Calculada |
| DVARP | Calculada |
| ENCODEURL | Calculada |
| EUROCONVERT | Não calculada no servidor (o Excel calcula ao abrir) |
| FIELDVALUE | Não calculada no servidor (o Excel calcula ao abrir) |
| FILTERXML | Calculada |
| ISTHAIDIGIT | Não calculada no servidor (o Excel calcula ao abrir) |
| NUMBERSTRING | Não calculada no servidor (o Excel calcula ao abrir) |
| PY | Não calculada no servidor (o Excel calcula ao abrir) |
| REGISTER.ID | Não calculada no servidor (o Excel calcula ao abrir) |
| ROUNDBAHTDOWN | Calculada |
| ROUNDBAHTUP | Calculada |
| SINGLE | Calculada |
| STOCKHISTORY | Não calculada no servidor (o Excel calcula ao abrir) |
| THAIDAYOFWEEK | Não calculada no servidor (o Excel calcula ao abrir) |
| THAIDIGIT | Não calculada no servidor (o Excel calcula ao abrir) |
| THAIMONTHOFYEAR | Não calculada no servidor (o Excel calcula ao abrir) |
| THAINUMSOUND | Não calculada no servidor (o Excel calcula ao abrir) |
| THAINUMSTRING | Não calculada no servidor (o Excel calcula ao abrir) |
| THAISTRINGLENGTH | Não calculada no servidor (o Excel calcula ao abrir) |
| THAIYEAR | Não calculada no servidor (o Excel calcula ao abrir) |
| WEBSERVICE | Não calculada no servidor (o Excel calcula ao abrir) |

### Data e hora

| Função | Situação |
|---|---|
| DATE | Calculada |
| DATEDIF | Calculada |
| DATEVALUE | Calculada |
| DAY | Calculada |
| DAYS | Calculada |
| DAYS360 | Calculada |
| EDATE | Calculada |
| EOMONTH | Calculada |
| HOUR | Calculada |
| ISOWEEKNUM | Calculada |
| MINUTE | Calculada |
| MONTH | Calculada |
| NETWORKDAYS | Calculada |
| NETWORKDAYS.INTL | Calculada |
| NOW | Calculada (volátil: o Excel recalcula ao abrir) |
| SECOND | Calculada |
| TIME | Calculada |
| TIMEVALUE | Calculada |
| TODAY | Calculada (volátil: o Excel recalcula ao abrir) |
| WEEKDAY | Calculada |
| WEEKNUM | Calculada |
| WORKDAY | Calculada |
| WORKDAY.INTL | Calculada |
| YEAR | Calculada |
| YEARFRAC | Calculada |

### Engenharia

| Função | Situação |
|---|---|
| BESSELI | Calculada |
| BESSELJ | Calculada |
| BESSELK | Calculada |
| BESSELY | Calculada |
| BIN2DEC | Calculada |
| BIN2HEX | Calculada |
| BIN2OCT | Calculada |
| BITAND | Calculada |
| BITLSHIFT | Calculada |
| BITOR | Calculada |
| BITRSHIFT | Calculada |
| BITXOR | Calculada |
| COMPLEX | Calculada |
| CONVERT | Calculada |
| DEC2BIN | Calculada |
| DEC2HEX | Calculada |
| DEC2OCT | Calculada |
| DELTA | Calculada |
| GESTEP | Calculada |
| HEX2BIN | Calculada |
| HEX2DEC | Calculada |
| HEX2OCT | Calculada |
| IMABS | Calculada |
| IMAGINARY | Calculada |
| IMARGUMENT | Calculada |
| IMCONJUGATE | Calculada |
| IMCOS | Calculada |
| IMCOSH | Calculada |
| IMCOT | Calculada |
| IMCSC | Calculada |
| IMCSCH | Calculada |
| IMDIV | Calculada |
| IMEXP | Calculada |
| IMLN | Calculada |
| IMLOG10 | Calculada |
| IMLOG2 | Calculada |
| IMPOWER | Calculada |
| IMPRODUCT | Calculada |
| IMREAL | Calculada |
| IMSEC | Calculada |
| IMSECH | Calculada |
| IMSIN | Calculada |
| IMSINH | Calculada |
| IMSQRT | Calculada |
| IMSUB | Calculada |
| IMSUM | Calculada |
| IMTAN | Calculada |
| OCT2BIN | Calculada |
| OCT2DEC | Calculada |
| OCT2HEX | Calculada |

### Estatística (distribuições e testes)

| Função | Situação |
|---|---|
| BETA.DIST | Calculada |
| BETA.INV | Calculada |
| BETADIST | Calculada |
| BETAINV | Calculada |
| BINOM.DIST | Calculada |
| BINOM.DIST.RANGE | Calculada |
| BINOM.INV | Calculada |
| BINOMDIST | Calculada |
| CHIDIST | Calculada |
| CHIINV | Calculada |
| CHISQ.DIST | Calculada |
| CHISQ.DIST.RT | Calculada |
| CHISQ.INV | Calculada |
| CHISQ.INV.RT | Calculada |
| CHISQ.TEST | Calculada |
| CHITEST | Calculada |
| CONFIDENCE | Calculada |
| CONFIDENCE.NORM | Calculada |
| CONFIDENCE.T | Calculada |
| CRITBINOM | Calculada |
| ERF | Calculada |
| ERF.PRECISE | Calculada |
| ERFC | Calculada |
| ERFC.PRECISE | Calculada |
| EXPON.DIST | Calculada |
| EXPONDIST | Calculada |
| F.DIST | Calculada |
| F.DIST.RT | Calculada |
| F.INV | Calculada |
| F.INV.RT | Calculada |
| F.TEST | Calculada |
| FDIST | Calculada |
| FINV | Calculada |
| FISHER | Calculada |
| FISHERINV | Calculada |
| FTEST | Calculada |
| GAMMA | Calculada |
| GAMMA.DIST | Calculada |
| GAMMA.INV | Calculada |
| GAMMADIST | Calculada |
| GAMMAINV | Calculada |
| GAMMALN | Calculada |
| GAMMALN.PRECISE | Calculada |
| GAUSS | Calculada |
| HYPGEOM.DIST | Calculada |
| HYPGEOMDIST | Calculada |
| LOGINV | Calculada |
| LOGNORM.DIST | Calculada |
| LOGNORM.INV | Calculada |
| LOGNORMDIST | Calculada |
| NEGBINOM.DIST | Calculada |
| NEGBINOMDIST | Calculada |
| NORM.DIST | Calculada |
| NORM.INV | Calculada |
| NORM.S.DIST | Calculada |
| NORM.S.INV | Calculada |
| NORMDIST | Calculada |
| NORMINV | Calculada |
| NORMSDIST | Calculada |
| NORMSINV | Calculada |
| PHI | Calculada |
| POISSON | Calculada |
| POISSON.DIST | Calculada |
| PROB | Calculada |
| STANDARDIZE | Calculada |
| T.DIST | Calculada |
| T.DIST.2T | Calculada |
| T.DIST.RT | Calculada |
| T.INV | Calculada |
| T.INV.2T | Calculada |
| T.TEST | Calculada |
| TDIST | Calculada |
| TINV | Calculada |
| TTEST | Calculada |
| WEIBULL | Calculada |
| WEIBULL.DIST | Calculada |
| Z.TEST | Calculada |
| ZTEST | Calculada |

### Estatística

| Função | Situação |
|---|---|
| AGGREGATE | Calculada |
| AVEDEV | Calculada |
| AVERAGE | Calculada |
| AVERAGEA | Calculada |
| AVERAGEIF | Calculada |
| AVERAGEIFS | Calculada |
| CORREL | Calculada |
| COUNT | Calculada |
| COUNTA | Calculada |
| COUNTBLANK | Calculada |
| COUNTIF | Calculada |
| COUNTIFS | Calculada |
| COVAR | Calculada |
| COVARIANCE.P | Calculada |
| COVARIANCE.S | Calculada |
| DEVSQ | Calculada |
| FORECAST | Calculada |
| FORECAST.ETS | Não calculada no servidor (o Excel calcula ao abrir) |
| FORECAST.ETS.CONFINT | Não calculada no servidor (o Excel calcula ao abrir) |
| FORECAST.ETS.SEASONALITY | Não calculada no servidor (o Excel calcula ao abrir) |
| FORECAST.ETS.STAT | Não calculada no servidor (o Excel calcula ao abrir) |
| FORECAST.LINEAR | Calculada |
| FREQUENCY | Calculada |
| GEOMEAN | Calculada |
| GROWTH | Calculada |
| HARMEAN | Calculada |
| INTERCEPT | Calculada |
| KURT | Calculada |
| LARGE | Calculada |
| LINEST | Calculada |
| LOGEST | Calculada |
| MAX | Calculada |
| MAXA | Calculada |
| MAXIFS | Calculada |
| MEDIAN | Calculada |
| MIN | Calculada |
| MINA | Calculada |
| MINIFS | Calculada |
| MODE | Calculada |
| MODE.MULT | Calculada |
| MODE.SNGL | Calculada |
| PEARSON | Calculada |
| PERCENTILE | Calculada |
| PERCENTILE.EXC | Calculada |
| PERCENTILE.INC | Calculada |
| PERCENTRANK | Calculada |
| PERCENTRANK.EXC | Calculada |
| PERCENTRANK.INC | Calculada |
| PERMUT | Calculada |
| PERMUTATIONA | Calculada |
| QUARTILE | Calculada |
| QUARTILE.EXC | Calculada |
| QUARTILE.INC | Calculada |
| RANK | Calculada |
| RANK.AVG | Calculada |
| RANK.EQ | Calculada |
| RSQ | Calculada |
| SKEW | Calculada |
| SKEW.P | Calculada |
| SLOPE | Calculada |
| SMALL | Calculada |
| STDEV | Calculada |
| STDEV.P | Calculada |
| STDEV.S | Calculada |
| STDEVA | Calculada |
| STDEVP | Calculada |
| STDEVPA | Calculada |
| STEYX | Calculada |
| SUBTOTAL | Calculada |
| TREND | Calculada |
| TRIMMEAN | Calculada |
| VAR | Calculada |
| VAR.P | Calculada |
| VAR.S | Calculada |
| VARA | Calculada |
| VARP | Calculada |
| VARPA | Calculada |

### Financeira

| Função | Situação |
|---|---|
| ACCRINT | Calculada |
| ACCRINTM | Calculada |
| AMORDEGRC | Calculada |
| AMORLINC | Calculada |
| COUPDAYBS | Calculada |
| COUPDAYS | Calculada |
| COUPDAYSNC | Calculada |
| COUPNCD | Calculada |
| COUPNUM | Calculada |
| COUPPCD | Calculada |
| CUMIPMT | Calculada |
| CUMPRINC | Calculada |
| DB | Calculada |
| DDB | Calculada |
| DISC | Calculada |
| DOLLARDE | Calculada |
| DOLLARFR | Calculada |
| DURATION | Calculada |
| EFFECT | Calculada |
| FV | Calculada |
| FVSCHEDULE | Calculada |
| INTRATE | Calculada |
| IPMT | Calculada |
| IRR | Calculada |
| ISPMT | Calculada |
| MDURATION | Calculada |
| MIRR | Calculada |
| NOMINAL | Calculada |
| NPER | Calculada |
| NPV | Calculada |
| ODDFPRICE | Calculada |
| ODDFYIELD | Calculada |
| ODDLPRICE | Calculada |
| ODDLYIELD | Calculada |
| PDURATION | Calculada |
| PMT | Calculada |
| PPMT | Calculada |
| PRICE | Calculada |
| PRICEDISC | Calculada |
| PRICEMAT | Calculada |
| PV | Calculada |
| RATE | Calculada |
| RECEIVED | Calculada |
| RRI | Calculada |
| SLN | Calculada |
| SYD | Calculada |
| TBILLEQ | Calculada |
| TBILLPRICE | Calculada |
| TBILLYIELD | Calculada |
| VDB | Calculada |
| XIRR | Calculada |
| XNPV | Calculada |
| YIELD | Calculada |
| YIELDDISC | Calculada |
| YIELDMAT | Calculada |

### Lógica e informação

| Função | Situação |
|---|---|
| AND | Calculada |
| CELL | Calculada |
| ERROR.TYPE | Calculada |
| FALSE | Calculada |
| IF | Calculada |
| IFERROR | Calculada |
| IFNA | Calculada |
| IFS | Calculada |
| INFO | Calculada |
| ISBLANK | Calculada |
| ISERR | Calculada |
| ISERROR | Calculada |
| ISEVEN | Calculada |
| ISFORMULA | Calculada |
| ISLOGICAL | Calculada |
| ISNA | Calculada |
| ISNONTEXT | Calculada |
| ISNUMBER | Calculada |
| ISODD | Calculada |
| ISOMITTED | Calculada |
| ISREF | Calculada |
| ISTEXT | Calculada |
| LAMBDA | Calculada |
| LET | Calculada |
| N | Calculada |
| NA | Calculada |
| NOT | Calculada |
| OR | Calculada |
| SHEET | Calculada |
| SHEETS | Calculada |
| SWITCH | Calculada |
| T | Calculada |
| TRUE | Calculada |
| TYPE | Calculada |
| XOR | Calculada |

### Matemática e trigonometria

| Função | Situação |
|---|---|
| ABS | Calculada |
| ACOS | Calculada |
| ACOSH | Calculada |
| ACOT | Calculada |
| ACOTH | Calculada |
| ARABIC | Calculada |
| ASIN | Calculada |
| ASINH | Calculada |
| ATAN | Calculada |
| ATAN2 | Calculada |
| ATANH | Calculada |
| BASE | Calculada |
| CEILING | Calculada |
| CEILING.MATH | Calculada |
| CEILING.PRECISE | Calculada |
| COMBIN | Calculada |
| COMBINA | Calculada |
| COS | Calculada |
| COSH | Calculada |
| COT | Calculada |
| COTH | Calculada |
| CSC | Calculada |
| CSCH | Calculada |
| DECIMAL | Calculada |
| DEGREES | Calculada |
| ECMA.CEILING | Calculada |
| EVEN | Calculada |
| EXP | Calculada |
| FACT | Calculada |
| FACTDOUBLE | Calculada |
| FLOOR | Calculada |
| FLOOR.MATH | Calculada |
| FLOOR.PRECISE | Calculada |
| GCD | Calculada |
| INT | Calculada |
| ISO.CEILING | Calculada |
| LCM | Calculada |
| LN | Calculada |
| LOG | Calculada |
| LOG10 | Calculada |
| MDETERM | Calculada |
| MINVERSE | Calculada |
| MMULT | Calculada |
| MOD | Calculada |
| MROUND | Calculada |
| MULTINOMIAL | Calculada |
| MUNIT | Calculada |
| ODD | Calculada |
| PI | Calculada |
| POWER | Calculada |
| PRODUCT | Calculada |
| QUOTIENT | Calculada |
| RADIANS | Calculada |
| RAND | Calculada (volátil: o Excel recalcula ao abrir) |
| RANDARRAY | Calculada (volátil: o Excel recalcula ao abrir) |
| RANDBETWEEN | Calculada (volátil: o Excel recalcula ao abrir) |
| ROMAN | Calculada |
| ROUND | Calculada |
| ROUNDDOWN | Calculada |
| ROUNDUP | Calculada |
| SEC | Calculada |
| SECH | Calculada |
| SEQUENCE | Calculada |
| SERIESSUM | Calculada |
| SIGN | Calculada |
| SIN | Calculada |
| SINH | Calculada |
| SQRT | Calculada |
| SQRTPI | Calculada |
| SUM | Calculada |
| SUMIF | Calculada |
| SUMIFS | Calculada |
| SUMPRODUCT | Calculada |
| SUMSQ | Calculada |
| SUMX2MY2 | Calculada |
| SUMX2PY2 | Calculada |
| SUMXMY2 | Calculada |
| TAN | Calculada |
| TANH | Calculada |
| TRUNC | Calculada |

### Pesquisa, referência e matriz dinâmica

| Função | Situação |
|---|---|
| ADDRESS | Calculada |
| AREAS | Calculada |
| BYCOL | Calculada |
| BYROW | Calculada |
| CHOOSE | Calculada |
| CHOOSECOLS | Calculada |
| CHOOSEROWS | Calculada |
| COLUMN | Calculada |
| COLUMNS | Calculada |
| DROP | Calculada |
| EXPAND | Calculada |
| FILTER | Calculada |
| FORMULATEXT | Calculada |
| GETPIVOTDATA | Não calculada no servidor (o Excel calcula ao abrir) |
| GROUPBY | Calculada |
| HLOOKUP | Calculada |
| HSTACK | Calculada |
| HYPERLINK | Calculada |
| IMAGE | Não calculada no servidor (o Excel calcula ao abrir) |
| INDEX | Calculada |
| INDIRECT | Calculada |
| LOOKUP | Calculada |
| MAKEARRAY | Calculada |
| MAP | Calculada |
| MATCH | Calculada |
| OFFSET | Calculada |
| PERCENTOF | Calculada |
| PIVOTBY | Calculada |
| REDUCE | Calculada |
| ROW | Calculada |
| ROWS | Calculada |
| RTD | Não calculada no servidor (o Excel calcula ao abrir) |
| SCAN | Calculada |
| SORT | Calculada |
| SORTBY | Calculada |
| TAKE | Calculada |
| TOCOL | Calculada |
| TOROW | Calculada |
| TRANSPOSE | Calculada |
| TRIMRANGE | Calculada |
| UNIQUE | Calculada |
| VLOOKUP | Calculada |
| VSTACK | Calculada |
| WRAPCOLS | Calculada |
| WRAPROWS | Calculada |
| XLOOKUP | Calculada |
| XMATCH | Calculada |

### Texto

| Função | Situação |
|---|---|
| ARRAYTOTEXT | Calculada |
| ASC | Calculada |
| BAHTTEXT | Não calculada no servidor (o Excel calcula ao abrir) |
| CHAR | Calculada |
| CLEAN | Calculada |
| CODE | Calculada |
| CONCAT | Calculada |
| CONCATENATE | Calculada |
| DBCS | Calculada |
| DETECTLANGUAGE | Não calculada no servidor (o Excel calcula ao abrir) |
| DOLLAR | Calculada |
| EXACT | Calculada |
| FIND | Calculada |
| FINDB | Calculada |
| FIXED | Calculada |
| JIS | Calculada |
| LEFT | Calculada |
| LEFTB | Calculada |
| LEN | Calculada |
| LENB | Calculada |
| LOWER | Calculada |
| MID | Calculada |
| MIDB | Calculada |
| NUMBERVALUE | Calculada |
| PHONETIC | Calculada |
| PROPER | Calculada |
| REGEXEXTRACT | Calculada |
| REGEXREPLACE | Calculada |
| REGEXTEST | Calculada |
| REPLACE | Calculada |
| REPLACEB | Calculada |
| REPT | Calculada |
| RIGHT | Calculada |
| RIGHTB | Calculada |
| SEARCH | Calculada |
| SEARCHB | Calculada |
| SUBSTITUTE | Calculada |
| TEXT | Calculada |
| TEXTAFTER | Calculada |
| TEXTBEFORE | Calculada |
| TEXTJOIN | Calculada |
| TEXTSPLIT | Calculada |
| TRANSLATE | Não calculada no servidor (o Excel calcula ao abrir) |
| TRIM | Calculada |
| UNICHAR | Calculada |
| UNICODE | Calculada |
| UPPER | Calculada |
| VALUE | Calculada |
| VALUETOTEXT | Calculada |

## Exemplo de função própria

```tlpp
//DOBRO(x) = x*2, aceitando intervalo (expande elemento a elemento)
oExcel:AddFuncFormula("DOBRO", {|oM,aArgs| oM:Num(aArgs[1])*2 }, 1, 1, "E")
oExcel:Pos(2,2):SetValue(, "DOBRO(A2)")
```
