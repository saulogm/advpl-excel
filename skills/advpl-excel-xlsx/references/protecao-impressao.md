# Proteção, senha e configuração de impressão

## Senha de abertura (criptografia MS-OFFCRYPTO Agile, AES-256)

- `SetPassword(cSenha)` → self: o arquivo é criptografado no `Save`. `""` remove a senha.
- `Open(cArq, cSenha)` → lógico. Em caso de falha, `GetError()` traz o motivo (inclusive senha incorreta).
- **Não** use `New(cFileOpen = arquivoComSenha)`: gera exceção. Use `New()` e depois `Open(cArq, cSenha)`.

```tlpp
//Gerar protegido
oExcel := custom.excel.Xlsx():New(cNomeFile = "Confidencial", cTipo = "A")
oExcel:ADDPlan(cNome = "Dados")
oExcel:Pos(1,1):SetValue("Conteúdo protegido")
oExcel:SetPassword(cSenha)
cArq := oExcel:Save("\temp")	//no servidor; use GetTempPath() para o client
oExcel:Destroy()
FreeObj(oExcel)

//Reabrir
oExcel := custom.excel.Xlsx():New()
If oExcel:Open(cArq, cSenha)
	xValor := oExcel:Pos(1,1):GetValue()
Else
	ConOut("Falha: " + oExcel:GetError())
EndIf
oExcel:Destroy()
FreeObj(oExcel)
```

Nunca deixe a senha fixa no fonte em código de produção. Ela deve vir de parâmetro (SX6/`SuperGetMV`), de uma pergunta ao usuário ou de um cofre.

## Proteção de planilha (edição)

`SetsheetProtection(cSenha, lsheet, lselectLockedCells, lselectUnlockedCells, lformatCells, lformatColumns, lformatRows, linsertColumns, linsertRows, linsertHyperlinks, ldeleteColumns, ldeleteRows, lsort, lautoFilter, lpivotTables, lobjects, lscenarios)`

- Protege a planilha **atual** contra edição. A senha pode ter até 255 caracteres. `lsheet` e as duas flags de seleção são `.T.` por padrão.
- Para liberar células, crie um estilo com `Cellprotection(llocked = .F.)`. Para esconder a fórmula, `Cellprotection(lhidden = .T.)`.
- Isso é diferente de `SetPassword`: a proteção de planilha não criptografa o arquivo.

```tlpp
oStyEditavel := oExcel:NewStyle():SetaValores({oExcel:Cellprotection(llocked = .F.)})
oStyFormOcul := oExcel:NewStyle():SetaValores({oExcel:Cellprotection(lhidden = .T.)})
oExcel:Pos(2,2):SetValue(0):SetStyle(oStyEditavel)
oExcel:Pos(3,2):SetValue(xValor = 0, cFormula = "B2*2"):SetStyle(oStyFormOcul)
oExcel:SetsheetProtection("123")
```

## Impressão e página

| Método | Descrição |
|---|---|
| `SetPrintArea(nLinha, nColuna, nLinha2, nColuna2, cRefPar, cPlanilha)` | Área de impressão |
| `SetPrintTitles(nLinha, nLinha2, cRefPar, cPlanilha)` | Linhas repetidas em cada página (ex.: cabeçalho, linha 1) |
| `SetPagOrientation(cOrientation)` | `"default"`, `"landscape"` (padrão da lib) ou `"portrait"`. As planilhas criadas depois herdam a orientação |
| `SetHeader(cLeft, cCenter, cRight)` / `SetFooter(cLeft, cCenter, cRight)` | Cabeçalho e rodapé |

Códigos de cabeçalho e rodapé:
- `&A` nome da planilha;
- `&D` data;
- `&T` hora;
- `&P` página;
- `&N` total de páginas;
- `&B` negrito;
- `&nn` tamanho da fonte;
- `&Krrggbb` cor;
- `CHR(10)` quebra de linha.

```tlpp
oExcel:SetPagOrientation("portrait")
oExcel:SetPrintTitles(nLinha = 1, nLinha2 = 1)
oExcel:SetHeader(cLeft = "&A", cCenter = "&BRelatório de Vendas", cRight = "&D" + CHR(10) + "&T")
oExcel:SetFooter(cRight = "Pág &P/&N")
```
