# Gravar em json e csv (Save2Json, Save2Csv, NewDados)

Fonte: `src/custom.excel.dados.tlpp` (classe `custom.excel.Dados`). Funciona nos tipos "A", "M" e "B", antes ou depois do `Save`/`Save2Html`, sempre antes do `Destroy()`. Lê as células em streaming pelo mesmo leitor do `Save2Html`; fórmulas sem valor são calculadas antes.

```tlpp
oDados := oExcel:NewDados()                      //Opcional: configuração (não chame FreeObj; o Destroy libera)
oDados:SetLinhaCab(3)                            //Cabeçalho do json "registros" (0 = primeira linha com valor)
cJson := oExcel:Save2Json(GetTempPath(), oDados) //Retorna o caminho do .json
aCsv  := oExcel:Save2Csv(GetTempPath(), oDados)  //Retorna ARRAY com um .csv por planilha
```

## Configuração (`NewDados()`, métodos encadeáveis)

| Método | Padrão | Observação |
|---|---|---|
| `SetPlanilhas(aPlanilhas)` | todas | números ou nomes |
| `SetLinhaCab(nLinha)` | 0 | só json "registros"; linhas anteriores ignoradas |
| `SetModoJson(cModo)` | `"registros"` | ou `"celulas"` |
| `SetFormatado(lFormatado)` | csv `.T.`, json `.F.` | `.T.` = texto como no Excel (`1.234,56`, `27/10/2026`, `VERDADEIRO`) |
| `SetSeparador(cSep)` | `";"` | csv |
| `SetEncoding(cEnc)` | `"utf8bom"` | csv: `utf8bom`, `utf8`, `cp1252`. Json sempre UTF-8 |

## Json

- **registros**: `{"arquivo":"Nome","planilhas":[{"nome":"Plan","indice":1,"linhaCab":1,"colunas":[...],"dados":[{"Produto":"Teclado","Qtd":10,"Data":"2026-10-27","Ativo":true}]}]}`. Todas as chaves em todos os registros (`null` sem valor); linhas sem valor são ignoradas; cabeçalho vazio vira a letra da coluna, repetido recebe `_2`.
- **celulas**: `"celulas":[{"ref":"D2","lin":2,"col":4,"valor":899,"tipo":"N","texto":"899,00","formula":"B2*C2"}]` (`tipo` C, N, D, T, L).
- Valor bruto: data `yyyy-mm-dd`, data e hora `yyyy-mm-ddThh:mm`, hora `hh:mm`.

## Csv

- Um arquivo por planilha: `<nome>.csv` com uma planilha, `<nome>_<planilha>.csv` com várias.
- Da primeira linha com valor até a última, colunas de A até a última usada, linhas vazias intermediárias mantidas; aspas RFC 4180; CRLF.
- Valor bruto (`SetFormatado(.F.)`): número com ponto, data ISO, `TRUE`/`FALSE`.
- O título acima do cabeçalho também vai para o csv (`SetLinhaCab` vale só para o json).
