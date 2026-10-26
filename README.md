# advpl-excel
Classe em ADVPL para criar, ler ou editar planilhas do excel no formato xlsx

Exemplo de uso:

[tstyexcel.prw](exemplo/tstyexcel.prw)<br>
[Olá Mundo](https://github.com/saulogm/advpl-excel/wiki/Ol%C3%A1-Mundo)

# Recursos disponíveis
* Definir células String,Numérica,data,DateTime,Logica,formula
* Cálculo automático de fórmulas: a classe calcula o valor das fórmulas (mais de 500 funções do Excel) ao salvar, sem precisar enviar o valor já calculado ([lista de funções](skills/advpl-excel-xlsx/references/formulas.md))
* Fórmula em português e L1C1 (`oExcel:Formula('SOMA(L2C:L[-1]C)')`): escrita uma vez e convertida para cada célula no `SetValue`
* Adicionar novas planilhas(Nome,Cor)
* Cor de preenchimento(simples,efeito de preenchimento)
* Alinhamento(Horizontal,Vertical,Reduzir para Caber,Quebra Texto,Angulo de Rotação)
* Formato da célula
* Mesclar células
* Auto Filtro
* hyperlink dentro da planilha
* Comentário
* Congelar painéis(colunas e linhas)
* Definir tamanho da linha / largura da coluna
* Formatar numeros(casas decimais)
* Letra: Fonte,Tamanho,Cor,Negrito,Italico,Sublinhado,Tachado
* Bordas: (Left,Right,Top,Bottom),Cor,Estilo
* Formatação condicional:(operador,formula)(font,fundo,bordas)
* Formatar como tabela(Estilos Predefinidos,Filtros,Totalizadores)
* Cria nome para referência de célula ou intervalo
* Agrupamento de linha e colunas
* Imagens
* Exibir/Oculta linhas de Grade
* Definir linha para repetir na impressão
* Definir orientação da pagina na impressão
* Cabeçalho e Ropadé
* Leitura de dados já gravados
* Gravação em html (Save2Html): uma tabela por planilha com estilos, filtros, agrupamento, formatação condicional, imagens, gráficos (Chart.js), colunas como input e blocos de código para alterar o html
* Fórmulas no html: recalculadas no navegador ao alterar inputs; SUBTOTAL acompanha o filtro e os gráficos se redesenham com os novos valores

* Leitura simples dos dados

# Testes
Arquivo excel gerado: [pasta2.xlsx](exemplo/pasta2.xlsx)

![Exemplo1](/exemplo/excel1.png)

Auto Filtros,Congelar painéis,Agrupar linhas e colunas, formatar com regras:

![Exemplo2](/exemplo/excel2.png)

Formatar como tabela:

![Exemplo2](/exemplo/excel3.png)

# Instalação
1. Compilar Fontes da pasta src

Pronto para usar a classe YExcel!

Problema conhecido:
* Para servidor Windows o arquivo gerado pela função FZip não é compatível com LibreOffice. Para contornar, realize a instalação do 7-zip no servidor Windows.

# Skill para Claude Code
A pasta [skills/advpl-excel-xlsx](skills/advpl-excel-xlsx) contém uma *skill* para o [Claude Code](https://claude.com/claude-code), que ensina o assistente de IA a gerar, ler e editar planilhas com a classe `custom.excel.Xlsx` (antiga `YExcel`).

A skill é acionada automaticamente quando você pede algo como "gerar relatório em Excel no Protheus", "exportar query para xlsx", "planilha com gráfico" ou "xlsx com senha", e quando você altera um fonte que usa `YExcel`.

O que a skill cobre:
* Código novo sempre com `custom.excel.Xlsx():New()`, com argumentos nomeados em TLPP e posicionais em ADVPL (.prw)
* Migração de fontes com `YExcel`, `YExcel_Style`, `YExcel_Table` etc. e da API v1 (`Cell`, `AddStyles`, `Gravar`, `OpenRead`) para o padrão atual
* Escolha do tipo de gravação (`"A"` arquivo, `"M"` memória, `"B"` banco/bulk) e a regra de escrita sequencial do tipo `"A"`
* Estilos, formatos numéricos, formatação condicional, tabelas formatadas, `Alias2Tab` com subtotais, bulk, gráficos, imagens, senha/proteção, impressão e leitura de arquivos
* Gravação em html com `Save2Html` (inputs por coluna, blocos de código por planilha/linha/coluna/célula)
* Ciclo de vida (`Save`, `OpenApp`, `Destroy`) e liberação de memória sem duplicidade
* Fontes em CP-1252 sem BOM, CRLF e TAB

Estrutura:
* `SKILL.md`: fluxo, regras e template base
* `references/`: referência da API por assunto (núcleo, estilos, tabelas/bulk/alias, gráficos, proteção/impressão, migração do YExcel e receitas completas)

### Instalação da skill
Copie a pasta `skills/advpl-excel-xlsx` para um dos locais abaixo:
* `~/.claude/skills/advpl-excel-xlsx` (Windows: `%USERPROFILE%\.claude\skills\advpl-excel-xlsx`): disponível em todos os seus projetos
* `<seu-projeto>/.claude/skills/advpl-excel-xlsx`: disponível só naquele projeto (e versionada com ele)

```powershell
Copy-Item -Recurse .\skills\advpl-excel-xlsx "$env:USERPROFILE\.claude\skills\"
```

Depois de copiar, abra uma nova sessão do Claude Code. A skill aparece na lista de skills disponíveis como `advpl-excel-xlsx`.

### Dúvidas
- https://github.com/saulogm/advpl-excel/issues

# Agradecimentos
[Samuel Gomes Martins] (https://github.com/samuelgmartins)
