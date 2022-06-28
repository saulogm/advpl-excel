# advpl-excel
Classe em ADVPL para criar, ler ou editar planilhas do excel no formato xlsx

Exemplo de uso:

[tstyexcel.prw](exemplo/tstyexcel.prw)<br>
[Olá Mundo](https://github.com/saulogm/advpl-excel/wiki/Ol%C3%A1-Mundo)

# Recursos disponíveis
* Definir células String,Numérica,data,DateTime,Logica,formula
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

### Dúvidas
- https://github.com/saulogm/advpl-excel/issues

# Agradecimentos
[Samuel Gomes Martins] (https://github.com/samuelgmartins)
