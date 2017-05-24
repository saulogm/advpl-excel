# advpl-excel
Classe em ADVPL para gerar planilhas do excel no formato xlsx com menor consumo de memória e mais otimizado.

# Recursos disponíveis
* Definir células String,Numérica,data,Logica,formula
* Adicionar novas planilhas(Nome,Cor)
* Cor de preenchimento(simples,efeito de preenchimento)
* Alinhamento(Horizontal,Vertical,Reduzir para Caber,Quebra Texto,Angulo de Rotação)
* Formato da célula
* Mesclar células
* Auto Filtro
* Congelar painéis(colunas e linhas)
* Definir tamanho da coluna
* Definir tamanho da linha
* Letra: Fonte,Tamanho,Cor,Negrito,Italico,Sublinhado,Tachado
* Bordas: (Left,Right,Top,Bottom),Cor,Estilo
* Formatação condicional:(operador,formula)(font,fundo,bordas)
* Formatar como tabela(Estilos Predefinidos,Filtros,Totalizadores)
* Cria nome para refencia de célula ou intervalo
* Agrupamento de linha

# Testes
![Exemplo1](/exemplo/excel1.png)

Auto Filtros,Congelar painéis,Agrupar linhas:

![Exemplo2](/exemplo/excel2.png)

Formatar como tabela:

![Exemplo2](/exemplo/excel3.png)

Exemplo de uso:

[tstyexcel.prw](exemplo/tstyexcel.prw)

# Instalação
1. Instalar o 7-Zip (http://www.7-zip.org/)
2. Configurar o appserver.ini com o caminho do 7-Zip.
```
[GENERAL]
LOCAL7ZIP=C:\Program Files\7-Zip\7z.exe
```
3. Aplicar patch [yexcel.ptm](patch/yexcel.ptm)

Pronto para usar a classe YExcel!

### Dúvidas
- Email: saulomax@gmail.com
- GitHub: https://github.com/saulogm
