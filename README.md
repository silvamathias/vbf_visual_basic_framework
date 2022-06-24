# vbf Visual Basic Framework

## Descrição

  __VBF__ é um grupo de funções criado para auxiliar no desenvolvimento de ferramentas em VBA. Embora a necessidade de contar com tais recursos tenha surgido na época em que atuei
no mercado financeiro a ideia de organizá-los em funções veio posteriormente e com a quarentena iniciada em 2020 foi possível começar seu desenvolvimento. Atuantes na área financeira  possuem o maior potencial de usufruir a totalidade dos benefícios deste recurso, contudo qualquer usuário do pacote __MS Office__ terá ganho ao usar o __VBF__ em:

* Acessar banco de dados;
* Manipular dados;
* Manipulação de arquivos e pastas salvos no computador;
* Envio de e-mail;
* Baixar dados disponível na web;
* Elaboração de relatórios.

  As funções estão separadas em grupos para facilitar o uso. Contam com um prefixo para facilitar a identificação enquanto se utiliza. Atualmente conta com 10 grupos de funções e 56 funções.

**prefixo** | **tipo** | **Descrição**
:-----:|:-----:|:-----
api|api_windows|Funções que usam as API's do windows
app|Excel app settings|Funções que manipulam o programa Excel.
def|directory and files settings|Funções usadas para manipular arquivos e diretórios do computador
dtg|work with datagroup|Funções que manipulam dados estruturados disponíveis em diversos formatos (txt, xls, recordset, etc)
eml|e-mail settings|Funções para manipular odjetos Outlook
fun|usinng in sheet|Funções que podem ser usadas em planilhas como uma função normal do Excel
msg|user interface|Funções para facilitar a comunicação com o usuário final
sql|sql connection|Funções para usar a linguagem sql em arquivos (csv, xls,etc) e bancos de dados
vld|data validate|Funções para validar os valores das variáveis
xls|Excel file and objects settings|Funções para manipular arquivos Excel e seus objetos

  O código utiliza do inglês em seu desenvolvimento, no entanto todas as funções que possuem uma mensagem de texto como retorno, esta é feita por padrão em português (PT-BR) porém
nestas existe a possibilidade de configurar o retorno em inglês (EN). Variáveis não possuem tradução, contudo este tutorial será feito inicialmente em português (PT-BR) visando
ajudar os falantes da língua em se desenvolverem como programador. O tutorial em inglês será disponibilizado/atualizado logo em seguida

## Tutorial

~~~vbnet
Dim num_col As Double
Dim j As Variant
Dim i As Variant

On Error GoTo error
num_row = UBound(array_ob)
num_col = UBound(array_ob(0))

ReDim transpor(0 To num_col)

For j = 0 To num_col
    ReDim row(0 To num_row)
    For i = 0 To num_row
        row(i) = array_ob(i)(j)
    Next i
    transpor(j) = row
    Erase row
Next j

~~~

***vbnet
Dim num_col As Double
Dim j As Variant
Dim i As Variant

On Error GoTo error
num_row = UBound(array_ob)
num_col = UBound(array_ob(0))

ReDim transpor(0 To num_col)

For j = 0 To num_col
    ReDim row(0 To num_row)
    For i = 0 To num_row
        row(i) = array_ob(i)(j)
    Next i
    transpor(j) = row
    Erase row
Next j
***


Em desenvolvimento...
