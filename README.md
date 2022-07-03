# vbf Visual Basic Framework

## Sumário:

1. Apresentação
    1. [Descrição](#descrição)
    1. [Requisitos](#requisitos)
2. Tutorial das Funções
    1. Funções [api_windows (api)](#api_windows)
        * [api_download_web_file](#api_download_web_file)
        * [api_user_windows](#api_user_windows)
    2. Funções [Excel app settings (app)](#Excel_app_settings)
        * [app_set_reference](#app_set_reference)
        * [app_config_off](#app_config_off)
        * [app_config_on](#app_config_on)
    3. Funções [Directory and files settings (def)](#directory_and_files_settings)
        * [def_copy_folder](#def_copy_folder)
        * [def_copy_file](#def_copy_file)
        * [def_find_file](#def_find_file)
        * [def_find_folder](#def_find_folder)
        * [def_folder_exist](#def_folder_exist)
        * [def_file_exist](#def_file_exist)
        * [def_create_folder](#def_create_folder)
        * [def_delete_folder](#def_delete_folder)
        * [def_delete_file](#def_delete_file)
        * [def_open_system_folder](#def_open_system_folder)
        * [def_list_folder_item](#def_list_folder_item)
    4. Funções [Work with datagroup (dtg)](#work_with_datagroup)
        * [dtg_sheet_to_array](#dtg_sheet_to_array)
        * [dtg_array_to_txt](#dtg_array_to_txt)
        * [dtg_list_to_txt](#dtg_list_to_txt)
        * [dtg_read_intire_txt](#dtg_read_intire_txt)
        * [dtg_read_array_txt](#dtg_read_array_txt)
        * [dtg_array_to_sheet](#dtg_array_to_sheet)
        * [dtg_recordset_to_sheet](#dtg_recordset_to_sheet)
        * [dtg_array_transpose](#dtg_array_transpose)
        * [dtg_array_to_html](#dtg_array_to_html)
        * [dtg_recordset_to_array](#dtg_recordset_to_array)
    5. Funções [E-mail settings (eml)](#e-mail_settings)
        * [eml_email_config](#eml_email_config)
    6. Funções [Usinng in sheet (fun)](#usinng_in_sheet)
        * [fun_symbol_off](#fun_symbol_off)
        * [fun_split_off](#fun_split_off)
        * [fun_concat_split_off](#fun_concat_split_off)
    7. Funções [User interface (msg)](#user_interface)
        * [msg_msg_config](#msg_msg_config)
    8. Funções [SQL connection (sql)](#sql_connection)
        * [sql_connection_access](#sql_connection_access)
        * [sql_connection_excel](#sql_connection_excel)
        * [sql_connection_txt](#sql_connection_txt)
        * [sql_connection_sharepoint](#sql_connection_sharepoint)
        * [sql_query](#sql_query)
    9. Funções [Data validate (vld)](#data_validate)
        * [vld_validate_date](#vld_validate_date)
        * [vld_validate_integer](#vld_validate_integer)
        * [vld_validate_double](#vld_validate_double)
        * [vld_validate_string](#vld_validate_string)
        * [vld_validate_not_blanc](#vld_validate_not_blanc)
    10. Funções [Excel file and objects settings (xls)](#excel_file_and_objects_settings)
        * [xls_delete_sheets](#xls_delete_sheets)
        * [xls_refresh_query](#xls_refresh_query)
        * [xls_create_sheet](#xls_create_sheet)
        * [xls_copy_sheet](#xls_copy_sheet)
        * [xls_hide_sheet](#xls_hide_sheet)
        * [xls_veryhide_sheet](#xls_veryhide_sheet)
        * [xls_show_sheet](#xls_show_sheet)
        * [xls_save_as_excel](#xls_save_as_excel)
        * [xls_save_excel](#xls_save_excel)
        * [xls_close_excel](#xls_close_excel)
        * [xls_open_excel](#xls_open_excel)
        * [xls_protect_sheet](#xls_protect_sheet)
        * [xls_unprotect_sheet](#xls_unprotect_sheet)
        * [xls_lock_file](#xls_lock_file)
        * [xls_unlock_file](#xls_unlock_file)


<a id="descrição"></a>

## Descrição

O _Framework_ **VBF** é um grupo de funções criado para auxiliar no desenvolvimento de ferramentas em *VBA*.
Embora a necessidadede contar com tais recursos tenha surgido na época em que atuei no mercado financeiro a ideia de 
organizá-los em funções veio posteriormente e com a quarentena iniciada em 2020 foi possível começar seu desenvolvimento. 
Atuantes na área financeira  possuem o maior potencial de usufruir a totalidade dos benefícios deste recurso, contudo 
qualquer usuário do pacote *MS Office* terá ganho ao usar o __VBF__ em:

* Acessar banco de dados;
* Manipular dados;
* Manipular arquivos e pastas salvos no computador;
* Envio de e-mail;
* Baixar dados disponível na web;
* Elabor relatórios automáticos.

As funções estão separadas em grupos para facilitar o uso. Contam com um prefixo para facilitar a identificação enquanto 
se utiliza. Atualmente conta com 10 grupos de funções e 56 funções.

**prefixo** | **tipo** | **Descrição**
:-----:|:-----:|:-----
api|api_windows|Funções que usam as API's do windows
app|Excel app settings|Funções que manipulam o *programa Excel*.
def|directory and files settings|Funções usadas para manipular arquivos e diretórios do computador
dtg|work with datagroup|Funções que manipulam dados estruturados disponíveis em diversos formatos (txt, xls, recordset, etc)
eml|e-mail settings|Funções para manipular odjetos Outlook
fun|usinng in sheet|Funções que podem ser usadas em planilhas como uma função normal do Excel
msg|user interface|Funções para facilitar a comunicação com o usuário final
sql|sql connection|Funções para usar a linguagem sql em arquivos (csv, xls,etc) e bancos de dados
vld|data validate|Funções para validar os valores das variáveis
xls|excel file and objects settings|Funções para manipular arquivos Excel e seus objetos

O código utiliza do inglês em seu desenvolvimento, no entanto todas as mensagem de texto que são retorno de funções estão por padrão em português (PT-BR) porém
nestas funções existe a possibilidade de configurar o retorno em inglês (EN). Variáveis não possuem tradução, contudo este tutorial será feito inicialmente em 
português (PT-BR) visando ajudar os falantes da língua em se desenvolverem como programador. O tutorial em inglês será disponibilizado/atualizado logo em seguida.

<a id="requisitos"></a>

## Requisitos 

É preciso habilitar três bibliotecas nas referências do *VBA*.



# Tutorial

<a id="api_windows"></a>

# API Windows

Funções que usam as API's do *Windows* para fazer diferentes tarefas.

<a id="api_download_web_file"></a>

### api_download_web_file

### Descrição:

Baixa arquivos disponíveis em sites.

### Sintaxe

~~~vbnet
vbf.api_download_web_file(url_file, file, path)
~~~

### Parâmetros

**Nome** | **Opcional** | **Tipo** | **Descrição**
:-----:|:-----:|:-----:|:-----
url_file|obrigatório|String|A *URL* completo do arquivo na internet
file|obrigatório|String|O nome que se deseja usar para salvar o arquivo no computador
path|opcional|String|O caminho no computador onde deseja salvar o arquivo. Caso não seja informado será salvo na mesma pasta onde o arquivo *Excel* está salvo.

### Retorno

*texto csv*, padão *Variant*

### Exemplo de uso

~~~vbnet
'Declara as variáveis
Dim dwld as Variant
Dim site_ind as String
Dim pt_temp as String
Dim arq as String

'URL completa do arquivo contendo o site e o nome do arquivo
site_ind = "https://www.anbima.com.br/informacoes/indicadores/arqs/indicadores.xls"

'Nome que será usado para salvar o arquivo no computador
arq = "Anbima_indicadores.xls"

'Nome da pasta onde o arquivo será salvo
pt_temp = ThisWorkbook.Path & "\temp\"

'Chama a função usando as variáveis declaradas acima
dwld = vbf.api_download_web_file(site_ind, arq, pt_temp)

~~~

<a id="api_user_windows"></a>

# api_user_windows

### Descrição:

Retorna o usuário atual.

### Sintaxe

~~~vbnet
vbf.api_user_windows()
~~~

### Parâmetros

Não aplicado

### Retorno

*String*, padão *Variant*

### Exemplo de uso

~~~vbnet
'Declara as variáveis
Dim user_id as String

'Chama a função para obter o usuário atual do sistema
user_id = vbf.api_user_windows()

~~~

<a id="Excel_app_settings"></a>

# Excel app settings

Funções que manipulam o *programa Excel*

<a id="app_set_reference"></a>

# app_set_reference

### Descrição:

Inclui as bibliotecas necessárias para o uso de todas as funões na lista de referências.

### Sintaxe

~~~vbnet
vbf.app_set_reference()
~~~

### Parâmetros

Não aplicado

### Retorno

*texto csv*, padão *Variant*

### Exemplo de uso

~~~vbnet
'Declara as variáveis
Sub test_app_set_reference()
Dim set_ref As String

'Neste caso a variável para receber o retorno da função é opcional
set_ref = vbf.app_set_reference

End Sub
~~~

<a id="app_config_off"></a>

# app_config_off

### Descrição:

Desabilita configurações do Excel para deixar as macros de automatização mais rápida e sem pausas. Esta função desabilita 
os seguintes parametros:

* Calculation: Torna o cálculo das céluas manual;
* EnableEvents: Desabilita os eventos;
* DisplayAlerts: Desabilita alertas de atividades coo deletar planilhas ou fechar *Excel* sem gravar;
* ScreenUpdating: Desabilita atualização, não mostrando o que é feito pela macro.

**OBS:** Estas alterações ficam salvas no *programa Excel* afetando outros arquivos *Excel*, mesmo depois de desligar o computador

### Sintaxe

~~~vbnet
vbf.app_config_off()
~~~

### Parâmetros

Não aplicado

### Retorno

*Boolean*

### Exemplo de uso

~~~vbnet
Sub test_app_config_off()
'Declara as variáveis
Dim app_off As Boolean

'Neste caso a variável para receber o retorno da função é opcional
app_off = vbf.app_config_off

End Sub
~~~ 

<a id="app_config_on"></a>

# app_config_on

### Descrição:

Retorna a configuração padrão do Excel podendo deixá-lo lento ao rodar macros de automatizações de tarefas. Faz o inverso 
da função [app_config_off](#app_config_off).

### Sintaxe

~~~vbnet
vbf.app_config_on()
~~~

### Parâmetros

Não aplicado

### Retorno

*Boolean*

### Exemplo de uso

~~~vbnet
Sub test_app_config_on()
'Declara as variáveis
Dim app_on As Boolean

'Neste caso a variável para receber o retorno da função é opcional
app_on = vbf.app_config_on

End Sub
~~~ 

<a id="directory_and_files_settings"></a>

# Directory and Files Settings

Funções usadas para manipular arquivos e diretórios do computador.  Estas funções funcionam para qualquer tipo de arquivo, *.accdb, .xlsx, .txt, .jpeg, .pdf*, etc.

<a id="def_copy_folder"></a>

# def_copy_folder

### Descrição:

Copia uma pasta para outro diretório podendo trocar seu nome. 

### Sintaxe

~~~vbnet
vbf.def_copy_folder( path, new_path, overwrite)
~~~

### Parâmetros

**Nome** | **Opcional** | **Tipo** | **Descrição**
:-----:|:-----:|:-----:|:-----
path|obrigatório|String|Pasta que deseja copiar
new_path|obrigatório|String|Diretório que deseja enviar a pasta
overwrite|opcional|Boolean|Define se deseja sobrescrever caso a pasta já exista. O valor padrão é Falso

### Retorno

*texto csv*, padão *Variant*

### Exemplo de uso

~~~vbnet
Sub test_def_copy_folder()
'Declara as variáveis
Dim cp_dir As Variant

'Copia a pasta informada trocando seu nome
cp_dir = vbf.def_copy_folder(ThisWorkbook.path & "\teste", ThisWorkbook.path & "\Cópia teste", True)

End Sub
~~~

<a id="def_copy_file"></a>

# def_copy_file

### Descrição:

Copia um arquivo para outro diretório podendo trocar seu nome.

### Sintaxe

~~~vbnet
vbf.def_copy_file(path_file, new_path_file, overwrite)
~~~

### Parâmetros

**Nome** | **Opcional** | **Tipo** | **Descrição**
:-----:|:-----:|:-----:|:-----
path_file|obrigatório|String|Diretório e arquivo que deseja copiar
new_path_file|obrigatório|String|Diretório e nome do arquivo de onde será salvo o arquivo
overwrite|opcional|Boolean|Define se deseja sobrescrever caso a pasta já exista. O valor padrão é Falso

### Retorno

*texto csv*, padão *Variant*

### Exemplo de uso

~~~vbnet
Sub test_def_copy_file()
'Declara as variáveis
Dim cp_arq As Variant

'Copia o arquivo 'banco.accdb' para a pasta 'Cópia teste' trocando seu nome para 'novo_db.accdb'
cp_arq = vbf.def_copy_file(ThisWorkbook.path & "\banco.accdb", ThisWorkbook.path & "\Cópia teste\novo_db.accdb", True)

End Sub
~~~

<a id="def_find_file"></a>

# def_find_file

### Descrição:

Exibe a caixa de diálogo padrão de abrir arquivos e retorna o nome do arquivo selecionado sem abri-lo.

### Sintaxe

~~~vbnet
vbf.def_find_file(verbose, EN_lang)
~~~

### Parâmetros

**Nome** | **Opcional** | **Tipo** | **Descrição**
:-----:|:-----:|:-----:|:-----
verbose|opcional|Boolean|Se for verdadeiro, retorna uma menssagem de erro em caso de falha
EN_lang|opcional|Boolean|Se verdadeiro, a menságem de retorno será feita em *inglês - EN*

### Retorno

*texto csv*, padão *Variant*

### Exemplo de uso

~~~vbnet
Sub test_def_find_file()
'Declara as variáveis
Dim file As Variant

'Chama a função optando que retorne mensságem de texto e que seja em inglês'
file = vbf.def_find_file(True, True)

End Sub
~~~

<a id="def_find_folder"></a>

# def_find_folder

### Descrição:

Exibe a caixa de diálogo padrão de selecionar pasta e retorna seu caminho sem abri-la.

### Sintaxe

~~~vbnet
vbf.def_find_folder(verbose, EN_lang)
~~~

### Parâmetros

**Nome** | **Opcional** | **Tipo** | **Descrição**
:-----:|:-----:|:-----:|:-----
verbose|opcional|Boolean|Se for verdadeiro, retorna uma menssagem de erro em caso de falha
EN_lang|opcional|Boolean|Se verdadeiro, a menságem de retorno será feita em *inglês - EN*

### Retorno

*texto csv*, padão *Variant*

### Exemplo de uso

~~~vbnet
Sub test_def_find_folder()
'Declara as variáveis
Dim folder As Variant

'Chama a função optando que retorne mensságem de texto e que seja em inglês'
folder = vbf.def_find_folder(True, True)

End Sub
~~~

<a id="def_folder_exist"></a>

# def_folder_exist

### Descrição:

Verificas se um diretório existe retornando verdadeiro ou falso.

### Sintaxe

~~~vbnet
vbf.def_folder_exist(path)
~~~

### Parâmetros

**Nome** | **Opcional** | **Tipo** | **Descrição**
:-----:|:-----:|:-----:|:-----
path|Obrigatório|String|Diretório que deseja verificar se existe

### Retorno

*Boolean*, padão *Variant*

### Exemplo de uso

~~~vbnet
Sub test_def_folder_exist()
'Declara as variáveis
Dim folder As Boolean

'Verifica se a pasta Windows existe no diretório 'c'
folder = vbf.def_folder_exist("c:\Windows")

End Sub
~~~

<a id="def_file_exist"></a>

# def_file_exist

### Descrição:

Verificas se um arquivo existe retornando verdadeiro ou falso.

### Sintaxe

~~~vbnet
vbf.def_file_exist(path)
~~~

### Parâmetros

**Nome** | **Opcional** | **Tipo** | **Descrição**
:-----:|:-----:|:-----:|:-----
path|Obrigatório|String|Diretório que deseja verificar se existe

### Retorno

*Boolean*, padão *Variant*

### Exemplo de uso

~~~vbnet
Sub test_def_file_exist()
'Declara as variáveis
Dim folder As Boolean

'Verifica se o arquivo novo_db.accdb existe no diretório informado
folder = vbf.def_file_exist(ThisWorkbook.path & "\Cópia teste\novo_db.accdb")

End Sub
~~~

<a id="def_create_folder"></a>

# def_create_folder

### Descrição:

Cria uma pasta no diretório informado.

### Sintaxe

~~~vbnet
vbf.def_create_folder(path_name)
~~~

### Parâmetros

**Nome** | **Opcional** | **Tipo** | **Descrição**
:-----:|:-----:|:-----:|:-----
path_name|Obrigatório|String|Diretório que deseja criar

### Retorno

*variant*, padão *Variant*

### Exemplo de uso

~~~vbnet
Sub test_def_create_folder()
'Declara as variáveis
Dim folder As Variant

'Cria uma pasta de nome 'teste' no mesmo diretório que o arquivo excel atual está salvo
folder = vbf.def_create_folder(ThisWorkbook.path & "\teste")

End Sub
~~~

<a id="def_delete_folder"></a>

# def_delete_folder

### Descrição:

Deleta a pasta informada. 

**OBS:** Desta forma a pasta não vai para a lixeira. Ela é excluida do computador completamente.

### Sintaxe

~~~vbnet
vbf.def_delete_folder(path_name)
~~~

### Parâmetros

**Nome** | **Opcional** | **Tipo** | **Descrição**
:-----:|:-----:|:-----:|:-----
path_name|Obrigatório|String|Diretório que deseja deletar

### Retorno

*variant*, padão *Variant*

### Exemplo de uso

~~~vbnet
Sub test_def_delete_folder()
'Declara as variáveis
Dim folder As Variant

'Deleta a pasta 'teste' localizada no mesmo diretório que o arquivo excel atual está salvo
folder = vbf.def_delete_folder(ThisWorkbook.path & "\teste")

End Sub
~~~

<a id="def_delete_file"></a>

# def_delete_file

### Descrição:

Deleta o arquivo informada. 

**OBS:** Desta forma o arquivo não vai para a lixeira. Ela é excluida do computador completamente.

### Sintaxe

~~~vbnet
vbf.def_delete_folder(path_name)
~~~

### Parâmetros

**Nome** | **Opcional** | **Tipo** | **Descrição**
:-----:|:-----:|:-----:|:-----
path_name|Obrigatório|String|Arquivo que deseja deletar

### Retorno

*variant*, padão *Variant*

### Exemplo de uso

~~~vbnet
Sub test_def_delete_file()
'Declara as variáveis
Dim folder As Variant

'Deleta o arquivo 'novo_db.accdb' localizado no diretório informado
folder = vbf.def_delete_file(ThisWorkbook.path & "\Cópia teste\novo_db.accdb")

End Sub
~~~

<a id="def_open_system_folder"></a>

# def_open_system_folder

### Descrição:

Abre o *Windows Explorer* na pasta informada.

### Sintaxe

~~~vbnet
vbf.def_open_system_folder(path)
~~~

### Parâmetros

**Nome** | **Opcional** | **Tipo** | **Descrição**
:-----:|:-----:|:-----:|:-----
path|Opcional|String|Pasta que deseja abrir. Caso não seja informada uma pasta será aberta a pasta onde o arquivo *Excel* atual está salvo

### Retorno

*variant*, padão *Variant*

### Exemplo de uso

~~~vbnet
Sub test_def_open_system_folder()
'Declara as variáveis
Dim folder As Variant

'Abre o diretório informado
folder = vbf.def_open_system_folder(ThisWorkbook.path & "\Cópia teste")

End Sub
~~~

<a id="def_list_folder_item"></a>

# def_list_folder_item

### Descrição:

Abre o *Windows Explorer* na pasta informada.

### Sintaxe

~~~vbnet
vbf.def_list_folder_item(ByVal path As String, Optional ByVal exclude_folder As Boolean, Optional ByVal exclude_file As Boolean) 
~~~

### Parâmetros

**Nome** | **Opcional** | **Tipo** | **Descrição**
:-----:|:-----:|:-----:|:-----
path|Obrigatório|String|Pasta que deseja listar seus itens
exclude_folder|Opcional|Boolean|Ignora as pastas contidas no diretório informado se verdadeiro, listando apenas os arquivos. Falso é o valor padrão
exclude_file|Opcional|Boolean|Ignora os arquivos contidos no diretório informado se verdadeiro, listando apenas as pastas. Falso é o valor padrão

### Retorno

*Array*, padão *Variant*

Na tabela abaixo estão os nomes das colunas e sua descrições 

Índice da coluna|Nome da coluna|Descrição
:----:|:----:|:----
0|item type|Informo se é uma pasta ou um arquivo
1|path|Diretório e o nome do arquivo ou pasta
2||name|Informa apenas o nome do arquivo ou pasta
3|date_created|Data de criação
4|date_last_accessed|Data do último acesso
5|date_last_modified|Data da última modificação
6|size|Tamanho em bytes
7|type|Tipo do arquivo ou pasta

### Exemplo de uso

~~~vbnet
Sub test_def_list_folder_item()
'Declara as variáveis
Dim folder As Variant

'Lista todos os arquivos e pastas do diretório informado
folder = vbf.def_list_folder_item(ThisWorkbook.path & "\Cópia teste")

End Sub
~~~

<a id="work_with_datagroup"></a>

# Work With Datagroup

Funções que manipulam dados estruturados disponíveis em diversos formatos (txt, xls, recordset, etc)

<a id="dtg_sheet_to_array"></a>

# dtg_sheet_to_array

<a id="dtg_array_to_txt"></a>

# dtg_array_to_txt

<a id="dtg_list_to_txt"></a>

# dtg_list_to_txt

<a id="dtg_read_intire_txt"></a>

# dtg_read_intire_txt

<a id="dtg_read_array_txt"></a>

# dtg_read_array_txt

<a id="dtg_array_to_sheet"></a>

# dtg_array_to_sheet

<a id="dtg_recordset_to_sheet"></a>

# dtg_recordset_to_sheet

<a id="dtg_array_transpose"></a>

# dtg_array_transpose

<a id="dtg_array_to_html"></a>

# dtg_array_to_htm

<a id="dtg_recordset_to_array"></a>

# dtg_recordset_to_array

<a id="e-mail_settings"></a>

# E-mail Settings

Funções para manipular odjetos Outlook

<a id="eml_email_config"></a>

# eml_email_config

<a id="usinng_in_sheet"></a>

# Usinng in Sheet

Funções que podem ser usadas em planilhas como uma função normal do Excel

<a id="fun_symbol_off"></a>

# fun_symbol_off

<a id="fun_split_off"></a>

# fun_split_off

<a id="fun_concat_split_off"></a>

# fun_concat_split_off

<a id="user_interface"></a>

# User Interface 

Funções para facilitar a comunicação com o usuário final

<a id="msg_msg_config"></a>

# msg_msg_config

<a id="sql_connection"></a>

# SQL Connection

Conecte a banco de dados e outros tipos de arquivos como *excel, Access, .txt, .csv* e listas no *Sharepoint* usando o *ODBC*. Através deste recurso será possivel usar 
as principais comondos como *SELECT, UPDATE, INSERT INTO, CREATE, DELETE* entre outros.

<a id="sql_connection_access"></a>

# sql_connection_access

### Descrição:

Cria uma conexão *ODBC* com um arquivo *Access*.

### Sintaxe

~~~vbnet
vbf.sql_connection_access(path_file, verbose, password)

~~~

### Parâmetros

**Nome** | **Opcional** | **Tipo** | **Descrição**
:-----:|:-----:|:-----:|:-----
path_file|obrigatório|String|Caminho e nome do arquivo que deseja fazer a conexão 
verbose|opcional|Boolean|Se for verdadeiro, retorna uma menssagem de erro em caso de falha
password|opcional|String|Usado para informar a senha do arquivo caso haja.

### Retorno

*ODBC.Connection*, padrão *Variant*

### Exemplo de uso

~~~vbnet
Sub test_sql_connection_access()
'Declara as variáveis
Dim cnn As ADODB.connection

'Conecta ao banco access, salvo na mesma pasta do arquivo Excel, de nome "banco.accdb"  
Set cnn = vbf.sql_connection_access(ThisWorkbook.path & "\banco.accdb", True, "1234")

End Sub
~~~

<a id="sql_connection_excel"></a>

# sql_connection_excel

### Descrição:

Cria uma conexão *ODBC* com um arquivo *Excel*.

### Sintaxe

~~~vbnet
vbf.sql_connection_excel(path_file, verbose)

~~~

### Parâmetros

**Nome** | **Opcional** | **Tipo** | **Descrição**
:-----:|:-----:|:-----:|:-----
path_file|obrigatório|String|Caminho e nome do arquivo que deseja fazer a conexão 
verbose|opcional|Boolean|Se for verdadeiro, retorna uma menssagem de erro em caso de falha

### Retorno

*ODBC.Connection*, padrão *Variant*

### Exemplo de uso

~~~vbnet
Sub test_sql_connection_excel()
'Declara as variáveis
Dim cnn1 As ADODB.connection
Dim cnn2 As ADODB.connection

'Conecta ao próprio arquivo excel
Set cnn1 = vbf.sql_connection_excel(ThisWorkbook.FullName, True)

'Conecta a um arquivo excel salvo em outa pasta
Set cnn2 = vbf.sql_connection_excel(ThisWorkbook.path & "\use_example\SQL_Excel.xlsm", True)

End Subb
~~~

<a id="sql_connection_txt"></a>

# sql_connection_txt

### Descrição:

Cria uma conexão *ODBC* com um arquivos de texto (*.csv, .txt, etc*).

**Obs:** Esta função só funciona no *Excel* de *32 Bits*. 

### Sintaxe

~~~vbnet
vbf.sql_connection_txt(path, verbose)

~~~

### Parâmetros

**Nome** | **Opcional** | **Tipo** | **Descrição**
:-----:|:-----:|:-----:|:-----
path|obrigatório|String|Caminho onde os arquivo que deseja conectar estão salvos 
verbose|opcional|Boolean|Se for verdadeiro, retorna uma menssagem de erro em caso de falha

Percebe-se que a conexão é feita a uma pasta e não a um arquivo. Desta forma todos os arquivos em formato de texto contidos na pasta estarão disponíveis para consulta, 
O nome do arquivo de interesse deve ser descriminado na consulta como se fosse o nome de uma tabela em um banco de dados.

Talves o resultado da consulta realizada não seja satisfatório por conta do formato do arquivo texto em questão. Isto ocorre porque a formatação do texto possa estar diferente 
ao esperado pelo *Driver* *ODBC*. Corrija isto criando um arquivo [schema.ini](https://docs.microsoft.com/pt-br/sql/odbc/microsoft/schema-ini-file-text-file-driver?view=sql-server-ver16) e salvando no mesmo local onde o arquivo exto se encontra salvo. Com ele será 
possivel trocar o separador de *vírgula* para *ponto e vírgula* por exemplo.

### Retorno

*ODBC.Connection*, padrão *Variant*

### Exemplo de uso

~~~vbnet
Sub test_sql_connection_txt()
'Declara as variáveis
Dim cnn As ADODB.connection

'Conecta ao próprio arquivo excel
Set cnn = vbf.sql_connection_excel(ThisWorkbook.path, True)

End Sub
~~~

<a id="sql_connection_sharepoint"></a>

# sql_connection_sharepoint

### Descrição:

Cria uma conexão *ODBC* com pastas do *sharepoint*.

### Sintaxe

~~~vbnet
vbf.sql_connection_sharepoint(sp_site, sp_list, verbose, password)

~~~

### Parâmetros

**Nome** | **Opcional** | **Tipo** | **Descrição**
:-----:|:-----:|:-----:|:-----
sp_site|obrigatório|String|Site do *sharepoint* 
sp_list|obrigatório|String|GUID da lista no *sharepoint*
verbose|opcional|Boolean|Se for verdadeiro, retorna uma menssagem de erro em caso de falha
password|opcional|String|Usado para informar a senha do arquivo caso haja.

Como visto acima é preciso de dois parametros incomuns. O site do *sharepoint* está disponível no navegador mas é preciso saber até onde se deve utilizar. 
Como no exemplo abaixo copie o endereço do site até o nome que estiver após "site/". Caso não funcione esperimente até ".com" inclusive.
Para pegar o GUID da lista é recomendado baixar a lista do *sharepoint* no formato *Excel# usando o botão que aparece na parte superior do *sharepoint* quando a lista 
estiver aberta. Após baixar a lista, abra usando o *Excel*, clique em *dados>conexões*. Na janela que abrir, selecione a conexão na lista a esquerda e depois clique em *propriedades*. Na nova janela que abrir clique na aba *destino*. No campo *texto de comando* terá uma instrução em *XML*. Copie a chave que estiver entre as *tags* 
`<LISTNAME>{copie o texto daqui}</LISTNAME>`

### Retorno

*ODBC.Connection*, padrão *Variant*

### Exemplo de uso

~~~vbnet
Sub test_sql_connection_sharepoint()
'Declara as variáveis
Dim cnn As ADODB.connection
Dim sharepoint_site As String
Dim sharepoint_listname As String

'atibuir a uma variável o site do sharepoint até o nome após "site/"
sharepoint_site = "https://suaempresa.sharepoint.com/sites/nome_da_lista"

'atribuir a uma variável o GUID da lista de interesse no sharepoint
sharepoint_listname = "2B057C59-68AA-43ED-97CC-D96852989222"
'
Set cnn = vbf.sql_connection_sharepoint(sharepoint_site, sharepoint_listname, True)

End Sub
~~~

<a id="sql_query"></a>

# sql_query

### Descrição:

Realiza uma consulta a um banco de dados. Com esta função é possível usar instruções como *SELECT, UPDATE, INSERT INTO, CREATE, DELETE*.

### Sintaxe

~~~vbnet
vbf.sql_query(connection, Query, verbose)

~~~

### Parâmetros

**Nome** | **Opcional** | **Tipo** | **Descrição**
:-----:|:-----:|:-----:|:-----
connection|obrigatório|ODBC.Connection|Informar a conexão com a fonte de dados 
Query|obrigatório|String|Qualquer instrução *SQL* que deseja realizar no banco
verbose|opcional|Boolean|Se for verdadeiro, retorna uma menssagem de erro em caso de falha

### Retorno

*ODBC.recordset*, padrão *Variant*

### Exemplo de uso

~~~vbnet
Sub test_sql_connection_query()
'Declara as variáveis
Dim cnn As ADODB.connection
Dim rc As ADODB.recordset

'Conecta ao banco access, salvo na mesma pasta do arquivo Excel, de nome "banco.accdb"
Set cnn = vbf.sql_connection_access(ThisWorkbook.path & "\banco2.accdb", True)

'Realiza uma consulta ao access ("select * from bk1")
Set rc = vbf.sql_query(cnn, "select * from bk1", True)

End Sub
~~~

**Obs:** Caso esta consulta do exemplo de uso estivesse sendo feita a uma tabela no *Excel* de nome *"bk1"*, seria preciso incluir *"$"* ao final do nome e entre colchetes.
conforme exemplificado abaixo.

~~~sql
select * from [bk1$]
~~~

<a id="data_validate"></a>

# Data Validate

Funções para validar os valores das variáveis

<a id="vld_validate_date"></a>

# vld_validate_date

<a id="vld_validate_integer"></a>

# vld_validate_integer

<a id="vld_validate_double"></a>

# vld_validate_double

<a id="vld_validate_string"></a>

# vld_validate_string

<a id="vld_validate_not_blanc"></a>

# vld_validate_not_blanc

<a id="vld_validate_date"></a>

# vld_validate_date

<a id="vld_validate_integer"></a>

# vld_validate_integer

<a id="vld_validate_double"></a>

# vld_validate_double

<a id="vld_validate_string"></a>

# vld_validate_string

<a id="vld_validate_not_blanc"></a>

# vld_validate_not_blanc

<a id="excel_file_and_objects_settings"></a>

# excel file and objects settings

Funções para manipular arquivos Excel e seus objetos

<a id="xls_delete_sheets"></a>

# xls_delete_sheets

<a id="xls_refresh_query"></a>

# xls_refresh_query

<a id="xls_create_sheet"></a>

# xls_create_sheet

<a id="xls_copy_sheet"></a>

# xls_copy_sheet

<a id="xls_hide_sheet"></a>

# xls_hide_sheet

<a id="xls_veryhide_sheet"></a>

# xls_veryhide_sheet

<a id="xls_show_sheet"></a>

# xls_show_sheet

<a id="xls_save_as_excel"></a>

# xls_save_as_excel

<a id="xls_save_excel"></a>

# xls_save_excel

<a id="xls_close_excel"></a>

# xls_close_excel

<a id="xls_open_excel"></a>

# xls_open_excel

<a id="xls_protect_sheet"></a>

# xls_protect_sheet

<a id="xls_unprotect_sheet"></a>

# xls_unprotect_sheet

<a id="xls_lock_file"></a>

# xls_lock_file

<a id="xls_unlock_file"></a>

# xls_unlock_file