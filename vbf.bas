'Copyright (c) 2022 silvamathias
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.
'---------------------------------------------------------------------------------------------------
'Project repository >> https://github.com/silvamathias/vbf_visual_basic_framework#api_windows
'---------------------------------------------------------------------------------------------------
Attribute VB_Name = "vbf"
Option Explicit
'----32bits download web API------------------------------------------------------------------------
'Private Declare Function api_download_from_web Lib "urlmon" Alias "URLDownloadToFileA" _
    (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, _
    ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
'----64bits download web API------------------------------------------------------------------------
Private Declare PtrSafe Function api_download_from_web Lib "urlmon" _
  Alias "URLDownloadToFileA" ( _
    ByVal pCaller As LongPtr, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As LongPtr _
  ) As Long

'---------------------------------------------------------------------------------------------------
'-------functions: api_windows----------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------
Function api_download_web_file(ByVal url_file As String, ByVal file As String, Optional ByVal path As String) As Variant
api_download_web_file = False
Dim download As Long

On Error GoTo error

If path = "" Then path = ThisWorkbook.path & "\"

download = api_download_from_web(0, url_file, path & file, 0, 0)

If download <> 0 Then GoTo error

api_download_web_file = True & ";api_download_web_file"
Exit Function
error:
api_download_web_file = "ERRO;" & "api_download_web_file;" & Err.Number & ";" & Err.Description

End Function


Function api_user_windows() As String
api_user_windows = False
Dim user As Object
On Error GoTo error
Set user = CreateObject("WScript.Network")


api_user_windows = user.UserName
Exit Function
error:
api_user_windows = "ERRO;" & "api_user_windows;" & Err.Number & ";" & Err.Description
End Function


'---------------------------------------------------------------------------------------------------
'-------functions: Excel app settings---------------------------------------------------------------
'---------------------------------------------------------------------------------------------------

Function app_set_reference() As Variant
Dim wb As Workbook
Dim ref(1 To 3) As String
Dim item As Variant
app_set_reference = False

On Error GoTo error

Set wb = ThisWorkbook

ref(1) = "{420B2830-E718-11CF-893D-00A0C9054228}"
ref(2) = "{00062FFF-0000-0000-C000-000000000046}"
ref(3) = "{B691E011-1797-432E-907A-4D8C69339129}"


For Each item In ref
    wb.VBProject.References.AddFromGuid GUID:=item, Major:=1, Minor:=0
Next item
app_set_reference = True & ";app_set_reference"
Exit Function
error:
app_set_reference = "ERRO;" & "app_set_reference;" & Err.Number & ";" & Err.Description
End Function


Function app_app_config_on()
Dim app As Application
Set app = Application
With app
    .Calculation = xlCalculationAutomatic
    .EnableEvents = True
    .DisplayAlerts = True
    .ScreenUpdating = True
End With
End Function


Function app_app_config_off()
Dim app As Application
Set app = Application
With app
    .Calculation = xlCalculationManual
    .EnableEvents = False
    .DisplayAlerts = False
    .ScreenUpdating = False
End With
End Function


'---------------------------------------------------------------------------------------------------
'-------functions: directory and files settings-----------------------------------------------------
'---------------------------------------------------------------------------------------------------
Function def_copy_folder(ByVal path As String, ByVal new_path As String, Optional ByVal overwrite As Boolean) As Variant
def_copy_folder = False

Dim srt As FileSystemObject
Dim var_type As Long
On Error GoTo error

var_type = VarType(overwrite)
If var_type = 10 Then overwrite = True

Set srt = New FileSystemObject
srt.CopyFolder path, new_path, overwrite
def_copy_folder = True
Exit Function
error:
def_copy_folder = "ERRO;" & "def_copy_folder;" & Err.Number & ";" & Err.Description
End Function


Function def_copy_file(ByVal path_file As String, ByVal new_path_file As String, Optional ByVal overwrite As Variant) As Variant
def_copy_file = False

Dim srt As FileSystemObject
Dim var_type As Long
On Error GoTo error

var_type = VarType(overwrite)
If var_type = 10 Then overwrite = True
Set srt = New FileSystemObject
srt.CopyFile path_file, new_path_file, overwrite
def_copy_file = True
Exit Function
error:
def_copy_file = "ERRO;" & "def_copy_file;" & Err.Number & ";" & Err.Description
End Function


Function def_find_file(Optional ByVal verbose As Boolean, Optional ByVal EN_lang As Boolean) As Variant

def_find_file = False
Dim Filter As String
Dim FilterIndex As Integer
Dim Title As String
Dim path As Variant
On Error GoTo error

Title = "Selecionar file"
ChDrive ("C")
ChDir (ThisWorkbook.path)

With Application
    path = .GetOpenFilename(Filter, FilterIndex, Title)
    ChDrive (Left(.DefaultFilePath, 1))
End With

If path = False Then
    If verbose = True Then
        If EN_lang = True Then
            MsgBox "No files selected.", vbExclamation + vbOKOnly, "Error selecting file"
        Else
            MsgBox "Nenhum arquivo selecionado.", vbExclamation + vbOKOnly, "Erro ao selecionar arquivo"
        End If
    End If
    Exit Function
End If
def_find_file = path
Exit Function
error:
def_find_file = "ERRO;" & "def_find_file;" & Err.Number & ";" & Err.Description
End Function


Function def_find_folder(Optional ByVal verbose As Boolean, Optional ByVal EN_lang As Boolean) As Variant
def_find_folder = False
On Error Resume Next
Dim fd As FileDialog
Dim path As Variant

Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    With fd
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        .Show
        path = .SelectedItems(1)
    End With

Set fd = Nothing

If path = False Or path = "" Then
    If verbose = True Then
        If EN_lang = True Then
            MsgBox "Nenhuma pasta selecionada.", vbExclamation + vbOKOnly, "Error selecting folder"
        Else
            MsgBox "Nenhuma pasta selecionada.", vbExclamation + vbOKOnly, "Erro ao selecionar pasta"
        End If
    End If
    Exit Function
End If
def_find_folder = path
Exit Function
error:
def_find_folder = "ERRO;" & "def_find_folder;" & Err.Number & ";" & Err.Description
End Function


Function def_folder_exist(ByVal path As String) As Boolean

def_folder_exist = False
Dim srt As FileSystemObject
Set srt = New FileSystemObject
On Error Resume Next
def_folder_exist = srt.FolderExists(path)
End Function


Function def_file_exist(path_file) As Boolean

def_file_exist = False
Dim srt As FileSystemObject
Set srt = New FileSystemObject
On Error Resume Next
def_file_exist = srt.FileExists(path_file)
End Function


Function def_create_folder(ByVal path_name As String) As Variant

def_create_folder = False
Dim srt As FileSystemObject
On Error GoTo error
Set srt = New FileSystemObject
srt.CreateFolder (path_name)
def_create_folder = True
Exit Function
error:
def_create_folder = "ERRO;" & "def_create_folder;" & Err.Number & ";" & Err.Description
End Function


Function def_delete_folder(ByVal path_name As String) As Variant
def_delete_folder = False

Dim srt As FileSystemObject
On Error GoTo error

Set srt = New FileSystemObject
srt.DeleteFolder (path_name)
def_delete_folder = True
Exit Function
error:
def_delete_folder = "ERRO;" & "def_delete_folder;" & Err.Number & ";" & Err.Description
End Function


Function def_delete_file(ByVal path_file As String) As Variant
def_delete_file = False

Dim srt As FileSystemObject
On Error GoTo error

Set srt = New FileSystemObject
srt.DeleteFile path_file
def_delete_file = True
Exit Function
error:
def_delete_file = "ERRO;" & "def_delete_file;" & Err.Number & ";" & Err.Description
End Function


Function def_open_system_folder(Optional ByVal path As String) As Variant

def_open_system_folder = False
Dim wb As Workbook
On Error GoTo error
Set wb = ThisWorkbook
If path = "" Then path = wb.path
Shell "C:\WINDOWS\explorer.exe " & path & """", vbNormalFocus
def_open_system_folder = True
Exit Function
error:
def_open_system_folder = "ERRO;" & "def_open_system_folder;" & Err.Number & ";" & Err.Description
End Function


Function def_list_folder_item(ByVal path As String, Optional ByVal exclude_folder As Boolean, Optional ByVal exclude_file As Boolean) As Variant
def_list_folder_item = False
Dim srt As FileSystemObject
Dim fl As Scripting.Folder
Dim item_obj As Object
Dim item As Variant
Dim row_item(7) As Variant
Dim item_type As String

Dim n As Integer
Dim i As Double

row_item(0) = "item type"
row_item(1) = "path"
row_item(2) = "name"
row_item(3) = "date_created"
row_item(4) = "date_last_accessed"
row_item(5) = "date_last_modified"
row_item(6) = "size"
row_item(7) = "type"

ReDim tabela(0 To 0)
tabela(0) = row_item

Erase row_item

Set srt = New FileSystemObject
Set fl = srt.GetFolder(path)
i = 1

For n = 1 To 2
    If n = 1 Then
        If exclude_file = True Then GoTo pular
        Set item_obj = fl.Files
        item_type = "File"
    ElseIf n = 2 Then
        If exclude_folder = True Then GoTo pular
        Set item_obj = fl.SubFolders
        item_type = "Folder"
    End If

    For Each item In item_obj
        row_item(0) = item_type
        row_item(1) = item.path
        row_item(2) = item.name
        row_item(3) = item.DateCreated
        row_item(4) = item.DateLastAccessed
        row_item(5) = item.DateLastModified
        row_item(6) = item.Size
        row_item(7) = item.Type

        ReDim Preserve tabela(0 To i)
        tabela(i) = row_item
        Erase row_item
        i = i + 1
        Next item
pular:
Next n
def_list_folder_item = tabela
Exit Function
error:
def_list_folder_item = "ERRO;" & "def_list_folder_item;" & Err.Number & ";" & Err.Description
End Function


'---------------------------------------------------------------------------------------------------
'-------functions: to work with datagroup-----------------------------------------------------------
'---------------------------------------------------------------------------------------------------

Function dtg_sheet_to_array(ByVal sheet As String, Optional ByVal row_ref As Long, Optional _
                            ByVal column_ref As Long, Optional ByVal columns As Long, Optional ByVal file As String) As Variant
Dim num_row As Double
Dim num_col As Double

Dim x As Double
Dim y As Double
Dim i As Double
Dim j As Double

Dim wb As Workbook
Dim ws As Worksheet

Dim cell_ref As String

dtg_sheet_to_array = False
On Error GoTo error

If file = "" Then
    Set wb = ThisWorkbook
Else
    Set wb = Workbooks(file)
End If

Set ws = wb.Sheets(sheet)

If row_ref = Empty Then row_ref = 1
If column_ref = Empty Then column_ref = 1

cell_ref = Cells(CInt(row_ref), CInt(column_ref)).Address
If ws.Range(cell_ref).Offset(1, 0).Value = "" Then
    num_row = 1
Else
    num_row = ws.Range(cell_ref).End(xlDown).row - ws.Range(cell_ref).row + 1
End If

If columns <> 0 Then
    num_col = 1
Else
    If ws.Range(cell_ref).Offset(0, 1).Value = "" Then
        num_col = 1
    Else
        num_col = ws.Range(cell_ref).End(xlToRight).Column - ws.Range(cell_ref).Column + 1
    End If
End If

If ws.Range(cell_ref).Value = "" And num_row = 1 And num_col = 1 Then
    dtg_sheet_to_array = "ERRO;" & "transformar_em_array_ob;" & "Erro" & ";" & "A colula informada como referência está vazia e não pertence a uma tabela."
    Exit Function
End If
x = 0
y = 0
ReDim array_ob(0)
For i = 0 To num_row - 1
    j = 0
    y = 0
    If ws.Range(cell_ref).Offset(i, j).EntireRow.Hidden = False Then

        ReDim row(0)
        For j = 0 To num_col - 1
            If ws.Range(cell_ref).Offset(i, j).EntireColumn.Hidden = False Then
                ReDim Preserve row(0 To y)
                row(y) = ws.Range(cell_ref).Offset(i, j).Value
                y = y + 1

            End If
        Next j
        ReDim Preserve array_ob(0 To x)
        array_ob(x) = row
        x = x + 1

        Erase row
    End If
Next i
dtg_sheet_to_array = array_ob
Exit Function
error:
dtg_sheet_to_array = "ERRO;" & "transformar_em_array_ob;" & Err.Number & ";" & Err.Description
End Function


Function dtg_array_to_txt(ByVal array_ob As Variant, Optional ByVal spacer As String, Optional ByVal path As String, Optional ByVal name_file As String) As Variant
Dim wb As Workbook
Dim file As FileSystemObject
Dim txt As TextStream
Dim num_row As Variant
Dim lin As Variant
Dim col As Variant
Dim row As Variant
Dim var_type As Variant

On Error GoTo error
dtg_array_to_txt = False
Set wb = ThisWorkbook

If spacer = Empty Then spacer = ";"
If path = Empty Then path = wb.path & "/"
If name_file = Empty Then name_file = "file-" & Format(Now, "yyyy-mm-dd_hhmmss") & ".txt"

Set file = New FileSystemObject
Set txt = file.CreateTextFile(path & name_file, True)

num_row = UBound(array_ob)

For Each lin In array_ob
    row = ""
    For Each col In lin
        var_type = VarType(col)
        Select Case var_type
            Case Is = 7
                If row = "" Then
                    row = Format(Day(col), "0#") & "/" & Format(Month(col), "0#") & "/" & Format(Year(col), "000#")
                Else
                    row = row & spacer & Format(Day(col), "0#") & "/" & Format(Month(col), "0#") & "/" & Format(Year(col), "000#")
                End If
            Case Is = 5
                If spacer = "," Then
                    col = Replace(Replace(CStr(col), ".", ""), ",", ".")
                End If
                If row = "" Then
                    row = col
                Else
                    row = row & spacer & col
                End If
            Case Else
                If spacer = ";" Then
                    col = Replace(col, ";", ":")
                ElseIf spacer = "," Then
                    col = Replace(col, ",", ";")
                End If
                If row = "" Then
                    row = col
                Else
                    row = row & spacer & col
                End If
        End Select
    Next col
    txt.WriteLine row
Next lin
txt.Close
dtg_array_to_txt = True
Exit Function
error:
dtg_array_to_txt = "ERRO;" & "dtg_array_to_txt;" & Err.Number & ";" & Err.Description
End Function


Function dtg_list_to_txt(ByVal list As Variant, Optional ByVal path As String, Optional ByVal name_file As String) As Variant
Dim wb As Workbook
Dim file As FileSystemObject
Dim txt As TextStream
Dim num_row As Variant
Dim num_col As Variant
Dim lin As Variant
Dim col As Variant
Dim row As Variant

Dim var_type As Variant
On Error GoTo error
dtg_list_to_txt = False
Set wb = ThisWorkbook

If path = Empty Then path = wb.path & "/"
If name_file = Empty Then name_file = "file-" & Format(Now, "yyyy-mm-dd_hhmmss") & ".txt"

Set file = New FileSystemObject
Set txt = file.CreateTextFile(path & name_file, True)

num_row = UBound(list)

For Each lin In list
    row = ""
        var_type = VarType(lin)
        Select Case var_type
            Case Is = 7
                    row = Format(Day(lin), "0#") & "/" & Format(Month(lin), "0#") & "/" & Format(Year(lin), "000#")
            Case Is = 5
                    row = lin
            Case Else
                    row = lin
        End Select
    txt.WriteLine row
Next lin
txt.Close
dtg_list_to_txt = True
Exit Function
error:
dtg_list_to_txt = "ERRO;" & "dtg_list_to_txt;" & Err.Number & ";" & Err.Description
End Function


Function dtg_read_intire_txt(ByVal path_file As String, Optional ByVal linear As Boolean) As String
dtg_read_intire_txt = False
Dim srt As FileSystemObject
Dim txt As TextStream
Dim arq As String
Dim row As String
Dim texto As String
On Error GoTo error
arq = path_file
Set srt = New FileSystemObject
Set txt = srt.OpenTextFile(arq, ForReading, True)
If linear = True Then
    Do While txt.AtEndOfLine = False
        row = txt.ReadLine
        If texto = "" Then
            texto = row
            Else
            texto = texto & ";" & row
        End If
    Loop
Else
    texto = txt.ReadAll
End If
txt.Close
dtg_read_intire_txt = texto

Exit Function
error:
dtg_read_intire_txt = "ERRO;" & "dtg_read_intire_txt;" & Err.Number & ";" & Err.Description
End Function


Function dtg_read_array_txt(ByVal path_file As String, Optional ByVal spacer As String) As Variant
dtg_read_array_txt = False
Dim srt As FileSystemObject
Dim txt As TextStream
Dim num_row As Double
Dim row As Variant
Dim list As Variant
Dim array_ob() As Variant
On Error GoTo error
Set srt = New FileSystemObject
Set txt = srt.OpenTextFile(path_file, ForReading, True)

If spacer = "" Then spacer = ";"
num_row = 0
Do While txt.AtEndOfLine = False
    row = txt.ReadLine
    list = Split(row, spacer)
    ReDim Preserve array_ob(0 To num_row)
    array_ob(num_row) = list
    num_row = num_row + 1
Loop
txt.Close
dtg_read_array_txt = array_ob
Debug.Print array_ob(8)(5)
Exit Function
error:
dtg_read_array_txt = "ERRO;" & "dtg_read_array_txt;" & Err.Number & ";" & Err.Description
End Function


Function dtg_array_to_sheet(ByVal array_ob As Variant, ByVal sheet As String, Optional ByVal row_ref As Long, Optional ByVal column_ref As Long, _
                            Optional ByVal ignore_null As Boolean) As Variant
dtg_array_to_sheet = False
Dim num_row As Double
Dim num_col As Double
Dim ws As Worksheet
Dim i As Long
Dim j As Long

On Error GoTo error

If row_ref = Empty Then row_ref = 1
If column_ref = Empty Then column_ref = 1
Set ws = Sheets(sheet)

num_row = UBound(array_ob)
num_col = UBound(array_ob(0))

For i = 0 To num_row
    For j = 0 To num_col
        If ignore_null = True Then
            If array_ob(i)(j) <> Empty Then
                ws.Cells(i + row_ref, j + column_ref) = array_ob(i)(j)
            End If

        Else
            ws.Cells(i + row_ref, j + column_ref) = array_ob(i)(j)
        End If
    Next j
Next i

dtg_array_to_sheet = True
Exit Function
error:
dtg_array_to_sheet = "ERRO;" & "dtg_array_to_sheet;" & Err.Number & ";" & Err.Description
End Function


Function dtg_recordset_to_sheet(ByVal recordset As Variant, ByVal sheet As String, Optional ByVal row_ref As Long, Optional ByVal column_ref As Long) As Variant
dtg_recordset_to_sheet = False

Dim i As Double
Dim j As Double

Dim ws As Worksheet
Dim recset As ADODB.recordset
On Error GoTo error

If row_ref = Empty Then row_ref = 1
If column_ref = Empty Then column_ref = 1
Set ws = Sheets(sheet)
Set recset = recordset

i = 0
For j = 0 To recset.Fields.Count - 1
    ws.Cells(i + row_ref, j + column_ref) = recset.Fields(j).name
Next j
i = i + 1
j = 0
ws.Cells(i + row_ref, j + column_ref).CopyFromRecordset recset

dtg_recordset_to_sheet = True
Exit Function
error:
dtg_recordset_to_sheet = "ERRO;" & "dtg_recordset_to_sheet;" & Err.Number & ";" & Err.Description
End Function


Function dtg_array_transpose(ByVal array_ob As Variant) As Variant
dtg_array_transpose = False
Dim row() As Variant
Dim num_row As Double
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

dtg_array_transpose = transpor
Exit Function
error:
dtg_array_transpose = "ERRO;" & "dtg_array_transpose;" & Err.Number & ";" & Err.Description
End Function


Function dtg_array_to_html(ByVal array_ob As Variant, Optional ByVal bkg_color_th As String, Optional ByVal font_color_th As String) As Variant
dtg_array_to_html = False

Dim table() As Variant
Dim num_row As Variant
Dim num_col As Variant
Dim n As Variant
Dim i As Variant
Dim j As Variant

On Error GoTo error


If bkg_color_th = "" Then
    bkg_color_th = "WHITE"
    font_color_th = "BLACK"
End If

If UCase(bkg_color_th) = "BLACK" And font_color_th = "" Then
    font_color_th = "WHITE"
End If

num_row = UBound(array_ob)
num_col = UBound(array_ob(0))


n = 0
ReDim Preserve table(0 To n)
table(n) = "<TABLE STYLE=""border: 1px solid black "">"


For i = 0 To num_row - 1
    n = n + 1
    ReDim Preserve table(0 To n)
    If i = 0 Then
        table(n) = "<TR BGCOLOR = " & bkg_color_th & ">"
        n = n + 1
        ReDim Preserve table(0 To n)
        For j = 0 To num_col - 1
            table(n) = table(n) & "<TH><FONT COLOR = " & font_color_th & ">" & array_ob(i)(j) & "</FONT></TH>"
        Next j
        table(n) = table(n) & "</TR>"
    Else
        table(n) = "<TR>"
        For j = 0 To num_col - 1
            table(n) = table(n) & "<TD>" & array_ob(i)(j) & "</TD>"
        Next j
        table(n) = table(n) & "</TR>"
    End If

Next i
n = n + 1
ReDim Preserve table(0 To n)
table(n) = "</TABLE>"

dtg_array_to_html = table
Exit Function
error:
dtg_array_to_html = "ERRO;" & "dtg_array_to_html;" & Err.Number & ";" & Err.Description
End Function


Function dtg_recordset_to_array(ByVal recordset As Variant) As Variant
dtg_recordset_to_array = False

Dim array_ob() As Variant
Dim RC As ADODB.recordset
Dim row() As Variant
Dim i As Long
Dim j As Long
Dim num_col As Long
On Error GoTo error

Set RC = recordset

num_col = RC.Fields.Count - 1
ReDim row(0 To num_col)
RC.MoveFirst

i = 0
Do Until RC.EOF <> False
    For j = 0 To num_col
        If i = 0 Then
            row(j) = RC.Fields(j).name
        Else
            row(j) = RC.Fields(j).Value
        End If
    Next j
    ReDim Preserve array_ob(0 To i)
    array_ob(i) = row
    i = i + 1
    RC.MoveNext
Loop
dtg_recordset_to_array = array_ob
Exit Function
error:
dtg_recordset_to_array = "ERRO;" & "dtg_recordset_to_array;" & Err.Number & ";" & Err.Description
End Function


'---------------------------------------------------------------------------------------------------
'-------functions: e-mail settings------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------

Function eml_email_config(ByVal send_to As String, ByVal send_subject As String, ByVal send_body As String, _
                        Optional ByVal send_copy As String, Optional ByVal send_add As String, _
                        Optional ByVal send_hide As String, Optional ByVal send_auto As Boolean) As Variant
eml_email_config = False

Dim ot As Outlook.Application
Dim email As Outlook.MailItem
Dim list_anexo As Variant
Dim nx As Variant
On Error GoTo error


Set ot = New Outlook.Application
Set email = ot.CreateItem(olMailItem)

With email
    .To = send_to
    .Subject = send_subject
    .HTMLBody = send_body
    If send_copy <> "" Then
        .CC = send_copy
    End If
    If send_add <> "" Then
        list_anexo = Split(send_add, ";")
        For Each nx In list_anexo
            email.Attachments.Add nx
        Next nx
    End If
    If send_hide <> "" Then
        .BCC = send_hide
    End If
End With

If send_auto = False Then
    email.Display
Else
    email.Send

End If

eml_email_config = True & ";eml-email_config"
Exit Function
error:
eml_email_config = "ERRO;" & "eml-email_config;" & Err.Number & ";" & Err.Description
End Function


'---------------------------------------------------------------------------------------------------
'-------functions: usinng in sheet------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------

Function fun_symbol_off(ByVal string_value As String) As Variant

Dim str_len As Double
Dim str_new As String
Dim item As Double
Dim str_caract As String
Dim num_asc As Integer
str_new = ""
str_len = Len(string_value)
If str_len < 1 Then Exit Function

For item = 1 To str_len
    str_caract = Mid(string_value, item, 1)
    num_asc = Asc(str_caract)
    If num_asc >= 192 And num_asc <= 198 Then
        str_new = str_new & Chr(65)
    ElseIf num_asc >= 224 And num_asc <= 230 Then
        str_new = str_new & Chr(97)
    ElseIf num_asc = 32 Then
        str_new = str_new & Chr(32)
    ElseIf num_asc >= 48 And num_asc <= 57 Then
        str_new = str_new & str_caract
    ElseIf num_asc >= 65 And num_asc <= 90 Then
        str_new = str_new & str_caract
    ElseIf num_asc >= 97 And num_asc <= 122 Then
        str_new = str_new & str_caract
    ElseIf num_asc >= 65 And num_asc <= 90 Then
        str_new = str_new & str_caract
    ElseIf num_asc = 216 Then
        str_new = str_new & Chr(48)
    ElseIf num_asc = 222 Or num_asc = 254 Then
        str_new = str_new & Chr(98)
    ElseIf num_asc = 162 Or num_asc = 199 Then
        str_new = str_new & Chr(67)
    ElseIf num_asc = 231 Then
        str_new = str_new & Chr(99)
    ElseIf num_asc = 208 Then
        str_new = str_new & Chr(68)
    ElseIf num_asc >= 200 And num_asc <= 203 Then
        str_new = str_new & Chr(69)
    ElseIf num_asc >= 232 And num_asc <= 235 Then
        str_new = str_new & Chr(101)
    ElseIf num_asc >= 204 And num_asc <= 207 Then
        str_new = str_new & Chr(73)
    ElseIf num_asc >= 236 And num_asc <= 239 Or num_asc = 161 Then
        str_new = str_new & Chr(105)
    ElseIf num_asc = 209 Then
        str_new = str_new & Chr(78)
    ElseIf num_asc = 241 Then
        str_new = str_new & Chr(110)
    ElseIf num_asc >= 211 And num_asc <= 211 Then
        str_new = str_new & Chr(79)
    ElseIf num_asc >= 214 And num_asc <= 214 Then
        str_new = str_new & Chr(79)
    ElseIf num_asc >= 240 And num_asc <= 246 Then
        str_new = str_new & Chr(111)
    ElseIf num_asc = 142 Then
        str_new = str_new & Chr(90)
    ElseIf num_asc = 158 Then
        str_new = str_new & Chr(122)
    ElseIf num_asc = 159 Or num_asc = 221 Then
        str_new = str_new & Chr(89)
    ElseIf num_asc = 253 Or num_asc = 255 Then
        str_new = str_new & Chr(121)
    ElseIf num_asc >= 217 And num_asc <= 220 Then
        str_new = str_new & Chr(85)
    ElseIf num_asc >= 249 And num_asc <= 252 Then
        str_new = str_new & Chr(117)
End If
Next item
On Error GoTo error

fun_symbol_off = str_new
Exit Function
error:
fun_symbol_off = "ERRO;" & "fun_symbol_off;" & Err.Number & ";" & Err.Description
End Function


Function fun_split_off(ByVal texto As String, Optional ByVal spacer As String) As Variant
Dim sp As Variant
sp = Split(texto, spacer)
fun_split_off = sp
End Function


Function fun_concat_split_off(ByVal list As Variant, Optional ByVal spacer As String) As String
Dim cl As Variant
Dim concatenate As String
Dim sp As String

If spacer = ";" Then
sp = ":"
ElseIf spacer = "," Then
sp = "."
End If

For Each cl In list
    If concatenate = "" Then
        concatenate = Replace(cl, spacer, sp)
    Else
        concatenate = concatenate & spacer & Replace(cl, spacer, sp)
    End If
Next
fun_concat_split_off = concatenate
End Function


'---------------------------------------------------------------------------------------------------
'-------functions: to user interface----------------------------------------------------------------
'---------------------------------------------------------------------------------------------------

Function msg_msg_config(ByVal answer As Variant, Optional ByVal EN_lang As Boolean) As Boolean
msg_msg_config = False
Dim list As Variant
Dim var_type As Variant
On Error GoTo error
var_type = VarType(answer)
If var_type = 8 Then
    list = Split(answer, ";")
ElseIf var_type = 11 Then
    ReDim list(0 To 0)
    list(0) = answer
ElseIf var_type = 2 Or var_type = 3 Or var_type = 4 Or var_type = 5 Or var_type = 6 Then
    ReDim list(0 To 0)
    If answer = 1 Then
        list(0) = True
    ElseIf answer = 0 Then
        list(0) = False
    Else
        GoTo error
    End If
End If
If UCase(list(0)) Like "ERRO" Then
    MsgBox list(2) & Chr(10) & list(3), vbCritical + vbOKOnly, "Erro em: " & list(1)
ElseIf list(0) = False Then
    If EN_lang = True Then
        MsgBox "Task possibly not achieved ", vbExclamation + vbOKOnly, "Executed without error. Task not completed"
    Else
        MsgBox "Tarefa possivelmente não alcançado", vbExclamation + vbOKOnly, "Executado sem erro. Tarefa não foi concluida"
    End If
ElseIf list(0) = True Then
    If EN_lang = True Then
        MsgBox "Process executed successfully.", vbInformation + vbOKOnly, "successfully executed"
    Else
        MsgBox "Processo executado com sucesso.", vbInformation + vbOKOnly, "Executado com sucesso"
    End If

    msg_msg_config = True
Else
    GoTo error
End If
Exit Function
error:
If EN_lang = True Then
    MsgBox "Enter only variables of type:" & Chr(10) & _
    "Bollean (True or False);" & Chr(10) & _
    "Integer values ??(1 for true or 0 for false);" & Chr(10) & _
    "Standardized text (See manual).", vbCritical + vbOKOnly, "Error in: msg_msg_config"
Else
    MsgBox "Informe apenas variáveis do tipo:" & Chr(10) & _
    "Bollean (Verdadeiro ou Faso);" & Chr(10) & _
    "Valores inteiros (1 para verdadeiro ou 0 para falso);" & Chr(10) & _
    "Texto padronizado (Consulte o manual).", vbCritical + vbOKOnly, "Erro em: msg_msg_config"
End If

End Function


'---------------------------------------------------------------------------------------------------
'-------functions: sql connection-------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------

Function sql_connection_access(ByVal path_file As String, Optional ByVal verbose As Boolean, Optional ByVal password As String) As Variant
sql_connection_access = False

Dim ole_version As String
Dim app As Application
Dim acCN As ADODB.connection
On Error GoTo error

Set app = Application
Set acCN = New ADODB.connection

ole_version = app.Version

If password = "" Then
    acCN.Open "Provider=Microsoft.ACE.OLEDB." & ole_version & ".0;Data Source=" & path_file & ";Persist Security Info=False"
Else
    acCN.ConnectionString = "Provider=Microsoft.ACE.OLEDB." & ole_version & ";Data Source=" & path_file & ";Jet OLEDB:Datsheetse Password=" & password
End If

Set sql_connection_access = acCN
Exit Function
error:
If verbose = True Then
    msg_msg_config ("ERRO;" & "sql_connection_access;" & Err.Number & ";" & Err.Description)
End If
Set sql_connection_access = acCN
End Function


Function sql_connection_excel(ByVal path_file As String, Optional ByVal verbose As Boolean) As Variant
sql_connection_excel = False

Dim ole_version As String
Dim app As Application
Dim exCN As ADODB.connection
On Error GoTo error


Set app = Application
Set exCN = New ADODB.connection

ole_version = app.Version


    If ole_version < 12 Then
        exCN.ConnectionString = _
          "Provider=Microsoft.Jet.OLEDB.4.0;" & _
          "Data Source=" & path_file & ";" & _
          "Extended Properties=Excel 8.0"
    Else
        exCN.ConnectionString = _
          "Provider=Microsoft.ACE.OLEDB." & ole_version & ";" & _
          "Data Source=" & path_file & ";" & _
          "Extended Properties=Excel 8.0"
    End If
    exCN.Open

Set sql_connection_excel = exCN
Exit Function
error:
If verbose = True Then
    msg_msg_config ("ERRO;" & "sql_connection_excel;" & Err.Number & ";" & Err.Description)
End If
Set sql_connection_excel = exCN
End Function


Function sql_connection_txt(ByVal path As String, Optional ByVal verbose As Boolean) As Variant
sql_connection_txt = False
Dim txtCN As ADODB.connection
Dim ole_version As String
Dim app As Application
On Error GoTo error

Set txtCN = New ADODB.connection
Set app = Application

ole_version = app.Version
txtCN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & path & _
";Extended Properties=""TEXT;HDR=Yes;FMT=Delimited"""

txtCN.Open

Set sql_connection_txt = txtCN
Exit Function
error:
If verbose = True Then
    msg_msg_config ("ERRO;" & "sql_connection_txt;" & Err.Number & ";" & Err.Description)
End If
Set sql_connection_txt = txtCN
End Function


Function sql_connection_sharepoint(ByVal sp_site As String, ByVal sp_list As String, Optional ByVal verbose As Boolean, Optional ByVal password As String) As Variant
sql_connection_sharepoint = False

Dim ole_version As String
Dim app As Application
Dim spCN As ADODB.connection
On Error GoTo error

Set app = Application
Set spCN = New ADODB.connection

ole_version = app.Version

spCN.ConnectionString = "Provider=Microsoft.ACE.OLEDB." & ole_version & ";WSS;IMEX=0;RetrieveIds=Yes;DATABASE=" & sp_site & ";LIST={" & sp_list & "};"
spCN.Open

Set sql_connection_sharepoint = spCN
Exit Function
error:
If verbose = True Then
    msg_msg_config ("ERRO;" & "sql_connection_sharepoint;" & Err.Number & ";" & Err.Description)
End If
Set sql_connection_sharepoint = spCN
End Function


Function sql_query(ByVal connection As Variant, ByVal Query As String, Optional ByVal verbose As Boolean) As Variant
sql_query = False

Dim RC As ADODB.recordset
On Error GoTo error

Set RC = connection.Execute(Query)

If UCase(Left(Query, 6)) = "SELECT" Then
    Set sql_query = RC
Else
    Set sql_query = Nothing
End If
Exit Function
error:
If verbose = True Then
    msg_msg_config ("ERRO;" & "sql_query;" & Err.Number & ";" & Err.Description)
End If
Set sql_query = Nothing
End Function


'---------------------------------------------------------------------------------------------------
'-------functions: data validate--------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------

Function vld_validate_date(ByVal date_value As Variant, Optional ByVal reference_text As String, Optional ByVal verbose As Boolean, Optional ByVal EN_lang As Boolean) As Variant
vld_validate_date = False

Dim chg_type As Variant
Dim msg As Variant
On Error GoTo error
On Error Resume Next

If reference_text = "" Then
    reference_text = "vld_validate_date"
End If
chg_type = CDate(date_value)

If verbose = True And Err.Number = 13 Then
    If EN_lang = True Then
        msg = MsgBox("Please enter a valid date in the field " & reference_text, vbExclamation + vbOKOnly, "Erro em: " & reference_text)
    Else
        msg = MsgBox("Favor informar uma data válida no campo " & reference_text, vbExclamation + vbOKOnly, "Erro em: " & reference_text)
    End If
    vld_validate_date = False
    Exit Function
ElseIf Err.Number = 0 Then
    vld_validate_date = chg_type
    Exit Function
ElseIf Err.Number <> 0 And Err.Number <> 13 Then
    GoTo error
End If
Exit Function

error:
vld_validate_date = "ERRO;" & "vld_validate_date;" & Err.Number & ";" & Err.Description
End Function


Function vld_validate_integer(ByVal integer_value As Variant, Optional ByVal reference_text As String, Optional ByVal verbose As Boolean, Optional ByVal EN_lang As Boolean) As Variant
vld_validate_integer = False

Dim chg_type As Variant
Dim msg As Variant
On Error GoTo error
On Error Resume Next

If reference_text = "" Then
    reference_text = "vld_validate_integer"
End If
chg_type = CLng(integer_value)

If verbose = True And Err.Number = 13 Then
    If EN_lang = True Then
        msg = MsgBox("Please enter a valid integer in the field " & reference_text, vbExclamation + vbOKOnly, "Erro em: " & reference_text)
    Else
        msg = MsgBox("Favor informar um número inteiro válido no campo " & reference_text, vbExclamation + vbOKOnly, "Erro em: " & reference_text)
    End If
    vld_validate_integer = False
    Exit Function
ElseIf Err.Number = 0 Then
    vld_validate_integer = chg_type
    Exit Function
ElseIf Err.Number <> 0 And Err.Number <> 13 Then
    GoTo error
End If
Exit Function

error:
vld_validate_integer = "ERRO;" & "vld_validate_integer;" & Err.Number & ";" & Err.Description
End Function


Function vld_validate_double(ByVal double_value As Variant, Optional ByVal reference_text As String, Optional ByVal verbose As Boolean, Optional ByVal EN_lang As Boolean) As Variant
vld_validate_double = False

Dim chg_type As Double
Dim msg As Variant
On Error GoTo error
On Error Resume Next

If reference_text = "" Then
    reference_text = "vld_validate_double"
End If
chg_type = CDbl(double_value)

If verbose = True And Err.Number = 13 Then
    If EN_lang = True Then
        msg = MsgBox("Please enter a valid number in the field " & reference_text, vbExclamation + vbOKOnly, "Erro em: " & reference_text)
    Else
        msg = MsgBox("Favor informar um número válido no campo " & reference_text, vbExclamation + vbOKOnly, "Erro em: " & reference_text)
    End If
    vld_validate_double = False
    Exit Function
ElseIf Err.Number = 0 Then
    vld_validate_double = chg_type
    Exit Function
ElseIf Err.Number <> 0 And Err.Number <> 13 Then
    GoTo error
End If
Exit Function

error:
vld_validate_double = "ERRO;" & "vld_validate_double;" & Err.Number & ";" & Err.Description
End Function


Function vld_validate_string(ByVal string_value As Variant, Optional ByVal reference_text As String, Optional ByVal verbose As Boolean, Optional ByVal EN_lang As Boolean) As Variant
vld_validate_string = False

Dim chg_type As Variant
Dim msg As Variant
On Error GoTo error
On Error Resume Next

If reference_text = "" Then
    reference_text = "vld_validate_string"
End If
chg_type = CStr(string_value)
If verbose = True And Err.Number = 13 Then

    If EN_lang = True Then
        msg = MsgBox("Please enter valid text in the field " & reference_text, vbExclamation + vbOKOnly, "Erro em: " & reference_text)
    Else
        msg = MsgBox("Favor informar texto válido no campo " & reference_text, vbExclamation + vbOKOnly, "Erro em: " & reference_text)
    End If
    vld_validate_string = False
    Exit Function
ElseIf verbose = True And chg_type = "" Then
    If EN_lang = True Then
        msg = MsgBox("The " & reference_text & " field cannot be empty " & reference_text, vbExclamation + vbOKOnly, "Erro em: " & reference_text)
    Else
        msg = MsgBox("O campo " & reference_text & " não pode ficar vazio", vbExclamation + vbOKOnly, "Erro em: " & reference_text)
    End If
    vld_validate_string = False
    Exit Function
ElseIf Err.Number = 0 Then
    vld_validate_string = chg_type
    Exit Function
ElseIf Err.Number <> 0 And Err.Number <> 13 Then
    GoTo error
End If
Exit Function

error:
vld_validate_string = "ERRO;" & "vld_validate_string;" & Err.Number & ";" & Err.Description
End Function


Function vld_validate_not_blanc(ByVal string_value As Variant, Optional ByVal reference_text As String, Optional ByVal verbose As Boolean, Optional ByVal EN_lang As Boolean) As Variant
vld_validate_not_blanc = False

Dim chg_type As Variant
Dim msg As Variant
On Error GoTo error
On Error Resume Next

If reference_text = "" Then
    reference_text = "vld_validate_string"
End If
chg_type = string_value

If verbose = True And chg_type = "" Then
    If EN_lang = True Then
        msg = MsgBox("The " & reference_text & " field cannot be empty", vbExclamation + vbOKOnly, "Erro em: " & reference_text)
    Else
        msg = MsgBox("O campo " & reference_text & " não pode ficar vazio", vbExclamation + vbOKOnly, "Erro em: " & reference_text)
    End If
    vld_validate_not_blanc = False
    Exit Function
ElseIf Err.Number = 0 Then
    vld_validate_not_blanc = chg_type
    Exit Function
ElseIf Err.Number <> 0 And Err.Number <> 13 Then
    GoTo error
End If
Exit Function

error:
vld_validate_not_blanc = "ERRO;" & "vld_validate_blanc;" & Err.Number & ";" & Err.Description
End Function


'---------------------------------------------------------------------------------------------------
'-------functions: excel file and objects settings--------------------------------------------------
'---------------------------------------------------------------------------------------------------

Function xls_delete_sheets(ByVal included As String, Optional ByVal file As String) As Variant

Dim wb As Workbook
Dim wbs As Workbooks
Dim st As Worksheet
Dim arq As Variant
Dim sheet As Variant
Dim name As String
On Error GoTo error
xls_delete_sheets = False

If file = Empty Then
    Set wb = ThisWorkbook
Else
    Set wbs = Workbooks
    For Each arq In wbs
        name = UCase(arq.name)
        If name Like UCase(file) Then
            Set wb = arq
            Exit For
        End If
    Next arq
End If
If wb Is Nothing Then
    Exit Function
End If

app_app_config_off

For Each sheet In wb.Sheets
    Set st = sheet
    If UCase(st.name) Like UCase(included) Then
        st.Delete
    End If
Next sheet
xls_delete_sheets = True
app_app_config_on
Exit Function
error:
app_app_config_on
xls_delete_sheets = "ERRO;" & "xls_delete_sheets;" & Err.Number & ";" & Err.Description
End Function


Function xls_refresh_query(ByVal sheet_name As String, Optional ByVal file As String) As Variant

Dim wb As Workbook
Dim wbs As Workbooks
Dim st As Worksheet

On Error GoTo error
xls_refresh_query = False

If file = Empty Then
    Set wb = ThisWorkbook
Else
    Set wbs = Workbooks
    Set wb = wbs(file)
End If
If wb Is Nothing Then
    Exit Function
End If

app_app_config_off

Set st = wb.Sheets(sheet_name)

st.ListObjects.item(1).QueryTable.Refresh BackgroundQuery:=False
wb.Save

xls_refresh_query = True
app_app_config_on
Exit Function
error:
app_app_config_on
xls_refresh_query = "ERRO;" & "xls_refresh_query;" & Err.Number & ";" & Err.Description
End Function


Function xls_create_sheet(Optional ByVal name_sheet As String, Optional location As Integer, Optional ByVal file As String) As Variant
Dim wb As Workbook
Dim wbs As Workbooks
Dim name As String
Dim arq As Variant

On Error GoTo error
xls_create_sheet = False

If file = Empty Then
    Set wb = ThisWorkbook
Else
    Set wbs = Workbooks
    For Each arq In wbs
        name = UCase(arq.name)
        If name Like UCase(file) Then
            Set wb = arq
            Exit For
        End If
    Next arq
End If

If name_sheet = Empty Then
    name_sheet = "sheet_" & Format(Time, "hh-mm-ss")
End If

If location = 1 Then         'a direita da sheet ativa
    wb.Sheets.Add(After:=wb.ActiveSheet).name = name_sheet
ElseIf location = 2 Then     'no in�cio
    wb.Sheets.Add(before:=wb.Sheets(1)).name = name_sheet
ElseIf location = 3 Then     'no final
    wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)).name = name_sheet
Else                        'a esquerda da sheet atual
    wb.Sheets.Add(before:=wb.ActiveSheet).name = name_sheet
End If

xls_create_sheet = True & ";" & name_sheet
Exit Function
error:

xls_create_sheet = "ERRO;" & "xls_create_sheet;" & Err.Number & ";" & Err.Description
End Function


Function xls_copy_sheet(ByVal name_sheet As String, Optional ByVal new_name_sheet As String, Optional location As Integer, Optional ByVal file As String) As Variant
Dim wb As Workbook
Dim wbs As Workbooks
Dim st As Worksheet
Dim name As String
Dim arq As Variant

On Error GoTo error
xls_copy_sheet = False

If file = Empty Then
    Set wb = ThisWorkbook
Else
    Set wbs = Workbooks
    For Each arq In wbs
        name = UCase(arq.name)
        If name Like UCase(file) Then
            Set wb = arq
            Exit For
        End If
    Next arq
End If

Set st = wb.Sheets(name_sheet)

If new_name_sheet = Empty Or new_name_sheet = "" Then
    st.Copy
    xls_copy_sheet = True & ";" & ActiveWorkbook.name
Else
    If location = 1 Then         'a direita da sheet ativa
        st.Copy After:=st
        wb.ActiveSheet.name = name_sheet
    ElseIf location = 2 Then     'no in�cio
        st.Copy before:=wb.Sheets(1)
        wb.ActiveSheet.name = name_sheet
    ElseIf location = 3 Then     'no final
        st.Copy After:=wb.Sheets(wb.Sheets.Count)
        wb.ActiveSheet.name = name_sheet
    Else                        'a esquerda da sheet atual
        st.Copy before:=st
        wb.ActiveSheet.name = name_sheet
    End If

    xls_copy_sheet = True & ";" & name_sheet
End If

Exit Function
error:

xls_copy_sheet = "ERRO;" & "xls_copy_sheet;" & Err.Number & ";" & Err.Description
End Function


Function xls_hide_sheet(ByVal name_sheet As String, Optional ByVal file As String) As Variant
xls_hide_sheet = False

Dim wb As Workbook
Dim wbs As Workbooks
Dim st As Worksheet
Dim name As String
Dim arq As Variant

On Error GoTo error
xls_hide_sheet = False

If file = Empty Then
    Set wb = ThisWorkbook
Else
    Set wbs = Workbooks
    For Each arq In wbs
        name = UCase(arq.name)
        If name Like UCase(file) Then
            Set wb = arq
            Exit For
        End If
    Next arq
End If

Set st = wb.Sheets(name_sheet)
st.Visible = xlSheetHidden
xls_hide_sheet = True & ";xls_hide_sheet"
Exit Function
error:
xls_hide_sheet = "ERRO;" & "xls_hide_sheet;" & Err.Number & ";" & Err.Description
End Function


Function xls_veryhide_sheet(ByVal name_sheet As String, Optional ByVal file As String) As Variant
xls_veryhide_sheet = False

Dim wb As Workbook
Dim wbs As Workbooks
Dim st As Worksheet
Dim name As String
Dim arq As Variant

On Error GoTo error
xls_veryhide_sheet = False

If file = Empty Then
    Set wb = ThisWorkbook
Else
    Set wbs = Workbooks
    For Each arq In wbs
        name = UCase(arq.name)
        If name Like UCase(file) Then
            Set wb = arq
            Exit For
        End If
    Next arq
End If

Set st = wb.Sheets(name_sheet)
st.Visible = xlSheetVeryHidden
xls_veryhide_sheet = True & ";xls_veryhide_sheet"
Exit Function
error:
xls_veryhide_sheet = "ERRO;" & "xls_veryhide_sheet;" & Err.Number & ";" & Err.Description
End Function


Function xls_show_sheet(ByVal name_sheet As String, Optional ByVal file As String) As Variant
xls_show_sheet = False

Dim wb As Workbook
Dim wbs As Workbooks
Dim st As Worksheet
Dim name As String
Dim arq As Variant

On Error GoTo error
xls_show_sheet = False

If file = Empty Then
    Set wb = ThisWorkbook
Else
    Set wbs = Workbooks
    For Each arq In wbs
        name = UCase(arq.name)
        If name Like UCase(file) Then
            Set wb = arq
            Exit For
        End If
    Next arq
End If

Set st = wb.Sheets(name_sheet)
st.Visible = xlSheetVisible
xls_show_sheet = True & ";xls_show_sheet"
Exit Function
error:
xls_show_sheet = "ERRO;" & "xls_show_sheet;" & Err.Number & ";" & Err.Description
End Function


Function xls_save_as_excel(Optional ByVal file As String, Optional ByVal path As String, Optional ByVal name_file As String) As Variant
xls_save_as_excel = False

Dim wb As Workbook
Dim dwb As Workbook
Dim wbs As Workbooks
Dim name As String
Dim arq As Variant

On Error GoTo error
Set dwb = ThisWorkbook

If path = Empty Then
    path = dwb.path & "\"
End If

If file = Empty Then
    Set wb = ActiveWorkbook 'atenção ! até agora é a única função fazendo referência ao file ativo e não a este file
Else
    Set wbs = Workbooks
    For Each arq In wbs
        name = UCase(arq.name)
        If name Like UCase(file) Then
            Set wb = arq
            Exit For
        End If
    Next arq
End If

If name_file = Empty Then
    name_file = wb.name
End If

wb.SaveAs (path & name_file)

xls_save_as_excel = True & ";xls_save_as_excel"
Exit Function
error:
xls_save_as_excel = "ERRO;" & "xls_save_as_excel;" & Err.Number & ";" & Err.Description
End Function


Function xls_save_excel(Optional ByVal file As String) As Variant
xls_save_excel = False


Dim wb As Workbook
Dim wbs As Workbooks
Dim name As String
Dim arq As Variant

On Error GoTo error

If file = Empty Then
    Set wb = ActiveWorkbook
Else
    Set wbs = Workbooks
    For Each arq In wbs
        name = UCase(arq.name)
        If name Like UCase(file) Then
            Set wb = arq
            Exit For
        End If
    Next arq
End If

wb.Save

xls_save_excel = True & ";xls_save_excel"
Exit Function
error:
xls_save_excel = "ERRO;" & "xls_save_excel;" & Err.Number & ";" & Err.Description
End Function


Function xls_close_excel(Optional ByVal file As String) As Variant

xls_close_excel = False

Dim wb As Workbook
Dim wbs As Workbooks
Dim name As String
Dim arq As Variant
Dim save_alteration As Boolean

save_alteration = False

On Error GoTo error

If file = Empty Then
    Set wb = ActiveWorkbook 'atenção ! até agora é a única função fazendo referência ao file ativo e não a este file
Else
    Set wbs = Workbooks
    For Each arq In wbs
        name = UCase(arq.name)
        If name Like UCase(file) Then
            Set wb = arq
            Exit For
        End If
    Next arq
    If wb Is Nothing Then GoTo error
End If

wb.Close savechanges:=save_alteration

xls_close_excel = True & ";xls_close_excel"
Exit Function
error:
xls_close_excel = "ERRO;" & "xls_close_excel;" & Err.Number & ";" & Err.Description
End Function


Function xls_open_excel(ByVal file As String, Optional ByVal path As String) As Variant
xls_open_excel = False

Dim wb As Workbook
Dim wbs As Workbooks
On Error GoTo error

Set wb = ThisWorkbook
Set wbs = Workbooks
If path = Empty Then
    path = wb.path & "\"
End If
wbs.Open Filename:=path & file
xls_open_excel = True & ";xls_open_excel"
Exit Function
error:
xls_open_excel = "ERRO;" & "xls_open_excel;" & Err.Number & ";" & Err.Description
End Function


Function xls_protect_sheet(ByVal password As String, Optional ByVal sheet As String, Optional ByVal file As String) As Variant

Dim wb As Workbook
Dim wbs As Workbooks
Dim st As Worksheet
Dim sts As Variant
Dim name As String
Dim arq As Variant

On Error GoTo error

If file = Empty Then
    Set wb = ActiveWorkbook 'atenção ! até agora é a única função fazendo referência ao file ativo e não a este file
Else
    Set wbs = Workbooks
    For Each arq In wbs
        name = UCase(arq.name)
        If name Like UCase(file) Then
            Set wb = arq
            Exit For
        End If
    Next arq
End If

If sheet = "" Then
    Set st = wb.ActiveSheet
    st.Protect password:=password, DrawingObjects:=True, Contents:=True, Scenarios:=True
Else
    Set st = wb.Sheets(sheet)
    st.Protect password:=password, DrawingObjects:=True, Contents:=True, Scenarios:=True
End If

xls_protect_sheet = True & ";xls_protect_sheet"
Exit Function
error:
xls_protect_sheet = "ERRO;" & "xls_protect_sheet;" & Err.Number & ";" & Err.Description
End Function


Function xls_unprotect_sheet(ByVal password As String, Optional ByVal sheet As String, Optional ByVal file As String) As Variant

Dim wb As Workbook
Dim wbs As Workbooks
Dim st As Worksheet
Dim sts As Variant
Dim name As String
Dim arq As Variant

On Error GoTo error

If file = Empty Then
    Set wb = ActiveWorkbook 'atenção ! até agora é a única função fazendo referência ao file ativo e não a este file
Else
    Set wbs = Workbooks
    For Each arq In wbs
        name = UCase(arq.name)
        If name Like UCase(file) Then
            Set wb = arq
            Exit For
        End If
    Next arq
End If

If sheet = "" Then
    Set st = wb.ActiveSheet
    st.Unprotect password:=password
Else
    Set st = wb.Sheets(sheet)
    st.Unprotect password:=password
End If

xls_unprotect_sheet = True & ";xls_unprotect_sheet"
Exit Function
error:
xls_unprotect_sheet = "ERRO;" & "xls_unprotect_sheet;" & Err.Number & ";" & Err.Description
End Function


Function xls_lock_file(ByVal password As String, Optional ByVal file As String) As Variant

Dim wb As Workbook
Dim wbs As Workbooks
Dim name As String
Dim arq As Variant

On Error GoTo error

If file = Empty Then
    Set wb = ActiveWorkbook 'atenção ! até agora é a única função fazendo referência ao file ativo e não a este file
Else
    Set wbs = Workbooks
    For Each arq In wbs
        name = UCase(arq.name)
        If name Like UCase(file) Then
            Set wb = arq
            Exit For
        End If
    Next arq
End If

wb.Protect password:=password, Structure:=True, Windows:=False

xls_lock_file = True & ";xls_lock_file"
Exit Function
error:
xls_lock_file = "ERRO;" & "xls_lock_file;" & Err.Number & ";" & Err.Description
End Function


Function xls_unlock_file(ByVal password As String, Optional ByVal file As String) As Variant

Dim wb As Workbook
Dim wbs As Workbooks
Dim name As String
Dim arq As Variant

On Error GoTo error

If file = Empty Then
    Set wb = ActiveWorkbook 'atenção ! até agora é a única função fazendo referência ao file ativo e não a este file
Else
    Set wbs = Workbooks
    For Each arq In wbs
        name = UCase(arq.name)
        If name Like UCase(file) Then
            Set wb = arq
            Exit For
        End If
    Next arq
End If

wb.Unprotect password:=password

xls_unlock_file = True & ";xls_unlock_file"
Exit Function
error:
xls_unlock_file = "ERRO;" & "xls_unlock_file;" & Err.Number & ";" & Err.Description
End Function