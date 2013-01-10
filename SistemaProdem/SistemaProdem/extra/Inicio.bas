Attribute VB_Name = "Inicio"
Public Banco_path As String, Banco_path_CEP As String
Public ocnBanco As New ADODB.Connection

Sub Main()

If App.Path <> "E:\Work Fire\Sistema\pf" And App.Path <> "E:\Work Fire\Sistema\extra" Then
'If App.Path <> "C:\temp\wf\Sistema\pf" Then
    'Em modo executável
    Banco_path = App.Path & "\banco\1bd_wf.mdb"
    Banco_path_CEP = App.Path & "\banco\cepbr.mdb"
Else
    'Em modo debug
    'Banco_path = "C:\temp\wf\" & "\banco\1bd_wf.mdb" 'para usar em modo debug
    'Banco_path_CEP = "C:\temp\wf\" & "\banco\cepbr.mdb" 'caminho
    
    Banco_path = "E:\Work Fire\" & "banco\1bd_wf.mdb" 'para usar em modo debug
    Banco_path_CEP = "E:\Work Fire\" & "\banco\cepbr.mdb" 'caminho
End If

If Open_Banco = True Then
Form1.Show
Else
MsgBox "Não foi possível abrir o Banco de Dados"
End
End If
End Sub

Function Open_Banco() As Boolean
' procedimento de abertura do banco de dados
On Error GoTo erro
'ocnBanco.Open "dsn=db_wf"
ocnBanco.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Banco_path

'ocnBanco.Open "Driver={SQL Server};Server=.;Uid=sa;Pwd=;Database=wf"


Open_Banco = True
Exit Function
erro:
Open_Banco = False
MsgBox Err.Description
End Function


