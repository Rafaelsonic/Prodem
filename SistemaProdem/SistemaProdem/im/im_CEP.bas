Attribute VB_Name = "CEP"
Option Explicit
Dim ocn_BancoCEP As New ADODB.Connection
Dim snap_Selecao As New ADODB.Recordset ' objeto da seleção
Dim SQL As String
Public CEP_Nome As String, CEP_Bairro As String, CEP_Cidade As String

Public Function Busca_Cep(ByVal NrCEP As String) As Boolean
Busca_Cep = False
If Open_BancoCEP = True Then
    SQL = "select * from cep_sp where cep =" & NrCEP
    Set snap_Selecao = ocn_BancoCEP.Execute(SQL)
    If Not (snap_Selecao.EOF And snap_Selecao.BOF) Then
        CEP_Nome = snap_Selecao("ABREVI") & ". " & snap_Selecao("NOME")
        CEP_Bairro = snap_Selecao("BAIRRO")
        CEP_Cidade = snap_Selecao("CIDADE")
        Busca_Cep = True
    Else
        MsgBox "CEP Não Encontrado no Estado de São Paulo!", , "Busca_CEP"
    End If
End If
End Function

Function Open_BancoCEP() As Boolean
Set ocn_BancoCEP = Nothing

' procedimento de abertura do banco de dados
On Error GoTo erro
ocn_BancoCEP.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Banco_path_CEP

Open_BancoCEP = True
Exit Function
erro:
Open_BancoCEP = False
MsgBox Err.Description
End Function

