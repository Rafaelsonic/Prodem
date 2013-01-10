VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_ReldesSint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Classe do relatório de despesa Sintético
'programa pf_RelDesSin

Option Explicit
Dim nomeTempTab(1) As String ' nome da tabela temporaria
Dim Funcoes As clsBasic ' classe dasprincipais funcoes
Function Mostra_emp(cod As String) As String
Dim snap_Selecao As New ADODB.Recordset
Dim Sql As String
If IsNumeric(cod) Then
    Sql = "select emp_rzsc from im_empresa " & vbCrLf
    Sql = Sql & "where emp_cod=" & cod
    Set snap_Selecao = ocnBanco.Execute(Sql)
    If Not (snap_Selecao.BOF And snap_Selecao.EOF) Then
        Mostra_emp = snap_Selecao(0)
    Else
        Mostra_emp = ""
    End If
Else
    Mostra_emp = ""
End If
End Function

Private Sub Class_Initialize()
nomeTempTab(0) = "temp_relSinH" 'header dados do cabeçalho
nomeTempTab(1) = "temp_relSinI" 'item
End Sub

Sub relatorio(str_cond As String, campo As String)
Dim Sql As String

On Error GoTo erro
DropTabelas
If campo = "des_ccusto" Then
    'Select para relatorio sintetico de despesas por Centro de Custo
    Sql = "SELECT des_emp, des_ccusto as des_cod, ccus_desc as des_desc, Month(des_comp) AS mes," & Chr(13) & _
            "Sum(pf_lancdes.des_vr) AS valor into " & Chr(13) & _
            nomeTempTab(1) & Chr(13) & _
            "From pf_lancDes, pf_ccusto" & Chr(13) & _
            "Where " & Chr(13) & _
            " des_ccusto= ccus_cod" & Chr(13) & _
            str_cond & Chr(13) & _
            "GROUP BY pf_lancdes.des_ccusto, Month(des_comp), ccus_desc, des_emp;"
ElseIf campo = "des_fornec" Then
    'Select para relatorio sintetico de despesas por Fornecedor
    Sql = "SELECT des_emp, des_fornec as des_cod, for_rzsc as des_desc, Month(des_comp) AS mes," & Chr(13) & _
            "Sum(pf_lancdes.des_vr) AS valor into " & Chr(13) & _
            nomeTempTab(1) & Chr(13) & _
            "From pf_lancDes, pf_fornec" & Chr(13) & _
            "Where " & Chr(13) & _
            " des_fornec= for_cod" & Chr(13) & _
            str_cond & Chr(13) & _
            "GROUP BY des_fornec, Month(des_comp), for_rzsc, des_emp;"
End If
ocnBanco.Execute Sql

erro:
Select Case Err.Number

Case Else
'Funcoes.Menssagem "s", "Titulo", 64

End Select
End Sub

Private Sub Class_Terminate()
DropTabelas
End Sub

Private Sub DropTabelas()
Dim Sql As String, X As Byte
On Error Resume Next
For X = 0 To 1
    Sql = "drop table " & nomeTempTab(X)
    ocnBanco.Execute Sql
Next X
End Sub

Sub Cabecalho(ByRef int_emp As Integer, int_ano As Integer, str_campo As String)
Dim Sql As String, titulo As String

Sql = "create table " & nomeTempTab(0)
Sql = Sql & " ( emp_fant varchar(255), titulo varchar(255));"
ocnBanco.Execute Sql

If str_campo = "des_ccusto" Then
    titulo = "RELATÓRIOS DE DESPESAS POR C.CUSTO DE " & int_ano
ElseIf str_campo = "des_fornec" Then
    titulo = "RELATÓRIOS DE DESPESAS POR FORNECEDOR DE " & int_ano
End If
Sql = "insert into " & nomeTempTab(0) & vbCrLf
Sql = Sql & "values('" & Mostra_emp(2) & "'," & vbCrLf
Sql = Sql & "'" & titulo & "')"
ocnBanco.Execute Sql
End Sub
