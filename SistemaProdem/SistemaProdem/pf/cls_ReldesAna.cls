VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_ReldesAna"
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
Dim sql As String
If IsNumeric(cod) Then
    sql = "select emp_rzsc from im_empresa " & vbCrLf
    sql = sql & "where emp_cod=" & cod
    Set snap_Selecao = ocnBanco.Execute(sql)
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
nomeTempTab(0) = "temp_relAnaH" 'header dados do cabeçalho
nomeTempTab(1) = "temp_relAnaI" 'item
End Sub

Sub Relatorio(str_cond As String, campo As String)
Dim sql As String

On Error GoTo erro
DropTabelas
If campo = "des_ccusto" Then
    'Select para relatorio sintetico de despesas por Centro de Custo
    sql = "SELECT des_emp, des_ccusto as des_cod, ccus_desc as des_desc, des_comp ," & Chr(13) & _
                "des_tpdoc, des_nrdoc, des_prod, des_dtemiss," & Chr(13) & _
                "des_fornec as des_cod2, for_rzsc as des_desc2," & Chr(13) 'dados do fornecedor
    sql = sql & "des_vr into " & Chr(13) & _
                nomeTempTab(1) & Chr(13) & _
                "From pf_lancDes, pf_ccusto, pf_fornec" & Chr(13) & _
                "Where " & Chr(13) & _
                " des_ccusto= ccus_cod and for_cod= des_fornec" & Chr(13) & _
                str_cond & Chr(13) & _
                ""
ElseIf campo = "des_fornec" Then
    'Select para relatorio sintetico de despesas por Fornecedor
    sql = "SELECT des_emp, des_fornec as des_cod2, for_rzsc as des_desc2,  des_comp ," & Chr(13) & _
                "des_tpdoc, des_nrdoc, des_prod, des_dtemiss," & Chr(13) & _
                "des_ccusto as des_cod, ccus_desc as des_desc," & Chr(13) 'dados do centro de custo
    sql = sql & "des_vr into " & Chr(13) & _
                nomeTempTab(1) & Chr(13) & _
                "From pf_lancDes, pf_ccusto, pf_fornec" & Chr(13) & _
                "Where " & Chr(13) & _
                " des_ccusto= ccus_cod and for_cod= des_fornec" & Chr(13) & _
                str_cond & Chr(13) & _
                ""
End If
ocnBanco.Execute sql
Exit Sub
erro:
Select Case Err.Number

Case Else
MsgBox Err.Description
'Funcoes.Menssagem "s", "Titulo", 64

End Select
End Sub

Private Sub Class_Terminate()
DropTabelas
End Sub

Private Sub DropTabelas()
Dim sql As String, X As Byte
On Error Resume Next
For X = 0 To 1
    sql = "drop table " & nomeTempTab(X)
    ocnBanco.Execute sql
Next X
End Sub

Sub Cabecalho(int_emp As Integer, competencia As Date, str_campo As String)
Dim sql As String, titulo As String

sql = "create table " & nomeTempTab(0)
sql = sql & " ( emp_fant varchar(50), titulo varchar(70));"
ocnBanco.Execute sql

If str_campo = "des_ccusto" Then
'FIXIT: Replace 'UCase' function with 'UCase$' function                                    FixIT90210ae-R9757-R1B8ZE
    titulo = "RELATÓRIOS DE DESPESAS ANALÍTICO POR C.CUSTO DE " & UCase(Format(competencia, "mmmm/yyyy"))
ElseIf str_campo = "des_fornec" Then
'FIXIT: Replace 'UCase' function with 'UCase$' function                                    FixIT90210ae-R9757-R1B8ZE
    titulo = "RELATÓRIOS DE DESPESAS ANALÍTICO POR FORNECEDOR DE " & UCase(Format(competencia, "mmmm/yyyy"))
End If
sql = "insert into " & nomeTempTab(0) & vbCrLf
'alterar esta linha
sql = sql & "values('" & Mostra_emp(1) & "'," & vbCrLf
sql = sql & "'" & titulo & "')"
ocnBanco.Execute sql
End Sub
