Attribute VB_Name = "pf_atestado"
Option Explicit
' cria a tabela temporaria dos dados
Dim snap_Selecao As New ADODB.Recordset
Const Pula_Linha = "<br>"
Dim Original() As String, Html() As String ' arrays para substituir
Dim quant_char As Integer
Public Function Atestado(codcerh As Integer) As String
Dim sql As String, Texto As String
Dim snap_atestado As New ADODB.Recordset
sql = "select pf_atest.* from pf_atest, pf_cerh, pf_tpcer " & Chr(13)
sql = sql & "where cerh_cod=" & codcerh & " and tpcer_cod = cerh_tpcert and tpcer_atestado = ate_cod"
Set snap_atestado = Nothing
snap_atestado.Open sql, ocnBanco

If Not (snap_atestado.BOF And snap_atestado.EOF) Then
    Texto = snap_atestado("ate_texto")
    sql = "select pf_cliente.* from pf_cerh, pf_cliente" & Chr(13)
    sql = sql & "where cerh_cli=cli_cod and cerh_cod=" & codcerh
    Set snap_Selecao = Nothing
    snap_Selecao.Open sql, ocnBanco
    'chama a sub formata e depois a Troca
    'Texto = Formata(Troca(Texto))
    Texto = Troca(Texto)
    'exclui dados da tabel a temporaria
    'sql = "delete * from temp_atest"
    'ocnBanco.Execute sql
    
    'deleta a tabela temp
    delTabTemp
    'cria tabela temporaia
    CriaTabTemp
    
    'gravar campo na tabela temporaria
    sql = "insert into temp_atest values(" & codcerh & ",null,'" & snap_atestado("ate_titulo") & "','" & Texto & "',1)"
    ocnBanco.Execute sql
    'chamar função para abrir o relaótio
Else
    MsgBox "Não existe atestedo para este certificado!", vbInformation, "function atestado"
End If
End Function

Private Function Formata(dados As String) As String
Dim cont As Integer
cont = 1
While cont <= Len(dados)
    If Mid$(dados, cont, 1) = Chr(13) Then
        dados = Mid$(dados, 1, cont - 1) & Pula_Linha & Mid$(dados, cont)
        cont = cont + 5
    Else
        cont = cont + 1
    End If
Wend
Formata = dados
End Function

Private Function Troca(dados As String) As String
Dim inicio As Integer, fim As Integer
Dim campo As String
Dim cont As Integer
inicio = 1
fim = 1
    While fim <= Len(dados)
        If InStr(1, dados, "{{") > 0 Then
            inicio = InStr(1, dados, "{{") + 2
            If InStr(fim, dados, "}}") > 0 Then
                'fim = InStr(fim, dados, "}}") - 1
                fim = InStr(1, dados, "}}") - 1
                campo = LCase$(Mid$(dados, inicio, fim - inicio + 1))
                fim = fim + 2
                
                Select Case LCase$(campo)
                    Case "empresa"
                        campo = snap_Selecao("cli_rzsc")
                    Case "cnpj"
                        campo = snap_Selecao("cli_cnpj")
                    Case "ie"
                        campo = snap_Selecao("cli_ie")
                    Case "endereco"
                        campo = snap_Selecao("cli_ende")
                    Case "bairro"
                        campo = snap_Selecao("cli_bairr")
                    Case "cidade"
                        campo = snap_Selecao("cli_cida")
                    Case "uf"
                        campo = snap_Selecao("cli_uf")
                    Case Else
                        campo = "erro"
                End Select
                
                'coloca o valor na string dados
                dados = Mid$(dados, 1, inicio - 3) & campo & Mid$(dados, fim + 1)
                inicio = fim
            Else
                fim = Len(dados) + 1
            End If
        Else
            'sair do while
            fim = Len(dados) + 1
        End If
    Wend

'Char_Spec ' carrega as variaveis na memória
'Dim swap As Integer 'variavel para dizer onde esta o caractere a ser alterado
'For cont = 0 To quant_char
'    swap = InStr(1, dados, Original(cont))
'    If swap > 0 Then dados = Mid$(dados, 1, swap - 1) & Html(cont) & Mid$(dados, swap)
'Next cont
Troca = dados
End Function

Sub Char_Spec()
quant_char = 10
ReDim Preserve Html(quant_char)
ReDim Preserve Original(quant_char)

Original(0) = "á"
Html(0) = "&aacute;"

Original(1) = "Á"
Html(1) = "&Aacute;"

Original(2) = "ã"
Html(2) = "&atilde;"

Original(3) = "Ã"
Html(3) = "&Atilde;"

Original(4) = "ê"
Html(4) = "&ecirc;"

Original(5) = "Ê"
Html(6) = "&Ecirc;"

Original(7) = "à"
Html(7) = "&agrave;"

Original(8) = "À"
Html(8) = "&Agrave;"

Original(9) = "ç"
Html(9) = "&ccedil;"

Original(10) = "Ç"
Html(10) = "&Ccedil;"

'Original() = ""
'Html() = ""


End Sub
Sub CriaTabTemp()
Dim sql As String
On Error GoTo Erro
sql = "CREATE TABLE temp_atest (" & vbCrLf & _
       "ate_cod              INTEGER ," & vbCrLf & _
       "ate_desc             VARCHAR(50)," & vbCrLf & _
       "ate_titulo           VARCHAR(70)," & vbCrLf & _
       "ate_texto            memo," & vbCrLf & _
       "ate_ident            INTEGER);" & vbCrLf
ocnBanco.Execute sql


'CREATE UNIQUE INDEX PrimaryKey ON temp_atest
'(
'       ate_cod,
'       ate_ident
');
'ALTER TABLE temp_atest
'       ADD PRIMARY KEY (ate_cod, ate_ident);
Exit Sub
Erro:
MsgBox Err.Description, vbCritical, "CriaTabTemp"
End Sub
Sub delTabTemp()
On Error Resume Next
Dim sql As String
sql = "drop table temp_atest"
ocnBanco.Execute sql
End Sub
