VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form pf_cadVendedor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Vendedores"
   ClientHeight    =   4845
   ClientLeft      =   105
   ClientTop       =   6060
   ClientWidth     =   8520
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1.077
   ScaleMode       =   0  'User
   ScaleWidth      =   0.862
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   3945
      Left            =   150
      TabIndex        =   14
      Top             =   270
      Width           =   7215
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   12
         Left            =   1110
         TabIndex        =   12
         Top             =   3510
         Width           =   4515
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   11
         Left            =   3810
         TabIndex        =   11
         Top             =   3060
         Width           =   1755
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   10
         Left            =   4350
         TabIndex        =   4
         Top             =   1230
         Width           =   1305
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   9
         Left            =   2640
         TabIndex        =   3
         Top             =   1230
         Width           =   735
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   8
         Left            =   1110
         TabIndex        =   10
         Top             =   3060
         Width           =   1665
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   7
         Left            =   3810
         TabIndex        =   9
         Top             =   2550
         Width           =   525
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   6
         Left            =   1110
         TabIndex        =   8
         Top             =   2550
         Width           =   1665
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   5
         Left            =   3810
         TabIndex        =   7
         Top             =   2070
         Width           =   1815
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   4
         Left            =   1110
         TabIndex        =   6
         Top             =   2070
         Width           =   1815
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   3
         Left            =   1110
         TabIndex        =   5
         Top             =   1590
         Width           =   4545
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   2
         Left            =   1110
         TabIndex        =   2
         Top             =   1230
         Width           =   615
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   1
         Left            =   1110
         TabIndex        =   1
         Top             =   840
         Width           =   4545
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   0
         Left            =   1110
         TabIndex        =   0
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "E-Mail"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   240
         TabIndex        =   27
         Top             =   3570
         Width           =   510
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Celular"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   3120
         TabIndex        =   26
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Conta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   3720
         TabIndex        =   25
         Top             =   1260
         Width           =   510
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Ag�ncia"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   1860
         TabIndex        =   24
         Top             =   1290
         Width           =   675
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Telefone"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   210
         TabIndex        =   23
         Top             =   3060
         Width           =   735
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "UF"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   3150
         TabIndex        =   22
         Top             =   2550
         Width           =   210
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Cidade"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   210
         TabIndex        =   21
         Top             =   2550
         Width           =   600
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "CEP"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   3150
         TabIndex        =   20
         Top             =   2100
         Width           =   345
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Bairo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   210
         TabIndex        =   19
         Top             =   2070
         Width           =   450
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Endere�o"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   18
         Top             =   1590
         Width           =   795
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Banco"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   17
         Top             =   1230
         Width           =   525
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   16
         Top             =   840
         Width           =   495
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "C�digo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   15
         Top             =   360
         Width           =   600
      End
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   4470
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "pf_cadVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Data de criacao:21/08/2003
'Criador :Rafael
'Ultima atualizacao:         por
Option Explicit
Dim nome_tab(0) As String
Dim cpo(12) As String 'NOME DOS CAMPOS
Dim tcpo(12) As String 'TIPO DOS CAMPOS
Dim Exibir As Boolean ' serve para ver se o formulario vai ficar aberto
Dim Selecao As Boolean 'verifica se exite alguma sele��o
Dim snap_Selecao As New ADODB.Recordset ' objeto da sele��o
'Descricao: manutencao

'Tabelas utilizadas: pf_ramativ



Private Sub Form_Activate()
Botoes
Me.Top = 0
Me.Left = 0
If Exibir = False Then Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Menu (8) ' esc
End Sub

Private Sub Form_Load()
Dim snap_tab As New ADODB.Recordset
Dim Sql As String
Declara

On Error GoTo erro

'pesquisa dados na tabela
Sql = "select count(*) from " & nome_tab(0)
Set snap_tab = ocnBanco.Execute(Sql)
If snap_tab(0) <= 0 Then
MsgBox "N�o exixte dados na Tabela!", vbInformation, "Aten��o"
End If
'status.Panels(1).Text = "INSERIR"
'Botoes
Exibir = True
Exit Sub
erro:
MsgBox Err.Description, vbCritical, "FORM LOAD " & Me.Name
Exibir = False
'Unload Me

End Sub

Sub Declara()

' declara todas as var�aveis de dados
nome_tab(0) = "pf_vendedor"
cpo(0) = "vend_cod"
cpo(1) = "vend_nome"
cpo(2) = "vend_banco"
cpo(3) = "vend_ende"
cpo(4) = "vend_bair"
cpo(5) = "vend_cep"
cpo(6) = "vend_cida"
cpo(7) = "vend_uf"
cpo(8) = "vend_tele"
cpo(9) = "vend_agencia"
cpo(10) = "vend_conta"
cpo(11) = "vend_celular"
cpo(12) = "vend_email"
tcpo(0) = "NUMERO"
tcpo(1) = "TEXTO"
tcpo(2) = "TEXTO"
tcpo(3) = "TEXTO"
tcpo(4) = "TEXTO"
tcpo(5) = "TEXTO"
tcpo(6) = "TEXTO"
tcpo(7) = "TEXTO"
tcpo(8) = "TEXTO"
tcpo(9) = "TEXTO"
tcpo(10) = "TEXTO"
tcpo(11) = "TEXTO"
tcpo(12) = "TEXTO"
End Sub

Sub Botoes() 'Bloqueia ou nao os bot�es
mdi.Toolbar.Buttons(5).Enabled = True
Select Case status.Panels(1).Text
Case "INSERIR"
    mdi.Toolbar.Buttons(1).Enabled = False
    mdi.Toolbar.Buttons(3).Enabled = False
    mdi.Toolbar.Buttons(2).Enabled = True
    mdi.Toolbar.Buttons(4).Enabled = False
    mdi.Toolbar.Buttons(6).Enabled = False
    mdi.Toolbar.Buttons(7).Enabled = False
    mdi.Toolbar.Buttons(8).Enabled = False
    mdi.Toolbar.Buttons(9).Enabled = False
    mdi.Toolbar.Buttons(10).Enabled = False

    
    txt(1).SetFocus
    txt(0).BackColor = &HC0FFFF
Case "SELECIONAR"
    mdi.Toolbar.Buttons(1).Enabled = False
    mdi.Toolbar.Buttons(3).Enabled = False
    mdi.Toolbar.Buttons(2).Enabled = False
    mdi.Toolbar.Buttons(4).Enabled = True
    mdi.Toolbar.Buttons(6).Enabled = True
    mdi.Toolbar.Buttons(7).Enabled = True
    mdi.Toolbar.Buttons(8).Enabled = True
    mdi.Toolbar.Buttons(9).Enabled = True
    mdi.Toolbar.Buttons(10).Enabled = True
    
    txt(0).Locked = False
    txt(0).SetFocus
Case ""
    mdi.Toolbar.Buttons(1).Enabled = True
    mdi.Toolbar.Buttons(3).Enabled = False
    mdi.Toolbar.Buttons(2).Enabled = False
    mdi.Toolbar.Buttons(4).Enabled = True
    mdi.Toolbar.Buttons(6).Enabled = False
    mdi.Toolbar.Buttons(7).Enabled = False
    mdi.Toolbar.Buttons(8).Enabled = False
    mdi.Toolbar.Buttons(9).Enabled = False
    mdi.Toolbar.Buttons(10).Enabled = False
    
Case "ALTERAR"
    mdi.Toolbar.Buttons(1).Enabled = False
    mdi.Toolbar.Buttons(3).Enabled = True
    mdi.Toolbar.Buttons(2).Enabled = True
    mdi.Toolbar.Buttons(4).Enabled = False
    mdi.Toolbar.Buttons(6).Enabled = True
    mdi.Toolbar.Buttons(7).Enabled = True
    mdi.Toolbar.Buttons(8).Enabled = True
    mdi.Toolbar.Buttons(9).Enabled = True
    mdi.Toolbar.Buttons(10).Enabled = True
    
    txt(0).Locked = True

End Select
End Sub
Sub Limpar()
Dim X As Integer
For X = 0 To txt.Count - 1
    txt(X).Text = ""
Next X
txt(0).BackColor = &H80000005
End Sub

Function conssiste_total() As Boolean
Dim X As Integer
On Error GoTo erro
For X = 0 To txt.Count - 1
Select Case tcpo(X)
Case "NUMERO"
    If Len(Trim$(txt(X))) > 0 Then
        If IsNumeric(txt(X)) = False Then
            conssiste_total = False
            txt(X).SetFocus
            Exit Function
        End If
    End If
'Case "TEXTO"
    
Case "DATA"
    If Len(txt(X)) > 0 Then
        If IsDate(txt(X)) = False Then
            conssiste_total = False
            Exit Function
        End If
    End If
End Select
Next X
conssiste_total = True
Exit Function
erro:
MsgBox Err.Description, vbCritical, "CONSSISTE_TOTAL"

End Function

Sub Movimentacao(botao As String)
Dim Sql As String, sql_where As String
Dim X As Integer
If Selecao = True Then
    If conssiste_total = True Then ' procurar os dados
        Sql = "select * from " & nome_tab(0)
        For X = 0 To txt.Count - 1
            If Len(txt(X)) > 0 Then
                If Len(sql_where) <= 0 Then
                    sql_where = " where "
                Else
                    sql_where = sql_where & " and " & Chr(13)
                End If
                sql_where = sql_where & Monta_SQL(cpo(X), tcpo(X), txt(X))
            End If
        Next X
    Sql = Sql & sql_where
    snap_Selecao.Open Sql, ocnBanco, adOpenKeyset, adLockOptimistic, adCmdText
    Selecao = True
    Else
        Exit Sub
    End If
End If

    Selecao = False
    If Not (snap_Selecao.EOF And snap_Selecao.BOF) Then
        Select Case UCase$(botao)
            Case "PRIMEIRO"
                snap_Selecao.MoveFirst
                'snap_Selecao.Requery
                Mostra_Dados
            Case "ANTERIOR"
                'snap_Selecao.Requery
                snap_Selecao.MovePrevious
                If snap_Selecao.BOF Then
                    MsgBox "N�o existe mais dados nesta dire��o!", vbInformation, "Aten��o"
                    snap_Selecao.MoveLast
                End If
                Mostra_Dados
            Case "PROXIMO"
                snap_Selecao.MoveNext
                If snap_Selecao.EOF Then
                    MsgBox "N�o existe mais dados nesta dire��o!", vbInformation, "Aten��o"
                    snap_Selecao.MoveFirst
                End If
                Mostra_Dados
            Case "ULTIMO"
                snap_Selecao.MoveLast
                Mostra_Dados
        End Select
        status.Panels(1).Text = "ALTERAR"
        Botoes
    Else
        MsgBox "N�o foi encontrado dados para a sele��o!", vbInformation, "Movimenta��o"
    End If

End Sub

Sub Mostra_Dados()
Dim X As Integer
For X = 0 To txt.Count - 1
    If IsNull(snap_Selecao(cpo(X))) = False Then txt(X) = snap_Selecao(cpo(X))
Next X
End Sub

Private Sub Form_Resize()
'Me.Caption = Me.Height & "   " & Me.Width
Me.Height = 5760
Me.Width = 7620
End Sub

Function conssiste() As Boolean
conssiste = False
If Len(txt(1)) > 0 Then
    conssiste = True
Else
    conssiste = False
    txt(1).SetFocus
    Exit Function
End If
If InStr(1, txt(1), "*") > 0 Then
    conssiste = False
    txt(1).SetFocus
End If

End Function

Private Sub Form_Unload(Cancel As Integer)
BOTDisable
End Sub

Private Sub txt_GotFocus(Index As Integer)
txt(Index).SelStart = 0
txt(Index).SelLength = Len(txt(Index))
End Sub
Sub Menu(Index As Integer)
Dim Sql As String
Dim snap_Cod As New ADODB.Recordset
On Error GoTo erro

Select Case Index
Case 0 ' botao inserir
    status.Panels(1).Text = "INSERIR"
    txt(0).Locked = True
    Botoes

Case 1 ' botao excluir
    'procura referencia
    'pede confirma��o
    If status.Panels(1).Text = "ALTERAR" Then
        If MsgBox("Deseja realmente excluir ?", vbQuestion + vbYesNo, "Aten��o") = vbYes Then
            Sql = "delete from " & nome_tab(0) & " where " & cpo(0) & "=" & txt(0)
            ocnBanco.Execute Sql
            Limpar
            Botoes
            'Mostra_Dados
        End If
    End If
        'apaga o registro atual
Case 2 ' botao gravar
    'atualiza os dados
    '
    If status.Panels(1).Text = "INSERIR" Then
        If conssiste_total = True Then 'verifica a consistencia
            If conssiste = True Then
            
                Sql = "SELECT MAX(vend_COD)+1 FROM " & nome_tab(0)
                Set snap_Cod = ocnBanco.Execute(Sql)
                If IsNull(snap_Cod(0)) Then
                    txt(0) = 1
                Else
                    txt(0) = snap_Cod(0)
                End If
                Set snap_Cod = Nothing
            
                Sql = "insert into " & nome_tab(0) & Chr(13)
                Sql = Sql & "values(" & txt(0) & ",'" & txt(1) & "','" & txt(3) & "','" & txt(4) & "','" & txt(5) & "','" & txt(6) & "','" & txt(7) & "','" & txt(8) & "','" & txt(11) & "','" & txt(12) & "','" & txt(2) & "','" & txt(9) & "','" & txt(10) & "')"
                ocnBanco.Execute (Sql)
                status.Panels(1).Text = "ALTERAR"
                MsgBox "Os dados foram gravados com sucesso!", vbInformation, "Grava��o"
                Botoes
            End If
        End If
    ElseIf status.Panels(1).Text = "ALTERAR" Then
        If conssiste = True Then
            Sql = "UPDATE " & nome_tab(0) & " SET " & Chr(13)
            
            Dim X As Integer
            For X = 1 To 11
                Sql = Sql & cpo(X) & "='" & txt(X) & "' ,"
            Next X
            Sql = Sql & cpo(12) & "='" & txt(12) & "'" & Chr(13)
            
            Sql = Sql & "WHERE " & cpo(0) & "=" & txt(0)
            ocnBanco.Execute Sql
            MsgBox "Os dados foram atualizados com sucesso !", vbInformation, "Altera��o"
        End If
    End If
Case 3 ' botao selecionar
    'coloca em modo sele��o
    status.Panels(1).Text = "SELECIONAR"
    Selecao = True
    Set snap_Selecao = Nothing
    Botoes
Case 4 ' botao primeiro
    Movimentacao ("PRIMEIRO")
Case 5 ' botao anterior
    Movimentacao ("ANTERIOR")
Case 6 ' botao proximo
    Movimentacao ("PROXIMO")
Case 7 ' botao ultimo
    Movimentacao ("ULTIMO")
Case 8 ' botao cancelar
    Limpar
    If status.Panels(1).Text = "" Then
        Unload Me
        Exit Sub
    End If
    status.Panels(1).Text = ""
    Botoes
    Set snap_Selecao = Nothing ' fecha objeto de selecao
    txt(0).Locked = False
    txt(0).SetFocus
End Select
Exit Sub
erro:
MsgBox Err.Description, vbCritical, "cmd_click"

End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)


Select Case Index
    Case 0
        If KeyAscii = 13 Then
            Menu (3)
            Menu (4)
        End If
    
    
    Case 5
        If KeyAscii = 13 Then
            If Busca_Cep(txt(Index)) = True Then
                txt(3) = CEP_Nome
                txt(4) = CEP_Bairro
                txt(6) = CEP_Cidade
                txt(7) = "SP"
                txt(3).SetFocus
            End If
        End If
End Select
End Sub
