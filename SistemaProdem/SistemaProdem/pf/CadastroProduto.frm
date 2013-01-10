VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form CadastroProduto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Produtos"
   ClientHeight    =   5595
   ClientLeft      =   1920
   ClientTop       =   2880
   ClientWidth     =   7710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1.244
   ScaleMode       =   0  'User
   ScaleWidth      =   0.781
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog Common 
      Left            =   4440
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   1095
      Left            =   1800
      TabIndex        =   14
      Top             =   5760
      Width           =   1095
      ExtentX         =   1931
      ExtentY         =   1931
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Frame Frame2 
      Height          =   5085
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   7425
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   1
         Left            =   1350
         MaxLength       =   100
         TabIndex        =   1
         Top             =   600
         Width           =   5895
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   0
         Left            =   1350
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   2
         Left            =   1320
         MaxLength       =   14
         TabIndex        =   2
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   3
         Left            =   1320
         MaxLength       =   14
         TabIndex        =   3
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   4
         Left            =   1320
         MaxLength       =   14
         TabIndex        =   4
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   5
         Left            =   1320
         MaxLength       =   14
         TabIndex        =   5
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   6
         Left            =   1320
         MaxLength       =   14
         TabIndex        =   6
         Top             =   2640
         Width           =   2175
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   7
         Left            =   1320
         MaxLength       =   14
         TabIndex        =   7
         Top             =   3030
         Width           =   2175
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   8
         Left            =   1320
         MaxLength       =   14
         TabIndex        =   8
         Top             =   3390
         Width           =   2175
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   9
         Left            =   1320
         MaxLength       =   14
         TabIndex        =   9
         Top             =   3720
         Width           =   2175
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   10
         Left            =   1320
         MaxLength       =   14
         TabIndex        =   10
         Top             =   4080
         Width           =   2175
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   11
         Left            =   1320
         MaxLength       =   14
         TabIndex        =   11
         Top             =   4440
         Width           =   2175
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Preço 9"
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
         Index           =   15
         Left            =   120
         TabIndex        =   26
         Top             =   4170
         Width           =   645
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Preço 8"
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
         Index           =   13
         Left            =   120
         TabIndex        =   25
         Top             =   3765
         Width           =   645
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Preço 7"
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
         Left            =   120
         TabIndex        =   24
         Top             =   3480
         Width           =   645
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Preço 4"
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
         Left            =   120
         TabIndex        =   23
         Top             =   2370
         Width           =   645
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Preço 6"
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
         Left            =   120
         TabIndex        =   22
         Top             =   3120
         Width           =   645
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Preço 5"
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
         Left            =   120
         TabIndex        =   21
         Top             =   2730
         Width           =   645
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Preço 3"
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
         Left            =   120
         TabIndex        =   20
         Top             =   2010
         Width           =   645
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Preço 2"
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
         Left            =   120
         TabIndex        =   19
         Top             =   1650
         Width           =   645
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Preço 1"
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
         Left            =   120
         TabIndex        =   18
         Top             =   1290
         Width           =   645
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Descricao"
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
         Left            =   150
         TabIndex        =   17
         Top             =   690
         Width           =   840
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Código"
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
         Left            =   150
         TabIndex        =   16
         Top             =   330
         Width           =   600
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Preço 10"
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
         Index           =   16
         Left            =   120
         TabIndex        =   15
         Top             =   4530
         Width           =   750
      End
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   5220
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "CadastroProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Data de criacao: 21/08/2003
'Criador :Rafael
'Ultima atualizacao:10/05/2012         por
Option Explicit
Dim nome_tab(1) As String
Dim cpo(27) As String 'NOME DOS CAMPOS
Dim tcpo(27) As String 'TIPO DOS CAMPOS
Dim Exibir As Boolean ' serve para ver se o formulario vai ficar aberto
Dim Selecao As Boolean 'verifica se exite alguma seleção
Dim snap_Selecao As New ADODB.Recordset ' objeto da seleção
Dim snap_Procura As New ADODB.Recordset ' objeto da seleção

Dim iQtdeCampo As Integer

'Descricao: manutencao


'Private Sub cmd_Click(Index As Integer)
'Dim relatorio As New cls_Relatorio
'Select Case Index
'Case 0 'enviar email...
    'WebBrowser.Navigate "mailto:" & Trim(txt(16))
'Case 1 ' exportar dados
    'relatorio.Banco = Banco_path
    'relatorio.relatorio = App.Path & "\reports\clientes.rpt"
    
'End Select
'End Sub

'Tabelas utilizadas: pf_cliente

Private Sub cmdP_Click(Index As Integer)

End Sub



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
MsgBox "Não exixte dados na Tabela!", vbInformation, "Atenção"
End If
'status.Panels(1).Text = "INSERIR"
'Botoes
Exibir = True
'SSTab.Tab = 0


Exit Sub
erro:
MsgBox Err.Description, vbCritical, "FORM LOAD " & Me.Name
Exibir = False
'Unload Me

End Sub

Sub Declara()
' declara todas as varíaveis de dados
nome_tab(0) = "Produto"
cpo(0) = "Produto_Codigo"
cpo(1) = "Produto_Descricao"
cpo(2) = "Produto_Preco1"
cpo(3) = "Produto_Preco2"
cpo(4) = "Produto_Preco3"
cpo(5) = "Produto_Preco4"
cpo(6) = "Produto_Preco5"
cpo(7) = "Produto_Preco6"
cpo(8) = "Produto_Preco7"
cpo(9) = "Produto_Preco8"
cpo(10) = "Produto_Preco9"
cpo(11) = "Produto_Preco10"

tcpo(0) = "NUMERO"
tcpo(1) = "TEXTO"
tcpo(2) = "NUMERO"
tcpo(3) = "NUMERO"
tcpo(4) = "NUMERO"
tcpo(5) = "NUMERO"
tcpo(6) = "NUMERO"
tcpo(7) = "NUMERO"
tcpo(8) = "NUMERO"
tcpo(9) = "NUMERO"
tcpo(10) = "NUMERO"
tcpo(11) = "NUMERO"


'Saber a quantidade de campo que vai no insert e update
iQtdeCampo = 11

End Sub

Sub Botoes() 'Bloqueia ou nao os botões
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
    
    txt(0).BackColor = &HC0FFFF
    txt(1).SetFocus
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
                    MsgBox "Não existe mais dados nesta direção!", vbInformation, "Atenção"
                    snap_Selecao.MoveLast
                End If
                Mostra_Dados
            Case "PROXIMO"
                snap_Selecao.MoveNext
                If snap_Selecao.EOF Then
                    MsgBox "Não existe mais dados nesta direção!", vbInformation, "Atenção"
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
        MsgBox "Não foi encontrado dados para a seleção!", vbInformation, "Movimentação"
    End If

End Sub

Sub Mostra_Dados()
Dim X As Integer
Dim Sql As String
'For x = 0 To txt.Count - 1
For X = 0 To txt.Count - 1
    If IsNull(snap_Selecao(cpo(X))) Then
        txt(X) = ""
    Else
        txt(X) = snap_Selecao(cpo(X))
    End If
Next X


'Sql = "select ram_desc from pf_cliente, pf_ramativ where cli_ramo = ram_cod and cli_cod=" & snap_Selecao("cli_cod")
'Set snap_Procura = ocnBanco.Execute(Sql)
'snap_Procura.MoveFirst
'lblramo = snap_Procura("ram_desc")


'Exibe_ClienteTipo snap_Selecao(cpo(27))
'If IsNull(snap_Selecao(cpo(27))) Then
'    cboClienteTipo.ListIndex = -1
'Else
'    Exibe_ClienteTipo snap_Selecao(cpo(27))
'End If

Set snap_Procura = Nothing

End Sub

Private Sub Form_Resize()
'Me.Caption = Me.Height & "   " & Me.Width
Me.Height = 6075
Me.Width = 7800
End Sub

Function conssiste() As Boolean
conssiste = False
' no txt 1 verificar cnpj
' no txt 2 veriicar inscricao estadual

'conssitencia especificas do formulario

If Len(txt(1).Text) < 1 Then
    txt(1).SetFocus
    Exit Function
End If

conssiste = True
End Function

Private Sub Form_Unload(Cancel As Integer)
BOTDisable
End Sub

Private Sub txt_GotFocus(Index As Integer)
txt(Index).SelStart = 0
txt(Index).SelLength = Len(txt(Index))

End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
    Case 0
        If KeyAscii = 13 Then
            Menu (3)
            Menu (4)
        End If


End Select

End Sub

Sub Menu(Index As Integer)
Dim Sql As String
Dim snap_Cod As New ADODB.Recordset
On Error GoTo erro

'Inserido  10/05/2012
Dim X As Integer
Select Case Index
Case 0 ' botao inserir
    status.Panels(1).Text = "INSERIR"
    txt(0).Locked = True
    Botoes
    

Case 1 ' botao excluir
    'procura referencia
    'pede confirmação
    If status.Panels(1).Text = "ALTERAR" Then
        If MsgBox("Deseja realmente excluir ?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
            Sql = "delete from " & nome_tab(0) & " where " & cpo(0) & "=" & txt(0)
            ocnBanco.Execute Sql
            Menu (8)
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
                
                Sql = "SELECT MAX(" & cpo(0) & ")+1 FROM " & nome_tab(0)
                Set snap_Cod = ocnBanco.Execute(Sql)
                If IsNull(snap_Cod(0)) Then
                    txt(0) = 1
                Else
                    txt(0) = snap_Cod(0)
                End If
                
                Set snap_Cod = Nothing
                
                
                'forma generica para inserir dados
                

                    Sql = "insert into " & nome_tab(0) & Chr(13)
                    Sql = Sql & "("
                    For X = 0 To iQtdeCampo
                        Sql = Sql & cpo(X)
                        If X < iQtdeCampo Then Sql = Sql & ","
                    Next X
                       
                    Sql = Sql & ") values("
                    For X = 0 To iQtdeCampo
                        Select Case tcpo(X)
                            Case "TEXTO"
                                If Len(txt(X)) > 0 Then
                                    Sql = Sql & "'" & txt(X) & "'"
                                Else
                                    Sql = Sql & "null"
                                End If
                            Case "NUMERO"
                                If Len(txt(X)) > 0 Then
                                    Sql = Sql & "'" & txt(X) & "'"
                                Else
                                    Sql = Sql & "null"
                                End If
                            Case "DATA"
                                If Len(txt(X)) > 0 Then
                                    Sql = Sql & "#" & Format(txt(X), "mm/dd/yyyy") & "#"
                                Else
                                    Sql = Sql & "null"
                                End If
                        End Select
                        
                        
                        If X < iQtdeCampo Then Sql = Sql & ","
                    Next X
                    
                    'Sql = Sql & ", " & cboClienteTipo.ItemData(cboClienteTipo.ListIndex)
                    Sql = Sql & ")" & vbCrLf
                
                ocnBanco.Execute (Sql)
                
                
                'Sql = "insert into " & nome_tab(0) & Chr(13)
                'Sql = Sql & "values(" & txt(0) & ",'" & txt(1) & "','" & txt(2) & "','" & txt(3) & "','" & txt(4) & "','" & txt(5) & "','" & txt(6) & "','" & txt(7) & "','" & txt(8) & "','" & txt(9) & "','" & txt(10) & "','" & txt(11) & "','" & txt(12) & "','" & txt(13) & "','" & txt(14) & "','" & txt(15) & "','" & txt(16) & "','" & txt(17) & "','" & txt(18) & "','" & txt(19) & "','" & txt(20) & "','" & txt(21) & "','" & txt(22) & "','" & cboMidia.ItemData(cboMidia.ListIndex) & "','" & txt(23) & "','" & cboClienteTipo.ItemData(cboClienteTipo.ListIndex) & "')"
                'ocnBanco.Execute (Sql)
                
                status.Panels(1).Text = "ALTERAR"
                MsgBox "Os dados foram gravados com sucesso!", vbInformation, "Gravação"
                Botoes
            End If
        End If
    ElseIf status.Panels(1).Text = "ALTERAR" Then
        If conssiste = True Then
        
            ' atualização genérica
            
            Sql = "UPDATE " & nome_tab(0) & " SET " & Chr(13)
                For X = 1 To iQtdeCampo
                    Sql = Sql & cpo(X) & "="
                        Select Case tcpo(X)
                            Case "TEXTO"
                                If Len(txt(X)) > 0 Then
                                    Sql = Sql & "'" & txt(X) & "'"
                                Else
                                    Sql = Sql & "null"
                                End If
                            Case "NUMERO"
                                If Len(txt(X)) > 0 Then
                                    Sql = Sql & "'" & txt(X) & "'"
                                Else
                                    Sql = Sql & "null"
                                End If
                            Case "DATA"
                                If Len(txt(X)) > 0 Then
                                    Sql = Sql & "#" & Format(txt(X), "mm/dd/yyyy") & "#"
                                Else
                                    Sql = Sql & "null"
                                End If
                        End Select
                        If X < iQtdeCampo Then Sql = Sql & ","
                    Next X
                    
                'Sql = Sql & "," & cpo(27) & "= " & cboClienteTipo.ItemData(cboClienteTipo.ListIndex)
                    
                Sql = Sql & vbCrLf
                Sql = Sql & "WHERE " & cpo(0) & "=" & txt(0)
        
            'Sql = "UPDATE " & nome_tab(0) & " SET " & cpo(1) & "='" & txt(1) & "'," & cpo(2) & "='" & txt(2) & "'," & cpo(3) & "='" & txt(3) & "'," & cpo(4) & "='" & txt(4) & "'," & cpo(5) & "='" & txt(5) & "'," & cpo(6) & "='" & txt(6) & "'," & cpo(7) & "='" & txt(7) & "'," & cpo(8) & "='" & txt(8) & "'," & cpo(9) & "='" & txt(9) & "'," & cpo(10) & "='" & txt(10) & "'," & cpo(11) & "='" & txt(11) & "'," & cpo(12) & "='" & txt(12) & "'," & cpo(13) & "='" & txt(13) & "'," & cpo(14) & "='" & txt(14) & "'," & cpo(15) & "='" & txt(15) & "'," & cpo(16) & "='" & txt(16) & "'," & cpo(17) & "='" & txt(17) & "'," & cpo(18) & "='" & txt(18) & "'," & cpo(19) & "='" & txt(19) & "'," & cpo(20) & "='" & txt(20) & "'," & cpo(21) & "='" & txt(21) & "'," & cpo(22) & "='" & txt(22) & "'," & cpo(25) & "=" & cboMidia.ItemData(cboMidia.ListIndex) & ", " & cpo(26) & "='" & txt(23) & "'," & cpo(27) & "=" & cboClienteTipo.ItemData(cboClienteTipo.ListIndex) & Chr(13)
            'Sql = Sql & "WHERE " & cpo(0) & "=" & txt(0)
            ocnBanco.Execute Sql
            MsgBox "Os dados foram atualizados com sucesso !", vbInformation, "Alteração"
        End If
    End If
Case 3 ' botao selecionar
    'coloca em modo seleção
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


