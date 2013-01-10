VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form pf_lancReceb 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contas à receber"
   ClientHeight    =   5430
   ClientLeft      =   1980
   ClientTop       =   2220
   ClientWidth     =   6855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1.206
   ScaleMode       =   0  'User
   ScaleWidth      =   0.693
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   4725
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6675
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   12
         Left            =   1200
         TabIndex        =   32
         Top             =   4320
         Width           =   1000
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   11
         Left            =   4320
         TabIndex        =   26
         Top             =   3960
         Width           =   1000
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   10
         Left            =   1200
         TabIndex        =   25
         Top             =   3960
         Width           =   1000
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   9
         Left            =   4320
         TabIndex        =   24
         Top             =   3480
         Width           =   1000
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   8
         Left            =   1200
         TabIndex        =   23
         Top             =   3480
         Width           =   1000
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   7
         Left            =   4320
         TabIndex        =   22
         Top             =   3000
         Width           =   1000
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   6
         Left            =   1200
         TabIndex        =   21
         Top             =   3000
         Width           =   1000
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   5
         Left            =   4320
         TabIndex        =   20
         Top             =   2520
         Width           =   1000
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   4
         Left            =   1800
         TabIndex        =   11
         Top             =   2520
         Width           =   1000
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   3
         Left            =   4320
         TabIndex        =   10
         Top             =   2160
         Width           =   1000
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   2
         Left            =   1800
         TabIndex        =   9
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   1
         Left            =   1650
         TabIndex        =   1
         Top             =   600
         Width           =   1000
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   0
         Left            =   1650
         TabIndex        =   0
         Top             =   240
         Width           =   1000
      End
      Begin VB.Label label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Realizado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   17
         Left            =   3000
         TabIndex        =   35
         Top             =   4320
         Width           =   1185
      End
      Begin VB.Label label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Carregar..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   16
         Left            =   2880
         TabIndex        =   34
         Top             =   600
         Width           =   1185
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Fatura"
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
         Left            =   240
         TabIndex        =   33
         Top             =   4320
         Width           =   540
      End
      Begin VB.Label lbl 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   2
         Left            =   1680
         TabIndex        =   31
         Top             =   1680
         Width           =   3615
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
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
         Index           =   14
         Left            =   240
         TabIndex        =   30
         Top             =   1800
         Width           =   750
      End
      Begin VB.Label lbl 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   1650
         TabIndex        =   29
         Top             =   1320
         Width           =   3615
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Prazo Pgto."
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
         Left            =   240
         TabIndex        =   28
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lbl 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   0
         Left            =   1650
         TabIndex        =   27
         Top             =   960
         Width           =   3615
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor"
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
         TabIndex        =   19
         Top             =   3600
         Width           =   825
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Descontos"
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
         Left            =   3000
         TabIndex        =   18
         Top             =   3960
         Width           =   885
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Nr. NF"
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
         Left            =   240
         TabIndex        =   17
         Top             =   3960
         Width           =   525
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Comissão"
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
         Left            =   3000
         TabIndex        =   16
         Top             =   3600
         Width           =   840
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Dt. Pagamento"
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
         Left            =   3000
         TabIndex        =   15
         Top             =   2640
         Width           =   1275
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Dt. Treinamento"
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
         Left            =   240
         TabIndex        =   14
         Top             =   2160
         Width           =   1395
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
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
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Lançamento"
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
         Left            =   240
         TabIndex        =   12
         Top             =   270
         Width           =   1035
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Valor Pago"
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
         Left            =   3000
         TabIndex        =   8
         Top             =   3120
         Width           =   930
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
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
         Left            =   240
         TabIndex        =   7
         Top             =   3000
         Width           =   450
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Dt. Vencimento"
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
         Left            =   240
         TabIndex        =   6
         Top             =   2550
         Width           =   1320
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Dt. Emissão"
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
         Left            =   3000
         TabIndex        =   5
         Top             =   2160
         Width           =   1020
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Nr. Orçamento"
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
         Left            =   240
         TabIndex        =   4
         Top             =   660
         Width           =   1260
      End
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   5055
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "pf_lancReceb"
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
Dim Selecao As Boolean 'verifica se exite alguma seleção
Dim snap_Selecao As New ADODB.Recordset ' objeto da seleção
Dim snap_Procura As New ADODB.Recordset ' objeto do mostra dados
'Descricao: manutencao


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
Exit Sub
erro:
MsgBox Err.Description, vbCritical, "FORM LOAD " & Me.Name
Exibir = False
'Unload Me

End Sub

Sub Declara()

' declara todas as varíaveis de dados
nome_tab(0) = "pf_lancreceb"
cpo(0) = "rec_cod"
cpo(1) = "rec_orcamento"
cpo(2) = "rec_dtrein"
cpo(3) = "rec_dtemiss"
cpo(4) = "rec_dtvenc"
cpo(5) = "rec_dtPGTO"
cpo(6) = "rec_vr"
cpo(7) = "rec_pago"
cpo(8) = "rec_vendedor"
cpo(9) = "rec_comis"
cpo(10) = "rec_nrnf"
cpo(11) = "rec_descontos"
cpo(12) = "rec_fatura"

tcpo(0) = "NUMERO"
tcpo(1) = "NUMERO"
tcpo(2) = "DATA"
tcpo(3) = "DATA"
tcpo(4) = "DATA"
tcpo(5) = "DATA"
tcpo(6) = "NUMERO"
tcpo(7) = "NUMERO"
tcpo(8) = "TEXTO"
tcpo(9) = "NUMERO"
tcpo(10) = "TEXTO"
tcpo(11) = "NUMERO"
tcpo(12) = "TEXTO"
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

    
    txt(0).SetFocus
    txt(0).BackColor = &HC0FFFF
    txt(0).Locked = True
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
    mdi.Toolbar.Buttons(2).Enabled = True 'alterado nao
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
On Error Resume Next
For X = 0 To txt.Count - 1
    txt(X).Text = ""
Next X
LBL(0) = ""
LBL(1) = ""
'txt(0).BackColor = &H80000005
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
    If Len(Trim$(txt(X).Text)) > 0 Then
        If IsDate(txt(X)) = False Then
            conssiste_total = False
            txt(X).SetFocus
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
        If conssiste(1) Then ' verifica a competencia
            Sql = "select * from " & nome_tab(0) & vbCrLf
    
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
            Sql = Sql & " order by " & cpo(0)
            snap_Selecao.Open Sql, ocnBanco, adOpenKeyset, adLockOptimistic, adCmdText
            Selecao = True
        Else
            Exit Sub
        End If
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
Limpar
Dim Sql As String
Dim X As Integer
For X = 0 To txt.Count - 1
    If IsNull(snap_Selecao(cpo(X))) = False Then txt(X) = snap_Selecao(cpo(X))
Next X
Exibe_orc txt(1)


Sql = "select emp_rzsc, ccus_desc, for_rzsc from im_empresa, pf_ccusto, pf_fornec, pf_lancdes" & vbCrLf
Sql = Sql & "where des_emp = emp_cod and des_ccusto = ccus_cod and des_fornec = for_cod"
Set snap_Procura = ocnBanco.Execute(Sql)
snap_Procura.MoveFirst
Set snap_Procura = Nothing

End Sub

Private Sub Form_Resize()
'Me.Caption = Me.Height & "   " & Me.Width
'Me.Height = 3760
'Me.Width = 7000
End Sub

Function conssiste(Index As Integer) As Boolean
Dim Sql As String
Select Case Index
Dim X As Integer
Case 0 'default
If Len(LBL(0)) > 0 Then
    Sql = "select orch_situacao from pf_orcamentoh where orch_cod=" & txt(1)
    Set snap_Procura = ocnBanco.Execute(Sql)
    If Not (snap_Procura.BOF And snap_Procura.EOF) Then
        If snap_Procura(0) = 3 Or snap_Procura(0) = 4 Then
            conssiste = True
        Else
            MsgBox "O orçamento não foi realizado!", vbInformation, "conssiste"
        End If
    Else
        MsgBox "Orçamento não encontrado!", vbInformation, "conssiste"
        conssiste = False
    End If
Else
    txt(1).SetFocus
    conssiste = False
End If

Case 1 ' DADOS PARA SELECAO
    conssiste = True

Case 2 ' DADOS PARA SALVAR
    For X = 2 To 12
        If Len(txt(X)) = 0 Then
            If X <> 5 And X <> 10 And X <> 12 Then
                txt(X).SetFocus
                conssiste = False
                Exit Function
            End If
        End If
    Next X
    conssiste = True
End Select
End Function

Private Sub Form_Unload(Cancel As Integer)
BOTDisable
End Sub



Private Sub label_Click(Index As Integer)
Select Case Index
    Case 16
        If IsNumeric(txt(1)) Then Carrega_Dados (txt(1))
    Case 17
    If MsgBox("Deseja alterar o orçamento?", vbYesNo + vbQuestion) = vbYes Then
        If IsDate(txt(1)) Then
            Confirma_Orcamento (txt(1))
        End If
    End If
End Select

End Sub

Private Sub label_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Select Case Index
Case 16, 17
    label(Index).BackStyle = 1
End Select

End Sub

Private Sub txt_GotFocus(Index As Integer)

txt(Index).SelStart = 0
txt(Index).SelLength = Len(txt(Index))

End Sub
Sub Menu(Index As Integer)
Dim Sql As String
Dim X As Integer
Dim snap_Cod As New ADODB.Recordset
On Error GoTo erro

Select Case Index
Case 0 ' botao inserir
    status.Panels(1).Text = "INSERIR"
    'txt(0).Locked = True
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
            If conssiste(2) = True Then ' VERIFICA DADOS PARA NAO SEREM NULOS
                If conssiste(0) = True Then
                    Set snap_Selecao = ocnBanco.Execute("select max(" & cpo(0) & ") +1 from " & nome_tab(0))
                    If IsNull(snap_Selecao(0)) = True Then
                        txt(0) = 1
                    Else
                        txt(0) = snap_Selecao(0)
                    End If
                    Set snap_Selecao = Nothing
            
            'forma generica para inserir dados
                    Sql = "insert into " & nome_tab(0) & Chr(13)
                    Sql = Sql & "("
                    For X = 0 To 12
                        Sql = Sql & cpo(X)
                        If X < 12 Then Sql = Sql & ","
                    Next X
                                    
                    Sql = Sql & ") values("
                    For X = 0 To 12
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
                        If X < 12 Then Sql = Sql & ","
                    Next X
                    
                    Sql = Sql & ")" & vbCrLf
                                                    
                    ocnBanco.Execute (Sql)
            
            
                    status.Panels(1).Text = "ALTERAR"
                    MsgBox "Os dados foram gravados com sucesso!", vbInformation, "Gravação"
                    Botoes
                End If
            End If
        End If
    ElseIf status.Panels(1).Text = "ALTERAR" Then
        If conssiste(0) = True Then
            If conssiste(2) = True Then
                Sql = "UPDATE " & nome_tab(0) & " SET " & Chr(13)
                ' atualização genérica
                For X = 1 To 12
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
                        If X < 12 Then Sql = Sql & ","
                    Next X
                    
                Sql = Sql & vbCrLf
                Sql = Sql & "WHERE " & cpo(0) & "=" & txt(0)
                ocnBanco.Execute Sql
                
                
                MsgBox "Os dados foram atualizados com sucesso !", vbInformation, "Alteração"
            End If
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
End Select
Exit Sub
erro:
MsgBox Err.Description, vbCritical, "cmd_click"

End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
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

Private Sub txt_LostFocus(Index As Integer)
Dim Sql As String, snap_Procura As New ADODB.Recordset
Select Case Index
    Case 0 ' empresa
'        If IsNumeric(txt(Index)) Then
            'sql = "Select emp_rzsc from im_empresa where emp_cod = " & txt(Index)
            'Set snap_Procura = ocnBanco.Execute(sql)
            'If snap_Procura.BOF And snap_Procura.EOF Then
                'Lbl(0).Caption = ""
            'Else
                'snap_Procura.MoveFirst
                'Lbl(0) = snap_Procura(0)
            'End If
            'Set snap_Procura = Nothing
        'Else
            'Lbl(0) = ""
        'End If
    
    Case 1 ' nr orcamneto
        If IsNumeric(txt(1)) Then
            Exibe_orc (txt(1))
        Else
            LBL(0) = ""
            LBL(1) = ""
            LBL(2) = ""
        End If
End Select
End Sub

Sub Exibe_orc(cod As Integer)
On Error GoTo erro
Dim Sql As String

Sql = "SELECT pf_cliente.cli_rzsc, pf_PrazoPgto.pgto_desc, emp_rzsc FROM pf_PrazoPgto, pf_Cliente, pf_orcamentoH, im_empresa "
Sql = Sql & "where orch_cli = cli_cod and OrcH_Prazopgto = pgto_cod and orch_emp = emp_cod "
Sql = Sql & "and orch_cod =" & cod
Set snap_Procura = ocnBanco.Execute(Sql)
If Not (snap_Procura.BOF And snap_Procura.EOF) Then
    'existe dados
    LBL(0) = snap_Procura("cli_rzsc")
    LBL(1) = snap_Procura("pgto_desc")
    LBL(2) = snap_Procura("emp_rzsc")
Else
    'dados em branco
    LBL(0) = ""
    LBL(1) = ""
    LBL(2) = ""
End If
Exit Sub
erro:
MsgBox Err.Description, vbCritical, "exibe_orc"
End Sub

Sub Carrega_Dados(orc As Integer)
On Error GoTo erro
Dim Sql As String
Sql = "SELECT Sum(pf_OrcamentoI.OrcI_ServQtdeP*pf_OrcamentoI.OrcI_ServVrVndP) AS totalbruto, Sum(((pf_OrcamentoI.OrcI_ServQtdeP*pf_OrcamentoI.OrcI_ServVrVndP)-(pf_OrcamentoI.OrcI_ServQtdeP*pf_OrcamentoI.OrcI_ServVrVndP*pf_orcamentoH.OrcH_Desconto/100))*orch_turma) AS liquido, pf_vendedor.vend_nome, pf_PrazoPgto.pgto_n1, max(pf_dtTrei.Orcdt_data) as data "
Sql = Sql & "FROM (pf_vendedor INNER JOIN (pf_PrazoPgto INNER JOIN (pf_orcamentoH INNER JOIN pf_OrcamentoI ON pf_orcamentoH.OrcH_cod = pf_OrcamentoI.OrcI_Orc) ON pf_PrazoPgto.pgto_cod = pf_orcamentoH.OrcH_Prazopgto) ON pf_vendedor.vend_cod = pf_orcamentoH.OrcH_vende) INNER JOIN pf_dtTrei ON pf_orcamentoH.OrcH_cod = pf_dtTrei.Orcdt_nrorc "
Sql = Sql & "Where (((pf_orcamentoH.OrcH_cod) = " & orc & ")) "
Sql = Sql & "GROUP BY pf_vendedor.vend_nome, pf_PrazoPgto.pgto_n1; "

Set snap_Procura = ocnBanco.Execute(Sql)
If Not (snap_Procura.BOF And snap_Procura.EOF) Then
    'existe dados
    txt(2) = snap_Procura("DATA")
    txt(4) = DateAdd("d", snap_Procura("pgto_n1"), snap_Procura("DATA"))
    txt(6) = snap_Procura("liquido")
    txt(8) = snap_Procura("vend_nome")
    'txt(12) = snap_Procura("pgto_n1")
    
Else
    'dados em branco
    txt(2) = ""
    txt(6) = ""
    txt(8) = ""
    
End If
Exit Sub
erro:
MsgBox Err.Description, vbCritical, "exibe_orc"
End Sub

Sub Confirma_Orcamento(orc As Integer)
On Error GoTo erro
Dim Sql As String
Sql = "update pf_orcamentoH set orch_situacao = 4 where orch_situacao = 3 and orcH_cod = " & orc

ocnBanco.Execute (Sql)
MsgBox "Ok."
'
Exit Sub
erro:
MsgBox Err.Description, vbCritical, "exibe_orc"

End Sub

