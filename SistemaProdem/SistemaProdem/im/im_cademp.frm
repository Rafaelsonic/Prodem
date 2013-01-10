VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form im_cademp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Empresa"
   ClientHeight    =   6225
   ClientLeft      =   465
   ClientTop       =   1710
   ClientWidth     =   7455
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1.383
   ScaleMode       =   0  'User
   ScaleWidth      =   0.754
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   5565
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   7215
      Begin VB.CommandButton cmdProc 
         Height          =   375
         Index           =   1
         Left            =   2760
         Picture         =   "im_cademp.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   4320
         Width           =   375
      End
      Begin VB.CommandButton cmdProc 
         Height          =   375
         Index           =   0
         Left            =   3120
         Picture         =   "im_cademp.frx":0420
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   1635
         Width           =   375
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   19
         Left            =   1560
         TabIndex        =   19
         Top             =   5100
         Width           =   3375
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   18
         Left            =   1560
         TabIndex        =   18
         Top             =   4770
         Width           =   1575
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   17
         Left            =   1560
         TabIndex        =   17
         Top             =   4410
         Width           =   1095
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   16
         Left            =   1320
         TabIndex        =   16
         Top             =   4020
         Width           =   3975
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   15
         Left            =   1320
         TabIndex        =   15
         Top             =   3690
         Width           =   3975
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   14
         Left            =   5160
         TabIndex        =   14
         Top             =   3360
         Width           =   1575
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   13
         Left            =   1320
         TabIndex        =   13
         Top             =   3360
         Width           =   1575
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   12
         Left            =   5160
         TabIndex        =   12
         Top             =   3030
         Width           =   1575
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   11
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   3030
         Width           =   1575
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   10
         Left            =   5160
         MaxLength       =   2
         TabIndex        =   10
         Top             =   2700
         Width           =   495
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   9
         Left            =   1320
         TabIndex        =   9
         Top             =   2700
         Width           =   1575
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   8
         Left            =   5160
         TabIndex        =   8
         Top             =   2370
         Width           =   1575
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   7
         Left            =   1320
         TabIndex        =   7
         Top             =   2370
         Width           =   2295
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   6
         Left            =   1320
         TabIndex        =   6
         Top             =   2040
         Width           =   3975
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   5
         Left            =   2040
         TabIndex        =   5
         Top             =   1650
         Width           =   975
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   4
         Left            =   1320
         TabIndex        =   4
         Top             =   1260
         Width           =   5535
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   3
         Left            =   1320
         TabIndex        =   3
         Top             =   930
         Width           =   5535
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   2
         Left            =   4680
         TabIndex        =   2
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "End. Eletronico"
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
         Index           =   19
         Left            =   120
         TabIndex        =   43
         Top             =   5100
         Width           =   1275
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
         Index           =   18
         Left            =   120
         TabIndex        =   42
         Top             =   4800
         Width           =   615
      End
      Begin VB.Label lblrisco 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3360
         TabIndex        =   41
         Top             =   4410
         Width           =   3735
      End
      Begin VB.Label lblramo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3720
         TabIndex        =   40
         Top             =   1650
         Width           =   3135
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Grau de Risco"
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
         Index           =   17
         Left            =   120
         TabIndex        =   39
         Top             =   4470
         Width           =   1200
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
         Index           =   16
         Left            =   120
         TabIndex        =   38
         Top             =   4080
         Width           =   510
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Contato"
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
         TabIndex        =   37
         Top             =   3750
         Width           =   675
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Fax"
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
         Left            =   4320
         TabIndex        =   36
         Top             =   3390
         Width           =   300
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
         Index           =   13
         Left            =   120
         TabIndex        =   35
         Top             =   3420
         Width           =   735
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Status"
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
         Left            =   4080
         TabIndex        =   34
         Top             =   3090
         Width           =   540
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Dt. Cadastro"
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
         TabIndex        =   33
         Top             =   3090
         Width           =   1095
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
         Index           =   10
         Left            =   4440
         TabIndex        =   32
         Top             =   2820
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
         Index           =   9
         Left            =   120
         TabIndex        =   31
         Top             =   2760
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
         Index           =   8
         Left            =   4320
         TabIndex        =   30
         Top             =   2490
         Width           =   345
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Bairro"
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
         TabIndex        =   29
         Top             =   2430
         Width           =   525
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Endereço"
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
         TabIndex        =   28
         Top             =   2100
         Width           =   795
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Ramo de Atividade"
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
         TabIndex        =   27
         Top             =   1710
         Width           =   1620
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Fantasia"
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
         TabIndex        =   26
         Top             =   1320
         Width           =   705
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Razão"
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
         TabIndex        =   25
         Top             =   990
         Width           =   525
      End
      Begin VB.Label label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "I.E."
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
         Left            =   4155
         TabIndex        =   24
         Top             =   600
         Width           =   300
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ"
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
         Left            =   120
         TabIndex        =   23
         Top             =   630
         Width           =   435
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
         Left            =   120
         TabIndex        =   22
         Top             =   300
         Width           =   600
      End
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   5850
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "im_cademp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Data de criacao: 21/08/2003
'Criador :Rafael
'Ultima atualizacao:         por
Option Explicit
Dim nome_tab(0) As String
Dim cpo(19) As String 'NOME DOS CAMPOS
Dim tcpo(19) As String 'TIPO DOS CAMPOS
Dim Exibir As Boolean ' serve para ver se o formulario vai ficar aberto
Dim Selecao As Boolean 'verifica se exite alguma seleção
Dim snap_Selecao As New ADODB.Recordset ' objeto da seleção
Dim snap_Procura As New ADODB.Recordset ' objeto da seleção
'Descricao: manutencao

'Tabelas utilizadas: pf_empresa

Private Sub cmd_Click(Index As Integer)
Menu Index
End Sub

Private Sub cmdProc_Click(Index As Integer)
Select Case Index
    Case 0
        Cod_con = 10
        IM_consulta.Show 1
        lblramo = Result_con
        txt(5) = CpoChave_con
    Case 1
        Cod_con = 20
        IM_consulta.Show 1
        lblrisco = Result_con
        txt(17) = CpoChave_con
End Select
End Sub

Private Sub Form_Activate()
Me.Top = 0
Me.Left = 0
Botoes
If Exibir = False Then Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then cmd_Click (8) ' esc
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
nome_tab(0) = "im_empresa"
cpo(0) = "emp_cod"
cpo(1) = "emp_cnpj"
cpo(2) = "emp_ie"
cpo(3) = "emp_rzsc"
cpo(4) = "emp_fant"
cpo(5) = "emp_ramo"
cpo(6) = "emp_ende"
cpo(7) = "emp_bairr"
cpo(8) = "emp_cep"
cpo(9) = "emp_cida"
cpo(10) = "emp_uf"
cpo(11) = "emp_dtcad"
cpo(12) = "emp_status"
cpo(13) = "emp_tele"
cpo(14) = "emp_fax"
cpo(15) = "emp_cont"
cpo(16) = "emp_mail"
cpo(17) = "emp_risc"
cpo(18) = "emp_CEL"
cpo(19) = "emp_endeletronico"

tcpo(0) = "NUMERO"
tcpo(1) = "TEXTO"
tcpo(2) = "TEXTO"
tcpo(3) = "TEXTO"
tcpo(4) = "TEXTO"
tcpo(5) = "NUMERO"
tcpo(6) = "TEXTO"
tcpo(7) = "TEXTO"
tcpo(8) = "TEXTO"
tcpo(9) = "TEXTO"
tcpo(10) = "TEXTO"
tcpo(11) = "DATA"
tcpo(12) = "NUMERO"
tcpo(13) = "TEXTO"
tcpo(14) = "TEXTO"
tcpo(15) = "TEXTO"
tcpo(16) = "TEXTO"
tcpo(17) = "NUMERO"
tcpo(18) = "TEXTO"
tcpo(19) = "TEXTO"
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
    
    txt(11) = Date
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
lblramo = ""
lblrisco = ""
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
For X = 0 To txt.Count - 1
    txt(X) = snap_Selecao(cpo(X))
Next X

Sql = "select ris_desc, ram_desc from im_empresa, pf_cliente, pf_grisco, pf_ramativ where emp_ramo = ram_cod and emp_risc=ris_cod and emp_cod=" & snap_Selecao("emp_cod")
Set snap_Procura = ocnBanco.Execute(Sql)
snap_Procura.MoveFirst
lblramo = snap_Procura("ram_desc")
lblrisco = snap_Procura("ris_desc")
Set snap_Procura = Nothing
End Sub

Private Sub Form_Resize()
'Me.Caption = Me.Height & "   " & Me.Width
Me.Height = 6965
Me.Width = 7560
End Sub

Function conssiste() As Boolean
conssiste = False
' no txt 1 verificar cnpj
' no txt 2 veriicar inscricao estadual
If Len(txt(3)) <= 0 Then
    conssiste = False
    txt(3).SetFocus
    Exit Function
End If
If Len(txt(4)) = 0 Then txt(4) = txt(3)
If Len(lblramo) <= 0 Then
    conssiste = False
    txt(5).SetFocus
    Exit Function
End If
If Len(lblrisco) <= 0 Then
    conssiste = False
    txt(17).SetFocus
    Exit Function
End If
If InStr(1, txt(1), "*") > 0 Then
    conssiste = False
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

Private Sub txt_LostFocus(Index As Integer)
Dim Sql As String
Select Case Index
    Case 5 ' ramo de atividade
        If IsNumeric(txt(Index)) Then
            Sql = "Select ram_desc from pf_ramativ where ram_cod = " & txt(Index)
            Set snap_Procura = ocnBanco.Execute(Sql)
            If snap_Procura.BOF And snap_Procura.EOF Then
                lblramo = ""
            Else
                snap_Procura.MoveFirst
                lblramo = snap_Procura("ram_desc")
            End If
            Set snap_Procura = Nothing
        Else
            lblramo = ""
        End If
        
    Case 17 ' grau de risco
        If IsNumeric(txt(Index)) Then
            Sql = "Select ris_desc from pf_grisco where ris_cod = " & txt(Index)
            Set snap_Procura = ocnBanco.Execute(Sql)
            If snap_Procura.BOF And snap_Procura.EOF Then
                lblrisco = ""
            Else
                snap_Procura.MoveFirst
                lblrisco = snap_Procura("ris_desc")
            End If
            Set snap_Procura = Nothing
        Else
            lblrisco = ""
        End If
    
End Select

End Sub


Sub Menu(Index As Integer)
Dim Sql As String
Dim snap_Cod As New ADODB.Recordset
On Error GoTo erro


Select Case Index
Case 0 ' botao inserir Novo
    status.Panels(1).Text = "INSERIR"
    txt(0).Locked = True
    Botoes

Case 1 ' botao excluir
    'procura referencia
    'pede confirmação
    If status.Panels(1).Text = "ALTERAR" Then
        If MsgBox("Deseja realmente excluir ?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
            Sql = "delete from " & nome_tab(0) & " where emp_cod=" & txt(0)
            ocnBanco.Execute Sql
            cmd_Click (8)
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
                
                Sql = "SELECT MAX(emp_COD)+1 FROM " & nome_tab(0)
                Set snap_Cod = ocnBanco.Execute(Sql)
                If IsNull(snap_Cod(0)) Then
                    txt(0) = 1
                Else
                    txt(0) = snap_Cod(0)
                End If
                Set snap_Cod = Nothing
            
            
                Sql = "insert into " & nome_tab(0) & Chr(13)
                Sql = Sql & "values(" & txt(0) & ",'" & txt(1) & "','" & txt(2) & "','" & txt(3) & "','" & txt(4) & "','" & txt(5) & "','" & txt(6) & "','" & txt(7) & "','" & txt(8) & "','" & txt(9) & "','" & txt(10) & "','" & txt(11) & "','" & txt(12) & "','" & txt(13) & "','" & txt(14) & "','" & txt(15) & "','" & txt(16) & "','" & txt(17) & "','" & txt(18) & "','" & txt(19) & "')"
                ocnBanco.Execute (Sql)
                status.Panels(1).Text = "ALTERAR"
                MsgBox "Os dados foram gravados com sucesso!", vbInformation, "Gravação"
                Botoes
            End If
        End If
    ElseIf status.Panels(1).Text = "ALTERAR" Then
        If conssiste = True Then
            Sql = "UPDATE " & nome_tab(0) & " SET " & cpo(1) & "='" & txt(1) & "'," & cpo(2) & "='" & txt(2) & "'," & cpo(3) & "='" & txt(3) & "'," & cpo(4) & "='" & txt(4) & "'," & cpo(5) & "='" & txt(5) & "'," & cpo(6) & "='" & txt(6) & "'," & cpo(7) & "='" & txt(7) & "'," & cpo(8) & "='" & txt(8) & "'," & cpo(9) & "='" & txt(9) & "'," & cpo(10) & "='" & txt(10) & "'," & cpo(11) & "='" & txt(11) & "'," & cpo(12) & "='" & txt(12) & "'," & cpo(13) & "='" & txt(13) & "'," & cpo(14) & "='" & txt(14) & "'," & cpo(15) & "='" & txt(15) & "'," & cpo(16) & "='" & txt(16) & "'," & cpo(17) & "='" & txt(17) & "'," & cpo(18) & "='" & txt(18) & "'," & cpo(19) & "='" & txt(19) & "'" & Chr(13)
            Sql = Sql & "WHERE " & cpo(0) & "=" & txt(0)
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
