VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form pf_cer_rela 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impressão de Certificados"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10545
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   10545
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   40
      Top             =   8010
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame5 
      Height          =   2055
      Left            =   9000
      TabIndex        =   33
      Top             =   2760
      Width           =   1455
      Begin VB.CommandButton cmd_tools 
         Caption         =   "Atestado"
         Height          =   495
         Index           =   2
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "0000"
         Top             =   1440
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7455
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin VB.CommandButton cmdProc 
         DownPicture     =   "pf_cer_rela.frx":0000
         Height          =   495
         Index           =   4
         Left            =   2820
         Picture         =   "pf_cer_rela.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   2940
         Width           =   495
      End
      Begin VB.CommandButton cmdProc 
         DownPicture     =   "pf_cer_rela.frx":25C4
         Height          =   495
         Index           =   3
         Left            =   2820
         Picture         =   "pf_cer_rela.frx":2E8E
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   2340
         Width           =   495
      End
      Begin VB.CommandButton cmdProc 
         DownPicture     =   "pf_cer_rela.frx":4B88
         Height          =   495
         Index           =   2
         Left            =   3540
         Picture         =   "pf_cer_rela.frx":5452
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1710
         Width           =   495
      End
      Begin VB.CommandButton cmdProc 
         DownPicture     =   "pf_cer_rela.frx":714C
         Height          =   495
         Index           =   1
         Left            =   2460
         Picture         =   "pf_cer_rela.frx":7A16
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton cmdProc 
         DownPicture     =   "pf_cer_rela.frx":9710
         Height          =   495
         Index           =   0
         Left            =   2460
         Picture         =   "pf_cer_rela.frx":9FDA
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   510
         Width           =   495
      End
      Begin VB.TextBox TXT 
         Height          =   315
         Index           =   8
         Left            =   1560
         TabIndex        =   30
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox TXT 
         Height          =   315
         Index           =   7
         Left            =   1200
         TabIndex        =   28
         Top             =   600
         Width           =   1095
      End
      Begin VB.Frame Frame4 
         Height          =   615
         Left            =   120
         TabIndex        =   23
         Top             =   3960
         Width           =   8175
         Begin VB.OptionButton opt 
            Caption         =   "Impresso"
            Height          =   255
            Index           =   2
            Left            =   5160
            TabIndex        =   26
            Top             =   240
            Width           =   2055
         End
         Begin VB.OptionButton opt 
            Caption         =   "Impresso Parcialmente"
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   25
            Top             =   240
            Width           =   2055
         End
         Begin VB.OptionButton opt 
            Caption         =   "Não Impresso"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   24
            Top             =   240
            Value           =   -1  'True
            Width           =   2055
         End
      End
      Begin VB.Frame fraALn 
         Height          =   2775
         Left            =   120
         TabIndex        =   16
         Top             =   4560
         Visible         =   0   'False
         Width           =   8175
         Begin MSComDlg.CommonDialog ctr 
            Left            =   7110
            Top             =   2280
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin Crystal.CrystalReport CR 
            Left            =   7680
            Top             =   2280
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin VB.CommandButton cmd_imp 
            Height          =   615
            Left            =   7200
            Picture         =   "pf_cer_rela.frx":BCD4
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox TXT 
            Height          =   315
            Index           =   6
            Left            =   960
            TabIndex        =   18
            Top             =   240
            Width           =   1095
         End
         Begin MSFlexGridLib.MSFlexGrid Grid 
            Height          =   1575
            Left            =   360
            TabIndex        =   17
            Top             =   1080
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   2778
            _Version        =   393216
            Rows            =   1
            Cols            =   3
            TextStyleFixed  =   4
            FocusRect       =   0
            FillStyle       =   1
            SelectionMode   =   1
         End
         Begin VB.Label label 
            Caption         =   "Aluno"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   15
            TabIndex        =   20
            Top             =   360
            Width           =   495
         End
         Begin VB.Label lbl 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   3240
            TabIndex        =   19
            Top             =   240
            Width           =   3855
         End
      End
      Begin VB.TextBox TXT 
         Height          =   315
         Index           =   5
         Left            =   2760
         TabIndex        =   12
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox TXT 
         Height          =   315
         Index           =   4
         Left            =   1560
         TabIndex        =   11
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox TXT 
         Height          =   315
         Index           =   3
         Left            =   2280
         TabIndex        =   10
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox TXT 
         Height          =   315
         Index           =   2
         Left            =   1200
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox TXT 
         Height          =   315
         Index           =   1
         Left            =   4440
         TabIndex        =   8
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox TXT 
         Height          =   315
         Index           =   0
         Left            =   1200
         TabIndex        =   7
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lbl 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   3480
         TabIndex        =   32
         Top             =   3000
         Width           =   3855
      End
      Begin VB.Label label 
         Caption         =   "Coordenador"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   240
         TabIndex        =   31
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label lbl 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   3120
         TabIndex        =   29
         Top             =   600
         Width           =   4215
      End
      Begin VB.Label label 
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   240
         TabIndex        =   27
         Top             =   720
         Width           =   960
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   5760
         TabIndex        =   22
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label lbl 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   3480
         TabIndex        =   15
         Top             =   2400
         Width           =   3855
      End
      Begin VB.Label lbl 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   4200
         TabIndex        =   14
         Top             =   1800
         Width           =   3135
      End
      Begin VB.Label lbl 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   3120
         TabIndex        =   13
         Top             =   1320
         Width           =   4215
      End
      Begin VB.Label label 
         Caption         =   "Instrutor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   255
         TabIndex        =   6
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label label 
         Caption         =   "Tipo do Treinamento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Width           =   1740
      End
      Begin VB.Label label 
         Caption         =   "Data do Treinamento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   3600
         Width           =   1785
      End
      Begin VB.Label label 
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   255
         TabIndex        =   3
         Top             =   1440
         Width           =   600
      End
      Begin VB.Label label 
         Caption         =   "Data"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   3840
         TabIndex        =   2
         Top             =   240
         Width           =   405
      End
      Begin VB.Label label 
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   255
         TabIndex        =   1
         Top             =   240
         Width           =   600
      End
   End
End
Attribute VB_Name = "pf_cer_rela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Data de criacao:21/08/2003
'Criador :Rafael
'Ultima atualizacao:         por
Option Explicit
Dim Nome_tab(1) As String
Dim cpo(8) As String 'NOME DOS CAMPOS
Dim Tcpo(8) As String 'TIPO DOS CAMPOS
Dim Exibir As Boolean ' serve para ver se o formulario vai ficar aberto
Dim Selecao As Boolean 'verifica se exite alguma seleção
Dim snap_Selecao As New ADODB.Recordset ' objeto da seleção
Dim snap_Procura As New ADODB.Recordset ' objeto da seleção auxiliar
Dim int_aluno As Integer
'Descricao: manutencao
'Tabelas utilizadas: pf_ramativ


Private Sub cmd_Click(index As Integer)
Dim SQL As String

On Error GoTo erro

Select Case index
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
    txt(1).Locked = False
End Select
Exit Sub
erro:
MsgBox Err.Description, vbCritical, "cmd_click"
End Sub


Private Sub cmd_sel_Click(index As Integer)
Select Case index
Case 0 ' Não Impressos
    
    Dim SQL As String, sql_where As String
    Dim X As Integer
    If Selecao = True Then
        If conssiste_total = True Then ' procurar os dados
            SQL = "select * from " & Nome_tab(0) & ", " & Nome_tab(1)
            For X = 0 To 5
                If Len(txt(X)) > 0 Then
                    If Len(sql_where) <= 0 Then
                        sql_where = " where "
                    Else
                        sql_where = sql_where & " and " & Chr(13)
                    End If
                    sql_where = sql_where & Monta_SQL(cpo(X), Tcpo(X), txt(X))
                End If
            Next X
        If Len(sql_where) > 0 Then
            sql_where = sql_where & " and cerh_cod = ceri_codcerh and ceri_status='0' "
        Else
            sql_where = " where cerh_cod = ceri_codcerh and ceri_status='0' "
        End If
        SQL = SQL & sql_where
        snap_Selecao.Open SQL, ocnBanco, adOpenKeyset, adLockOptimistic, adCmdText
        Selecao = True
        Else
            Exit Sub
        End If
    End If
    
    Selecao = False
    If Not (snap_Selecao.EOF And snap_Selecao.BOF) Then
        snap_Selecao.MoveFirst
        Mostra_Dados
        status.Panels(1).Text = "ALTERAR"
        Botoes
    Else
        MsgBox "Não foi encontrado dados para a seleção!", vbInformation, "Movimentação"
    End If


    
    
    
    
    
    
    
End Select
End Sub

Private Sub cmd_imp_Click()
Dim SQL As String
Dim int_nimp As Integer, int_imp As Integer  ' nao impresso e impresso respectivamente
If Impressora = True Then
    If Len(lbl(3)) > 0 Then
        SQL = "update " & Nome_tab(1) & Chr(13)
        SQL = SQL & "set ceri_status=1" & Chr(13)
        SQL = SQL & "where ceri_codcerh=" & txt(0) & " and ceri_codaln=" & txt(6)
        ocnBanco.Execute SQL
    
        'verificar staus total
        
        SQL = "select count(*) from " & Nome_tab(1) & Chr(13)
        SQL = SQL & "where ceri_status=0 and ceri_codcerh=" & txt(0)
        Set snap_Procura = ocnBanco.Execute(SQL)
        int_nimp = Int(snap_Procura(0))
        SQL = "select count(*) from " & Nome_tab(1) & Chr(13)
        SQL = SQL & "where ceri_status=1 and ceri_codcerh=" & txt(0)
        Set snap_Procura = ocnBanco.Execute(SQL)
        int_imp = Int(snap_Procura(0))
        Set snap_Procura = Nothing
        
        If int_nimp = (int_nimp + int_imp) Then
            'pedido nao impresso
            SQL = "update " & Nome_tab(0) & " set cerh_status=0" & Chr(13)
            lbl(4) = "0 - Não Impresso"
        ElseIf int_imp = (int_imp + int_nimp) Then
            'pedido impresso
            SQL = "update " & Nome_tab(0) & " set cerh_status=2" & Chr(13)
            lbl(4) = "2 - Impresso"
            
        Else
            'pedido impresso parcialmente
            SQL = "update " & Nome_tab(0) & " set cerh_status=1" & Chr(13)
            lbl(4) = "1 - Impresso Parcialmente"
        End If
        SQL = SQL & "where cerh_cod=" & txt(0)
        ocnBanco.Execute (SQL)
        view_aluno
        Print_Cert txt(0), txt(6)
    Else
        
        SQL = "update " & Nome_tab(1) & Chr(13)
        SQL = SQL & "set ceri_status=1" & Chr(13)
        SQL = SQL & "where ceri_codcerh=" & txt(0)
        ocnBanco.Execute SQL
        
        SQL = "update " & Nome_tab(0) & " set cerh_status=2" & Chr(13)
        lbl(4) = "2 - Impresso"
            
        SQL = SQL & "where cerh_cod=" & txt(0)
        ocnBanco.Execute (SQL)
        view_aluno
        Print_Cert txt(0), 0
    End If
        
End If
'atualizar o grid e o status
End Sub

Private Sub cmd_tools_Click(index As Integer)
Select Case index
    Case 0
    










    Case 2
        Atestado (txt(0))
End Select
End Sub

Private Sub cmdProc_Click(index As Integer)
Select Case index
    Case 0
        Cod_con = 30
        IM_consulta.Show 1
        lbl(5) = Result_con
        txt(7) = CpoChave_con
    Case 1
        Cod_con = 40
        IM_consulta.Show 1
        lbl(0) = Result_con
        txt(2) = CpoChave_con
    Case 2
        Cod_con = 50
        IM_consulta.Show 1
        lbl(1) = Result_con
        txt(3) = CpoChave_con
    Case 3
        Cod_con = 60
        IM_consulta.Show 1
        lbl(2) = Result_con
        txt(4) = CpoChave_con
    Case 4
        Cod_con = 70
        IM_consulta.Show 1
        lbl(6) = Result_con
        txt(8) = CpoChave_con
    

End Select

End Sub

Private Sub Form_Activate()
Botoes
Me.Top = 0
Me.Left = 0
If Exibir = False Then Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then cmd_Click (8) ' esc
End Sub

Private Sub Form_Load()
Dim snap_tab As New ADODB.Recordset
Dim SQL As String
Declara

On Error GoTo erro

'pesquisa dados na tabela
SQL = "select count(*) from " & Nome_tab(0)
Set snap_tab = ocnBanco.Execute(SQL)
If snap_tab(0) <= 0 Then
MsgBox "Não exixte dados para impressão!", vbInformation, "Atenção"
End If
Exibir = True
Exit Sub
erro:
MsgBox Err.Description, vbCritical, "FORM LOAD " & Me.Name
Exibir = False

Cabecalho_grid
End Sub

Sub Declara()
' declara todas as varíaveis de dados
Nome_tab(0) = "pf_cerh"
Nome_tab(1) = "pf_ceri"
cpo(0) = "cerH_cod"
cpo(1) = "cerH_dtcri"
cpo(2) = "cerH_cli"
cpo(3) = "cerH_tpcert"
cpo(4) = "cerH_ins"
cpo(5) = "cerH_dttrein"
cpo(6) = "ceri_codaln"
cpo(7) = "cerh_emp"
cpo(8) = "cerh_coord"
Tcpo(0) = "NUMERO"
Tcpo(1) = "DATA"
Tcpo(2) = "NUMERO"
Tcpo(3) = "NUMERO"
Tcpo(4) = "NUMERO"
Tcpo(5) = "DATA"
Tcpo(6) = "NUMERO"
Tcpo(7) = "NUMERO"
Tcpo(8) = "NUMERO"
End Sub

Sub Botoes() 'Bloqueia ou nao os botões
mdi.Toolbar.Buttons(5).Enabled = True
Dim SQL As String
Dim SNAP_COD As New ADODB.Recordset
Select Case status.Panels(1).Text
Case "Impressao"
    opt(0).Enabled = False
    opt(1).Enabled = False
    opt(2).Enabled = False
    txt(0).Locked = True
    txt(1).Locked = True
    txt(2).Locked = True
    txt(3).Locked = True
    txt(4).Locked = True
    txt(5).Locked = True
    txt(6).Locked = False
    txt(7).Locked = True
    fraALn.Visible = True
    view_aluno
Case "SELECIONAR"
    mdi.Toolbar.Buttons(4).Enabled = True
    mdi.Toolbar.Buttons(6).Enabled = True
    mdi.Toolbar.Buttons(7).Enabled = True
    mdi.Toolbar.Buttons(8).Enabled = True
    mdi.Toolbar.Buttons(9).Enabled = True: mdi.Toolbar.Buttons(10).Enabled = True

    opt(0).Enabled = True
    opt(1).Enabled = True
    opt(2).Enabled = True
    txt(0).Locked = False
    txt(1).Locked = False
    txt(2).Locked = False
    txt(3).Locked = False
    txt(4).Locked = False
    txt(5).Locked = False
    txt(6).Locked = False
    txt(7).Locked = False
    
    
    fraALn.Visible = False
    txt(0).SetFocus

End Select
End Sub
Sub Limpar()
Dim X As Integer
For X = 0 To txt.Count - 1
    txt(X).Text = ""
Next X
For X = 0 To 6
lbl(X) = ""
Next X
grid.Rows = 1
fraALn.Visible = False
End Sub

Function conssiste_total() As Boolean
Dim X As Integer
On Error GoTo erro
For X = 0 To txt.Count - 1
Select Case Tcpo(X)
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
    Else
        txt(X) = Empty
    End If
End Select
Next X
conssiste_total = True
Exit Function
erro:
MsgBox Err.Description, vbCritical, "CONSSISTE_TOTAL"

End Function

Sub Movimentacao(botao As String)
Dim SQL As String, sql_where As String
Dim X As Integer
If Selecao = True Then
    If conssiste_total = True Then ' procurar os dados
        SQL = "select * from " & Nome_tab(0) & Chr(13)
        For X = 0 To 2
            If opt(X).Value = True Then
                sql_where = "WHERE CERH_STATUS=" & X & Chr(13)
                Exit For
            End If
        Next X
        For X = 0 To 5
            If Len(txt(X)) > 0 Then
                sql_where = sql_where & " and " & Monta_SQL(cpo(X), Tcpo(X), txt(X))
            End If
        Next X
    sql_where = sql_where & " and ((select count(*) from pf_ceri where ceri_codcerh=cerh_cod)>0)"
    SQL = SQL & sql_where
    
    snap_Selecao.Open SQL, ocnBanco, adOpenKeyset, adLockOptimistic, adCmdText
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
        status.Panels(1).Text = "Impressao"

        Botoes
    Else
        MsgBox "Não foi encontrado dados para a seleção!", vbInformation, "Movimentação"
        cmd_Click (8)
    End If

End Sub

Sub Mostra_Dados()
Dim X As Integer
Dim SQL As String
For X = 0 To 5
    txt(X) = snap_Selecao(cpo(X))
Next X
txt(6) = ""
txt(7) = snap_Selecao(cpo(7))
txt(8) = snap_Selecao(cpo(8))
lbl(3) = ""
SQL = "select cli_fant, tpcer_desc, instrutor.tec_nome , emp_rzsc, coordenador.tec_nome from pf_tpcer, pf_cliente, pf_tecn as instrutor, pf_cerh, im_empresa, pf_tecn as coordenador where cerh_cod=" & txt(0) & Chr(13)
SQL = SQL & "and cerh_cli = cli_cod and cerh_tpcert = tpcer_cod and cerh_ins = instrutor.tec_cod and cerh_emp = emp_cod and cerh_coord = coordenador.tec_cod"
Set snap_Procura = ocnBanco.Execute(SQL)
snap_Procura.MoveFirst
lbl(0) = snap_Procura(0)
lbl(1) = snap_Procura(1)
lbl(2) = snap_Procura(2)
lbl(5) = snap_Procura(3)
lbl(6) = snap_Procura(4)
Select Case snap_Selecao("cerh_status")
Case 0
    lbl(4) = "0 - Não Impresso"
Case 1
    lbl(4) = "1 - Impresso Parcialmente"
Case 2
    lbl(4) = "2 - Impresso"
End Select
Set snap_Procura = Nothing
'view_aluno
End Sub

Private Sub Form_Resize()
'Me.Caption = Me.Height
Me.Width = 11200
Me.Height = 8500
End Sub

Function conssiste(index As Integer) As Boolean
Select Case index
Case 0
If Len(txt(2)) <= 0 Then
    conssiste = False
    txt(2).SetFocus
    Exit Function
End If
If Len(txt(3)) <= 0 Then
    conssiste = False
    txt(3).SetFocus
    Exit Function
End If
If Len(txt(4)) <= 0 Then
    conssiste = False
    txt(4).SetFocus
    Exit Function
End If
If Len(txt(5)) <= 0 Then
    txt(5) = Date
End If
Case 1
    MsgBox 1
End Select
conssiste = True
End Function


Private Sub Grid_Click()
If grid.RowSel > 0 Then
    grid.Row = grid.RowSel
    grid.col = 0
    txt(6) = grid.Text
    grid.col = 1
    lbl(3) = grid.Text
    grid.ColSel = 2
End If
End Sub


Private Sub txt_GotFocus(index As Integer)
txt(index).SelStart = 0
txt(index).SelLength = Len(txt(index))
End Sub

Private Sub txt_LostFocus(index As Integer)
Dim SQL As String
Select Case index
    Case 2 ' codigo do cliente
        If IsNumeric(txt(index)) Then
            SQL = "Select cli_fant from pf_cliente where cli_cod = " & txt(index)
            Set snap_Procura = ocnBanco.Execute(SQL)
            If snap_Procura.BOF And snap_Procura.EOF Then
                lbl(0) = ""
            Else
                snap_Procura.MoveFirst
                lbl(0) = snap_Procura(0)
            End If
            Set snap_Procura = Nothing
        Else
            lbl(0) = ""
        End If
        
    Case 3 ' Tipo do trinamento
        If IsNumeric(txt(index)) Then
            SQL = "Select tpcer_desc from pf_tpcer where tpcer_cod = " & txt(index)
            Set snap_Procura = ocnBanco.Execute(SQL)
            If snap_Procura.BOF And snap_Procura.EOF Then
                lbl(1) = ""
            Else
                snap_Procura.MoveFirst
                lbl(1) = snap_Procura(0)
            End If
            Set snap_Procura = Nothing
        Else
            lbl(1) = ""
        End If
        
    Case 4 ' Instrutor
        If IsNumeric(txt(index)) Then
            SQL = "Select tec_nome from pf_tecn where tec_cod = " & txt(index)
            Set snap_Procura = ocnBanco.Execute(SQL)
            If snap_Procura.BOF And snap_Procura.EOF Then
                lbl(2) = ""
            Else
                snap_Procura.MoveFirst
                lbl(2) = snap_Procura(0)
            End If
            Set snap_Procura = Nothing
        Else
            lbl(2) = ""
        End If
    
    Case 6 ' Aluno
        If IsNumeric(txt(index)) Then
            SQL = "Select aln_nome from pf_aluno where aln_cod = " & txt(index)
            SQL = SQL & " and aln_cli=" & txt(2)
            Set snap_Procura = ocnBanco.Execute(SQL)
            If snap_Procura.BOF And snap_Procura.EOF Then
                lbl(3) = ""
            Else
                snap_Procura.MoveFirst
                lbl(3) = snap_Procura(0)
            End If
            Set snap_Procura = Nothing
        Else
            lbl(3) = ""
        End If
    
    Case 7 ' Empresa
        If IsNumeric(txt(index)) Then
            SQL = "Select emp_rzsc from im_empresa where emp_cod = " & txt(index)
            Set snap_Procura = ocnBanco.Execute(SQL)
            If snap_Procura.BOF And snap_Procura.EOF Then
                lbl(5) = ""
            Else
                snap_Procura.MoveFirst
                lbl(5) = snap_Procura(0)
            End If
            Set snap_Procura = Nothing
        Else
            lbl(5) = ""
        End If
    Case 8 ' coordenador
        If IsNumeric(txt(index)) Then
            SQL = "Select tec_nome from pf_tecn where tec_cod = " & txt(index)
            Set snap_Procura = ocnBanco.Execute(SQL)
            If snap_Procura.BOF And snap_Procura.EOF Then
                lbl(6) = ""
            Else
                snap_Procura.MoveFirst
                lbl(6) = snap_Procura(0)
            End If
            Set snap_Procura = Nothing
        Else
            lbl(6) = ""
        End If
End Select

End Sub

Private Sub Cabecalho_grid()
With grid
    .Cols = 4
    .Rows = 1
    .ColWidth(0) = 1000
    .ColWidth(1) = 4700
    .ColWidth(2) = 1500
    .Row = 0
    .col = 0
    .Text = "Código"
    .col = 1
    .Text = "Nome"
    .col = 2
    .Text = "Status"
    .col = 3
    .Text = "Aprovado"
End With
End Sub


Sub view_aluno()
Dim SQL As String
Dim Impressao As String, Exame As String
SQL = "Select aln_cod, aln_nome, ceri_status, ceri_aprovado from pf_aluno, pf_cerh, pf_ceri" & Chr(13)
SQL = SQL & "where cerh_cod=" & txt(0) & " and aln_cli= cerh_cli and aln_cod = ceri_codaln and ceri_codcerh = cerh_cod"
snap_Procura.Open SQL, ocnBanco
'snap_Procura.MoveFirst
grid.Clear
Cabecalho_grid
int_aluno = 2
grid.Rows = 2
grid.Row = 1
If Not (snap_Procura.BOF And snap_Procura.EOF) Then snap_Procura.MoveFirst
While Not snap_Procura.EOF
    grid.col = 0
    grid.Text = snap_Procura(0)
    grid.CellAlignment = 3
    grid.col = 1
    grid.Text = snap_Procura(1)
    grid.CellAlignment = 1
    grid.col = 2
    If snap_Procura(2) = 0 Then
        Impressao = "Não Impresso"
        
        grid.CellBackColor = &HFF&
    Else
        Impressao = "Impresso"
        grid.CellBackColor = 65280
    End If
    grid.Text = Impressao
    grid.CellAlignment = 1
    
    If Len(snap_Procura(3)) > 0 Then
        If snap_Procura(3) = 1 Then
            Exame = "SIM"
        ElseIf snap_Procura(3) = 0 Then
            Exame = "NAO"
        End If
    End If
    grid.col = 3
    grid.Text = Exame
    grid.CellAlignment = 1
    
    snap_Procura.MoveNext
    int_aluno = int_aluno + 1
    grid.Rows = int_aluno
    grid.Row = int_aluno - 1
    
    
Wend
grid.Rows = grid.Rows - 1
int_aluno = int_aluno - 1
Set snap_Procura = Nothing
End Sub

Sub Print_Cert(Cer As Integer, aln As Integer)
Dim SQL As String
SQL = "select tpcer_relat from pf_tpcer, pf_cerh" & Chr(13)
SQL = SQL & ", pf_ceri where cerh_cod=" & Cer & " and cerh_tpcert = tpcer_cod "
Set snap_Procura = ocnBanco.Execute(SQL)
If Len(snap_Procura(0)) > 0 Then
    CR.CopiesToPrinter = ctr.Copies
    CR.WindowShowPrintBtn = False
    If aln = 0 Then 'se vedadeiro imprimir todos
        CR.SelectionFormula = "{PF_CERH.CERH_COD}=" & Cer
        '& " AND {pf_cerI.ceri_status}=0"
    Else
        CR.SelectionFormula = "{PF_ALUNO.ALN_COD}=" & aln & " AND {PF_CERH.CERH_COD}=" & Cer
    End If
    CR.ReportFileName = snap_Procura(0)
    CR.WindowState = crptMaximized
    CR.Destination = crptToPrinter
    'CR.Destination = crptToWindow
    CR.Action = 0
End If
Set snap_Procura = Nothing
End Sub

