VERSION 5.00
Begin VB.Form pf_cadOrcCli 
   Caption         =   "Clientes"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   5760
   Begin VB.ComboBox cboMidia 
      Height          =   315
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   1320
      TabIndex        =   3
      Top             =   1320
      Width           =   3495
   End
   Begin VB.CommandButton cmd 
      Height          =   405
      Index           =   0
      Left            =   240
      Picture         =   "pf_cadOrcCli.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   405
   End
   Begin VB.CommandButton cmd 
      Height          =   405
      Index           =   1
      Left            =   720
      Picture         =   "pf_cadOrcCli.frx":0464
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   1320
      TabIndex        =   2
      Top             =   960
      Width           =   3495
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   3495
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label Lbl 
      Caption         =   "Conheceu a Empresa.."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   360
      TabIndex        =   11
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Lbl 
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
      Height          =   315
      Index           =   3
      Left            =   360
      TabIndex        =   10
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Lbl 
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
      Height          =   315
      Index           =   2
      Left            =   360
      TabIndex        =   9
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Lbl 
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
      Height          =   315
      Index           =   1
      Left            =   360
      TabIndex        =   8
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Lbl 
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
      Height          =   315
      Index           =   0
      Left            =   360
      TabIndex        =   7
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "pf_cadOrcCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sql As String
Dim cpo(5) As String
Dim tcpo(5) As String
Dim nome_tab(0) As String
Dim ID As Integer
Dim snap_Procura As New ADODB.Recordset
Sub Declara()
' declara todas as varíaveis de dados
nome_tab(0) = "pf_cliente"

cpo(0) = "cli_cod"
cpo(1) = "cli_rzsc"
cpo(2) = "cli_cont"
cpo(3) = "cli_tele"
cpo(4) = "cli_mail"
cpo(5) = "cli_midia"

tcpo(1) = "NUMERO"
tcpo(1) = "TEXTO"
tcpo(2) = "TEXTO"
tcpo(3) = "TEXTO"
tcpo(4) = "TEXTO"
tcpo(5) = "NUMERO"
End Sub



Sub GravarDados()
On Error GoTo erro
Dim snap_Cod As New ADODB.Recordset
Sql = "SELECT MAX(" & cpo(0) & ")+1 FROM " & nome_tab(0)
Set snap_Cod = ocnBanco.Execute(Sql)
If IsNull(snap_Cod(0)) Then
    ID = 1
Else
    ID = snap_Cod(0)
End If
Set snap_Cod = Nothing
'forma generica para inserir dados
Sql = "insert into " & nome_tab(0) & Chr(13)
Sql = Sql & "("
For x = 0 To 4
    Sql = Sql & cpo(x)
    If x < 4 Then Sql = Sql & ","
Next x
Sql = Sql & " ,cli_midia "
Sql = Sql & ",cli_risc, cli_ramo, cli_dtcad,clienteTipo) "
Sql = Sql & "values( " & ID & ","
For x = 1 To 4
    Select Case tcpo(x)
        Case "TEXTO"
            If Len(txt(x)) > 0 Then
                Sql = Sql & "'" & txt(x) & "'"
            Else
                Sql = Sql & "null"
            End If
        Case "NUMERO"
            If Len(txt(x)) > 0 Then
                Sql = Sql & txt(x) & ""
            Else
                Sql = Sql & "null"
            End If
        Case "DATA"
            If Len(txt(x)) > 0 Then
                Sql = Sql & "#" & Format(txt(x), "mm/dd/yyyy") & "#"
            Else
                Sql = Sql & "null"
            End If
    End Select
    If x < 4 Then Sql = Sql & ","
Next x
Sql = Sql & "," & cboMidia.ItemData(cboMidia.ListIndex)
Sql = Sql & ",1,1,"
Sql = Sql & "#" & Format(Date, "mm/dd/yyyy") & "#,1)" & vbCrLf
                                
ocnBanco.Execute (Sql)
Exit Sub
erro:
MsgBox Err.Description, "Gravar dados"
ID = 0
End Sub


Private Sub cmd_Click(Index As Integer)
Select Case Index
Case 0
    If Len(txt(1)) > 0 Then
        If cboMidia.ListIndex <> -1 Then
            GravarDados
            pf_Orcamento.Tag = ID
            Unload Me
        Else
            cboMidia.SetFocus
        End If
    Else
        txt(1).SetFocus
    End If
Case 1
    'GravarDados
    pf_Orcamento.Tag = 0
    Unload Me
End Select
End Sub

Private Sub Form_Load()
Declara
Preenche_Midia
End Sub

Sub Preenche_Midia()
Set snap_Procura = Nothing
Dim Sql As String
Dim i As Integer
Sql = "select * from pf_midia order by mid_desc"
snap_Procura.Open Sql, ocnBanco, adOpenKeyset, adLockOptimistic, adCmdText
cboMidia.Clear
i = 0
If Not (snap_Procura.BOF And snap_Procura.EOF) Then
    snap_Procura.MoveFirst
    While Not snap_Procura.EOF
        cboMidia.AddItem snap_Procura("mid_desc")
        cboMidia.ItemData(i) = snap_Procura("mid_cod")
        i = i + 1
        snap_Procura.MoveNext
    Wend
End If
End Sub

