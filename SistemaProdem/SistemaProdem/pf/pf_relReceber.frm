VERSION 5.00
Begin VB.Form pf_relReceber 
   Caption         =   "Relatório- Contas à Receber"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   4515
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   4215
      Begin VB.CommandButton cmdProc 
         Height          =   375
         Index           =   0
         Left            =   2760
         Picture         =   "pf_relReceber.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1680
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
         Height          =   285
         Index           =   3
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   5
         Top             =   2160
         Width           =   1215
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
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   3
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CheckBox Check 
         Caption         =   "Títulos Pendentes"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   2640
         Width           =   1695
      End
      Begin VB.ComboBox cbo 
         Height          =   315
         ItemData        =   "pf_relReceber.frx":0420
         Left            =   1440
         List            =   "pf_relReceber.frx":0430
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1200
         Width           =   1695
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
         Height          =   285
         Index           =   1
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   1
         Top             =   720
         Width           =   1215
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
         Height          =   285
         Index           =   0
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   0
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Vizualizar"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   3120
         Width           =   975
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
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Nr. da Nota"
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
         TabIndex        =   12
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Pesquisa Por:"
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
         TabIndex        =   11
         Top             =   1200
         Width           =   1170
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Data Final"
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
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Data Inicial"
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
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "pf_relReceber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim relatorio As New cls_Relatorio

Private Sub cmd_Click()
Dim campo As String
Dim strFormula As String
If cbo.ListIndex = -1 Then
    cbo.SetFocus
Else
    Select Case cbo.ListIndex
    Case 0
        campo = "{pf_LancReceb.rec_dtrein}"
    Case 1
        campo = "{pf_LancReceb.rec_dtvenc}"
    Case 2
        campo = "{pf_LancReceb.rec_dtPGTO}"
    Case 3
        campo = "{pf_LancReceb.rec_dtemiss}"
    End Select
    strFormula = campo & " >=#" & Data(txt(0)) & "# and " & campo & " <=#" & Data(txt(1)) & "#"
    
    If Len(txt(2)) > 0 Then 'se existe alguma coisa no cliente
        strFormula = strFormula & " and {pf_cliente.cli_cod} in [" & Replace(txt(2), ";", ",") & "]"
    End If
    
    If Len(txt(3)) > 0 Then 'existe alguma nota
        strFormula = strFormula & " and {pf_LancReceb.rec_nrnf} = '" & txt(3) & "'"
    End If
    
    If Check.Value = 1 Then
        'strFormula = strFormula & "and (CStr ({pf_LancReceb.rec_dtPGTO}) ='' or not isnull({pf_LancReceb.rec_dtpgto}) or {pf_LancReceb.rec_dtvenc} <CurrentDate) and {@Diferenca}>0"
        strFormula = strFormula & " and {@Diferenca}>0"
    End If
    relatorio.Banco = Banco_path
    relatorio.relatorio = App.Path & "\reports\ContasReceb.rpt"
    relatorio.titulo = "Contas à receber " & txt(0) & " até " & txt(1)
    relatorio.Formula = strFormula
    
    relatorio.Vizualizar
End If
End Sub

Private Sub cmdProc_Click(Index As Integer)
Select Case Index
    Case 0
        Cod_con = 40
        IM_consulta.Show 1
        'lblramo = Result_con
        txt(2) = txt(2) & CpoChave_con & ";"
End Select

End Sub

Private Sub Form_Activate()
Me.Top = 0
Me.Left = 0
End Sub

Private Sub Form_Resize()
Me.Width = 5000
Me.Height = 5000
End Sub

Function Consiste() As Boolean
Consiste = True
If IsDate(txt(0)) = False Then Consiste = False
If IsDate(txt(1)) = False Then Consiste = False

End Function

