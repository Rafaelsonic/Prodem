VERSION 5.00
Begin VB.Form pf_clietq 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Etiquetas"
   ClientHeight    =   3000
   ClientLeft      =   4500
   ClientTop       =   4770
   ClientWidth     =   6615
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   6615
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
      Left            =   990
      TabIndex        =   6
      Text            =   "1"
      Top             =   1380
      Width           =   975
   End
   Begin VB.CommandButton cmdProc 
      Height          =   375
      Index           =   0
      Left            =   5460
      Picture         =   "pf_clietq.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   300
      Width           =   405
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
      Index           =   0
      Left            =   1020
      TabIndex        =   2
      Top             =   330
      Width           =   4005
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Visualizar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   6225
   End
   Begin VB.Label lbl 
      Caption         =   "Etq. Iicial"
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
      Left            =   120
      TabIndex        =   5
      Top             =   1410
      Width           =   825
   End
   Begin VB.Label lbl 
      Caption         =   "Separe com ( ; )  o intervalo do código dos participantes"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   930
      TabIndex        =   3
      Top             =   750
      Width           =   4245
   End
   Begin VB.Label lbl 
      Caption         =   "Clientes"
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
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   825
   End
End
Attribute VB_Name = "pf_clietq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim relatorio As New cls_Relatorio


Private Sub cmd_Click()
On Error GoTo erro
If Consiste = True Then
    Dim StrCod() As String, cont As Integer, val As Integer, fim As Integer, Sql As String, x As Integer
    ' -- excliuir tabela
    DelTab
    
    ' -- criar tabela temporaria
    CriaTab
    
    ' -- colocar um loop de acordo com o text no máximo 10
        If txt(1).Text > 1 Then
            For x = 1 To txt(1).Text - 1
                    Sql = "insert into temp_EtiquetaCli" & vbCrLf
                    Sql = Sql & "(cli_rzsc) values('')"
                    ocnBanco.Execute Sql
            Next x
        End If
    
    
    If Len(txt(0).Text) = 0 Then ' exibir todos os clientes
        ' -- iserir os registos na tabela temporaria
        Sql = "insert into temp_EtiquetaCli" & vbCrLf
        Sql = Sql & "SELECT  cli_rzsc, cli_ende, cli_bairr, cli_cep, cli_cida, cli_uf, cli_cont FROM pf_cliente" & vbCrLf
        ocnBanco.Execute Sql
    Else
        
        'cria lista
        cont = 1
        fim = 1
        For val = 1 To Len(txt(0).Text)
            Do While Mid$(txt(0).Text, fim, 1) <> ";"
                If fim >= Len(txt(0).Text) Then
                    fim = Len(txt(0).Text) + 1
                    Exit Do
                End If
                fim = fim + 1
            Loop
            ReDim Preserve StrCod(cont)
            StrCod(cont) = Mid$(txt(0).Text, val, fim - val)
            val = fim
            fim = fim + 1
            cont = cont + 1
        Next val
        'criterios de seleção
        
        Sql = ""
        For val = 1 To cont - 1
            ' -- iserir os registos na tabela temporaria
            Sql = "insert into temp_EtiquetaCli" & vbCrLf
            Sql = Sql & "SELECT  cli_rzsc, cli_ende, cli_bairr, cli_cep, cli_cida, cli_uf, cli_cont FROM pf_cliente" & vbCrLf
            Sql = Sql & "where cli_cod=" & StrCod(val)
            ocnBanco.Execute Sql
        Next val
    End If
        
    Sql = Sql & ""
    
    
    'relatorio.relatorio = App.Path & "\reports\etiquetacli2.rpt"
    relatorio.relatorio = App.Path & "\reports\etiquetacli.rpt"
    relatorio.Banco = Banco_path
    relatorio.titulo = "Etiqueta de Clientes"
    'relatorio.Formula = "{pf_cliente.cli_status} ="
    relatorio.Vizualizar
    
End If
Exit Sub
erro:
MsgBox Err.Description, vbCritical, "cmd_click"
End Sub

Private Sub cmdProc_Click(Index As Integer)
Select Case Index
    Case 0
        Cod_con = 40
        IM_consulta.Show 1
        'lbl(0) = Result_con
        txt(0) = txt(0) & CpoChave_con & ";"
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me ' esc
End Sub

Function Consiste() As Boolean
Consiste = False
If IsNumeric(txt(1).Text) = True Then
    If txt(1).Text >= 1 And txt(1).Text <= 10 Then
    Consiste = True
    Else
    Consiste = False
    MsgBox "O campo etiqueta inicial deve estar entre 1 a 10." & vbCrLf & _
           "Caso o local não esteja no intervalo, inverta o papel da etiqueta !", vbInformation, "consiste"
    End If
Else
    Consiste = False
    MsgBox "O campo etiqueta inicial deve ser numérico", vbInformation, "consiste"
End If
End Function

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
End Sub

Sub DelTab()
Dim Sql As String
On Error Resume Next
Sql = "drop table temp_EtiquetaCli;"
ocnBanco.Execute Sql
End Sub
Sub CriaTab()
Dim Sql As String
On Error GoTo erro
Sql = "CREATE TABLE temp_EtiquetaCli (" & vbCrLf & _
       "cli_rzsc             VARCHAR(100)," & vbCrLf & _
       "cli_ende             VARCHAR(50)," & vbCrLf & _
       "cli_bairr            VARCHAR(50)," & vbCrLf & _
       "cli_cep              VARCHAR(50)," & vbCrLf & _
       "cli_cida             VARCHAR(50)," & vbCrLf & _
       "cli_uf               VARCHAR(50)," & vbCrLf & _
       "cli_cont VarChar(50)           );" & vbCrLf
ocnBanco.Execute Sql
Exit Sub
erro:
MsgBox Err.Description, vbCritical, "CriaTab"
End Sub
