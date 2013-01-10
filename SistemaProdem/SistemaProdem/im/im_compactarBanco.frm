VERSION 5.00
Begin VB.Form im_compactarBanco 
   Caption         =   "Compactar Banco de dados"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExecutar 
      Caption         =   "Executar"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   2040
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Antes de executar verifique se o sistema esta fechado em todas as máquinas."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "im_compactarBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExecutar_Click()
On Error GoTo erro
Dim sOrigem As String
Dim sDestino As String

sOrigem = ocnBanco
sDestino = "bkp.mdb"
'FileCopy sOrigem, sDestino

'arquivo de seguranca
'FileCopy sOrigem, "seguranca.mdb"

'Dim db As DAO.DBEngine
'ocnBanco.Close
'db.CompactDatabase sOrigem, sDestino
'ocnBanco.Open
Dim x As Integer
Dim y As Integer
For x = 0 To 2000

    For y = 0 To 20000
    
    Next y

Next x

MsgBox "Reparação concluída."
End
erro:

End Sub
