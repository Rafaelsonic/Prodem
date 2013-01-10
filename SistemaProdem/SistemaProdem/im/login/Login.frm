VERSION 5.00
Begin VB.Form Login 
   ClientHeight    =   2700
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5625
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd 
      Caption         =   "Cancelar"
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      Caption         =   "OK"
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   3120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   0
      Left            =   3120
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1875
      ScaleWidth      =   1875
      TabIndex        =   5
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label LBL 
      Caption         =   "Senha:"
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   7
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label LBL 
      Caption         =   "Usuario:"
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   6
      Top             =   600
      Width           =   735
   End
   Begin VB.Label LBL 
      AutoSize        =   -1  'True
      Caption         =   "Sistema desenvolvido"
      Height          =   195
      Index           =   0
      Left            =   2400
      TabIndex        =   4
      Top             =   2280
      Width           =   1545
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Rafael Implementacao da validação do usuario
Dim snap_Procura As New ADODB.Recordset ' objeto da seleção
Dim SQL As String

Private Sub cmd_Click(Index As Integer)
Select Case Index
Case 0
    Valida_User
Case 1
    End
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then End
If KeyAscii = 13 Then Valida_User
End Sub

Sub Valida_User()

SQL = "select usuarioid, login, nome, senha, administrador, ativo from usuario where ativo = True " & _
" and login='" & txt(0) & "' and senha='" & txt(1) & "'"
Set snap_Procura = ocnBanco.Execute(SQL)

If snap_Procura.EOF And snap_Procura.BOF Then
    MsgBox "Usuario não validado"
Else
    'buscar as funcionalidades
    UsuarioID = snap_Procura("usuarioID")
    Unload Me
End If

'SomenteSitio = False

End Sub
