VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   2760
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Executar"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Alterar o orcçamento 18 para o cliente 70"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   5415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
ocnBanco.Execute "update pf_orcamentoH set orch_cli = 70 where orch_cod=18"
MsgBox "Dados alterados com sucesso!"
End Sub
