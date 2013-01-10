VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm mdi 
   BackColor       =   &H8000000C&
   Caption         =   "RFL - Systems"
   ClientHeight    =   3060
   ClientLeft      =   1320
   ClientTop       =   3390
   ClientWidth     =   4650
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog ctr 
      Left            =   2100
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   582
      ButtonWidth     =   2011
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Novo"
            Key             =   "Novo"
            Object.ToolTipText     =   "Novo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gravar"
            Key             =   "Gravar"
            Object.ToolTipText     =   "Gravar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excluir"
            Key             =   "Excluir"
            Object.ToolTipText     =   "Excluir"
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Selecionar"
            Key             =   "Selecionar"
            Object.ToolTipText     =   "Selecionar"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar"
            Key             =   "Cancelar"
            Object.ToolTipText     =   "Cancelar"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Primeiro"
            Key             =   "Primeiro"
            Object.ToolTipText     =   "Primeiro"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Anterior"
            Key             =   "Anterior"
            Object.ToolTipText     =   "Anterior"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Proximo"
            Key             =   "Proximo"
            Object.ToolTipText     =   "Proximo"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Último"
            Key             =   "ultimo"
            Object.ToolTipText     =   "Ultimo"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList 
      Index           =   0
      Left            =   630
      Top             =   1620
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":077A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":1054
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":1936
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":20B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":282A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":2FA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":371E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":3E98
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu0 
      Caption         =   "Cadastros"
      Index           =   0
      Begin VB.Menu mnucadcli 
         Caption         =   "Clientes"
         Index           =   5
      End
      Begin VB.Menu mnucadfornec 
         Caption         =   "Fornecedores"
         Index           =   22
      End
      Begin VB.Menu mnu_cadprod 
         Caption         =   "Produtos"
         Index           =   25
      End
      Begin VB.Menu mnuSedex 
         Caption         =   "Sedex"
      End
   End
   Begin VB.Menu mnuvendas 
      Caption         =   "Vendas"
      Begin VB.Menu mnu_orcamento 
         Caption         =   "Orçamentos"
      End
      Begin VB.Menu mnurelat 
         Caption         =   "Relatórios"
         Begin VB.Menu mnuOrcSituacao 
            Caption         =   "Status Orçamento"
         End
         Begin VB.Menu mnuRelReserva 
            Caption         =   "Reservas"
         End
         Begin VB.Menu mnuTreiReali 
            Caption         =   "Treinamentos Realizados"
         End
      End
   End
   Begin VB.Menu mnuuti 
      Caption         =   "Utilitários"
      Visible         =   0   'False
      Begin VB.Menu Backup 
         Caption         =   "Backup"
         Index           =   15
      End
      Begin VB.Menu compactar 
         Caption         =   "Compactar BD"
         Index           =   16
      End
   End
   Begin VB.Menu mnureceb 
      Caption         =   "Contas Receber"
      Begin VB.Menu mnulancReceb 
         Caption         =   "Lançamentos"
      End
      Begin VB.Menu mnuRel 
         Caption         =   "Relatórios"
         Begin VB.Menu mnuRelReceber 
            Caption         =   "Contas a Receber"
         End
      End
   End
End
Attribute VB_Name = "mdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim snap_Procura As New ADODB.Recordset ' objeto da seleção
Dim Sql As String


Private Sub CEP_Click(Index As Integer)
im_cadcep.Show
im_cadcep.SetFocus
End Sub




Private Sub MDIForm_Load()
BOTDisable
Login.Show (1)
OcultaMenus


End Sub

Sub OcultaMenus()
Dim oFuncionalidadeMenu() As String
Dim X As Integer

'mnu0.Count
'Utilizado para ocultar as outras opções para levar o sistema para o sitio
Sql = "select * from usuario where usuarioID=" & UsuarioID
Set snap_Procura = ocnBanco.Execute(Sql)
If snap_Procura.EOF And snap_Procura.BOF Then
    'nao encontrou
Else
    If snap_Procura("administrador") = True Then
        'todos acessos
    Else
        Sql = "select b.nome, a.funcionalidadeid, a.usuarioid from usuariofuncionalidade a, funcionalidade b where a.funcionalidadeid = b.funcionalidadeid and usuarioid = " & UsuarioID
         Set snap_Procura = Nothing
         Set snap_Procura = ocnBanco.Execute(Sql)
        
        'snap_Procura.Open SQL, ocnBanco, adOpenKeyset, adLockOptimistic, adCmdTable
        If Not (snap_Procura.EOF And snap_Procura.BOF) Then
            
            snap_Procura.MoveFirst
            
            'colocar o loop para habilitar os menus
            While Not snap_Procura.EOF
                If Not IsNull(snap_Procura("Nome")) Then
                    ReDim oFuncionalidadeMenu(X)
                    oFuncionalidadeMenu(X) = snap_Procura("Nome")
                    X = X + 1
                End If
                snap_Procura.MoveNext
            Wend
            
            
            '
            
            Dim dados As Object
            Dim mnuGlobal As Menu
            
            For Each dados In Me.Controls
                If TypeOf dados Is Menu Then
                    Set mnuGlobal = dados
                    'mnuGlobal.Name
                    Beep
                    'Procura na lista...
                    For Each Fun In oFuncionalidadeMenu
                        If mnuGlobal.Name = Fun Then
                            mnuGlobal.Visible = True
                        Else
                            'mnuGlobal.Visible = False
                        End If
                    Next Fun
                End If
            Next dados

            
        End If
    End If
End If
End Sub



Private Sub mnu_cadprod_Click(Index As Integer)
CadastroProduto.Show
CadastroProduto.SetFocus
End Sub



Private Sub mnu_orcamento_Click()
pf_Orcamento2.Show
pf_Orcamento2.SetFocus
End Sub


Private Sub mnucadcli_Click(Index As Integer)
pf_cadcli.Show
pf_cadcli.SetFocus
End Sub

Private Sub mnucadfornec_Click(Index As Integer)
pf_cadfornec.Show
pf_cadfornec.SetFocus
End Sub


Private Sub mnuconsulta_Click(Index As Integer)
IM_consulta.Show
IM_consulta.SetFocus
End Sub


Private Sub mnulancReceb_Click()
pf_lancReceb.Show
pf_lancReceb.SetFocus
End Sub


Private Sub mnuRelReceber_Click()
pf_relReceber.Show
pf_relReceber.SetFocus
End Sub



Private Sub mnusinc_Click()
On Error GoTo erro
    Shell App.Path & "\xml.exe"
Exit Sub
erro:
    MsgBox Err.Description, vbCritical, "mnuSinc"
End Sub



Private Sub mnuSedex_Click()
pf_sedex.Show
pf_sedex.SetFocus
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case LCase$(Button.Key)
Case "novo"
'FIXIT: Whenever possible replace ActiveForm or ActiveControl with an early-bound variable     FixIT90210ae-R1614-RCFE85
    ActiveForm.Menu 0 ' novo
Case "gravar"
'FIXIT: Whenever possible replace ActiveForm or ActiveControl with an early-bound variable     FixIT90210ae-R1614-RCFE85
    ActiveForm.Menu 2 'gravar
Case "excluir"
'FIXIT: Whenever possible replace ActiveForm or ActiveControl with an early-bound variable     FixIT90210ae-R1614-RCFE85
    ActiveForm.Menu 1 'excluir
Case "selecionar" ' selecionar
'FIXIT: Whenever possible replace ActiveForm or ActiveControl with an early-bound variable     FixIT90210ae-R1614-RCFE85
    ActiveForm.Menu 3 'selecionar
Case "separador" ' separador
Case "primeiro" ' primeiro
'FIXIT: Whenever possible replace ActiveForm or ActiveControl with an early-bound variable     FixIT90210ae-R1614-RCFE85
    ActiveForm.Menu 4
Case "anterior" 'anterior
'FIXIT: Whenever possible replace ActiveForm or ActiveControl with an early-bound variable     FixIT90210ae-R1614-RCFE85
    ActiveForm.Menu 5
Case "proximo" 'proximo
'FIXIT: Whenever possible replace ActiveForm or ActiveControl with an early-bound variable     FixIT90210ae-R1614-RCFE85
    ActiveForm.Menu 6
Case "ultimo" 'ultimo
'FIXIT: Whenever possible replace ActiveForm or ActiveControl with an early-bound variable     FixIT90210ae-R1614-RCFE85
    ActiveForm.Menu 7
Case "cancelar"
'FIXIT: Whenever possible replace ActiveForm or ActiveControl with an early-bound variable     FixIT90210ae-R1614-RCFE85
    ActiveForm.Menu 8
End Select
End Sub

