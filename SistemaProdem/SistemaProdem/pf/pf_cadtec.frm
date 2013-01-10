VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form pf_cadtec 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Técnicos"
   ClientHeight    =   4710
   ClientLeft      =   1815
   ClientTop       =   1995
   ClientWidth     =   7470
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1.046
   ScaleMode       =   0  'User
   ScaleWidth      =   0.756
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog common 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Selecione a assinatura"
      Filter          =   ".jpg"
   End
   Begin VB.Frame Frame2 
      Height          =   4005
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   7215
      Begin VB.CommandButton cmdFig 
         Caption         =   "Limpar"
         Height          =   345
         Index           =   0
         Left            =   5820
         TabIndex        =   14
         Top             =   2820
         Width           =   645
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   4
         Left            =   1560
         TabIndex        =   4
         Top             =   1800
         Width           =   4455
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   3
         Left            =   1560
         TabIndex        =   3
         Top             =   1380
         Width           =   4455
      End
      Begin VB.CommandButton cmdProc 
         Caption         =   "..."
         Height          =   255
         Left            =   5790
         TabIndex        =   5
         Top             =   2460
         Width           =   495
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   2
         Top             =   990
         Width           =   4455
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   1
         Top             =   600
         Width           =   4455
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Complemento"
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
         Top             =   1920
         Width           =   1200
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Fomação"
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
         TabIndex        =   12
         Top             =   1500
         Width           =   765
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1020
         Left            =   1560
         OLEDropMode     =   2  'Automatic
         Stretch         =   -1  'True
         Top             =   2460
         Width           =   3900
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Assinatura"
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
         Left            =   360
         TabIndex        =   11
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Cargo"
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
         TabIndex        =   10
         Top             =   1110
         Width           =   525
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
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
         TabIndex        =   9
         Top             =   720
         Width           =   495
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
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   600
      End
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   4335
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "pf_cadtec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Data de criacao:21/08/2003
'Criador :Rafael
'Ultima atualizacao:         por
Option Explicit
Dim nome_tab(0) As String
Dim cpo(5) As String 'NOME DOS CAMPOS
Dim tcpo(5) As String 'TIPO DOS CAMPOS
Dim Selecao As Boolean 'verifica se exite alguma seleção
Dim snap_Selecao As New ADODB.Recordset ' objeto da seleção
'Descricao: manutencao

'Tabelas utilizadas: pf_tecn


Private Sub cmd_Click(Index As Integer)
Dim Sql As String
Dim snap_Cod As New ADODB.Recordset
On Error GoTo erro

Select Case Index
Case 0 ' botao inserir
    Status.Panels(1).Text = "INSERIR"
    TXT(0).Locked = True
    Botoes

Case 1 ' botao excluir
    'procura referencia
    'pede confirmação
    If Status.Panels(1).Text = "ALTERAR" Then
        If MsgBox("Deseja realmente excluir ?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
            Sql = "delete from " & nome_tab(0) & " where tec_cod=" & TXT(0)
            ocnBanco.Execute Sql
            
            cmd_Click (8)
            'Mostra_Dados
        End If
    End If
        'apaga o registro atual
Case 2 ' botao gravar
    'atualiza os dados
    '
    If Status.Panels(1).Text = "INSERIR" Then
        If conssiste_total = True Then 'verifica a consistencia
            If conssiste = True Then
            
                Sql = "SELECT MAX(tec_COD)+1 FROM " & nome_tab(0)
                Set snap_Cod = ocnBanco.Execute(Sql)
                If IsNull(snap_Cod(0)) Then
                    TXT(0) = 1
                Else
                    TXT(0) = snap_Cod(0)
                End If
                Set snap_Cod = Nothing
            
                Sql = "insert into " & nome_tab(0) & Chr(13)
                Sql = Sql & "values(" & TXT(0) & ",'" & TXT(1) & "','" & TXT(2) & "','" & TXT(3) & "','" & TXT(4) & "',null)"
                ocnBanco.Execute (Sql)
                Status.Panels(1).Text = "ALTERAR"
                Gravar_img
                MsgBox "Os dados foram gravados com sucesso!", vbInformation, "Gravação"
                Botoes
            End If
        End If
    ElseIf Status.Panels(1).Text = "ALTERAR" Then
        If conssiste = True Then
            Sql = "UPDATE " & nome_tab(0) & " SET " & cpo(1) & "='" & TXT(1) & "'," & cpo(2) & "='" & TXT(2) & "'," & cpo(3) & "='" & TXT(3) & "'," & cpo(4) & "='" & TXT(4) & "'" & Chr(13)
            Sql = Sql & "WHERE " & cpo(0) & "=" & TXT(0)
            ocnBanco.Execute Sql
            Gravar_img
            MsgBox "Os dados foram atualizados com sucesso !", vbInformation, "Alteração"
        End If
    End If
Case 3 ' botao selecionar
    'coloca em modo seleção
    Status.Panels(1).Text = "SELECIONAR"
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
    If Status.Panels(1).Text = "" Then
        Unload Me
        Exit Sub
    End If
    Status.Panels(1).Text = ""
    Botoes
    Set snap_Selecao = Nothing ' fecha objeto de selecao
    TXT(0).Locked = False
End Select
Exit Sub
erro:
MsgBox Err.Description, vbCritical, "cmd_click"
End Sub

Private Sub cmdFig_Click(Index As Integer)
img.Picture = LoadPicture()
End Sub

Private Sub cmdProc_Click()
Dim PicName As String, pic As String
Common.InitDir = App.Path
'Set the Filters
  Common.Filter = "Gif Files (*.gif)|*.gif|Jpg Files" & _
               "(*.jpg)|*.jpg|Bmp Files (*.bmp)|*.bmp"
  ' Specify default filter
  Common.FilterIndex = 2
  'set starting Path
        'common.InitDir = Path1
  'Show the Open Dialog
  Common.ShowOpen
  'If Canceled is Pressed
  If Common.FileName = "" Then Exit Sub
  'Put the File Title in a Virable
  PicName = Common.FileTitle
  'put The FilePathName in a Virable
  pic = Common.FileName
  'Load the Image in the ImageBox
  img.Picture = LoadPicture(pic)


End Sub

Private Sub Form_Activate()
Botoes
Me.Top = 0
Me.Left = 0

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

Exit Sub
erro:
MsgBox Err.Description, vbCritical, "FORM LOAD " & Me.Name

'Unload Me

End Sub

Sub Declara()
' declara todas as varíaveis de dados
nome_tab(0) = "pf_tecn"
cpo(0) = "tec_cod"
cpo(1) = "tec_nome"
cpo(2) = "tec_cargo"
cpo(3) = "tec_formacao"
cpo(4) = "tec_complformacao"
cpo(5) = "tec_ass"
tcpo(0) = "NUMERO"
tcpo(1) = "TEXTO"
tcpo(2) = "TEXTO"
tcpo(3) = "TEXTO"
tcpo(4) = "TEXTO"
tcpo(5) = "IMAGEM"
End Sub

Sub Botoes() 'Bloqueia ou nao os botões
mdi.Toolbar.Buttons(5).Enabled = True
Select Case Status.Panels(1).Text
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
    
    
    TXT(0).BackColor = &HC0FFFF
    TXT(1).SetFocus
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
    
    TXT(0).Locked = False
    TXT(0).SetFocus
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
    
    TXT(0).Locked = True

End Select
End Sub
Sub Limpar()
Dim x As Integer
For x = 0 To TXT.Count - 1
    TXT(x).Text = ""
Next x
img.Picture = LoadPicture()
TXT(0).BackColor = &H80000005
End Sub

Function conssiste_total() As Boolean
Dim x As Integer
On Error GoTo erro
For x = 0 To TXT.Count - 1
Select Case tcpo(x)
Case "NUMERO"
    If Len(Trim$(TXT(x))) > 0 Then
        If IsNumeric(TXT(x)) = False Then
            conssiste_total = False
            TXT(x).SetFocus
            Exit Function
        End If
    End If
'Case "TEXTO"
    
Case "DATA"
    If Len(TXT(x)) > 0 Then
        If IsDate(TXT(x)) = False Then
            conssiste_total = False
            Exit Function
        End If
    End If
End Select
Next x
conssiste_total = True
Exit Function
erro:
MsgBox Err.Description, vbCritical, "CONSSISTE_TOTAL"

End Function

Sub Movimentacao(botao As String)
Dim Sql As String, sql_where As String
Dim x As Integer
If Selecao = True Then
    If conssiste_total = True Then ' procurar os dados
        Sql = "select * from " & nome_tab(0)
        For x = 0 To TXT.Count - 1
            If Len(TXT(x)) > 0 Then
                If Len(sql_where) <= 0 Then
                    sql_where = " where "
                Else
                    sql_where = sql_where & " and " & Chr(13)
                End If
                sql_where = sql_where & Monta_SQL(cpo(x), tcpo(x), TXT(x))
            End If
        Next x
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
        Status.Panels(1).Text = "ALTERAR"
        Botoes
    Else
        MsgBox "Não foi encontrado dados para a seleção!", vbInformation, "Movimentação"
    End If

End Sub

Sub Mostra_Dados()
Dim x As Integer
For x = 0 To TXT.Count - 1
    TXT(x) = snap_Selecao(cpo(x))
Next x
Exibir_img
End Sub

Private Sub Form_Resize()
'Me.Caption = Me.Height & "   " & Me.Width
Me.Height = 5100
Me.Width = 7635
End Sub

Function conssiste() As Boolean
conssiste = False
If Len(TXT(1)) > 0 Then conssiste = True
If InStr(1, TXT(1), "*") > 0 Then
    conssiste = False
    TXT(1).SetFocus
End If

End Function

Function Open_Img(campo As String) As String
Dim IngImageSize As Long
Dim Ingoffset As Long
Dim bytChunck() As Byte
Dim intFile As Integer
Dim strTempPic As String
Const conchunksize = 100
Dim strImage As String

strTempPic = "c:\temp\img.bmp"
If Len(Dir(strTempPic)) > 0 Then Kill strTempPic
intFile = FreeFile
Open strTempPic For Binary As #intFile
IngImageSize = snap_Selecao(campo).ActualSize
'Do While Ingoffset < IngImageSize
'    bytChunck = snap_Selecao(campo).GetChunk(conchunksize)
    bytChunck = snap_Selecao(campo).GetChunk(IngImageSize)
    Put #intFile, , bytChunck()
'    Ingoffset = Ingoffset + conchunksize
 '   If Ingoffset > IngImageSize Then
'    MsgBox ""
'    End If
'Loop
Close #intFile
End Function

Private Sub Form_Unload(Cancel As Integer)
BOTDisable
End Sub

Private Sub txt_GotFocus(Index As Integer)
TXT(Index).SelStart = 0
TXT(Index).SelLength = Len(TXT(Index))
End Sub
Sub Exibir_img()
Dim imag() As Byte, i As Long
If IsNull(snap_Selecao(5)) Then
    img.Picture = LoadPicture()
Else
    imag = snap_Selecao(5)
    Close #10
    Open "c:\a.jpg" For Binary As #10
        Put #10, , imag
    Close #10
    img.Picture = LoadPicture("c:\a.jpg")
    Kill "c:\a.jpg"
End If
End Sub
Sub Gravar_img()
Dim pic As String
Dim snap_Img As New ADODB.Recordset
Set snap_Img = Nothing
snap_Img.Open "select " & cpo(5) & " from " & nome_tab(0) & " where " & cpo(0) & "=" & TXT(0).Text, ocnBanco, adOpenKeyset, adLockOptimistic, adCmdText
pic = Common.FileName
Dim imag() As Byte, i As Long
Close #1
If Len(pic) > 0 Then
    Open pic For Binary As 1
    ReDim imag(LOF(1))
    Do While Not EOF(1)
    Get #1, , imag
    i = i + 1
    Loop
    Close #1
    snap_Img(cpo(5)) = imag
    snap_Img.Update
End If
End Sub
Sub Menu(Index As Integer)
Dim Sql As String
Dim snap_Cod As New ADODB.Recordset
On Error GoTo erro

Select Case Index
Case 0 ' botao inserir
    Status.Panels(1).Text = "INSERIR"
    TXT(0).Locked = True
    Botoes

Case 1 ' botao excluir
    'procura referencia
    'pede confirmação
    If Status.Panels(1).Text = "ALTERAR" Then
        If MsgBox("Deseja realmente excluir ?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
            Sql = "delete from " & nome_tab(0) & " where tec_cod=" & TXT(0)
            ocnBanco.Execute Sql
            
            cmd_Click (8)
            'Mostra_Dados
        End If
    End If
        'apaga o registro atual
Case 2 ' botao gravar
    'atualiza os dados
    '
    If Status.Panels(1).Text = "INSERIR" Then
        If conssiste_total = True Then 'verifica a consistencia
            If conssiste = True Then
            
                Sql = "SELECT MAX(tec_COD)+1 FROM " & nome_tab(0)
                Set snap_Cod = ocnBanco.Execute(Sql)
                If IsNull(snap_Cod(0)) Then
                    TXT(0) = 1
                Else
                    TXT(0) = snap_Cod(0)
                End If
                Set snap_Cod = Nothing
            
                Sql = "insert into " & nome_tab(0) & Chr(13)
                Sql = Sql & "values(" & TXT(0) & ",'" & TXT(1) & "','" & TXT(2) & "','" & TXT(3) & "','" & TXT(4) & "',null)"
                ocnBanco.Execute (Sql)
                Status.Panels(1).Text = "ALTERAR"
                Gravar_img
                MsgBox "Os dados foram gravados com sucesso!", vbInformation, "Gravação"
                Botoes
            End If
        End If
    ElseIf Status.Panels(1).Text = "ALTERAR" Then
        If conssiste = True Then
            Sql = "UPDATE " & nome_tab(0) & " SET " & cpo(1) & "='" & TXT(1) & "'," & cpo(2) & "='" & TXT(2) & "'," & cpo(3) & "='" & TXT(3) & "'," & cpo(4) & "='" & TXT(4) & "'" & Chr(13)
            Sql = Sql & "WHERE " & cpo(0) & "=" & TXT(0)
            ocnBanco.Execute Sql
            Gravar_img
            MsgBox "Os dados foram atualizados com sucesso !", vbInformation, "Alteração"
        End If
    End If
Case 3 ' botao selecionar
    'coloca em modo seleção
    Status.Panels(1).Text = "SELECIONAR"
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
    If Status.Panels(1).Text = "" Then
        Unload Me
        Exit Sub
    End If
    Status.Panels(1).Text = ""
    Botoes
    Set snap_Selecao = Nothing ' fecha objeto de selecao
    TXT(0).Locked = False
    TXT(0).SetFocus
End Select
Exit Sub
erro:
MsgBox Err.Description, vbCritical, "cmd_click"

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
