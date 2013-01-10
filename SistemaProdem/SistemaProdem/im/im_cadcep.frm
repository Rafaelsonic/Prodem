VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form im_cadcep 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de CEP"
   ClientHeight    =   4500
   ClientLeft      =   375
   ClientTop       =   2160
   ClientWidth     =   9165
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1
   ScaleMode       =   0  'User
   ScaleWidth      =   0.927
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   3255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   7215
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   5
         Left            =   1770
         TabIndex        =   13
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   4
         Left            =   1770
         TabIndex        =   12
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   3
         Left            =   1770
         TabIndex        =   11
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   2
         Left            =   1770
         TabIndex        =   10
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   1
         Left            =   1770
         TabIndex        =   1
         Top             =   720
         Width           =   4455
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   0
         Left            =   1770
         MaxLength       =   8
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Abreviação"
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
         Left            =   360
         TabIndex        =   9
         Top             =   2340
         Width           =   960
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
         Index           =   4
         Left            =   360
         TabIndex        =   8
         Top             =   1980
         Width           =   1200
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
         Index           =   3
         Left            =   360
         TabIndex        =   7
         Top             =   1620
         Width           =   600
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
         Index           =   2
         Left            =   360
         TabIndex        =   6
         Top             =   1260
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
         Left            =   360
         TabIndex        =   5
         Top             =   780
         Width           =   495
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
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   300
         Width           =   345
      End
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   4125
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "im_cadcep"
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
Dim Exibir As Boolean ' serve para ver se o formulario vai ficar aberto
Dim Selecao As Boolean 'verifica se exite alguma seleção
Dim ocn_BancoCEP As New ADODB.Connection
Dim snap_Selecao As New ADODB.Recordset ' objeto da seleção
'Descricao: manutencao

'Banco utilizadas: ceps



Private Sub cmd_Click(Index As Integer)
Dim Sql As String
Dim snap_Cod As New ADODB.Recordset
On Error GoTo erro

Select Case Index
Case 0 ' botao inserir
    status.Panels(1).Text = "INSERIR"
    txt(0).Locked = True
    Botoes

Case 1 ' botao excluir
    'procura referencia
    'pede confirmação
    If status.Panels(1).Text = "ALTERAR" Then
        If MsgBox("Deseja realmente excluir ?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
            Sql = "delete from " & nome_tab(0) & " where ris_cod=" & txt(0)
            ocn_BancoCEP.Execute Sql
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
            
                            
                Sql = "insert into " & nome_tab(0) & Chr(13)
                Sql = Sql & "values(" & txt(1) & ",'" & txt(0) & "'," & txt(4) & ",'" & txt(3) & ",'" & txt(5) & ",'" & txt(2) & ",')"
                ocn_BancoCEP.Execute (Sql)
                status.Panels(1).Text = "ALTERAR"
                MsgBox "Os dados foram gravados com sucesso!", vbInformation, "Gravação"
                Botoes
            End If
        End If
    ElseIf status.Panels(1).Text = "ALTERAR" Then
        If conssiste = True Then
            Sql = "UPDATE " & nome_tab(0) & " SET " & cpo(1) & "='" & txt(1) & "', " & cpo(2) & "='" & txt(2) & "', " & cpo(3) & "='" & txt(3) & "', " & cpo(4) & "='" & txt(4) & "', " & cpo(5) & "='" & txt(5) & "'" & Chr(13)
            Sql = Sql & "WHERE " & cpo(0) & "=" & txt(0)
            ocn_BancoCEP.Execute Sql
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
End Select
Exit Sub
erro:
MsgBox Err.Description, vbCritical, "cmd_click"
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
Dim Sql As String
If Open_BancoCEP = True Then
    Declara
    On Error GoTo erro
    'pesquisa dados na tabela
    Sql = "select count(*) from " & nome_tab(0)
    Set snap_tab = ocn_BancoCEP.Execute(Sql)
    If snap_tab(0) <= 0 Then
    MsgBox "Não exixte dados na Tabela!", vbInformation, "Atenção"
    End If
    Exibir = True
End If
Exit Sub
erro:
MsgBox Err.Description, vbCritical, "FORM LOAD " & Me.Name
Exibir = False
End Sub

Sub Declara()
' declara todas as varíaveis de dados
nome_tab(0) = "CEP_SP"
cpo(0) = "cep"
cpo(1) = "nome"
cpo(2) = "bairro"
cpo(3) = "cidade"
cpo(4) = "complemento"
cpo(5) = "abrevi"
tcpo(0) = "TEXTO"
tcpo(1) = "TEXTO"
tcpo(2) = "TEXTO"
tcpo(3) = "TEXTO"
tcpo(4) = "TEXTO"
tcpo(5) = "TEXTO"
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
        
    txt(0).BackColor = &HC0FFFF
    txt(1).SetFocus
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
Dim x As Integer
For x = 0 To txt.Count - 1
    txt(x).Text = ""
Next x
txt(0).BackColor = &H80000005
End Sub

Function conssiste_total() As Boolean
Dim x As Integer
On Error GoTo erro
For x = 0 To txt.Count - 1
Select Case tcpo(x)
Case "NUMERO"
    If Len(Trim$(txt(x))) > 0 Then
        If IsNumeric(txt(x)) = False Then
            conssiste_total = False
            txt(x).SetFocus
            Exit Function
        End If
    End If
'Case "TEXTO"
    
Case "DATA"
    If Len(txt(x)) > 0 Then
        If IsDate(txt(x)) = False Then
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
        Sql = "select top 50 * from " & nome_tab(0)
        For x = 0 To txt.Count - 1
            If Len(txt(x)) > 0 Then
                If Len(sql_where) <= 0 Then
                    sql_where = " where "
                Else
                    sql_where = sql_where & " and " & Chr(13)
                End If
                sql_where = sql_where & Monta_SQL(cpo(x), tcpo(x), txt(x))
            End If
        Next x
    Sql = Sql & sql_where
    snap_Selecao.Open Sql, ocn_BancoCEP, adOpenKeyset, adLockOptimistic, adCmdText
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
Dim x As Integer
For x = 0 To txt.Count - 1
    If IsNull(snap_Selecao(cpo(x))) = False Then txt(x) = snap_Selecao(cpo(x))
Next x
End Sub

Private Sub Form_Resize()
'Me.Caption = Me.Height
Me.Height = 4760
Me.Width = 7620
End Sub

Function conssiste() As Boolean
conssiste = False
If Len(txt(1)) > 0 Then conssiste = True
If InStr(1, txt(1), "*") > 0 Then
    conssiste = False
    txt(1).SetFocus
End If

End Function

Private Sub Form_Unload(Cancel As Integer)
BOTDisable
End Sub

Private Sub txt_GotFocus(Index As Integer)
txt(Index).SelStart = 0
txt(Index).SelLength = Len(txt(Index))
End Sub
Sub Menu(Index As Integer)
Dim Sql As String
Dim snap_Cod As New ADODB.Recordset
On Error GoTo erro

Select Case Index
Case 0 ' botao inserir
    status.Panels(1).Text = "INSERIR"
    txt(0).Locked = True
    Botoes

Case 1 ' botao excluir
    'procura referencia
    'pede confirmação
    If status.Panels(1).Text = "ALTERAR" Then
        If MsgBox("Deseja realmente excluir ?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
            Sql = "delete from " & nome_tab(0) & " where ris_cod=" & txt(0)
            ocn_BancoCEP.Execute Sql
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
            
                Sql = "SELECT MAX(RIS_COD)+1 FROM " & nome_tab(0)
                Set snap_Cod = ocn_BancoCEP.Execute(Sql)
                If IsNull(snap_Cod(0)) Then
                    txt(0) = 1
                Else
                    txt(0) = snap_Cod(0)
                End If
                Set snap_Cod = Nothing
            
                Sql = "insert into " & nome_tab(0) & Chr(13)
                Sql = Sql & "values(" & txt(0) & ",'" & txt(1) & "')"
                ocn_BancoCEP.Execute (Sql)
                status.Panels(1).Text = "ALTERAR"
                MsgBox "Os dados foram gravados com sucesso!", vbInformation, "Gravação"
                Botoes
            End If
        End If
    ElseIf status.Panels(1).Text = "ALTERAR" Then
        If conssiste = True Then
            Sql = "UPDATE " & nome_tab(0) & " SET " & cpo(1) & "='" & txt(1) & "'" & Chr(13)
            Sql = Sql & "WHERE " & cpo(0) & "=" & txt(0)
            ocn_BancoCEP.Execute Sql
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
End Select
Exit Sub
erro:
MsgBox Err.Description, vbCritical, "menu"

End Sub


Function Open_BancoCEP() As Boolean
Set ocn_BancoCEP = Nothing

' procedimento de abertura do banco de dados
On Error GoTo erro
ocn_BancoCEP.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Banco_path_CEP

Open_BancoCEP = True
Exit Function
erro:
Open_BancoCEP = False
MsgBox Err.Description
End Function
