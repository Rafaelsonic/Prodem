VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form pf_cadPrazoPgto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro Prazo de Pagamento"
   ClientHeight    =   5115
   ClientLeft      =   2550
   ClientTop       =   3330
   ClientWidth     =   7095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1.137
   ScaleMode       =   0  'User
   ScaleWidth      =   0.718
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   4065
      Left            =   210
      TabIndex        =   3
      Top             =   240
      Width           =   6375
      Begin VB.Frame Frame1 
         Caption         =   " Quat. Dias "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   390
         TabIndex        =   8
         Top             =   1650
         Width           =   3135
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   12
            Left            =   2130
            TabIndex        =   28
            Top             =   1590
            Width           =   645
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   11
            Left            =   2130
            TabIndex        =   27
            Top             =   1260
            Width           =   645
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   10
            Left            =   2130
            TabIndex        =   26
            Top             =   930
            Width           =   645
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   9
            Left            =   2130
            TabIndex        =   25
            Top             =   600
            Width           =   645
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   8
            Left            =   2130
            TabIndex        =   24
            Top             =   270
            Width           =   645
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   7
            Left            =   570
            TabIndex        =   23
            Top             =   1620
            Width           =   645
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   6
            Left            =   570
            TabIndex        =   22
            Top             =   1290
            Width           =   645
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   5
            Left            =   570
            TabIndex        =   21
            Top             =   960
            Width           =   645
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   4
            Left            =   600
            TabIndex        =   20
            Top             =   630
            Width           =   645
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   3
            Left            =   570
            TabIndex        =   19
            Top             =   300
            Width           =   645
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "10"
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
            Index           =   12
            Left            =   1740
            TabIndex        =   18
            Top             =   1650
            Width           =   210
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "9"
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
            Index           =   11
            Left            =   1830
            TabIndex        =   17
            Top             =   1290
            Width           =   105
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "8"
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
            Index           =   10
            Left            =   1830
            TabIndex        =   16
            Top             =   960
            Width           =   105
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "7"
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
            Index           =   9
            Left            =   1860
            TabIndex        =   15
            Top             =   660
            Width           =   105
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "6"
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
            Index           =   8
            Left            =   1860
            TabIndex        =   14
            Top             =   330
            Width           =   105
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "5"
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
            Index           =   7
            Left            =   300
            TabIndex        =   13
            Top             =   1650
            Width           =   105
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "4"
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
            Index           =   6
            Left            =   300
            TabIndex        =   12
            Top             =   1320
            Width           =   105
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "3"
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
            Left            =   300
            TabIndex        =   11
            Top             =   990
            Width           =   105
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "2"
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
            Left            =   300
            TabIndex        =   10
            Top             =   660
            Width           =   105
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "1"
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
            Left            =   300
            TabIndex        =   9
            Top             =   360
            Width           =   255
         End
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   2
         Left            =   1530
         TabIndex        =   7
         Top             =   1050
         Width           =   1575
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   1
         Top             =   630
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
         Caption         =   "Nr. Parcelas"
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
         Top             =   1140
         Width           =   1035
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
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
         Top             =   750
         Width           =   840
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
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   600
      End
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   4740
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "pf_cadPrazoPgto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Data de criacao:21/08/2003
'Criador :Rafael
'Ultima atualizacao:         por
Option Explicit
Dim nome_tab(0) As String
Dim cpo(12) As String 'NOME DOS CAMPOS
Dim tcpo(12) As String 'TIPO DOS CAMPOS
Dim Exibir As Boolean ' serve para ver se o formulario vai ficar aberto
Dim Selecao As Boolean 'verifica se exite alguma seleção
Dim snap_Selecao As New ADODB.Recordset ' objeto da seleção
'Descricao: manutencao

'Tabelas utilizadas: pf_ramativ


Private Sub Form_Activate()
Botoes
Me.Top = 0
Me.Left = 0
If Exibir = False Then Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Menu (8) ' esc
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
Exibir = True
Exit Sub
erro:
MsgBox Err.Description, vbCritical, "FORM LOAD " & Me.Name
Exibir = False
'Unload Me

End Sub

Sub Declara()
' declara todas as varíaveis de dados
nome_tab(0) = "pf_PrazoPgto"
cpo(0) = "pgto_cod"
cpo(1) = "pgto_desc"
cpo(2) = "pgto_nrParc"
cpo(3) = "pgto_n1"
cpo(4) = "pgto_n2"
cpo(5) = "pgto_n3"
cpo(6) = "pgto_n4"
cpo(7) = "pgto_n5"
cpo(8) = "pgto_n6"
cpo(9) = "pgto_n7"
cpo(10) = "pgto_n8"
cpo(11) = "pgto_n9"
cpo(12) = "pgto_n10"
tcpo(0) = "NUMERO"
tcpo(1) = "TEXTO"
tcpo(2) = "NUMERO"
tcpo(3) = "NUMERO"
tcpo(4) = "NUMERO"
tcpo(5) = "NUMERO"
tcpo(6) = "NUMERO"
tcpo(7) = "NUMERO"
tcpo(8) = "NUMERO"
tcpo(9) = "NUMERO"
tcpo(10) = "NUMERO"
tcpo(11) = "NUMERO"
tcpo(12) = "NUMERO"

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

    
    txt(1).SetFocus
    txt(0).BackColor = &HC0FFFF
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
Dim X As Integer
For X = 0 To txt.Count - 1
    txt(X).Text = ""
Next X
txt(0).BackColor = &H80000005
End Sub

Function conssiste_total() As Boolean
Dim X As Integer
On Error GoTo erro
For X = 0 To txt.Count - 1
Select Case tcpo(X)
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
    End If
End Select
Next X
conssiste_total = True
Exit Function
erro:
MsgBox Err.Description, vbCritical, "CONSSISTE_TOTAL"

End Function

Sub Movimentacao(botao As String)
Dim Sql As String, sql_where As String
Dim X As Integer
If Selecao = True Then
    If conssiste_total = True Then ' procurar os dados
        Sql = "select * from " & nome_tab(0)
        For X = 0 To txt.Count - 1
            If Len(txt(X)) > 0 Then
                If Len(sql_where) <= 0 Then
                    sql_where = " where "
                Else
                    sql_where = sql_where & " and " & Chr(13)
                End If
                sql_where = sql_where & Monta_SQL(cpo(X), tcpo(X), txt(X))
            End If
        Next X
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
        status.Panels(1).Text = "ALTERAR"
        Botoes
    Else
        MsgBox "Não foi encontrado dados para a seleção!", vbInformation, "Movimentação"
    End If

End Sub

Sub Mostra_Dados()
Dim X As Integer
For X = 0 To txt.Count - 1
    If IsNull(snap_Selecao(cpo(X))) = False Then txt(X) = snap_Selecao(cpo(X))
Next X
Trava_txt
End Sub

Private Sub Form_Resize()
'Me.Caption = Me.Height & "   " & Me.Width
Me.Height = 5760
Me.Width = 7620
End Sub

Function conssiste() As Boolean
conssiste = False
If Len(txt(1)) > 0 Then conssiste = True
If InStr(1, txt(1), "*") > 0 Then
    conssiste = False
    txt(1).SetFocus
End If
Dim X As Integer
For X = 3 To 12
    If Len(txt(X)) = 0 Then txt(X) = 0
Next X
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
            Sql = "delete from " & nome_tab(0) & " where pgto_cod=" & txt(0)
            ocnBanco.Execute Sql
            Menu (8)
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
            
                Sql = "SELECT MAX(pgto_COD)+1 FROM " & nome_tab(0)
                Set snap_Cod = ocnBanco.Execute(Sql)
                If IsNull(snap_Cod(0)) Then
                    txt(0) = 1
                Else
                    txt(0) = snap_Cod(0)
                End If
                Set snap_Cod = Nothing
            
                Sql = "insert into " & nome_tab(0) & Chr(13)
                Sql = Sql & "values(" & txt(0) & ",'" & txt(1) & "'," & txt(2)
                Sql = Sql & ", " & txt(3) & ", " & txt(4) & ", " & txt(5) & ", " & txt(6) & ", " & txt(7) & ", " & txt(8)
                Sql = Sql & ", " & txt(9) & ", " & txt(10) & ", " & txt(11) & ", " & txt(12) & " )"
                ocnBanco.Execute (Sql)
                status.Panels(1).Text = "ALTERAR"
                MsgBox "Os dados foram gravados com sucesso!", vbInformation, "Gravação"
                Botoes
            End If
        End If
    ElseIf status.Panels(1).Text = "ALTERAR" Then
        If conssiste = True Then
            Sql = "UPDATE " & nome_tab(0) & " SET " & Chr(13)
            Dim X As Integer
            For X = 1 To 11
                If tcpo(X) = "NUMERO" Then
                    Sql = Sql & cpo(X) & "=" & txt(X) & " ,"
                Else
                    Sql = Sql & cpo(X) & "='" & txt(X) & "' ,"
                End If
            Next X
            Sql = Sql & cpo(12) & "=" & txt(12) & Chr(13)
            
            
            Sql = Sql & "WHERE " & cpo(0) & "=" & txt(0)
            ocnBanco.Execute Sql
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
    txt(0).SetFocus
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

Private Sub txt_LostFocus(Index As Integer)
Trava_txt
End Sub

Private Sub Trava_txt()
Dim X As Integer
If IsNumeric(txt(2).Text) Then
    If txt(2).Text > 10 Then txt(2).Text = 10
    For X = 3 To 12
        If X <= txt(2) + 2 Then
            'txt(x) = "s"
            txt(X).Enabled = True
        Else
            'txt(x) = "n"
            txt(X).Enabled = False
        End If
    Next X
End If


End Sub
