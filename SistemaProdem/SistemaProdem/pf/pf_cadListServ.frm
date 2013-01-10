VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form pf_cadListServ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Listas Orçamentos Padrões"
   ClientHeight    =   7485
   ClientLeft      =   2085
   ClientTop       =   3105
   ClientWidth     =   9165
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1.663
   ScaleMode       =   0  'User
   ScaleWidth      =   0.927
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm 
      Caption         =   "Produtos / Serviços"
      Height          =   4695
      Left            =   180
      TabIndex        =   10
      Top             =   1530
      Width           =   7395
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   3
         Left            =   1440
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton cmdProc 
         Height          =   375
         Index           =   0
         Left            =   3210
         Picture         =   "pf_cadListServ.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   300
         Width           =   375
      End
      Begin VB.CommandButton cmdP 
         Height          =   420
         Index           =   0
         Left            =   6300
         Picture         =   "pf_cadListServ.frx":0420
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   390
         Width           =   420
      End
      Begin VB.CommandButton cmdP 
         Height          =   420
         Index           =   1
         Left            =   6780
         Picture         =   "pf_cadListServ.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   390
         Width           =   420
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   2
         Left            =   1500
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid grid 
         Height          =   3255
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   5741
         _Version        =   393216
         FixedCols       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade"
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
         TabIndex        =   15
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblProd 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3690
         TabIndex        =   14
         Top             =   360
         Width           =   2565
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
         Index           =   5
         Left            =   300
         TabIndex        =   11
         Top             =   360
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1305
      Left            =   150
      TabIndex        =   7
      Top             =   120
      Width           =   7425
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   0
         Top             =   180
         Width           =   1575
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   1
         Top             =   750
         Width           =   4455
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
         Index           =   6
         Left            =   360
         TabIndex        =   9
         Top             =   810
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
         Index           =   7
         Left            =   360
         TabIndex        =   8
         Top             =   240
         Width           =   600
      End
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   7110
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
Attribute VB_Name = "pf_cadListServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Data de criacao:21/08/2003
'Criador :Rafael
'Ultima atualizacao:         por
Option Explicit
Dim nome_tab(1) As String
Dim cpo(3) As String 'NOME DOS CAMPOS
Dim tcpo(3) As String 'TIPO DOS CAMPOS
Dim Exibir As Boolean ' serve para ver se o formulario vai ficar aberto
Dim Selecao As Boolean 'verifica se exite alguma seleção
Dim snap_Selecao As New ADODB.Recordset ' objeto da seleção
Dim snap_Procura As New ADODB.Recordset
'Descricao: manutencao

'Tabelas utilizadas: pf_ramativ


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
            Sql = "delete from " & nome_tab(0) & " where ram_cod=" & txt(0)
            ocnBanco.Execute Sql
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
            
                Sql = "SELECT MAX(ram_COD)+1 FROM " & nome_tab(0)
                Set snap_Cod = ocnBanco.Execute(Sql)
                If IsNull(snap_Cod(0)) Then
                    txt(0) = 1
                Else
                    txt(0) = snap_Cod(0)
                End If
                Set snap_Cod = Nothing
            
                Sql = "insert into " & nome_tab(0) & Chr(13)
                Sql = Sql & "values(" & txt(0) & ",'" & txt(1) & "')"
                ocnBanco.Execute (Sql)
                status.Panels(1).Text = "ALTERAR"
                MsgBox "Os dados foram gravados com sucesso!", vbInformation, "Gravação"
                Botoes
            End If
        End If
    ElseIf status.Panels(1).Text = "ALTERAR" Then
        If conssiste = True Then
            Sql = "UPDATE " & nome_tab(0) & " SET " & cpo(1) & "='" & txt(1) & "'" & Chr(13)
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
End Select
Exit Sub
erro:
MsgBox Err.Description, vbCritical, "cmd_click"
End Sub




Private Sub cmdP_Click(Index As Integer)
Select Case Index
    Case 0
        add_produto
    Case 1
        rem_produto
End Select
End Sub

Private Sub cmdProc_Click(Index As Integer)
Select Case Index
    Case 0
        Cod_con = 120
        IM_consulta.Show 1
        lblProd = Result_con
        txt(2) = CpoChave_con
End Select
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
Cabecalho
Exit Sub
erro:
MsgBox Err.Description, vbCritical, "FORM LOAD " & Me.Name
Exibir = False
'Unload Me

End Sub

Sub Declara()
' declara todas as varíaveis de dados
nome_tab(0) = "pf_listaprodH"
nome_tab(1) = "pf_listaprodI"
cpo(0) = "listah_cod"
cpo(1) = "listah_desc"
cpo(2) = "listai_prod"

tcpo(0) = "NUMERO"
tcpo(1) = "TEXTO"
tcpo(2) = "NUMERO"

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
    frm.Enabled = False
    
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
    frm.Enabled = True
    
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
    frm.Enabled = False
    
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
    frm.Enabled = True
End Select
End Sub
Sub Limpar()
Dim X As Integer
For X = 0 To txt.Count - 1
    txt(X).Text = ""
Next X
txt(0).BackColor = &H80000005
lblProd.Caption = ""
grid.Rows = 1
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
For X = 0 To 1
    txt(X) = snap_Selecao(cpo(X))
Next X
view_produto
End Sub

Private Sub Form_Resize()
'Me.Caption = Me.Height & "   " & Me.Width
Me.Height = 7200
Me.Width = 8000
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



Private Sub grid_DblClick()
If grid.Rows > 1 Then
    If grid.Row > 0 Then
        grid.col = 0
        txt(2).Text = grid.Text
        grid.col = 1
        lblProd.Caption = grid.Text
        grid.col = 2
        txt(3).Text = grid.Text
    End If
End If
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
            Sql = "delete from " & nome_tab(0) & " where " & cpo(0) & "=" & txt(0)
            ocnBanco.Execute Sql
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
            
                Sql = "SELECT MAX(" & cpo(0) & ")+1 FROM " & nome_tab(0)
                Set snap_Cod = ocnBanco.Execute(Sql)
                If IsNull(snap_Cod(0)) Then
                    txt(0) = 1
                Else
                    txt(0) = snap_Cod(0)
                End If
                Set snap_Cod = Nothing
            
                Sql = "insert into " & nome_tab(0) & Chr(13)
                Sql = Sql & "values(" & txt(0) & ",'" & txt(1) & "')"
                ocnBanco.Execute (Sql)
                status.Panels(1).Text = "ALTERAR"
                MsgBox "Os dados foram gravados com sucesso!", vbInformation, "Gravação"
                Botoes
            End If
        End If
    ElseIf status.Panels(1).Text = "ALTERAR" Then
        If conssiste = True Then
            Sql = "UPDATE " & nome_tab(0) & " SET " & cpo(1) & "='" & txt(1) & "'" & Chr(13)
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










Sub add_produto()
Dim Sql As String
On Error GoTo erro

'procurar se o aluno ja existe no grid
If lblProd.Caption <> "" Then
    If Len(txt(3)) > 0 And IsNumeric(txt(3)) Then
        Sql = "select " & cpo(2) & " from " & nome_tab(1) & vbCrLf & _
              "where listai_lista=" & txt(0) & " and listai_prod=" & txt(2)
        Set snap_Procura = Nothing
        snap_Procura.Open Sql, ocnBanco, adOpenKeyset, adLockReadOnly, adCmdText
        If snap_Procura.RecordCount = 0 Then
            Sql = "insert into " & nome_tab(1) & " (listaI_lista, listaI_prod, listaI_quant ) values(" & txt(0) & "," & txt(2) & ", '" & txt(3) & "')"
            ocnBanco.Execute (Sql)
            view_produto
            txt(2) = ""
            lblProd = ""
            txt(3) = ""
            txt(2).SetFocus
        Else
            MsgBox "Já existe este produto!", vbInformation, "add_produto"
        End If
    Else
        txt(3).SetFocus
    End If
End If
Exit Sub
erro:
MsgBox Err.Description, vbCritical, "add_aluno"
End Sub

Sub view_produto()
Dim Sql As String, Exame As String ' a veriavel exame é para dizer se esta aprovado ou nao
    Cabecalho ' cria o cabeçalho do grid
    Sql = "Select serv_cod, serv_desc, listai_quant from pf_orcservico,  pf_listaprodi" & vbCrLf
    Sql = Sql & "where listai_lista=" & txt(0) & vbCrLf
    Sql = Sql & " and listai_prod = serv_cod" & vbCrLf
        
    Set snap_Procura = Nothing
    snap_Procura.Open Sql, ocnBanco
    grid.Rows = 2
    grid.Row = 1
    grid.Cols = 3
    ' se tem aluma coisa ....
    If Not (snap_Procura.BOF And snap_Procura.EOF) Then snap_Procura.MoveFirst
    While Not snap_Procura.EOF
        grid.col = 0
        grid.Text = snap_Procura(0)
        grid.CellAlignment = 3
        grid.col = 1
        grid.Text = snap_Procura(1)
        grid.CellAlignment = 1
        grid.col = 2
        grid.Text = snap_Procura(2)
        grid.CellAlignment = 3
        grid.Rows = grid.Rows + 1
        grid.Row = grid.Rows - 1
        snap_Procura.MoveNext
    Wend
    grid.Rows = grid.Rows - 1
    Set snap_Procura = Nothing
End Sub
Sub rem_produto()
Dim Sql As String
On Error GoTo erro

'procurar se o aluno ja existe no grid
If lblProd.Caption <> "" Then
    
    Sql = "select " & cpo(2) & " from " & nome_tab(1) & vbCrLf & _
          "where listai_lista=" & txt(0) & " and listai_prod=" & txt(2)
    Set snap_Procura = Nothing
    snap_Procura.Open Sql, ocnBanco, adOpenKeyset, adLockReadOnly, adCmdText
    If snap_Procura.RecordCount > 0 Then
        Sql = "delete from " & nome_tab(1) & vbCrLf & _
              "where listai_lista=" & txt(0) & " and listai_prod=" & txt(2)
        ocnBanco.Execute (Sql)
        view_produto
        txt(2) = ""
        txt(3) = ""
        lblProd = ""
    Else
        MsgBox "Não existe este produto!", vbInformation, "add_produto"
    End If
End If
Exit Sub
erro:
MsgBox Err.Description, vbCritical, "add_aluno"
End Sub

Sub Cabecalho()
grid.Rows = 1
grid.Cols = 3
grid.Row = 0
grid.col = 0
grid.Text = "Código"
grid.col = 1
grid.Text = "Descrição"
grid.col = 2
grid.Text = "Quant."
grid.ColWidth(0) = 1000
grid.ColWidth(1) = 4800
grid.ColWidth(2) = 1100
End Sub

Function Exibe_Produto(cod As String) As String
Dim Sql As String
Exibe_Produto = ""
If IsNumeric(cod) Then
    Sql = "select serv_desc from pf_orcServico where serv_cod=" & cod
    Set snap_Procura = ocnBanco.Execute(Sql)
    If Not (snap_Procura.EOF And snap_Procura.BOF) Then
        Exibe_Produto = snap_Procura(0)
    End If
End If
End Function

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
Select Case Index
Case 2
    lblProd = Exibe_Produto(txt(Index))
End Select
End Sub
