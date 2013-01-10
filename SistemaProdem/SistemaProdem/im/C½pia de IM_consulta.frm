VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form IM_consulta 
   Caption         =   "Consulta de Dados"
   ClientHeight    =   4410
   ClientLeft      =   4395
   ClientTop       =   3765
   ClientWidth     =   7620
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   7620
   Begin VB.TextBox txt 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   1680
      Width           =   5415
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   2175
      Left            =   0
      TabIndex        =   3
      Top             =   2160
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   3836
      _Version        =   393216
      FixedCols       =   0
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancelar"
      Height          =   495
      Index           =   1
      Left            =   5640
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&OK"
      Height          =   495
      Index           =   0
      Left            =   5640
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label label 
      Caption         =   "Critério de Seleção"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1815
   End
End
Attribute VB_Name = "IM_consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim snap_Selecao As New ADODB.Recordset
Dim codigo As Integer 'codigo da seleção para teste
Dim sql_consulta As String
Const QtdeLinhas = "Top 20"
Dim cpos() As String 'campos do grid
Dim cpoPrin As String ' campo onde vai ser feit a pesquisa
Dim Tama() As String 'tamanho dos campos do grid
Dim Desc() As String ' descricao do campo
Dim cols_grid As Integer ' Quantidade de colunas

Private Sub cmd_Click(Index As Integer)
Select Case Index
Case 0
    Monta_grd
Case 1
    Unload Me
End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmd_Click (0)
End Sub

Private Sub Form_Load()
Dim SQL As String

codigo = 1
SQL = "select * from im_conH, im_conI where conh_cod = coni_codconh" & Chr(13)
SQL = SQL & " and conh_cod=" & codigo & Chr(13)
SQL = SQL & "order by coni_cod"
snap_Selecao.Open SQL, ocnBanco
If Not (snap_Selecao.EOF And snap_Selecao.BOF) Then ' caso tenha dados
    snap_Selecao.MoveFirst
    sql_consulta = snap_Selecao("conh_sql")
    sql_consulta = Mid(sql_consulta, 1, InStr(1, sql_consulta, " ", vbTextCompare)) & QtdeLinhas & Mid(sql_consulta, InStr(1, sql_consulta, " ", vbTextCompare))
    Me.Caption = snap_Selecao("conh_desc")
    cpoPrin = snap_Selecao("conh_cpocon")
    cols_grid = 0
    While Not snap_Selecao.EOF
        ReDim Preserve cpos(cols_grid)
        ReDim Preserve Tama(cols_grid)
        ReDim Preserve Desc(cols_grid)
        cpos(cols_grid) = snap_Selecao("coni_cpo")
        Tama(cols_grid) = snap_Selecao("coni_tam")
        Desc(cols_grid) = snap_Selecao("coni_desc")
        cols_grid = cols_grid + 1
        snap_Selecao.MoveNext
    Wend
    grid.Cols = cols_grid
    cols_grid = cols_grid - 1 'saber o numero de colunas
    Cab_grid
    Set snap_Selecao = Nothing
Else
    MsgBox "Não foi encontrado referencia para essa consulta", vbCritical, "form_load"
End If
Set snap_Selecao = Nothing


End Sub

Private Sub Form_Resize()
'IM_consulta.Caption = Me.Width & "," & Me.Height
Me.Width = 7695
Me.Height = 4980

End Sub

Sub Monta_grd()
Dim SQL As String
SQL = sql_consulta & Chr(13) & "where " & cpoPrin & " like '" & txt & "*'"
snap_Selecao.Open SQL, ocnBanco

'abre um objeto
'verifica se tem dados
'ja sei quais os campos
End Sub
Sub Cab_grid()
Dim x As Integer
'escrever na coluna de título
'definir tamanho das colunas
' colocar tudo dessa funcao no form load
grid.Row = 0
For x = 0 To cols_grid
grid.Col = x
If cpoPrin = cpos(x) Then grid.CellFontBold = True
grid.Text = Desc(x)
grid.ColWidth(x) = Tama(x)
Next x
End Sub
