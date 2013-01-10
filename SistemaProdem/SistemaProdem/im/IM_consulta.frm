VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form IM_consulta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Dados"
   ClientHeight    =   4410
   ClientLeft      =   2790
   ClientTop       =   3180
   ClientWidth     =   7620
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid grid 
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   3201
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc ado 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=db_w"
      OLEDBString     =   "DSN=db_w"
      OLEDBFile       =   ""
      DataSourceName  =   "db_wf"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *  from im_empresa"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txt 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   1680
      Width           =   5415
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
      TabIndex        =   3
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
Dim cpochave As String 'campo chave primaria
Dim Tama() As String 'tamanho dos campos do grid
Dim Desc() As String ' descricao do campo
Dim cols_grid As Integer ' Quantidade de colunas
Dim con As Control


Private Sub cmd_Click(Index As Integer)
Select Case Index
Case 0
    dados
Case 1
    CpoChave_con = ""
    Result_con = ""
    Unload Me
End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmd_Click (0) ' enter
If KeyAscii = 27 Then cmd_Click (1) ' esc
End Sub

Private Sub Form_Load()
ado.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Banco_path & ";Persist Security Info=False" & ";Jet OLEDB:Database Password=" & Banco_senha
Set grid.DataSource = ado
Dim Sql As String

codigo = Cod_con
Sql = "select * from im_conH where " & Chr(13)
Sql = Sql & "conh_cod=" & codigo & Chr(13)

'SQL = "select * from im_conH, im_conI where conh_cod = coni_codconh" & Chr(13)
'SQL = SQL & " and conh_cod=" & codigo & Chr(13)
'SQL = SQL & "order by coni_cod"
snap_Selecao.Open Sql, ocnBanco
If Not (snap_Selecao.EOF And snap_Selecao.BOF) Then ' caso tenha dados
    snap_Selecao.MoveFirst
    sql_consulta = snap_Selecao("conh_sql")
    sql_consulta = Mid$(sql_consulta, 1, InStr(1, sql_consulta, " ", vbTextCompare)) & QtdeLinhas & Mid$(sql_consulta, InStr(1, sql_consulta, " ", vbTextCompare))
    Me.Caption = snap_Selecao("conh_desc")
    cpoPrin = snap_Selecao("conh_cpocon")
    cpochave = snap_Selecao("conh_cpochave")
    'cols_grid = 0
    'While Not snap_Selecao.EOF
        'ReDim Preserve cpos(cols_grid)
        'ReDim Preserve Tama(cols_grid)
        'ReDim Preserve Desc(cols_grid)
        'cpos(cols_grid) = snap_Selecao("coni_cpo")
        'Tama(cols_grid) = snap_Selecao("coni_tam")
        'Desc(cols_grid) = snap_Selecao("coni_desc")
        'cols_grid = cols_grid + 1
        'snap_Selecao.MoveNext
    'Wend
    ado.RecordSource = sql_consulta
    ado.Refresh
    grid.Columns(1).Width = 5040
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
Dim Sql As String
If InStr(1, sql_consulta, "where") > 0 Then
    Sql = sql_consulta & Chr(13) & "and " & cpoPrin & " like '" & txt & "%'"
Else
    Sql = sql_consulta & Chr(13) & "where " & cpoPrin & " like '" & txt & "%'"
End If
ado.RecordSource = Sql
ado.Refresh

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
grid.col = x
'If cpoPrin = cpos(X) Then grid.
grid.Text = Desc(x)
'grid.ColWidth(X) = Tama(X)
Next x
End Sub

Private Sub grid_DblClick()
dados
End Sub

Private Sub txt_KeyUp(KeyCode As Integer, Shift As Integer)

'MsgBox KeyCode
On Error GoTo erro
If KeyCode = 40 Then grid.Row = grid.Row + 1: Exit Sub
If KeyCode = 38 Then grid.Row = grid.Row - 1: Exit Sub
Monta_grd
Exit Sub
erro:
    MsgBox Err.Description, vbCritical, "im_consulta _ keyup"
End Sub


Sub dados()
On Error GoTo erro

Dim col As Integer
For col = 0 To grid.VisibleCols - 1
    If grid.Columns(col).Caption = cpoPrin Then
        grid.col = col
        Result_con = grid.Text
        
    End If
    If grid.Columns(col).Caption = cpochave Then
        grid.col = col
        CpoChave_con = grid.Text
    End If
Next col
Unload Me
Exit Sub
erro:
MsgBox Err.Description, vbCritical, "dados"
End Sub

