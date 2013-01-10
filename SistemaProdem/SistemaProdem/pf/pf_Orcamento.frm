VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form pf_Orcamento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Orçamento / Vendas"
   ClientHeight    =   7125
   ClientLeft      =   630
   ClientTop       =   960
   ClientWidth     =   9180
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1.583
   ScaleMode       =   0  'User
   ScaleWidth      =   0.929
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6705
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   8565
      Begin VB.CommandButton cmd 
         Caption         =   "Novo"
         Height          =   375
         Index           =   0
         Left            =   3720
         TabIndex        =   10
         Top             =   600
         Width           =   855
      End
      Begin MSComDlg.CommonDialog dlg 
         Left            =   7560
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.ComboBox cboStatus 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "pf_Orcamento.frx":0000
         Left            =   4680
         List            =   "pf_Orcamento.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   1755
      End
      Begin VB.TextBox txt 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1410
         TabIndex        =   0
         Top             =   270
         Width           =   1545
      End
      Begin VB.TextBox txt 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1380
         TabIndex        =   1
         Top             =   660
         Width           =   1545
      End
      Begin VB.CommandButton cmdProc 
         Height          =   375
         Index           =   0
         Left            =   3180
         Picture         =   "pf_Orcamento.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   600
         Width           =   375
      End
      Begin TabDlg.SSTab SSTab 
         Height          =   5505
         Left            =   120
         TabIndex        =   11
         Top             =   1110
         Width           =   8235
         _ExtentX        =   14526
         _ExtentY        =   9710
         _Version        =   393216
         TabOrientation  =   1
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   5
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Header"
         TabPicture(0)   =   "pf_Orcamento.frx":0424
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "txt(10)"
         Tab(0).Control(1)=   "txt(9)"
         Tab(0).Control(2)=   "cmd(2)"
         Tab(0).Control(3)=   "txt(23)"
         Tab(0).Control(4)=   "txt(24)"
         Tab(0).Control(5)=   "Label5(0)"
         Tab(0).Control(6)=   "Label5(20)"
         Tab(0).Control(7)=   "Label5(23)"
         Tab(0).Control(8)=   "Label5(3)"
         Tab(0).Control(9)=   "Label5(28)"
         Tab(0).Control(10)=   "lbl(8)"
         Tab(0).Control(11)=   "Label5(15)"
         Tab(0).ControlCount=   12
         TabCaption(1)   =   "Itens"
         TabPicture(1)   =   "pf_Orcamento.frx":0440
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "lbl(9)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label5(27)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label5(16)"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Label5(11)"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Label5(10)"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "Label5(9)"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "Label5(7)"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "ado"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "grdItens"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "grid"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "txt(20)"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "cmd_aln(1)"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "cmd_aln(0)"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "cmdProc(6)"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "txt(18)"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).Control(15)=   "txt(17)"
         Tab(1).Control(15).Enabled=   0   'False
         Tab(1).Control(16)=   "txt(15)"
         Tab(1).Control(16).Enabled=   0   'False
         Tab(1).ControlCount=   17
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   15
            Left            =   1290
            TabIndex        =   31
            Top             =   720
            Width           =   1545
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   17
            Left            =   1200
            TabIndex        =   30
            Top             =   1140
            Width           =   1545
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   18
            Left            =   3540
            TabIndex        =   29
            Top             =   1140
            Width           =   1545
         End
         Begin VB.CommandButton cmdProc 
            Height          =   375
            Index           =   6
            Left            =   3660
            Picture         =   "pf_Orcamento.frx":045C
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton cmd_aln 
            Height          =   405
            Index           =   0
            Left            =   420
            Picture         =   "pf_Orcamento.frx":087C
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   1590
            Width           =   405
         End
         Begin VB.CommandButton cmd_aln 
            Height          =   405
            Index           =   1
            Left            =   870
            Picture         =   "pf_Orcamento.frx":0CE0
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   1590
            Width           =   375
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   20
            Left            =   1260
            TabIndex        =   25
            Top             =   240
            Width           =   1545
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   10
            Left            =   -73440
            TabIndex        =   18
            Top             =   90
            Width           =   1545
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   9
            Left            =   -73800
            TabIndex        =   16
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Vizualizar"
            Height          =   375
            Index           =   2
            Left            =   -68040
            TabIndex        =   14
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   23
            Left            =   -68520
            TabIndex        =   13
            Top             =   600
            Width           =   1545
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2805
            Index           =   24
            Left            =   -74760
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   2040
            Width           =   7545
         End
         Begin MSFlexGridLib.MSFlexGrid grid 
            Height          =   1815
            Left            =   240
            TabIndex        =   15
            Top             =   3060
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   3201
            _Version        =   393216
            Cols            =   7
         End
         Begin MSDataGridLib.DataGrid grdItens 
            Bindings        =   "pf_Orcamento.frx":1155
            CausesValidation=   0   'False
            Height          =   2055
            Left            =   240
            TabIndex        =   17
            Top             =   2760
            Visible         =   0   'False
            Width           =   7725
            _ExtentX        =   13626
            _ExtentY        =   3625
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   0   'False
            Appearance      =   0
            ColumnHeaders   =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            RowDividerStyle =   1
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Produtos"
            ColumnCount     =   7
            BeginProperty Column00 
               DataField       =   "orci_item"
               Caption         =   "Item"
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
               DataField       =   "orci_servcod"
               Caption         =   "Código"
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
            BeginProperty Column02 
               DataField       =   "orci_servdesc"
               Caption         =   "Descrição"
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
            BeginProperty Column03 
               DataField       =   "orci_servtexto"
               Caption         =   "Texto"
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
            BeginProperty Column04 
               DataField       =   "orci_servqtdep"
               Caption         =   "Qtde"
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
            BeginProperty Column05 
               DataField       =   "orci_servvrvndp"
               Caption         =   "Valor"
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
            BeginProperty Column06 
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
                  ColumnWidth     =   599,811
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   900,284
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   2505,26
               EndProperty
               BeginProperty Column03 
               EndProperty
               BeginProperty Column04 
               EndProperty
               BeginProperty Column05 
               EndProperty
               BeginProperty Column06 
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc ado 
            Height          =   330
            Left            =   210
            Top             =   4350
            Visible         =   0   'False
            Width           =   7785
            _ExtentX        =   13732
            _ExtentY        =   582
            ConnectMode     =   1
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
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
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
         Begin VB.Label Label5 
            Caption         =   "Observação"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   -74760
            TabIndex        =   39
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Codigo"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   120
            TabIndex        =   38
            Top             =   330
            Width           =   1995
         End
         Begin VB.Label Label5 
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
            Height          =   285
            Index           =   9
            Left            =   180
            TabIndex        =   37
            Top             =   810
            Width           =   1995
         End
         Begin VB.Label Label5 
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
            Height          =   285
            Index           =   10
            Left            =   150
            TabIndex        =   36
            Top             =   1140
            Width           =   1995
         End
         Begin VB.Label Label5 
            Caption         =   "Valor"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   11
            Left            =   2910
            TabIndex        =   35
            Top             =   1140
            Width           =   495
         End
         Begin VB.Label Label5 
            Caption         =   "[não pode alterar]"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   16
            Left            =   3210
            TabIndex        =   34
            Top             =   720
            Visible         =   0   'False
            Width           =   1995
         End
         Begin VB.Label Label5 
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   27
            Left            =   5460
            TabIndex        =   33
            Top             =   1140
            Width           =   495
         End
         Begin VB.Label lbl 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   9
            Left            =   6060
            TabIndex        =   32
            Top             =   1140
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "Data Emissão"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   20
            Left            =   -74790
            TabIndex        =   24
            Top             =   120
            Width           =   1275
         End
         Begin VB.Label Label5 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   23
            Left            =   -72720
            TabIndex        =   23
            Top             =   600
            Width           =   405
         End
         Begin VB.Label Label5 
            Caption         =   "Desconto"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   -74760
            TabIndex        =   22
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "Total Liq."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   28
            Left            =   -69480
            TabIndex        =   21
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lbl 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   8
            Left            =   -71160
            TabIndex        =   20
            Top             =   600
            Width           =   1515
         End
         Begin VB.Label Label5 
            Caption         =   "Total Bruto"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   15
            Left            =   -72240
            TabIndex        =   19
            Top             =   600
            Width           =   1155
         End
      End
      Begin VB.Label Label5 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   3960
         TabIndex        =   9
         Top             =   240
         Width           =   1995
      End
      Begin VB.Label lbl 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   4980
         TabIndex        =   7
         Tag             =   "1"
         Top             =   690
         Width           =   3285
      End
      Begin VB.Label Label2 
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Orçamento"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   300
         Width           =   1425
      End
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   6750
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "pf_Orcamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Data de criacao:21/08/2003
'Criador :Rafael
'Ultima atualizacao:         por
Option Explicit
Dim nome_tab(0) As String
Dim cpo(24) As String 'NOME DOS CAMPOS
Dim tcpo(24) As String 'TIPO DOS CAMPOS
Dim Exibir As Boolean ' serve para ver se o formulario vai ficar aberto
Dim Selecao As Boolean 'verifica se exite alguma seleção
Dim snap_Selecao As New ADODB.Recordset ' objeto da seleção
Dim snap_Procura As New ADODB.Recordset ' objeto para exibcao dos labels

Dim relatorio As New cls_Relatorio

Private Sub cmd_aln_Click(index As Integer)
Dim Sql As String
Dim cod

Select Case index

Case 0 ' incluir
    If IsNumeric(txt(20)) And IsNumeric(txt(17)) And IsNumeric(txt(18)) Then
        If val(txt(20)) > 0 And val(txt(17)) > 0 Then
            If Len(Trim(txt(15))) > 0 Then
                
                Inserir_Produto (txt(20)) 'inserir com código
                MsgBox "Produto cadastrado", vbInformation, "cmd_aln"
                'GridItens
                Mostra_Grid txt(0)
                'limpar campos do produto
                txt(20).SetFocus
            Else
                MsgBox "O campo descição deve ser preenchido.", , "cmd_aln"
                txt(15).SetFocus
            End If
        Else
            MsgBox "Os campos numéricos devem ser maior que zero", , "cmd_aln"
        End If
    Else
        MsgBox "Os alguns campos devem ser numéricos ", , "cmd_aln"
    End If
Case 1 'excluir
'excluir pelo numero do item
    If Len(lbl(7)) > 0 Then ' existe algum item selecionado....
        Sql = "delete from pf_orcamentoi where orci_orc=" & txt(0)
        Sql = Sql & " and orci_item=" & lbl(7)
        ocnBanco.Execute Sql
        MsgBox "Item excluído", vbInformation, "cmd_aln"
        'GridItens
        Mostra_Grid txt(0)
        Limpar_produto
    End If
    
Case 2
    Inserir_Produto txt(19), 1
    txt(19) = ""
    lbl(6) = ""
    'GridItens
    Mostra_Grid txt(0)
    SSTab.Tab = 1
End Select

'atualizar total
'atualizar desconto
Total_Orcamento
End Sub



Private Sub cmd_Click(index As Integer)
Dim Sql As String
Select Case index
Case 0 ' cadastrar cliente
    pf_cadOrcCli.Show (1)
    txt(1) = Me.Tag
    Me.Tag = 0
    txt(1).SetFocus

Case 1 'Enviar por email
    If IsNumeric(txt(1)) Then
        If Len(lbl(4)) > 0 Then
            Sql = "select cli_mail from pf_cliente where cli_cod= " & txt(1)
            Set snap_Procura = ocnBanco.Execute(Sql)
            If Not (snap_Procura.EOF And snap_Procura.BOF) Then
                If Len(snap_Procura(0)) > 0 Then
                    relatorio.relatorio = App.Path & "\Reports\orcamento.rpt" ' alterar quando gerar exe
                    relatorio.Formula = "{PF_ORCAMENTOH.ORCH_COD}=" & txt(0).Text
                    'relatorio.Formula = "{PF_ORCAMENTOH.ORCH_COD}=" & txt(0).Text & " and {@Teste} = False"
                    
                    relatorio.Banco = Banco_path
                    'relatorio.EMailPara = "rafael.sp@gmail.com" 'snap_Procura(0)
                    'relatorio.EMailTitulo = "Orçamento Nr." & txt(0)
                    'relatorio.EMailMenssasgem = "Sr. Cliente, " & Chr(13) & Chr(13) & "     Favor confirmar o recebimento desse e-mail, aprovando ou reprovando." & Chr(13) & Chr(13) & "Obrigado," & Chr(13) & Chr(13) & LBL(4) & Chr(13) & "Work Fire" & Chr(13) & "Departamento de Vendas" & Chr(13) & "+55 11 6468-4088" & Chr(13) & "" & Chr(13) & Chr(13) & Chr(13) & "Para visualizar esse pedido você necessita do Acrobat Reader. Faça o download gratuito: " & Chr(13) & "http://ardownload.adobe.com/pub/adobe/reader/win/6.x/6.0/ptb/AdbeRdr60_ptb_full.exe" & Chr(13) & Chr(13)
                    'relatorio.EMail
                    
                    relatorio.Export ("c:\temp\" & txt(0).Text & ".pdf")
        
                Else
                    MsgBox "O e-mail do Cliente está em branco."
                End If
            Else
                MsgBox "Não foi encontrado o cliente"
            End If
        Else
            txt(8).SetFocus
        End If
    Else
        txt(1).SetFocus
    End If

Case 2 'vizualizar
    relatorio.relatorio = App.Path & "\Reports\orcamento.rpt" ' alterar quando gerar exe
    'relatorio.Formula = "{PF_ORCAMENTOH.ORCH_COD}=" & txt(0).Text & " and {@Teste} = 'False'"
    relatorio.Formula = "{PF_ORCAMENTOH.ORCH_COD}=" & txt(0).Text
    relatorio.Banco = Banco_path
    relatorio.SenhaBanco = Banco_senha
    relatorio.Vizualizar
End Select
End Sub

'Descricao: manutencao


Private Sub cmdProc_Click(index As Integer)
Select Case index
    Case 0 ' cliente
        Cod_con = 40
        IM_consulta.Show 1
        lbl(0) = Result_con
        txt(1) = CpoChave_con
 
    
    
    
    Case 5 ' grupo de produtos
        Cod_con = 130
        IM_consulta.Show 1
        lbl(6) = Result_con
        txt(19) = CpoChave_con
    Case 6 ' produtos
        Cod_con = 120
        IM_consulta.Show 1
        txt(15) = Result_con
        txt(20) = CpoChave_con
        txt(20).SetFocus
   
    
End Select
End Sub


'Tabelas utilizadas: pf_ramativ






Private Sub Form_Activate()
Botoes
Me.Top = 0
Me.Left = 0
If Exibir = False Then Unload Me
End Sub

Private Sub Form_DblClick()
PreencheCombo
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Menu (8) ' esc
End Sub

Private Sub Form_Load()
Dim snap_tab As New ADODB.Recordset
Dim Sql As String
Declara

On Error GoTo erro


'sql = "select * from pf_orcamentoi where orci_orc=0 "
'ado.ConnectionString = ocnBanco.ConnectionString
'ado.RecordSource = sql
'ado.Refresh




'pesquisa dados na tabela
Sql = "select count(*) from " & nome_tab(0)
Set snap_tab = ocnBanco.Execute(Sql)
If snap_tab(0) <= 0 Then
MsgBox "Não exixte dados na Tabela!", vbInformation, "Atenção"
End If
'status.Panels(1).Text = "INSERIR"
'Botoes
Exibir = True
SSTab.Tab = 0
PreencheCombo
Exit Sub
erro:
MsgBox Err.Description, vbCritical, "FORM LOAD " & Me.Name
Exibir = False
'Unload Me

End Sub

Sub Declara()
' declara todas as varíaveis de dados
nome_tab(0) = "pf_OrcamentoH"

cpo(0) = "OrcH_cod"
cpo(1) = "OrcH_cli"
cpo(2) = "OrcH_dtval"
cpo(3) = "OrcH_qtdeParticip"
cpo(4) = "OrcH_texto"
cpo(5) = "OrcH_PrazoPgto"
cpo(6) = "OrcH_dttrei"
cpo(7) = "OrcH_obs"
cpo(8) = "OrcH_vende"
cpo(9) = "OrcH_desconto"
cpo(10) = "OrcH_dtemiss"
cpo(11) = "OrcH_comiss"
cpo(12) = "OrcH_emp"
cpo(13) = "OrcH_esclarecimento"

cpo(14) = "OrcH_Situacao"
cpo(15) = "OrcH_Periodo"
cpo(22) = "OrcH_turma"

cpo(24) = "OrcH_Histo"

tcpo(0) = "NUMERO"
tcpo(1) = "NUMERO"
tcpo(2) = "DATA"
tcpo(3) = "NUMERO"
tcpo(4) = "NUMERO"
tcpo(5) = "NUMERO"
tcpo(6) = "DATA"
tcpo(7) = "TEXTO"
tcpo(8) = "NUMERO"
tcpo(9) = "NUMERO"
tcpo(10) = "DATA"
tcpo(11) = "NUMERO"
tcpo(12) = "NUMERO"
tcpo(13) = "TEXTO"

tcpo(14) = "NUMERO"
tcpo(15) = "TEXTO"
tcpo(22) = "NUMERO"

tcpo(24) = "TEXTO"
End Sub

Sub Botoes() 'Bloqueia ou nao os botões
'cmd(1).Visible = False ' esconde botao do e-mail
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
    
    cboStatus.ListIndex = 0
    txt(10) = Date
    txt(2) = CDate(txt(10)) + 15
    cmd(0).Enabled = True
    cmd(1).Enabled = False
    cmd(2).Enabled = False
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
    
    cmd(0).Enabled = False
   
    cmd(2).Enabled = False
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

    cmd(0).Enabled = False
    cmd(1).Enabled = True ' botão do e-mail bloqueado para bloquear coloque false
    cmd(2).Enabled = True
    
    If cboStatus.ItemData(cboStatus.ListIndex) = 4 Then ' orcamento realizado
        mdi.Toolbar.Buttons(2).Enabled = False
        cmd_aln(0).Enabled = False
        cmd_aln(1).Enabled = False
        cmd_aln(2).Enabled = False
    Else
        mdi.Toolbar.Buttons(2).Enabled = True
        cmd_aln(0).Enabled = True
        cmd_aln(1).Enabled = True
        cmd_aln(2).Enabled = True
    End If

End Select
End Sub
Sub Limpar()
On Error Resume Next
Dim X As Integer
For X = 0 To txt.Count - 1
    txt(X).Text = ""
Next X
For X = 0 To lbl.Count - 1
    lbl(X).Caption = Empty
Next X
txt(0).BackColor = &H80000005
PreencheCombo

grid.Clear
'ado.RecordSource = ado.RecordSource & " AND 1=2"
'ado.Refresh
SSTab.Tab = 0
End Sub

Function conssiste_total() As Boolean
Dim X As Integer
On Error GoTo erro
'On Error Resume Next
For X = 0 To 13
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
    
If Len(Trim$(txt(22))) > 0 Then
    If IsNumeric(txt(22)) = False Then
        conssiste_total = False
        txt(22).SetFocus
        Exit Function
    End If
End If
    
Case "DATA"
    If Len(txt(X)) > 0 Then
        If IsDate(txt(X)) = False Then
            conssiste_total = False
            txt(X).SetFocus
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
        
        If Me.cboStatus.ListIndex >= 0 Then
            If Len(sql_where) <= 0 Then
                sql_where = " where "
            Else
                sql_where = sql_where & " and " & Chr(13)
            End If
            sql_where = sql_where & Monta_SQL("OrcH_Situacao", "numero", cboStatus.ItemData(cboStatus.ListIndex))
        End If
        
        
        
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
SSTab.Tab = 0
cboStatus.ListIndex = snap_Selecao(cpo(14)) - 1
Dim X As Integer
For X = 0 To 13
    If IsNull(snap_Selecao(cpo(X))) = False Then
        txt(X) = snap_Selecao(cpo(X))
    Else
        txt(X) = ""
    End If
Next X

txt(22) = snap_Selecao(cpo(22))
If IsNull(snap_Selecao(cpo(24))) = False Then
    txt(24) = snap_Selecao(cpo(24))
Else
    txt(24) = ""
End If
'GridItens
Mostra_Grid txt(0)
MostraRotulo (99)
Total_Orcamento
End Sub

Private Sub Form_Resize()
'Me.Caption = Me.Height & "   " & Me.Width
Me.Height = 7600
Me.Width = 9500
End Sub

Function conssiste() As Boolean
Dim X As Integer
conssiste = False
If Len(txt(1)) > 0 Then conssiste = True
If InStr(1, txt(1), "*") > 0 Then
    conssiste = False
    txt(1).SetFocus
End If

If Len(txt(3)) > 0 Then
    conssiste = True
    If InStr(1, txt(3), "*") > 0 Then
        conssiste = False
        txt(3).SetFocus
    End If
Else
    conssiste = False
    txt(3).SetFocus
    Exit Function
End If
If Len(txt(9)) > 0 Then
    If txt(9) < 100 Then
        conssiste = True
    Else
        conssiste = False
        txt(9).SetFocus
        Exit Function
    End If
Else
    txt(9) = 0
End If

If IsNumeric(txt(22)) = False Then
    conssiste = False
    txt(22).SetFocus
End If

If cboStatus.ListIndex = -1 Then
    conssiste = False
    cboStatus.SetFocus
End If

If status.Panels(1).Text = "INSERIR" Or status.Panels(1).Text = "ALTERAR" Then
    For X = 0 To 4
        If lbl(X) = "" Then
            conssiste = False
            txt(lbl(X).Tag).SetFocus
            Exit Function
        End If
    Next X
End If

If lbl(11) = "" Then
    conssiste = False
    txt(lbl(11).Tag).SetFocus
    Exit Function
End If
End Function

Private Sub Form_Unload(Cancel As Integer)
BOTDisable
End Sub


Private Sub grdItens_DblClick()
On Error GoTo erro

lbl(7) = grdItens.Columns(0).Text
txt(15) = grdItens.Columns(2).Text
txt(16) = grdItens.Columns(3).Text
txt(17) = grdItens.Columns(4).Text
txt(18) = grdItens.Columns(5).Text
txt(20) = grdItens.Columns(1).Text
SSTab.Tab = 3
Total_Item
Exit Sub
erro:
MsgBox Err.Description, vbCritical, "grdItens"
End Sub



Private Sub grid_DblClick()
On Error GoTo erro
With grid
    .col = 0
    lbl(7) = .Text
    .col = 2
    txt(15) = .Text
    .col = 6
    txt(16) = .Text
    .col = 3
    txt(17) = .Text
    .col = 4
    txt(18) = .Text
    .col = 1
    txt(20) = .Text
End With
SSTab.Tab = 3
Total_Item
Exit Sub
erro:
MsgBox Err.Description, vbCritical, "grdItens"
End Sub

Private Sub txt_GotFocus(index As Integer)
txt(index).SelStart = 0
txt(index).SelLength = Len(txt(index))
End Sub
Sub Menu(index As Integer)
Dim Sql As String
Dim snap_Cod As New ADODB.Recordset
On Error GoTo erro
Dim X As Integer 'variavel para loop

Select Case index
Case 0 ' botao inserir
    status.Panels(1).Text = "INSERIR"
    txt(0).Locked = True
    Botoes
    txt(9) = 0
    txt(14) = 0

Case 1 ' botao excluir
    'procura referencia
    'pede confirmação
    If status.Panels(1).Text = "ALTERAR" Then
        If MsgBox("Deseja realmente excluir ?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
            Sql = "delete from " & nome_tab(0) & " where " & cpo(0) & "=" & txt(0)
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
            
                Sql = "SELECT MAX(" & cpo(0) & ")+1 FROM " & nome_tab(0)
                Set snap_Cod = ocnBanco.Execute(Sql)
                If IsNull(snap_Cod(0)) Then
                    txt(0) = 1
                Else
                    txt(0) = snap_Cod(0)
                End If
                Set snap_Cod = Nothing
        'forma generica para inserir dados
                Sql = "insert into " & nome_tab(0) & Chr(13)
                Sql = Sql & "("
                For X = 0 To 14
                    Sql = Sql & cpo(X)
                    If X < 14 Then Sql = Sql & ","
                Next X
                Sql = Sql & "," & cpo(22) & "," & cpo(24) & ")" & vbCrLf
                
                Sql = Sql & "values("
                For X = 0 To 13
                    Select Case tcpo(X)
                        Case "TEXTO"
                            If Len(txt(X)) > 0 Then
                                Sql = Sql & "'" & txt(X) & "'"
                            Else
                                Sql = Sql & "null"
                            End If
                        Case "NUMERO"
                            If Len(txt(X)) > 0 Then
                                Sql = Sql & "'" & txt(X) & "'"
                            Else
                                Sql = Sql & "null"
                            End If
                        Case "DATA"
                            If Len(txt(X)) > 0 Then
                                Sql = Sql & "#" & Format(txt(X), "mm/dd/yyyy") & "#"
                            Else
                                Sql = Sql & "null"
                            End If
                    End Select
                    If X < 14 Then Sql = Sql & ","
                Next X
                Sql = Sql & cboStatus.ItemData(cboStatus.ListIndex) & vbCrLf
                Sql = Sql & "," & txt(22) & ",'" & txt(24) & "')" & vbCrLf
                                                
                ocnBanco.Execute (Sql)
                status.Panels(1).Text = "ALTERAR"
                MsgBox "Os dados foram gravados com sucesso!", vbInformation, "Gravação"
                Botoes
            End If
        End If
    ElseIf status.Panels(1).Text = "ALTERAR" Then 'alterar dados
        If conssiste = True Then
            
            Sql = "UPDATE " & nome_tab(0) & " SET " & Chr(13)
            
            ' atualização genérica
            For X = 1 To 13
                Sql = Sql & cpo(X) & "="
                    Select Case tcpo(X)
                        Case "TEXTO"
                            If Len(txt(X)) > 0 Then
                                Sql = Sql & "'" & txt(X) & "'"
                            Else
                                Sql = Sql & "null"
                            End If
                        Case "NUMERO"
                            If Len(txt(X)) > 0 Then
                                Sql = Sql & "'" & txt(X) & "'"
                            Else
                                Sql = Sql & "null"
                            End If
                        Case "DATA"
                            If Len(txt(X)) > 0 Then
                                Sql = Sql & "#" & Format(txt(X), "mm/dd/yyyy") & "#"
                            Else
                                Sql = Sql & "null"
                            End If
                    End Select
                    Sql = Sql & ","
                Next X
                
            Sql = Sql & cpo(14) & "=" & cboStatus.ItemData(cboStatus.ListIndex) & vbCrLf
            Sql = Sql & "," & cpo(22) & "=" & txt(22) & vbCrLf
            Sql = Sql & "," & cpo(24) & "='" & txt(24) & "'" & vbCrLf
            Sql = Sql & vbCrLf
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

Private Sub PreencheCombo()

'Combo de status
cboStatus.Clear
cboStatus.AddItem "1 - Pendente"
cboStatus.AddItem "2 - Negociando"
cboStatus.AddItem "3 - Aprovado"
cboStatus.AddItem "4 - Realizado"
cboStatus.AddItem "5 - Cancelado"
cboStatus.ItemData(0) = "1"
cboStatus.ItemData(1) = "2"
cboStatus.ItemData(2) = "3"
cboStatus.ItemData(3) = "4"
cboStatus.ItemData(4) = "5"

'Periodo
'cboPer.Clear
'cboPer.AddItem "Manha"
'cboPer.AddItem "Tarde"
'cboPer.AddItem "Integral"


'formataçao do grid itens

End Sub

Private Sub GridItens()
Dim Sql As String

'sql = "select orci_servcod as [Codigo], orci_servdesc as [Descrição] from pf_orcamentoi where orci_orc= " & txt(0)
Sql = "select * from pf_orcamentoi where orci_orc= " & txt(0)
ado.ConnectionString = ocnBanco.ConnectionString
ado.RecordSource = Sql
ado.Refresh


End Sub

Private Sub TXT_KeyPress(index As Integer, KeyAscii As Integer)
Select Case index
    Case 0
        If KeyAscii = 13 Then
            Menu (3)
            Menu (4)
        End If
    
End Select

End Sub

Private Sub txt_LostFocus(index As Integer)
MostraRotulo (index)
Select Case index

Case 9
    Desconto_Orcamento
Case 17, 18
    Total_Item
Case 22
    Total_Orcamento
Case 23
    Desconto_OrcamentoPreco
    txt(23) = FormatCurrency(txt(23), 2)
End Select
End Sub
Sub MostraRotulo(index As Integer)
Dim Sql As String
Select Case index
    Case 1 ' cliente
        If IsNumeric(txt(index)) Then
            Sql = "Select cli_rzsc from pf_cliente where cli_cod = " & txt(index)
            Set snap_Procura = ocnBanco.Execute(Sql)
            
            If snap_Procura.BOF And snap_Procura.EOF Then
                lbl(0) = ""
            Else
                snap_Procura.MoveFirst
                lbl(0) = snap_Procura(0)
            End If
            Set snap_Procura = Nothing
        Else
            lbl(0) = ""
        End If
        
        
    Case 12 ' Empresa
        If IsNumeric(txt(index)) Then
            Sql = "Select emp_fant from im_empresa where emp_cod = " & txt(index)
            Set snap_Procura = ocnBanco.Execute(Sql)
            If snap_Procura.BOF And snap_Procura.EOF Then
                lbl(1) = ""
            Else
                snap_Procura.MoveFirst
                lbl(1) = snap_Procura(0)
            End If
            Set snap_Procura = Nothing
        Else
            lbl(1) = ""
        End If
        
        
        
    Case 4 ' Texto
        If IsNumeric(txt(index)) Then
            Sql = "Select txt_desc from pf_orctexto where txt_cod = " & txt(index)
            Set snap_Procura = ocnBanco.Execute(Sql)
            If snap_Procura.BOF And snap_Procura.EOF Then
                lbl(2) = ""
            Else
                snap_Procura.MoveFirst
                lbl(2) = snap_Procura(0)
            End If
            Set snap_Procura = Nothing
        Else
            lbl(2) = ""
        End If
        
        
    Case 5 ' Prazo pagamento
        If IsNumeric(txt(index)) Then
            Sql = "Select pgto_desc from pf_prazopgto where pgto_cod = " & txt(index)
            Set snap_Procura = ocnBanco.Execute(Sql)
            If snap_Procura.BOF And snap_Procura.EOF Then
                lbl(3) = ""
            Else
                snap_Procura.MoveFirst
                lbl(3) = snap_Procura(0)
            End If
            Set snap_Procura = Nothing
        Else
            lbl(3) = ""
        End If
        
        
    Case 8 ' vendedor
        If IsNumeric(txt(index)) Then
            Sql = "Select vend_nome from pf_vendedor where vend_cod = " & txt(index)
            Set snap_Procura = ocnBanco.Execute(Sql)
            If snap_Procura.BOF And snap_Procura.EOF Then
                lbl(4) = ""
            Else
                snap_Procura.MoveFirst
                lbl(4) = snap_Procura(0)
            End If
            Set snap_Procura = Nothing
        Else
            lbl(4) = ""
        End If
        
    Case 13 ' Esclarecimento
        If IsNumeric(txt(index)) Then
            Sql = "Select esc_desc from pf_orcEsclarecimento where esc_cod = " & txt(index)
            Set snap_Procura = ocnBanco.Execute(Sql)
            If snap_Procura.BOF And snap_Procura.EOF Then
                lbl(11) = ""
            Else
                snap_Procura.MoveFirst
                lbl(11) = snap_Procura(0)
            End If
            Set snap_Procura = Nothing
        Else
            lbl(11) = ""
        End If
        
    Case 19 ' Lista de Produtos
        If IsNumeric(txt(index)) Then
            Sql = "Select listah_desc from pf_listaprodh where listah_cod = " & txt(index)
            Set snap_Procura = ocnBanco.Execute(Sql)
            If snap_Procura.BOF And snap_Procura.EOF Then
                lbl(6) = ""
            Else
                snap_Procura.MoveFirst
                lbl(6) = snap_Procura(0)
            End If
            Set snap_Procura = Nothing
        Else
            lbl(6) = ""
        End If
    
    Case 20 ' Produtos / Serviços
        If IsNumeric(txt(index)) Then
            Sql = "Select * from pf_orcservico where serv_cod = " & txt(index)
            Set snap_Procura = ocnBanco.Execute(Sql)
            If snap_Procura.BOF And snap_Procura.EOF Then
                
                txt(15) = ""
                txt(16) = ""
                txt(18) = ""
            Else
                snap_Procura.MoveFirst
                txt(20) = snap_Procura("serv_cod")
                txt(15) = snap_Procura("serv_desc")
                txt(16) = snap_Procura("serv_texto")
                txt(18) = snap_Procura("serv_vrVenda")
            End If
            Set snap_Procura = Nothing
        Else
            'lbl(6) = ""
        End If
    
    
    
    Case 99 'todos
    
                Sql = "Select cli_rzsc from pf_cliente where cli_cod = " & txt(1)
            Set snap_Procura = ocnBanco.Execute(Sql)
                snap_Procura.MoveFirst
                lbl(0) = snap_Procura(0)

            Sql = "Select emp_fant from im_empresa where emp_cod = " & txt(12)
            Set snap_Procura = ocnBanco.Execute(Sql)
                snap_Procura.MoveFirst
                lbl(1) = snap_Procura(0)

            Sql = "Select txt_desc from pf_orctexto where txt_cod = " & txt(4)
            Set snap_Procura = ocnBanco.Execute(Sql)
                snap_Procura.MoveFirst
                lbl(2) = snap_Procura(0)

            Sql = "Select pgto_desc from pf_prazopgto where pgto_cod = " & txt(5)
            Set snap_Procura = ocnBanco.Execute(Sql)
                snap_Procura.MoveFirst
                lbl(3) = snap_Procura(0)

            Sql = "Select vend_nome from pf_vendedor where vend_cod = " & txt(8)
            Set snap_Procura = ocnBanco.Execute(Sql)
                snap_Procura.MoveFirst
                lbl(4) = snap_Procura(0)

            Sql = "Select esc_desc from pf_orcEsclarecimento where esc_cod = " & txt(13)
            Set snap_Procura = ocnBanco.Execute(Sql)
            snap_Procura.MoveFirst
            lbl(11) = snap_Procura(0)
            
            
            
            Set snap_Procura = Nothing
End Select

End Sub

Sub Inserir_Produto(codigo As Integer, Optional tipo As Integer = 0)
On Error GoTo erro
Dim Item As Integer
Dim snap
Dim snap_item
Dim snap_Servico
Dim vrCusto As Double
Dim vrPrev As Double
Dim Sql As String
Dim Qtde As Double ' quantidade  do item
'verifico se o item a foi cadastrado
If Len(Trim(lbl(7))) = 0 Then
     'se tipo = 0 cadastro de um produto somente
    If tipo = 0 Then
        ' caso nao exista o item .......
        Sql = "select max(orcI_item)+1 from pf_orcamentoI where " & vbCrLf
        Sql = Sql & "orci_orc =" & txt(0) & "" ' codigo do orçamento
        snap_item = ocnBanco.Execute(Sql)
        If IsNull(snap_item(0)) Then
            Item = 1
        Else
            Item = snap_item(0)
        End If
        'seleciono o produto gravar o valor vr do produto no previsto
        Sql = "SELECT serv_cod, serv_vrVenda FROM pf_OrcServico" & vbCrLf
        Sql = Sql & "where serv_cod=" & codigo
        snap = ocnBanco.Execute(Sql)
    
   
        'se vim do botaão gravar.....
        Sql = "insert into pf_OrcamentoI (orci_orc, orci_item, orci_servcod, orci_servdesc, orci_servtexto, OrcI_ServQtdeP, orci_servVrvndP) values("
        Sql = Sql & "'" & txt(0) & "','" & Item & "','" & txt(20) & "','" & txt(15) & "','" & txt(16) & "','" & txt(17) & "','" & txt(18) & "')"
        ocnBanco.Execute Sql
        'adiciona ao grid
    Else
        Sql = "select * from pf_listaProdI where listaI_lista =" & codigo
        Set snap = ocnBanco.Execute(Sql)
        If Not (snap.EOF And snap.BOF) Then
            While Not snap.EOF
                'pega nr item
                Sql = "select max(orcI_item)+1 from pf_orcamentoI where " & vbCrLf
                Sql = Sql & "orci_orc =" & txt(0) & "" ' codigo do orçamento
                snap_item = ocnBanco.Execute(Sql)
                If IsNull(snap_item(0)) Then
                    Item = 1
                Else
                    Item = snap_item(0)
                End If
                
                'seleciono dados do produto
                Sql = "select * from pf_orcServico where serv_cod =" & snap("listai_prod")
                Set snap_Servico = ocnBanco.Execute(Sql)
                
                'verifico a quantidade a ser inserida
                If snap("listai_quant") = 0 Then
                    'pego o nr de participantes
                    Qtde = txt(3)
                Else
                    Qtde = snap("listai_quant")
                End If
                
                Sql = "insert into pf_orcamentoI ( orcI_orc, orci_item, orci_servcod, orci_servDesc, orci_servtexto, orci_servqtdep, orci_servVrvndp)" & vbCrLf
                Sql = Sql & "values (" & txt(0) & "," & Item & "," & snap_Servico("serv_cod") & ",'" & snap_Servico("serv_desc") & "','" & snap_Servico("serv_texto") & "','" & Qtde & "','" & snap_Servico("serv_vrVenda") & "')"
                ocnBanco.Execute Sql
                snap.MoveNext
                'se tipo = 1 cadastro de grupo
                'se vim do grupo.....
                'sql = xxxxxxxxxxx
            Wend
        Else
            'nao encontrou dados
        End If
    End If
Else ' o item ja existe
    Sql = "update pf_orcamentoi set orci_servcod = '" & txt(20) & "', orci_servdesc = '" & txt(15) & "', orci_servtexto = '" & txt(16) & "', orci_servqtdep = '" & txt(17) & "', orci_servvrvndp = '" & txt(18) & "' "
    Sql = Sql & "where orci_orc =" & txt(0) & " and orci_item = " & lbl(7)
    ocnBanco.Execute Sql
End If
Limpar_produto
Exit Sub
erro:
MsgBox Err.Description, vbCritical, "inserir_produto"
End Sub

Sub Limpar_produto()
lbl(7) = ""
lbl(9) = ""
txt(15) = ""
txt(16) = ""
txt(17) = ""
txt(18) = ""
txt(20) = ""
End Sub

Sub Total_Item()
If IsNumeric(txt(17)) And IsNumeric(txt(18)) Then
    lbl(9) = CDbl(txt(17)) * CDbl(txt(18))
Else
    lbl(9) = ""
End If

End Sub

Sub Total_Orcamento()
Dim vrTotal As Double
Dim Sql As String
If IsNumeric(txt(0)) Then
    If IsNumeric(txt(22)) = True Then
        Sql = "SELECT orci_orc, sum( orci_servvrvndp * orci_servqtdep ) as Total" & vbCrLf
        Sql = Sql & "From pf_OrcamentoI" & vbCrLf
        Sql = Sql & "where orci_orc=" & txt(0) & vbCrLf
        Sql = Sql & "group by orci_orc" & vbCrLf
        Set snap_Procura = ocnBanco.Execute(Sql)
        If Not (snap_Procura.BOF And snap_Procura.EOF) Then  ' existe dados
            vrTotal = snap_Procura("total") * Int(txt(22))
            lbl(8) = FormatCurrency(vrTotal, 2)
            Desconto_Orcamento
        Else
            lbl(8) = FormatCurrency(0, 2)
            txt(23) = FormatCurrency(0, 2)
        End If
    Else
        lbl(8) = 0
    End If
End If
End Sub

Sub Desconto_Orcamento()
Dim vrDesc As Double
If Len(lbl(8)) > 0 Then
    If IsNumeric(txt(9)) Then
        vrDesc = CDbl(lbl(8)) - (CDbl(lbl(8)) * (CDbl(txt(9)) / 100))
        txt(23) = FormatCurrency(vrDesc, 2)
    End If
End If
End Sub

Sub Desconto_OrcamentoPreco()
Dim vrDesc As Double
If Len(lbl(8)) > 0 Then
    If IsNumeric(txt(9)) Then
        vrDesc = 100 - (CDbl(txt(23) / CDbl(lbl(8)) * 100))
        txt(9) = FormatNumber(vrDesc, 5)
    End If
End If
End Sub


Sub Mostra_Grid(cod As Integer)
Set snap_Procura = Nothing
Dim Linha As Integer
Dim Sql As String
grid.Clear

With grid
    .Row = 0
    .col = 0
    .ColWidth(0) = 600
    .Text = "Item"
    .col = 1
    .ColWidth(1) = 900
    .Text = "Código"
    .col = 2
    .ColWidth(2) = 2500
    .Text = "Descrição"
    .col = 3
    .ColWidth(3) = 1500
    .Text = "Qtde"
    .col = 4
    .ColWidth(4) = 1500
    .Text = "Valor"
    .col = 5
    .ColWidth(5) = 1500
    .Text = "Total"
    .col = 6
    .Text = "Texto..."
End With
Sql = "select * from pf_orcamentoi where orci_orc = " & cod
snap_Procura.Open Sql, ocnBanco, adOpenKeyset, adLockOptimistic, adCmdText
If Not (snap_Procura.BOF And snap_Procura.EOF) Then
    snap_Procura.MoveFirst
    grid.Rows = snap_Procura.RecordCount + 1
    Linha = 1
    While Not snap_Procura.EOF
        With grid
            .Row = Linha
            .col = 0
            .Text = snap_Procura("orci_item")
            .col = 1
            .Text = snap_Procura("orci_servCod")
            .col = 2
            .Text = snap_Procura("orci_Servdesc")
            .col = 3
            .Text = snap_Procura("orci_servQtdeP")
            .col = 4
            .Text = snap_Procura("orci_servVrvndP")
            .col = 5
            .Text = CDbl(snap_Procura("orci_servQtdeP")) * CDbl(snap_Procura("orci_servVrvndP"))
            .col = 6
            .Text = snap_Procura("orci_servtexto")
        End With
        Linha = Linha + 1
        snap_Procura.MoveNext
    Wend
End If
End Sub


