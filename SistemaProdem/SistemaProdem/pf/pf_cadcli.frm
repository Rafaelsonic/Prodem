VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form pf_cadcli 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   6765
   ClientLeft      =   1920
   ClientTop       =   2880
   ClientWidth     =   8685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1.504
   ScaleMode       =   0  'User
   ScaleWidth      =   0.879
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog Common 
      Left            =   8040
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   1095
      Left            =   7920
      TabIndex        =   45
      Top             =   480
      Width           =   1095
      ExtentX         =   1931
      ExtentY         =   1931
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Frame Frame2 
      Height          =   6285
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   7665
      Begin TabDlg.SSTab SSTab 
         Height          =   5985
         Left            =   120
         TabIndex        =   18
         Top             =   180
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   10557
         _Version        =   393216
         TabOrientation  =   1
         Tabs            =   2
         Tab             =   1
         TabHeight       =   520
         TabCaption(0)   =   "Dados Gerais"
         TabPicture(0)   =   "pf_cadcli.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "lblramo"
         Tab(0).Control(1)=   "label(16)"
         Tab(0).Control(2)=   "label(15)"
         Tab(0).Control(3)=   "label(14)"
         Tab(0).Control(4)=   "label(13)"
         Tab(0).Control(5)=   "label(12)"
         Tab(0).Control(6)=   "label(11)"
         Tab(0).Control(7)=   "label(10)"
         Tab(0).Control(8)=   "label(9)"
         Tab(0).Control(9)=   "label(8)"
         Tab(0).Control(10)=   "label(7)"
         Tab(0).Control(11)=   "label(6)"
         Tab(0).Control(12)=   "label(5)"
         Tab(0).Control(13)=   "label(4)"
         Tab(0).Control(14)=   "label(3)"
         Tab(0).Control(15)=   "label(2)"
         Tab(0).Control(16)=   "label(1)"
         Tab(0).Control(17)=   "label(0)"
         Tab(0).Control(18)=   "label(31)"
         Tab(0).Control(19)=   "label(17)"
         Tab(0).Control(20)=   "cmdProc(0)"
         Tab(0).Control(21)=   "txt(16)"
         Tab(0).Control(22)=   "txt(15)"
         Tab(0).Control(23)=   "txt(14)"
         Tab(0).Control(24)=   "txt(13)"
         Tab(0).Control(25)=   "txt(12)"
         Tab(0).Control(26)=   "txt(11)"
         Tab(0).Control(27)=   "txt(10)"
         Tab(0).Control(28)=   "txt(9)"
         Tab(0).Control(29)=   "txt(8)"
         Tab(0).Control(30)=   "txt(7)"
         Tab(0).Control(31)=   "txt(6)"
         Tab(0).Control(32)=   "txt(5)"
         Tab(0).Control(33)=   "txt(4)"
         Tab(0).Control(34)=   "txt(3)"
         Tab(0).Control(35)=   "txt(2)"
         Tab(0).Control(36)=   "txt(1)"
         Tab(0).Control(37)=   "txt(0)"
         Tab(0).Control(38)=   "cboClienteTipo"
         Tab(0).Control(39)=   "cboListaPrecoPadrao"
         Tab(0).ControlCount=   40
         TabCaption(1)   =   "Financeiro / Ref"
         TabPicture(1)   =   "pf_cadcli.frx":001C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "label(18)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "label(19)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "label(20)"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "label(21)"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "label(22)"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "label(23)"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "Line1"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "label(24)"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "label(25)"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "lblDUPagamento"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "lblSaldo"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "txt(17)"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "txt(18)"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "txt(19)"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "txt(20)"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).Control(15)=   "txt(21)"
         Tab(1).Control(15).Enabled=   0   'False
         Tab(1).ControlCount=   16
         Begin VB.ComboBox cboListaPrecoPadrao 
            Height          =   315
            ItemData        =   "pf_cadcli.frx":0038
            Left            =   -72600
            List            =   "pf_cadcli.frx":005A
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   4800
            Width           =   735
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   21
            Left            =   5040
            TabIndex        =   52
            Top             =   1560
            Width           =   1935
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   20
            Left            =   5040
            TabIndex        =   51
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   19
            Left            =   1080
            TabIndex        =   50
            Top             =   1560
            Width           =   2895
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   18
            Left            =   1080
            TabIndex        =   49
            Top             =   1080
            Width           =   2895
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   17
            Left            =   1080
            TabIndex        =   48
            Top             =   600
            Width           =   5895
         End
         Begin VB.ComboBox cboClienteTipo 
            Height          =   315
            Left            =   -72600
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   4320
            Width           =   2535
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   0
            Left            =   -73650
            TabIndex        =   0
            Top             =   120
            Width           =   1575
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   1
            Left            =   -73650
            MaxLength       =   14
            TabIndex        =   1
            Top             =   480
            Width           =   2175
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   2
            Left            =   -70290
            MaxLength       =   13
            TabIndex        =   2
            Top             =   480
            Width           =   2175
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   3
            Left            =   -73650
            TabIndex        =   3
            Top             =   810
            Width           =   5535
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   4
            Left            =   -73650
            TabIndex        =   4
            Top             =   1140
            Width           =   5535
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   5
            Left            =   -72930
            TabIndex        =   5
            Top             =   1530
            Width           =   975
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   6
            Left            =   -73650
            TabIndex        =   8
            Top             =   2190
            Width           =   5475
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   7
            Left            =   -73650
            TabIndex        =   9
            Top             =   2550
            Width           =   2295
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   8
            Left            =   -73650
            MaxLength       =   8
            TabIndex        =   6
            Top             =   1860
            Width           =   1575
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   9
            Left            =   -69780
            TabIndex        =   10
            Top             =   2550
            Width           =   1575
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   10
            Left            =   -69720
            MaxLength       =   2
            TabIndex        =   7
            Top             =   1830
            Width           =   495
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   11
            Left            =   -73650
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   2880
            Width           =   1575
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   12
            Left            =   -69810
            TabIndex        =   12
            Top             =   2880
            Width           =   1575
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   13
            Left            =   -73650
            TabIndex        =   13
            Top             =   3210
            Width           =   1575
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   14
            Left            =   -69810
            TabIndex        =   14
            Top             =   3210
            Width           =   1575
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   15
            Left            =   -73650
            TabIndex        =   15
            Top             =   3540
            Width           =   3975
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   16
            Left            =   -73650
            TabIndex        =   16
            Top             =   3870
            Width           =   3975
         End
         Begin VB.CommandButton cmdProc 
            Height          =   375
            Index           =   0
            Left            =   -71850
            Picture         =   "pf_cadcli.frx":007D
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label lblSaldo 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   840
            TabIndex        =   58
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label lblDUPagamento 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2400
            TabIndex        =   57
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Saldo"
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
            Index           =   25
            Left            =   120
            TabIndex        =   56
            Top             =   2400
            Width           =   480
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Data Ultimo Pagamento"
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
            Index           =   24
            Left            =   120
            TabIndex        =   55
            Top             =   2040
            Width           =   2025
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Tabela Preço Padrão"
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
            Index           =   17
            Left            =   -74880
            TabIndex        =   53
            Top             =   4800
            Width           =   1770
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cliente"
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
            Index           =   31
            Left            =   -74880
            TabIndex        =   46
            Top             =   4320
            Width           =   1020
         End
         Begin VB.Line Line1 
            X1              =   240
            X2              =   7080
            Y1              =   3720
            Y2              =   3720
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Dados para Cobrança"
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
            Index           =   23
            Left            =   150
            TabIndex        =   44
            Top             =   180
            Width           =   1875
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "UF"
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
            Index           =   22
            Left            =   4470
            TabIndex        =   43
            Top             =   1530
            Width           =   210
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
            Index           =   21
            Left            =   150
            TabIndex        =   42
            Top             =   1530
            Width           =   600
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
            Index           =   20
            Left            =   4350
            TabIndex        =   41
            Top             =   1110
            Width           =   345
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
            Index           =   19
            Left            =   150
            TabIndex        =   40
            Top             =   1110
            Width           =   525
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Endereço"
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
            Index           =   18
            Left            =   150
            TabIndex        =   39
            Top             =   660
            Width           =   795
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
            Left            =   -74850
            TabIndex        =   38
            Top             =   210
            Width           =   600
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "CNPJ"
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
            Left            =   -74850
            TabIndex        =   37
            Top             =   570
            Width           =   435
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "I.E."
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
            Left            =   -70785
            TabIndex        =   36
            Top             =   510
            Width           =   300
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Razão"
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
            Left            =   -74850
            TabIndex        =   35
            Top             =   900
            Width           =   525
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Fantasia"
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
            Left            =   -74850
            TabIndex        =   34
            Top             =   1230
            Width           =   705
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Ramo de Atividade"
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
            Left            =   -74850
            TabIndex        =   33
            Top             =   1590
            Width           =   1620
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Endereço"
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
            Left            =   -74850
            TabIndex        =   32
            Top             =   2250
            Width           =   795
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
            Index           =   7
            Left            =   -74850
            TabIndex        =   31
            Top             =   2640
            Width           =   525
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
            Index           =   8
            Left            =   -74820
            TabIndex        =   30
            Top             =   1920
            Width           =   345
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
            Index           =   9
            Left            =   -70920
            TabIndex        =   29
            Top             =   2610
            Width           =   600
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "UF"
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
            Left            =   -70590
            TabIndex        =   28
            Top             =   1890
            Width           =   210
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Dt. Cadastro"
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
            Left            =   -74850
            TabIndex        =   27
            Top             =   2940
            Width           =   1095
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   12
            Left            =   -70890
            TabIndex        =   26
            Top             =   2940
            Width           =   540
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Telefone"
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
            Index           =   13
            Left            =   -74850
            TabIndex        =   25
            Top             =   3270
            Width           =   735
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Fax"
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
            Index           =   14
            Left            =   -70650
            TabIndex        =   24
            Top             =   3270
            Width           =   300
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Contato"
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
            Index           =   15
            Left            =   -74850
            TabIndex        =   23
            Top             =   3600
            Width           =   675
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "E-Mail"
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
            Index           =   16
            Left            =   -74850
            TabIndex        =   22
            Top             =   3900
            Width           =   510
         End
         Begin VB.Label lblramo 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   -71250
            TabIndex        =   21
            Top             =   1530
            Width           =   3135
         End
      End
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   6390
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "pf_cadcli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Data de criacao: 21/08/2003
'Criador :Rafael
'Ultima atualizacao:10/05/2012         por
Option Explicit
Dim nome_tab(1) As String
Dim cpo(30) As String 'NOME DOS CAMPOS
Dim tcpo(30) As String 'TIPO DOS CAMPOS
Dim Exibir As Boolean ' serve para ver se o formulario vai ficar aberto
Dim Selecao As Boolean 'verifica se exite alguma seleção
Dim snap_Selecao As New ADODB.Recordset ' objeto da seleção
Dim snap_Procura As New ADODB.Recordset ' objeto da seleção
'Descricao: manutencao


'Private Sub cmd_Click(Index As Integer)
'Dim relatorio As New cls_Relatorio
'Select Case Index
'Case 0 'enviar email...
    'WebBrowser.Navigate "mailto:" & Trim(txt(16))
'Case 1 ' exportar dados
    'relatorio.Banco = Banco_path
    'relatorio.relatorio = App.Path & "\reports\clientes.rpt"
    
'End Select
'End Sub

'Tabelas utilizadas: pf_cliente

Private Sub cmdP_Click(Index As Integer)

End Sub

Private Sub cmdProc_Click(Index As Integer)
Select Case Index
    Case 0
        Cod_con = 10
        IM_consulta.Show 1
        lblramo = Result_con
        txt(5) = CpoChave_con
End Select

End Sub

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
SSTab.Tab = 0

Preenche_ClienteTipo

Exit Sub
erro:
MsgBox Err.Description, vbCritical, "FORM LOAD " & Me.Name
Exibir = False
'Unload Me

End Sub

Sub Declara()
' declara todas as varíaveis de dados
nome_tab(0) = "pf_cliente"
nome_tab(1) = "pf_Aluno"
cpo(0) = "cli_cod"
cpo(1) = "cli_cnpj"
cpo(2) = "cli_ie"
cpo(3) = "cli_rzsc"
cpo(4) = "cli_fant"
cpo(5) = "cli_ramo"
cpo(6) = "cli_ende"
cpo(7) = "cli_bairr"
cpo(8) = "cli_cep"
cpo(9) = "cli_cida"
cpo(10) = "cli_uf"
cpo(11) = "cli_dtcad"
cpo(12) = "cli_status"
cpo(13) = "cli_tele"
cpo(14) = "cli_fax"
cpo(15) = "cli_cont"
cpo(16) = "cli_mail"

cpo(17) = "cli_endefinan"
cpo(18) = "cli_bairrfinan"
cpo(19) = "cli_cidafinan"
cpo(20) = "cli_cepfinan"
cpo(21) = "cli_uffinan"

cpo(27) = "ClienteTipo"

cpo(28) = "cli_dupg"
cpo(29) = "cli_total"
cpo(30) = "cli_TabelaPrecoPadrao"


tcpo(0) = "NUMERO"
tcpo(1) = "TEXTO"
tcpo(2) = "TEXTO"
tcpo(3) = "TEXTO"
tcpo(4) = "TEXTO"
tcpo(5) = "NUMERO"
tcpo(6) = "TEXTO"
tcpo(7) = "TEXTO"
tcpo(8) = "TEXTO"
tcpo(9) = "TEXTO"
tcpo(10) = "TEXTO"
tcpo(11) = "DATA"
tcpo(12) = "NUMERO"
tcpo(13) = "TEXTO"
tcpo(14) = "TEXTO"
tcpo(15) = "TEXTO"
tcpo(16) = "TEXTO"

tcpo(17) = "NUMERO"
tcpo(18) = "TEXTO"
tcpo(19) = "TEXTO"
tcpo(20) = "NUMERO"
tcpo(21) = "TEXTO"


tcpo(22) = "TEXTO"
tcpo(23) = "TEXTO"
tcpo(24) = "TEXTO"
tcpo(25) = "NUMERO"
tcpo(26) = "TEXTO"
tcpo(27) = "NUMERO"

tcpo(28) = "DATA"
tcpo(29) = "NUMERO"
tcpo(30) = "NUMERO"

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
Dim X As Integer
For X = 0 To txt.Count - 1
    txt(X).Text = ""
Next X
lblramo = ""

txt(0).BackColor = &H80000005
lblDUPagamento = ""
lblSaldo = ""

cboClienteTipo.ListIndex = -1

cboListaPrecoPadrao.ListIndex = -1
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
Dim Sql As String
'For x = 0 To txt.Count - 1
For X = 0 To txt.Count - 1
    If IsNull(snap_Selecao(cpo(X))) Then
        txt(X) = ""
    Else
        txt(X) = snap_Selecao(cpo(X))
    End If
Next X


'Sql = "select ram_desc from pf_cliente, pf_ramativ where cli_ramo = ram_cod and cli_cod=" & snap_Selecao("cli_cod")
'Set snap_Procura = ocnBanco.Execute(Sql)
'snap_Procura.MoveFirst
'lblramo = snap_Procura("ram_desc")

MostraRotulo (99)
'Exibe_ClienteTipo snap_Selecao(cpo(27))
If IsNull(snap_Selecao(cpo(27))) Then
    cboClienteTipo.ListIndex = -1
Else
    Exibe_ClienteTipo snap_Selecao(cpo(27))
End If

'Lista de preco
If IsNull(snap_Selecao(cpo(30))) Then
    cboListaPrecoPadrao.ListIndex = -1
Else
    Exibe_ListaPadrao snap_Selecao(cpo(30))
End If


'dataUltimoPagamento
If IsNull(snap_Selecao(cpo(28))) Then
    lblDUPagamento = ""
    
Else
    lblDUPagamento = snap_Selecao(cpo(28))
End If


'saldo
If IsNull(snap_Selecao(cpo(29))) Then
    lblSaldo = ""
Else
    lblSaldo = snap_Selecao(cpo(29))
End If

Set snap_Procura = Nothing

End Sub

Private Sub Form_Resize()
'Me.Caption = Me.Height & "   " & Me.Width
Me.Height = 6705
Me.Width = 7860

End Sub

Function conssiste() As Boolean
conssiste = False
' no txt 1 verificar cnpj
' no txt 2 veriicar inscricao estadual
If Len(txt(3)) <= 0 Then
    conssiste = False
    txt(3).SetFocus
    Exit Function
End If
If Len(txt(4)) = 0 Then txt(4) = txt(3)
If Len(lblramo) <= 0 Then
    conssiste = False
    txt(5).SetFocus
    Exit Function
End If
If InStr(1, txt(1), "*") > 0 Then
    conssiste = False
    txt(1).SetFocus
    Exit Function
End If
'If Len(txt(1)) > 0 Then
    'If CNPJ(txt(1)) = False Then
        'conssiste = False
        'txt(1).SetFocus
        'Exit Function
    'End If
'End If
'If Len(txt(2)) > 0 Then
    'If ChecaInscrE(txt(10), txt(2)) = False Then
        'conssiste = False
        'txt(2).SetFocus
        'Exit Function
    'End If
'End If

If cboClienteTipo.ListIndex = -1 Then
    conssiste = False
    cboClienteTipo.SetFocus
    Exit Function
End If

If cboListaPrecoPadrao.ListIndex = -1 Then
    conssiste = False
    cboListaPrecoPadrao.SetFocus
    Exit Function
End If


conssiste = True
End Function

Private Sub Form_Unload(Cancel As Integer)
BOTDisable
End Sub

Private Sub txt_GotFocus(Index As Integer)
txt(Index).SelStart = 0
txt(Index).SelLength = Len(txt(Index))

End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
    Case 0
        If KeyAscii = 13 Then
            Menu (3)
            Menu (4)
        End If


Case 8 'endereço do cliente
    If KeyAscii = 13 Then
        If Busca_Cep(txt(Index)) = True Then
            txt(6) = CEP_Nome
            txt(7) = CEP_Bairro
            txt(9) = CEP_Cidade
            txt(10) = "SP"
            txt(6).SetFocus
        End If
    End If
Case 20 'cep do financeiro
    If KeyAscii = 13 Then
        If Busca_Cep(txt(Index)) = True Then
            txt(17) = CEP_Nome
            txt(18) = CEP_Bairro
            txt(19) = CEP_Cidade
            txt(21) = "SP"
            txt(18).SetFocus
        End If
    End If
End Select

End Sub

Private Sub txt_LostFocus(Index As Integer)
Dim Sql As String
Select Case Index
    Case 5 ' ramo de atividade
        MostraRotulo (1)
        
    
End Select

End Sub

Sub MostraRotulo(Index As Integer)
Dim Sql As String
Select Case Index
    Case 1 ' RamoAtividade
    
        If IsNumeric(txt(5)) Then
            Sql = "Select ram_desc from pf_ramativ where ram_cod = " & txt(5)
            Set snap_Procura = ocnBanco.Execute(Sql)
            If snap_Procura.BOF And snap_Procura.EOF Then
                lblramo = ""
            Else
                snap_Procura.MoveFirst
                lblramo = snap_Procura("ram_desc")
            End If
            Set snap_Procura = Nothing
        Else
            lblramo = ""
        End If
        
        
    Case 2 ' SEDEX
    
    Case 3 'Produto
    
    Case 99 'todos
    
        'ramo atividade
        If IsNumeric(txt(5)) Then
            Sql = "Select ram_desc from pf_ramativ where ram_cod = " & txt(5)
            Set snap_Procura = ocnBanco.Execute(Sql)
            If snap_Procura.BOF And snap_Procura.EOF Then
                lblramo = ""
            Else
                snap_Procura.MoveFirst
                lblramo = snap_Procura("ram_desc")
            End If
            Set snap_Procura = Nothing
        Else
            lblramo = ""
        End If
    
    

End Select

End Sub


Sub Menu(Index As Integer)
Dim Sql As String
Dim snap_Cod As New ADODB.Recordset
On Error GoTo erro

'Inserido  10/05/2012
Dim iQtdeCampo As Integer
Dim X As Integer
Select Case Index
Case 0 ' botao inserir
    status.Panels(1).Text = "INSERIR"
    txt(0).Locked = True
    Botoes
    txt(11) = Date

Case 1 ' botao excluir
    'procura referencia
    'pede confirmação
    If status.Panels(1).Text = "ALTERAR" Then
        If MsgBox("Deseja realmente excluir ?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
            Sql = "delete from " & nome_tab(0) & " where cli_cod=" & txt(0)
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
                
                Sql = "SELECT MAX(cli_COD)+1 FROM " & nome_tab(0)
                Set snap_Cod = ocnBanco.Execute(Sql)
                If IsNull(snap_Cod(0)) Then
                    txt(0) = 1
                Else
                    txt(0) = snap_Cod(0)
                End If
                txt(11) = Date
                Set snap_Cod = Nothing
                
                
                'forma generica para inserir dados
                
                iQtdeCampo = 21
                    Sql = "insert into " & nome_tab(0) & Chr(13)
                    Sql = Sql & "("
                    For X = 0 To iQtdeCampo
                        Sql = Sql & cpo(X)
                        If X < iQtdeCampo Then Sql = Sql & ","
                    Next X
                       Sql = Sql & ", " & cpo(27) & ", " & cpo(30)
                    Sql = Sql & ") values("
                    For X = 0 To iQtdeCampo
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
                        
                        
                        If X < iQtdeCampo Then Sql = Sql & ","
                    Next X
                    
                    Sql = Sql & ", " & cboClienteTipo.ItemData(cboClienteTipo.ListIndex)
                    Sql = Sql & ", " & cboListaPrecoPadrao.ItemData(cboListaPrecoPadrao.ListIndex)
                    Sql = Sql & ")" & vbCrLf
                
                ocnBanco.Execute (Sql)
                
                
                'Sql = "insert into " & nome_tab(0) & Chr(13)
                'Sql = Sql & "values(" & txt(0) & ",'" & txt(1) & "','" & txt(2) & "','" & txt(3) & "','" & txt(4) & "','" & txt(5) & "','" & txt(6) & "','" & txt(7) & "','" & txt(8) & "','" & txt(9) & "','" & txt(10) & "','" & txt(11) & "','" & txt(12) & "','" & txt(13) & "','" & txt(14) & "','" & txt(15) & "','" & txt(16) & "','" & txt(17) & "','" & txt(18) & "','" & txt(19) & "','" & txt(20) & "','" & txt(21) & "','" & txt(22) & "','" & cboMidia.ItemData(cboMidia.ListIndex) & "','" & txt(23) & "','" & cboClienteTipo.ItemData(cboClienteTipo.ListIndex) & "')"
                'ocnBanco.Execute (Sql)
                
                status.Panels(1).Text = "ALTERAR"
                MsgBox "Os dados foram gravados com sucesso!", vbInformation, "Gravação"
                Botoes
            End If
        End If
    ElseIf status.Panels(1).Text = "ALTERAR" Then
        If conssiste = True Then
        
            ' atualização genérica
            iQtdeCampo = 21
            Sql = "UPDATE " & nome_tab(0) & " SET " & Chr(13)
                For X = 1 To iQtdeCampo
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
                        If X < iQtdeCampo Then Sql = Sql & ","
                    Next X
                    
                Sql = Sql & "," & cpo(27) & "= " & cboClienteTipo.ItemData(cboClienteTipo.ListIndex)
                
                Sql = Sql & "," & cpo(30) & "= " & cboListaPrecoPadrao.ItemData(cboListaPrecoPadrao.ListIndex)
                    
                Sql = Sql & vbCrLf
                Sql = Sql & "WHERE " & cpo(0) & "=" & txt(0)
        
            'Sql = "UPDATE " & nome_tab(0) & " SET " & cpo(1) & "='" & txt(1) & "'," & cpo(2) & "='" & txt(2) & "'," & cpo(3) & "='" & txt(3) & "'," & cpo(4) & "='" & txt(4) & "'," & cpo(5) & "='" & txt(5) & "'," & cpo(6) & "='" & txt(6) & "'," & cpo(7) & "='" & txt(7) & "'," & cpo(8) & "='" & txt(8) & "'," & cpo(9) & "='" & txt(9) & "'," & cpo(10) & "='" & txt(10) & "'," & cpo(11) & "='" & txt(11) & "'," & cpo(12) & "='" & txt(12) & "'," & cpo(13) & "='" & txt(13) & "'," & cpo(14) & "='" & txt(14) & "'," & cpo(15) & "='" & txt(15) & "'," & cpo(16) & "='" & txt(16) & "'," & cpo(17) & "='" & txt(17) & "'," & cpo(18) & "='" & txt(18) & "'," & cpo(19) & "='" & txt(19) & "'," & cpo(20) & "='" & txt(20) & "'," & cpo(21) & "='" & txt(21) & "'," & cpo(22) & "='" & txt(22) & "'," & cpo(25) & "=" & cboMidia.ItemData(cboMidia.ListIndex) & ", " & cpo(26) & "='" & txt(23) & "'," & cpo(27) & "=" & cboClienteTipo.ItemData(cboClienteTipo.ListIndex) & Chr(13)
            'Sql = Sql & "WHERE " & cpo(0) & "=" & txt(0)
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





Sub Preenche_ClienteTipo()
Set snap_Procura = Nothing
Dim Sql As String
Dim i As Integer
Sql = "select ClienteTipoID, Descricao from ClienteTipo order by Descricao"
snap_Procura.Open Sql, ocnBanco, adOpenKeyset, adLockOptimistic, adCmdText
cboClienteTipo.Clear
i = 0
If Not (snap_Procura.BOF And snap_Procura.EOF) Then
    snap_Procura.MoveFirst
    While Not snap_Procura.EOF
        cboClienteTipo.AddItem snap_Procura("Descricao")
        cboClienteTipo.ItemData(i) = snap_Procura("ClienteTipoID")
        i = i + 1
        snap_Procura.MoveNext
    Wend
End If
End Sub

Sub Exibe_ClienteTipo(cod As Integer)
Dim X As Integer
For X = 0 To cboClienteTipo.ListCount - 1
    If cboClienteTipo.ItemData(X) = cod Then
        cboClienteTipo.ListIndex = X
        Exit For
    End If
Next X
End Sub



Sub Exibe_ListaPadrao(cod As Integer)
Dim X As Integer
For X = 0 To cboListaPrecoPadrao.ListCount - 1
    If cboListaPrecoPadrao.ItemData(X) = cod Then
        cboListaPrecoPadrao.ListIndex = X
        Exit For
    End If
Next X
End Sub

