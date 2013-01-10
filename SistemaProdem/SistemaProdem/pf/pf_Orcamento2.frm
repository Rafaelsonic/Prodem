VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form pf_Orcamento2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Orcamento / Vendas V2"
   ClientHeight    =   7920
   ClientLeft      =   465
   ClientTop       =   1080
   ClientWidth     =   8730
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1.761
   ScaleMode       =   0  'User
   ScaleWidth      =   0.883
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   7425
      Left            =   0
      TabIndex        =   12
      Top             =   120
      Width           =   8655
      Begin VB.CommandButton cmdRelatorio 
         Caption         =   "Impressão"
         Height          =   375
         Left            =   6840
         TabIndex        =   51
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmd_confirmacao 
         Caption         =   "Cancelar Pedido"
         Enabled         =   0   'False
         Height          =   405
         Index           =   4
         Left            =   6840
         Picture         =   "pf_Orcamento2.frx":0000
         TabIndex        =   47
         Top             =   960
         Width           =   1485
      End
      Begin VB.CommandButton cmd_confirmacao 
         Caption         =   "Finaliza Pedido"
         Enabled         =   0   'False
         Height          =   405
         Index           =   1
         Left            =   6840
         Picture         =   "pf_Orcamento2.frx":0464
         TabIndex        =   46
         Top             =   360
         Width           =   1485
      End
      Begin VB.ComboBox cboSedex 
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
         ItemData        =   "pf_Orcamento2.frx":08C8
         Left            =   1320
         List            =   "pf_Orcamento2.frx":08D5
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2280
         Width           =   1395
      End
      Begin VB.CommandButton cmdProc 
         Height          =   375
         Index           =   0
         Left            =   2520
         Picture         =   "pf_Orcamento2.frx":08F4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   720
         Width           =   375
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4095
         Left            =   120
         TabIndex        =   20
         Top             =   3240
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   7223
         _Version        =   393216
         TabOrientation  =   1
         Tab             =   1
         TabHeight       =   520
         TabCaption(0)   =   "Produto"
         TabPicture(0)   =   "pf_Orcamento2.frx":0D14
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "label(7)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "label(6)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "label(5)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "label(4)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "lbl(1)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "label(9)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "lbl(5)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "lbl(7)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "label(19)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "lbl(10)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "grid(0)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "txt(8)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "txt(7)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "txt(5)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "txt(6)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "cmdProc(2)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "cmd_confirmacao(0)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "cmd_Cancela(1)"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "txt(9)"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "cmdGerarDevolucao"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "cmd_confirmacao(5)"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).ControlCount=   21
         TabCaption(1)   =   "Itens Devolvidos"
         TabPicture(1)   =   "pf_Orcamento2.frx":0D30
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "label(10)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "label(11)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "lbl(2)"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "label(17)"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "lbl(8)"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "label(18)"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "lbl(9)"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "label(20)"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "lbl(11)"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "grid(1)"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "cmd_confirmacao(2)"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "cmd_confirmacao(3)"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "txt(10)"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).ControlCount=   13
         TabCaption(2)   =   "Observações"
         TabPicture(2)   =   "pf_Orcamento2.frx":0D4C
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "txt(4)"
         Tab(2).ControlCount=   1
         Begin VB.CommandButton cmd_confirmacao 
            Caption         =   "Limpar"
            Height          =   405
            Index           =   5
            Left            =   -67800
            Picture         =   "pf_Orcamento2.frx":0D68
            TabIndex        =   60
            Top             =   120
            Width           =   885
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   10
            Left            =   1320
            TabIndex        =   53
            Top             =   1080
            Width           =   825
         End
         Begin VB.CommandButton cmdGerarDevolucao 
            Caption         =   "Gerar Devolução"
            Height          =   375
            Left            =   -72600
            TabIndex        =   52
            Top             =   1200
            Width           =   1455
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   9
            Left            =   -72360
            TabIndex        =   6
            Top             =   240
            Width           =   2505
         End
         Begin VB.TextBox txt 
            Height          =   2925
            Index           =   4
            Left            =   -74880
            MultiLine       =   -1  'True
            TabIndex        =   39
            Top             =   240
            Width           =   7935
         End
         Begin VB.CommandButton cmd_confirmacao 
            Height          =   405
            Index           =   3
            Left            =   840
            Picture         =   "pf_Orcamento2.frx":11CC
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   1440
            Width           =   375
         End
         Begin VB.CommandButton cmd_confirmacao 
            Height          =   405
            Index           =   2
            Left            =   360
            Picture         =   "pf_Orcamento2.frx":1641
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   1440
            Width           =   405
         End
         Begin VB.CommandButton cmd_Cancela 
            Enabled         =   0   'False
            Height          =   405
            Index           =   1
            Left            =   -68190
            Picture         =   "pf_Orcamento2.frx":1AA5
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton cmd_confirmacao 
            Enabled         =   0   'False
            Height          =   405
            Index           =   0
            Left            =   -68640
            Picture         =   "pf_Orcamento2.frx":1F1A
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   120
            Width           =   405
         End
         Begin VB.CommandButton cmdProc 
            Height          =   375
            Index           =   2
            Left            =   -72840
            Picture         =   "pf_Orcamento2.frx":237E
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   120
            Width           =   375
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   6
            Left            =   -73800
            TabIndex        =   7
            Text            =   "1"
            Top             =   720
            Width           =   705
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   5
            Left            =   -73800
            TabIndex        =   5
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   7
            Left            =   -71880
            TabIndex        =   8
            Top             =   690
            Width           =   525
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   8
            Left            =   -70200
            TabIndex        =   9
            Top             =   690
            Width           =   825
         End
         Begin MSFlexGridLib.MSFlexGrid grid 
            Height          =   2055
            Index           =   0
            Left            =   -74880
            TabIndex        =   25
            Top             =   1680
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   3625
            _Version        =   393216
            Rows            =   1
            Cols            =   6
         End
         Begin MSFlexGridLib.MSFlexGrid grid 
            Height          =   2055
            Index           =   1
            Left            =   120
            TabIndex        =   28
            Top             =   2040
            Width           =   8055
            _ExtentX        =   14208
            _ExtentY        =   3625
            _Version        =   393216
            Cols            =   7
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
            Index           =   11
            Left            =   4200
            TabIndex        =   62
            Tag             =   "1"
            Top             =   1080
            Width           =   1485
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade Atual"
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
            Left            =   120
            TabIndex        =   61
            Top             =   600
            Width           =   1470
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
            Index           =   10
            Left            =   -73800
            TabIndex        =   59
            Tag             =   "1"
            Top             =   1200
            Width           =   765
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Nr. Item"
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
            Left            =   -74880
            TabIndex        =   58
            Top             =   1320
            Width           =   720
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
            Left            =   5160
            TabIndex        =   57
            Tag             =   "1"
            Top             =   600
            Width           =   1485
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Valor Devolvido"
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
            Left            =   2520
            TabIndex        =   56
            Top             =   1080
            Width           =   1365
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
            Left            =   1800
            TabIndex        =   55
            Tag             =   "1"
            Top             =   600
            Width           =   1485
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Valor Produto"
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
            Left            =   3720
            TabIndex        =   54
            Top             =   600
            Width           =   1170
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
            Index           =   7
            Left            =   -74160
            TabIndex        =   50
            Tag             =   "1"
            Top             =   240
            Visible         =   0   'False
            Width           =   285
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
            Index           =   5
            Left            =   -68640
            TabIndex        =   44
            Tag             =   "1"
            Top             =   720
            Width           =   1005
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   9
            Left            =   -69240
            TabIndex        =   34
            Top             =   720
            Width           =   450
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
            Index           =   2
            Left            =   1320
            TabIndex        =   31
            Tag             =   "1"
            Top             =   240
            Width           =   5325
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
            Index           =   1
            Left            =   -69720
            TabIndex        =   30
            Tag             =   "1"
            Top             =   240
            Width           =   645
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Produto"
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
            Left            =   240
            TabIndex        =   27
            Top             =   240
            Width           =   660
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
            Index           =   10
            Left            =   120
            TabIndex        =   26
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Produto"
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
            Left            =   -74880
            TabIndex        =   24
            Top             =   240
            Width           =   660
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
            Index           =   5
            Left            =   -72960
            TabIndex        =   23
            Top             =   720
            Width           =   975
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Desconto %"
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
            Left            =   -71280
            TabIndex        =   22
            Top             =   720
            Width           =   1035
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Lista Preço"
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
            Left            =   -74880
            TabIndex        =   21
            Top             =   720
            Width           =   930
         End
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   3
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   3
         Left            =   4440
         TabIndex        =   13
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   8520
         Y1              =   3120
         Y2              =   3120
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
         Index           =   16
         Left            =   3120
         TabIndex        =   49
         Top             =   1320
         Width           =   600
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
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
         Left            =   120
         TabIndex        =   48
         Top             =   1320
         Width           =   570
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
         Index           =   6
         Left            =   4320
         TabIndex        =   45
         Tag             =   "1"
         Top             =   360
         Width           =   2205
      End
      Begin VB.Label lblCidade 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3960
         TabIndex        =   43
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label lblUF 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1320
         TabIndex        =   42
         Top             =   1320
         Width           =   855
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
         Index           =   4
         Left            =   4440
         TabIndex        =   41
         Tag             =   "1"
         Top             =   2400
         Width           =   1365
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Valor Frete"
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
         Left            =   3120
         TabIndex        =   40
         Top             =   2400
         Width           =   945
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
         Index           =   3
         Left            =   1320
         TabIndex        =   33
         Tag             =   "1"
         Top             =   2760
         Width           =   1365
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Total Pedido"
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
         Left            =   120
         TabIndex        =   32
         Top             =   2760
         Width           =   1050
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
         Left            =   3120
         TabIndex        =   29
         Tag             =   "1"
         Top             =   840
         Width           =   3405
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   945
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   600
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   1950
         Width           =   1170
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Data Prevista"
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
         Left            =   3120
         TabIndex        =   16
         Top             =   1920
         Width           =   1155
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Entrega"
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
         Left            =   120
         TabIndex        =   15
         Top             =   2280
         Width           =   660
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
         Index           =   15
         Left            =   3480
         TabIndex        =   14
         Top             =   360
         Width           =   540
      End
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   7545
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "pf_Orcamento2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Data de criacao:21/08/2003
'Criador :Rafael
'Ultima atualizacao:         por
Option Explicit
Dim nome_tab(0) As String
Dim cpo(16) As String 'NOME DOS CAMPOS
Dim tcpo(16) As String 'TIPO DOS CAMPOS
Dim Exibir As Boolean ' serve para ver se o formulario vai ficar aberto
Dim Selecao As Boolean 'verifica se exite alguma seleção
Dim snap_Selecao As New ADODB.Recordset ' objeto da seleção
Dim snap_Procura As New ADODB.Recordset
Dim relatorio As New cls_Relatorio


Private Sub cmd_Click(index As Integer)
Dim Sql As String
Dim snap_Cod As New ADODB.Recordset
On Error GoTo erro

Select Case index
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
    txt(0).SetFocus
    txt(2).Locked = False
End Select
Exit Sub
erro:
MsgBox Err.Description, vbCritical, "cmd_click"
End Sub

Private Sub cboSedex_Change()
MostraRotulo (2)
End Sub

Private Sub cboSedex_LostFocus()
MostraRotulo (2)
End Sub

Private Sub cmd_Cancela_Click(index As Integer)
Select Case index
Case 1
    If IsNumeric(txt(5)) Then
        'remove os dados do grid
        remover_Produto (txt(5))
        Limpar_produto
        Total_Orcamento
    Else
        MsgBox "Selecionar o produto", , "cmd_cancela"
    End If
End Select

End Sub

Private Sub cmd_confirmacao_Click(index As Integer)
Dim Sql As String
Select Case index
Case 0
    'Insere os dados no grid
        If IsNumeric(txt(5)) And IsNumeric(txt(6)) And IsNumeric(txt(7)) And (IsNumeric(txt(8)) Or txt(8) = "") Then
            If val(txt(6)) > 0 And val(txt(7)) > 0 Then
                If Len(Trim(lbl(1))) > 0 Then
                
                    '10/01/2013
                    'Valido o campo de desconto de 0 a 100
                    If IsNumeric(txt(8)) Then
                        If Not (txt(8) >= 0 And txt(8) < 100) Then
                            MsgBox "Desconto inválido." & vbCrLf & "O valor será ignorado.", vbCritical, "Atenção"
                            txt(8) = 0
                        End If
                    Else
                        txt(8) = 0
                    End If
                    
                    Inserir_Produto (txt(5)) 'inserir com código
                    Total_Orcamento
                    'GridItens
                    Mostra_Grid txt(0)
                    'limpar campos do produto
                    txt(5).SetFocus
                Else
                    MsgBox "Selecionar o produto", , "cmd_confirmacao"
                    txt(5).SetFocus
                End If
            Else
                MsgBox "Os campos numéricos devem ser maior que zero", , "cmd_aln"
            End If
        Else
            MsgBox "Deve ser numérico", , "cmd_confirmacao"
        End If

Case 1
    If MsgBox("Deseja finalizar?", vbYesNo + vbQuestion, "Cancelar Pedido") = vbYes Then
        'cancelar com o enter
        Sql = "update pf_OrcamentoH set orch_situacao = 2 where orch_cod=" & txt(0) & " and orch_situacao = 1 "
        ocnBanco.Execute (Sql)
        MsgBox "Pedido finalizado com sucesso.", vbInformation, "cmd_confirmacao"
        
        Mostra_Dados
        Botoes
        
    End If

    
Case 4
    If MsgBox("Deseja realmente cancelar?", vbYesNo + vbQuestion, "Cancelar Pedido") = vbYes Then
        'cancelar com o enter
        Sql = "update pf_OrcamentoH set orch_situacao = 3 where orch_cod=" & txt(0) & " and orch_situacao = 1 "
        ocnBanco.Execute (Sql)
        MsgBox "Pedido cancelado com sucesso.", vbInformation, "cmd_confirmacao"
        
        Mostra_Dados
        Botoes
        
    End If

'10/01/2012 Limpar Produto
Case 5
    Limpar_produto
End Select
End Sub

Private Sub cmdProc_Click(index As Integer)
Select Case index
Case 0
    Cod_con = 40
    IM_consulta.Show 1
    lbl(0) = Result_con
    txt(1) = CpoChave_con
    MostraRotulo (1)
'Case 1
'    Cod_con = 200
'    IM_consulta.Show 1
'    lbl(4) = Result_con
'    txt(4) = CpoChave_con
Case 2
    Cod_con = 190
    IM_consulta.Show 1
    lbl(1) = Result_con
    txt(9) = Result_con
    txt(5) = CpoChave_con
    
    txt(9).SetFocus
    
    
End Select

End Sub

Private Sub cmdRelatorio_Click()
If IsNumeric(txt(0)) Then
    relatorio.Banco = Banco_path
    relatorio.relatorio = App.Path & "\reports\Pedido.rpt"
    relatorio.titulo = "Pedido " & txt(0)
    relatorio.SenhaBanco = "sistema"
    relatorio.Formula = "{pf_orcamentoH.OrcH_cod} =" & txt(0)
    relatorio.Vizualizar
End If
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
Exit Sub
erro:
MsgBox Err.Description, vbCritical, "FORM LOAD " & Me.Name
Exibir = False
'Unload Me

End Sub

Sub Declara()
'declara todas as varíaveis de dados
nome_tab(0) = "pf_OrcamentoH"
cpo(0) = "Orch_cod"
cpo(1) = "Orch_cli"
cpo(2) = "Orch_dtEmissao"
cpo(3) = "orch_dtEntregaPrevista"
cpo(4) = "observacao"
cpo(5) = "Entrega"
cpo(6) = "orch_situacao"

tcpo(0) = "NUMERO"
tcpo(1) = "NUMERO"
tcpo(2) = "DATA"
tcpo(3) = "DATA"
tcpo(4) = "TEXTO"
tcpo(5) = "NUMERO"
tcpo(6) = "NUMERO"
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
    
    
    cmd_Cancela(1).Enabled = False
    cmd_confirmacao(0).Enabled = False
    
    SSTab1.Visible = False

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
    txt(2).Locked = False
    
    
    cmd_Cancela(1).Enabled = False
    cmd_confirmacao(0).Enabled = False

    SSTab1.Visible = False

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
    
    
    cmd_Cancela(1).Enabled = False
    cmd_confirmacao(0).Enabled = False

    SSTab1.Visible = False
    
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
    txt(2).Locked = True

    cmd_Cancela(1).Enabled = True
    cmd_confirmacao(0).Enabled = True

    SSTab1.Visible = True
End Select

Select Case Mid(lbl(6), 1, 1)
Case 1
    cmd_confirmacao(1).Enabled = True
    cmd_confirmacao(4).Enabled = True
    
    cmd_confirmacao(0).Enabled = True
    cmd_Cancela(1).Enabled = True
    
Case 2
    cmd_confirmacao(1).Enabled = False
    cmd_confirmacao(4).Enabled = False
    
    mdi.Toolbar.Buttons(3).Enabled = False
    mdi.Toolbar.Buttons(2).Enabled = False
    
    cmd_confirmacao(0).Enabled = False
    cmd_Cancela(1).Enabled = False
    
Case 3
    cmd_confirmacao(1).Enabled = False
    cmd_confirmacao(4).Enabled = False
    
    mdi.Toolbar.Buttons(3).Enabled = False
    mdi.Toolbar.Buttons(2).Enabled = False
    
    cmd_confirmacao(0).Enabled = False
    cmd_Cancela(1).Enabled = False
End Select

End Sub
Sub Limpar()
Dim X As Integer
For X = 0 To txt.Count - 1
    txt(X).Text = ""
Next X
txt(0).BackColor = &H80000005
cboSedex.ListIndex = -1

lbl(0) = ""
lbl(3) = ""
lbl(4) = ""
lbl(6) = ""
lblCidade = ""
lblUF = ""

grid(0).Clear
grid(0).Rows = 1
Limpar_produto
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
        Else ' verifico o ano
            If Year(txt(X)) < 1990 Or Year(txt(X)) > 2090 Then
                txt(X).SetFocus
                conssiste_total = False
                Exit Function
            End If
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
        'Rafael alterado para 4
        For X = 0 To 4 ''txt.Count - 1
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
Limpar
Dim X As Integer
For X = 0 To 4 'txt.Count - 1
    If IsNull(snap_Selecao(cpo(X))) = False Then txt(X) = snap_Selecao(cpo(X))
Next X

Select Case snap_Selecao(cpo(6))
Case 1
    lbl(6) = "1 - Pendente"
Case 2
    lbl(6) = "2 - Finalizado"
Case 3
    lbl(6) = "3 - Cancelado"
End Select

If IsNull(snap_Selecao(cpo(5))) = False Then
    Select Case snap_Selecao(cpo(5))
        Case 1
            cboSedex.Text = "A RETIRAR"
        Case 2
            cboSedex.Text = "SEDEX"
        Case 3
            cboSedex.Text = "SEDEX10"
    End Select
Else
    cboSedex.ListIndex = -1
End If

Mostra_Grid txt(0)
MostraRotulo 99
Total_Orcamento
End Sub

Private Sub Form_Resize()
'Me.Caption = Me.Height & "   " & Me.Width
Me.Height = 8400
Me.Width = 8790
End Sub

Function conssiste() As Boolean
conssiste = False
If Len(txt(1)) > 0 Then
    conssiste = True
Else
    conssiste = False
    txt(1).SetFocus
    Exit Function
End If
If InStr(1, txt(1), "*") > 0 Then
    conssiste = False
    txt(1).SetFocus
End If

End Function

Private Sub Form_Unload(Cancel As Integer)
BOTDisable
End Sub


Private Sub grid_Click(index As Integer)
'10/01/13
txt(5) = "" 'codigo do produto
txt(9) = "" 'descricao
'lbl(1) = snap_Procura("produto_descricao")
'txt(9) = snap_Procura("produto_descricao")
'lbl(5) = snap_Procura("Preco")
'lbl(7) = snap_Procura("produto_codigo")
'6 listaPreco 7 qtde 8 desc

Me.grid(0).Row = Me.grid(0).MouseRow

Me.grid(0).col = 0
Me.lbl(10) = Me.grid(0).Text

Me.grid(0).col = 1
Me.txt(5) = Me.grid(0).Text
lbl(7) = Me.grid(0).Text


Me.grid(0).col = 2
txt(9) = Me.grid(0).Text
Me.lbl(1) = Me.grid(0).Text


Me.grid(0).col = 3
Me.txt(7) = Me.grid(0).Text

'valor
Me.grid(0).col = 4
Me.lbl(5) = Me.grid(0).Text



End Sub

Private Sub txt_GotFocus(index As Integer)
txt(index).SelStart = 0
txt(index).SelLength = Len(txt(index))
End Sub
Sub Menu(index As Integer)
Dim Sql As String
Dim snap_Cod As New ADODB.Recordset
Dim X As Integer
On Error GoTo erro

Select Case index
Case 0 ' botao inserir
    status.Panels(1).Text = "INSERIR"
    txt(0).Locked = True
    txt(2).Locked = True
    txt(2).Text = Format(Now, "dd/mm/yyyy")
    'cboStatus.Text = "Aberto"
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
            
            'forma generica para inserir dados
                    Sql = "insert into " & nome_tab(0) & Chr(13)
                    Sql = Sql & "("
                    For X = 0 To 5
                        Sql = Sql & cpo(X)
                        If X < 5 Then Sql = Sql & ","
                    Next X
                    Sql = Sql & ", " & cpo(6)
                    Sql = Sql & ") values("
                    For X = 0 To 4
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
                        If X < 4 Then Sql = Sql & ","
                    Next X
                    
                    
                    Select Case cboSedex.Text
                    
                    Case "A RETIRAR"
                        Sql = Sql & ",1 "
                    Case "SEDEX"
                        Sql = Sql & ",2 "
                    Case "SEDEX10"
                        Sql = Sql & ",3 "
                    Case Else
                        Sql = Sql & ",NULL "
                    End Select
                    
                    
                    Sql = Sql & ",1)" & vbCrLf
                
                ocnBanco.Execute (Sql)
                status.Panels(1).Text = "ALTERAR"
                MsgBox "Os dados foram gravados com sucesso!", vbInformation, "Gravação"
                lbl(6) = "1 - Pendente"
                Botoes
            End If
        End If
    ElseIf status.Panels(1).Text = "ALTERAR" Then
        If conssiste_total = True Then
            If conssiste = True Then
                
                ' atualização genérica
                Sql = "UPDATE " & nome_tab(0) & " SET " & Chr(13)
                    For X = 1 To 4
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
                            If X < 4 Then Sql = Sql & ","
                        Next X
                    If cboSedex.Text = "A RETIRAR" Then Sql = Sql & ",Entrega=1 "
                    If cboSedex.Text = "SEDEX" Then Sql = Sql & ",Entrega=2 "
                    If cboSedex.Text = "SEDEX10" Then Sql = Sql & ",Entrega=3 "
                    Sql = Sql & vbCrLf
                    Sql = Sql & "WHERE " & cpo(0) & "=" & txt(0)
                
                ocnBanco.Execute Sql
                MsgBox "Os dados foram atualizados com sucesso !", vbInformation, "Alteração"
            End If
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
    txt(2).Locked = False
    
End Select
Exit Sub
erro:

MsgBox Err.Description, vbCritical, "cmd_click"

End Sub

Private Sub TXT_KeyPress(index As Integer, KeyAscii As Integer)
Select Case index
    Case 0
        If KeyAscii = 13 Then
            Menu (3)
            Menu (4)
        End If
    Case 5
        If KeyAscii = 13 Then
            
        End If
End Select
End Sub



Sub MostraRotulo(index As Integer)
Dim Sql As String
Select Case index
    Case 1 ' cliente
        If IsNumeric(txt(index)) Then
            Sql = "Select cli_rzsc, cli_uf, cli_cida from pf_cliente where cli_cod = " & txt(index)
            Set snap_Procura = ocnBanco.Execute(Sql)
            
            If snap_Procura.BOF And snap_Procura.EOF Then
                lbl(0) = ""
                lblUF = ""
                lblCidade = ""
            Else
                snap_Procura.MoveFirst
                lbl(0) = snap_Procura(0)
                
                If IsNull(snap_Procura(1)) = False Then
                    lblUF = snap_Procura(1)
                Else
                    lblUF = ""
                End If
                
                If IsNull(snap_Procura(2)) = False Then
                    lblCidade = snap_Procura(2)
                Else
                    lblCidade = ""
                End If
                
                
            End If
            Set snap_Procura = Nothing
        Else
            lbl(0) = ""
        End If
        
        
    Case 2 ' SEDEX
    
        If cboSedex.ListIndex <> -1 Then
            Sql = "Select Cidade, Valor1, Valor2 from sedex where uf = '" & lblUF.Caption & "'"
            If cboSedex.Text = "SEDEX" Then
                Sql = Sql & " and tipo =1"
            ElseIf cboSedex.Text = "SEDEX10" Then
                Sql = Sql & " and tipo =2"
            Else
                Sql = Sql & " and tipo =99"
            End If
            
            Set snap_Procura = ocnBanco.Execute(Sql)
            If snap_Procura.BOF And snap_Procura.EOF Then
                lbl(4) = FormatCurrency(0, 2)
            Else
            
                snap_Procura.MoveFirst
                If snap_Procura("cidade") = lblCidade Then
                    lbl(4) = FormatCurrency(snap_Procura("valor1"), 2)
                Else
                    lbl(4) = FormatCurrency(snap_Procura("valor2"), 2)
                End If
                
            End If
            Set snap_Procura = Nothing
        Else
            lbl(4) = ""
        End If
        
    Case 3 'Produto
    
        If IsNumeric(txt(5).Text) And IsNumeric(txt(6)) Then
            If txt(6).Text > 0 And txt(6).Text < 11 Then
            
                lbl(1) = ""
                lbl(5) = ""
                lbl(7) = "" ' codigo do produto
            
               Sql = "select produto_codigo, produto_descricao, produto_preco" & txt(6).Text & _
                " as Preco from Produto where produto_codigo=" & txt(5).Text
                
                Set snap_Procura = ocnBanco.Execute(Sql)
                If snap_Procura.BOF And snap_Procura.EOF Then
                
                Else
                
                    snap_Procura.MoveFirst
                    lbl(1) = snap_Procura("produto_descricao")
                    txt(9) = snap_Procura("produto_descricao")
                    lbl(5) = snap_Procura("Preco")
                    lbl(7) = snap_Procura("produto_codigo")
                    
                End If
                
            End If
        End If
    
    
    Case 99 'todos
    
        'cliente
        Sql = "Select cli_rzsc, cli_uf, cli_cida from pf_cliente where cli_cod = " & txt(1)
        Set snap_Procura = ocnBanco.Execute(Sql)
        
        If snap_Procura.BOF And snap_Procura.EOF Then
            lbl(0) = ""
            lblUF = ""
            lblCidade = ""
        Else
            snap_Procura.MoveFirst
            lbl(0) = snap_Procura(0)
            
            If IsNull(snap_Procura(1)) = False Then
                lblUF = snap_Procura(1)
            Else
                lblUF = ""
            End If
            
            If IsNull(snap_Procura(2)) = False Then
                lblCidade = snap_Procura(2)
            Else
                lblCidade = ""
            End If
            
            
        End If
        Set snap_Procura = Nothing
    

        'sedex
        Sql = "Select Cidade, Valor1, Valor2 from sedex where uf = '" & lblUF.Caption & "'"
            If cboSedex.Text = "SEDEX" Then
                Sql = Sql & " and tipo =1"
            ElseIf cboSedex.Text = "SEDEX10" Then
                Sql = Sql & " and tipo =2"
            Else
                Sql = Sql & " and tipo =99"
            End If
            
            Set snap_Procura = ocnBanco.Execute(Sql)
            
            If snap_Procura.BOF And snap_Procura.EOF Then
                lbl(4) = FormatCurrency(0, 2)
            Else
            
                snap_Procura.MoveFirst
                If snap_Procura("cidade") = lblCidade Then
                    lbl(4) = FormatCurrency(snap_Procura("valor1"), 2)
                Else
                    lbl(4) = FormatCurrency(snap_Procura("valor2"), 2)
                End If
                
            End If
            Set snap_Procura = Nothing

End Select
End Sub

Private Sub txt_LostFocus(index As Integer)
Select Case index
Case 1 'cliente
    MostraRotulo (1)

Case 5, 6
    MostraRotulo (3)
End Select
End Sub

Sub Mostra_Grid(cod As Integer)
Set snap_Procura = Nothing
Dim Linha As Integer
Dim Sql As String
grid(0).Clear
grid(0).Rows = 1
With grid(0)
    .Row = 0
    .col = 0
    .ColWidth(0) = 600
    .Text = "Item"
    .col = 1
    .ColWidth(1) = 800
    .Text = "Código"
    .col = 2
    .ColWidth(2) = 4000
    .Text = "Descrição"
    .col = 3
    .ColWidth(3) = 800
    .Text = "Qtde"
    .col = 4
    .ColWidth(4) = 800
    .Text = "Valor"
    .col = 5
    .ColWidth(5) = 800
    .Text = "Total"
    
End With
Sql = "select * from pf_orcamentoi where orci_orc = " & cod
snap_Procura.Open Sql, ocnBanco, adOpenKeyset, adLockOptimistic, adCmdText
If Not (snap_Procura.BOF And snap_Procura.EOF) Then
    snap_Procura.MoveFirst
    grid(0).Rows = snap_Procura.RecordCount + 1
    Linha = 1
    While Not snap_Procura.EOF
        With grid(0)
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
        End With
        Linha = Linha + 1
        snap_Procura.MoveNext
    Wend
End If
End Sub

'Sub Inserir_Produto(codigo As Integer, Optional tipo As Integer = 0)
Sub Inserir_Produto(codigo As Integer)
On Error GoTo erro
Dim Item As Integer
Dim snap
Dim snap_item
Dim snap_Servico
Dim vrCusto As Double
Dim vrPrev As Double
Dim Sql As String
Dim Qtde As Double ' quantidade  do item
Dim ValorProduto As Double

Dim bEncontrouItem As Boolean

Dim X As Integer


'verifico se o item a foi cadastrado
For X = 1 To grid(0).Rows - 1
    grid(0).col = 0
    grid(0).Row = X
    If Int(grid(0).Text) = lbl(10) Then
        'MsgBox "Ja existe esse produto.", vbInformation, "Inserir Produto"
        bEncontrouItem = True
        Exit Sub
    End If
Next X

'10/01/2013 Implementado o update
If bEncontrouItem Then
    'Sql = "update pf_orcamentoi set orci_servcod = '" & txt(20) & "', orci_servdesc = '" & txt(15) & "', orci_servtexto = '" & txt(16) & "', orci_servqtdep = '" & txt(17) & "', orci_servvrvndp = '" & txt(18) & "' "
    'Sql = Sql & "where orci_orc =" & txt(0) & " and orci_item = " & LBL(7)
    'ocnBanco.Execute Sql
Else
    ' caso nao exista o item .......
    Sql = "select max(orcI_item)+1 from pf_orcamentoI where " & vbCrLf
    Sql = Sql & "orci_orc =" & txt(0) & "" ' codigo do orçamento
    snap_item = ocnBanco.Execute(Sql)
    If IsNull(snap_item(0)) Then
        Item = 1
    Else
        Item = snap_item(0)
    End If
    
    ValorProduto = CDbl(lbl(5)) - CDbl(lbl(5)) / 100 * CDbl(txt(8))
    
    'se vim do botaão gravar.....
    Sql = "insert into pf_OrcamentoI (orci_orc, orci_item, orci_servcod, orci_servdesc,  OrcI_ServQtdeP, orci_servVrvndP, orci_percentualDesconto, orcI_ListaPreco) values("
    Sql = Sql & "'" & txt(0) & "','" & Item & "','" & lbl(7) & "','" & txt(9) & "','" & txt(7) & "','" & ValorProduto & "','" & txt(8) & "','" & txt(6) & "')"
    ocnBanco.Execute Sql
    'adiciona ao grid
End If

Limpar_produto
Exit Sub
erro:
MsgBox Err.Description, vbCritical, "inserir_produto"
End Sub

Sub Limpar_produto()
lbl(1) = "" 'desc produto
lbl(5) = "" 'valor
lbl(7) = "" 'cod produto
lbl(10) = "" 'itemid
txt(5) = ""
txt(7) = ""
txt(8) = ""
txt(9) = ""
End Sub

Sub remover_Produto(int_Produto As Double)
Dim Sql As String
Dim X As Integer
'procurar se o aluno ja existe no grid
grid(0).col = 1
For X = 1 To grid(0).Rows - 1
    grid(0).Row = X
    grid(0).col = 1
    If Int(grid(0).Text) = int_Produto Then
        'excluir
        Sql = "delete from pf_OrcamentoI where orci_orc=" & txt(0) & " and orci_servcod=" & int_Produto
        ocnBanco.Execute (Sql)
        Mostra_Grid txt(0)
        'view_aluno 1
        Exit For
    End If
Next X
End Sub

Sub Total_Orcamento()
Dim vrTotal As Double
Dim Sql As String
If IsNumeric(txt(0)) Then
    'If IsNumeric(txt(22)) = True Then
        Sql = "SELECT orci_orc, sum( orci_servvrvndp * orci_servqtdep ) as Total" & vbCrLf
        Sql = Sql & "From pf_OrcamentoI" & vbCrLf
        Sql = Sql & "where orci_orc=" & txt(0) & vbCrLf
        Sql = Sql & "group by orci_orc" & vbCrLf
        Set snap_Procura = ocnBanco.Execute(Sql)
        If Not (snap_Procura.BOF And snap_Procura.EOF) Then  ' existe dados
            vrTotal = snap_Procura("total")
            lbl(3) = CDbl(FormatCurrency(vrTotal, 2)) + CDbl(lbl(4))
            'Desconto_Orcamento
        Else
            lbl(3) = FormatCurrency(0, 2)
        End If
    'Else
        'lbl(3) = 0
    'End If
End If
End Sub

