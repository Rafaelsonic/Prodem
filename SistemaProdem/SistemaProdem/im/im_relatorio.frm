VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form im_relatorio 
   Caption         =   "Form1"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   0   'False
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "im_relatorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
If Me.Width > 150 Then CRViewer1.Width = Me.Width - 100
If Me.Height > 550 Then CRViewer1.Height = Me.Height - 500
End Sub
