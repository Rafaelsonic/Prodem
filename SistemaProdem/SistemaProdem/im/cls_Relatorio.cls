VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Relatorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarBanco As String 'local copy
Private mvarBancoSenha As String 'local copy
Private mvarRelatorio As String 'local copy
Private mvarFormula As String 'local copy
'local variable(s) to hold property value(s)
Private mvarTitulo As String 'local copy
'local variable(s) to hold property value(s)
Private mvarEMailPara As String 'local copy
Private mvarEMailTitulo As String 'local copy
Private mvarEMailMenssasgem As String 'local copy
'local variable(s) to hold property value(s)



Public Property Let EMailMenssasgem(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EMailMenssasgem = 5
    mvarEMailMenssasgem = vData
End Property


Public Property Get EMailMenssasgem() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EMailMenssasgem
    EMailMenssasgem = mvarEMailMenssasgem
End Property



Public Property Let EMailTitulo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EMailTitulo = 5
    mvarEMailTitulo = vData
End Property


Public Property Get EMailTitulo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EMailTitulo
    EMailTitulo = mvarEMailTitulo
End Property



Public Property Let EMailPara(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EMailPara = 5
    mvarEMailPara = vData
End Property


Public Property Get EMailPara() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EMailPara
    EMailPara = mvarEMailPara
End Property








Public Sub Export(sDestino As String)
On Error GoTo erro
    Dim report As CRAXDRT.report
    
    Dim FormReport As New im_relatorio
    Dim CrxApplication As New CRAXDRT.Application
    Dim CrxDatabaseTable As CRAXDRT.DatabaseTable
    Dim Crxsubreport As CRAXDRT.report
    Set report = CrxApplication.OpenReport(mvarRelatorio, 1)
    
    For Each CrxDatabaseTable In report.Database.Tables
        CrxDatabaseTable.Location = mvarBanco
    Next CrxDatabaseTable
    
    report.RecordSelectionFormula = mvarFormula
    
    report.ExportOptions.DestinationType = crEDTDiskFile
    report.ExportOptions.DiskFileName = sDestino ' "c:\a.pdf"
    report.ExportOptions.FormatType = crEFTPortableDocFormat
    report.ExportOptions.PDFExportAllPages = True
    
'report.e
    report.Export False
    
    
    Set CrxApplication = Nothing

    Set CrxDatabaseTable = Nothing

    Set Crxsubreport = Nothing
Exit Sub
erro:
MsgBox Err.Description, vbCritical, "Vizualizar"

End Sub


Public Sub Vizualizar()
On Error GoTo erro
    Dim x As Integer
    Dim report As CRAXDRT.report
    
    Dim FormReport As New im_relatorio
    Dim CrxApplication As New CRAXDRT.Application
    Dim CrxDatabaseTable As CRAXDRT.DatabaseTable
    Dim Crxsubreport As CRAXDRT.report
    Set report = CrxApplication.OpenReport(mvarRelatorio, 1)
    
    For Each CrxDatabaseTable In report.Database.Tables
        CrxDatabaseTable.Location = mvarBanco
        CrxDatabaseTable.SetLogOnInfo "", mvarBanco, "", mvarBancoSenha
    Next CrxDatabaseTable
    
    
    
Dim nomeSUB(12) As String

nomeSUB(0) = "cliente"
nomeSUB(1) = "orcamentos"
nomeSUB(2) = "Receber"
nomeSUB(3) = "titulo"
nomeSUB(4) = ""
nomeSUB(5) = ""
nomeSUB(6) = ""
nomeSUB(7) = ""
nomeSUB(8) = ""
nomeSUB(9) = ""
nomeSUB(10) = ""

For x = 0 To 10
    Set Crxsubreport = SubRelat(report, nomeSUB(x))
Next x
    
    'Set Crxsubreport = report.OpenSubreport("cliente")
    'For Each CrxDatabaseTable In Crxsubreport.Database.Tables
        'CrxDatabaseTable.Location = mvarBanco
    'Next CrxDatabaseTable
    
    report.RecordSelectionFormula = mvarFormula
    
    
    'report.PrinterSetup mvarhWnd
    report.PrinterSetup 0
    FormReport.Caption = mvarTitulo
               
    FormReport.CRViewer1.ReportSource = report
    FormReport.CRViewer1.Top = 0
    FormReport.CRViewer1.Left = 0
    FormReport.CRViewer1.Height = ScaleHeight
    FormReport.CRViewer1.Width = ScaleWidth
    Screen.MousePointer = vbArrow
    FormReport.CRViewer1.ViewReport
    FormReport.WindowState = 2
    FormReport.Show 0
    
    
    Set CrxApplication = Nothing

    Set CrxDatabaseTable = Nothing

    Set Crxsubreport = Nothing
Exit Sub
erro:
MsgBox Err.Description, vbCritical, "Vizualizar"
End Sub


Public Sub EMail()
On Error GoTo erro
    Dim report As CRAXDRT.report
    
    Dim FormReport As New im_relatorio
    Dim CrxApplication As New CRAXDRT.Application
    Dim CrxDatabaseTable As CRAXDRT.DatabaseTable
    Dim Crxsubreport As CRAXDRT.report
    Set report = CrxApplication.OpenReport(mvarRelatorio, 1)
    
    For Each CrxDatabaseTable In report.Database.Tables
        CrxDatabaseTable.Location = mvarBanco
    Next CrxDatabaseTable
    
    report.RecordSelectionFormula = mvarFormula
    
    FormReport.Caption = mvarTitulo
               
    FormReport.CRViewer1.ReportSource = report
    FormReport.CRViewer1.Top = 0
    FormReport.CRViewer1.Left = 0
    FormReport.CRViewer1.Height = ScaleHeight
    FormReport.CRViewer1.Width = ScaleWidth
    Screen.MousePointer = vbArrow
    FormReport.CRViewer1.ViewReport
    FormReport.WindowState = 2
    'FormReport.Show 0
             
    report.ExportOptions.DestinationType = crEDTEMailMAPI
    report.ExportOptions.FormatType = crEFTPortableDocFormat
    report.ExportOptions.MailToList = mvarEMailPara
    report.ExportOptions.MailSubject = mvarEMailTitulo
    report.ExportOptions.MailMessage = mvarEMailMenssasgem
    'FormReport.Show (modal)
    report.Export False
    Set report = Nothing
    
    Set CrxApplication = Nothing
    Set CrxDatabaseTable = Nothing
    Set Crxsubreport = Nothing
    MsgBox "E-MAIL Enviado com sucesso", vbInformation, "EMAIL"
    
Exit Sub
erro:
MsgBox Err.Description, vbCritical, "E-Mail"
End Sub







Public Property Let SenhaBanco(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Banco = 5
    mvarBancoSenha = vData
End Property


Public Property Get SenhaBanco() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Banco
    Banco = mvarBancoSenha
End Property






Public Property Let Formula(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Formula = 5
    mvarFormula = vData
End Property


Public Property Get Formula() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Formula
    Formula = mvarFormula
End Property



Public Property Let relatorio(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Relatorio = 5
    mvarRelatorio = vData
End Property


Public Property Get relatorio() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Relatorio
    relatorio = mvarRelatorio
End Property



Public Property Let Banco(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Banco = 5
    mvarBanco = vData
End Property


Public Property Get Banco() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Banco
    Banco = mvarBanco
End Property



Public Property Let titulo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Titulo = 5
    mvarTitulo = vData
End Property


Public Property Get titulo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Titulo
    titulo = mvarTitulo
End Property

Function SubRelat(relatorio As CRAXDRT.report, nomeSUB As String) As CRAXDRT.report
On Error Resume Next
Dim CrxDatabaseTable As CRAXDRT.DatabaseTable
Dim Crxsubreport As CRAXDRT.report

Set Crxsubreport = relatorio.OpenSubreport(nomeSUB)
    For Each CrxDatabaseTable In Crxsubreport.Database.Tables
        CrxDatabaseTable.Location = mvarBanco
    Next CrxDatabaseTable

Set SubRelat = Crxsubreport
End Function
