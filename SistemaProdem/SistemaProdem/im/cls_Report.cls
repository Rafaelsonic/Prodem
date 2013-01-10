VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Private sPath_Banco As String
'Private sPath_relat  As String
'Private sFormula As String
Private report As CRAXDRT.report
Sub Vizualizar(sPath_Banco As String, sPath_relat As String, sFormula As String, Optional sTitulo As String)
'StrBanco = "E:\Work Fire\Reports\orcamento.rpt"
    Dim FormReport As New Form2

    Dim CrxApplication As New CRAXDRT.Application
    Dim CrxDatabaseTable As CRAXDRT.DatabaseTable
    Dim Crxsubreport As CRAXDRT.report
    
    Set report = CrxApplication.OpenReport(sPath_relat, 1)
    
    For Each CrxDatabaseTable In report.Database.Tables
        CrxDatabaseTable.Location = sPath_Banco
    Next CrxDatabaseTable
    
    report.RecordSelectionFormula = sFormula
    
    
    
    FormReport.Caption = sTitulo
    
    'Report.ParameterFields.Item(1).AddCurrentValue "Main Report Parameter"
    '        For Each CrxDatabaseTable In Crxsubreport.Database.Tables'
    '          CrxDatabaseTable.Location = "E:\Work Fire\banco\1bd_wf.mdb"
    '       Next CrxDatabaseTable
    
    
    ' dados para exportaçaõ
    report.ExportOptions.DestinationType = crEDTEMailMAPI
    report.ExportOptions.FormatType = crEFTPortableDocFormat
    report.ExportOptions.MailToList = "rafael.sp@gmail.com"
    report.ExportOptions.MailSubject = "Orçamento"
    report.ExportOptions.MailMessage = "Sr. Cliente, " & Chr(13) & Chr(13) & "     Favor confirmar o recebimento desse e-mail, aprovando ou reprovando." & Chr(13) & Chr(13) & "Obrigado," & Chr(13) & Chr(13) & Comprador & Chr(13) & "Colauto Adesivos e Massas Ltda" & Chr(13) & "Departamento de Compras" & Chr(13) & "+55 11 6946-4011" & Chr(13) & "+55 11 6969-7740 Fax" & Chr(13) & Chr(13) & Chr(13) & "Para visualizar esse pedido você necessita do Acrobat Reader. Faça o download gratuito: " & Chr(13) & "http://ardownload.adobe.com/pub/adobe/reader/win/6.x/6.0/ptb/AdbeRdr60_ptb_full.exe" & Chr(13) & Chr(13)

        
        'fjdklfdklfjdkfjdkfldjfkldjfklsdjfklds
        FormReport.CRViewer1.ReportSource = report
        FormReport.CRViewer1.Top = 0
        FormReport.CRViewer1.Left = 0
        FormReport.CRViewer1.Height = ScaleHeight
        FormReport.CRViewer1.Width = ScaleWidth
        Screen.MousePointer = vbArrow
        FormReport.CRViewer1.ViewReport
        FormReport.WindowState = 2
        FormReport.Show 0
        Exit Sub
        
        
        Form2.Show (modal)

                             

        MsgResp = MsgBox("Deseja realmente enviar este e-mail ?", vbYesNo, "Envio de E-Mail")

       

        If MsgResp = 6 Then

            report.Export False

        End If

        

        Set report = Nothing

        Set CrxApplication = Nothing

        Set CrxDatabaseTable = Nothing

        Set Crxsubreport = Nothing


End Sub

Private Sub Class_Initialize()
Set Exprotar = report
End Sub
