Attribute VB_Name = "Steps_Calculations"


'--- Calculations ---


Public Sub EmailToPDF(Item As Outlook.MailItem, outputPath As String)
    On Error GoTo ErrorHandler
    
    Dim fso As New FileSystemObject
    
    'Save email as MIME HTML Archive file
    Dim tempFilePath, tempFileName As String
    tempFilePath = Environ("temp") & "\Scotia\Calculations\"
    
    'create directory if not exists
    If Not fso.FolderExists(tempFilePath) Then
        FSOCreateFolder2 CStr(tempFilePath)
    End If
          
    tempFileName = tempFilePath & Format(CStr(Now), "yyyy-mm-dd_hh.mm.ss") & ".mht"
    
    DisplayWindowsNotification "Saving", "email"
    Item.SaveAs tempFileName, olMHTML
    
    'Convert MHT to PDF
    Dim wordapp As Word.Application
    Set wordapp = New Word.Application
    wordapp.Visible = False
    
    wordapp.Documents.Open tempFileName, ConfirmConversions:=False
    
    tempFilePath = outputPath & "\Calculations\"
    
    'create directory if not exists
    If Not fso.FolderExists(tempFilePath) Then
        FSOCreateFolder2 CStr(tempFilePath)
    End If
          
    tempFileName = tempFilePath & "ThisEmail_" & Format(CStr(Now), "yyyy-mm-dd_hh.mm.ss") & ".pdf"
    
    wordapp.ActiveDocument.ExportAsFixedFormat _
        OutputFileName:=tempFileName, ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, BitmapMissingFonts:=True
        
    wordapp.Documents.Close
    wordapp.Quit
    
    'Delete MHT file
    'Kill tempFileName
    
ExitSub:
        Exit Sub

ErrorHandler:
    DisplayWindowsNotification "Error - " & Err.Number, "Email to PDF failed - " & Err.Description
    Resume ExitSub
End Sub


