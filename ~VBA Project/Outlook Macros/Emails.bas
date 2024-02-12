Attribute VB_Name = "Emails"
Public Sub EmailToPDF(Item As Outlook.MailItem, outputPath As String)
    On Error GoTo ErrorHandler
    
    Dim FSO As New FileSystemObject
    
    'Save email as MIME HTML Archive file
    Dim tempFilePath, tempFileName As String
    tempFilePath = Environ("temp") & "\Scotia\Calculations\"
    
    'create directory if not exists
    If Not FSO.FolderExists(tempFilePath) Then
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
    If Not FSO.FolderExists(tempFilePath) Then
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
    DisplayWindowsNotification "Error", "Email to PDF failed"
    DisplayWindowsNotification Err.Number, Err.Description
    Resume ExitSub
End Sub

Public Sub TestEmail()
    Dim FSO As New FileSystemObject
    
    'Save email as MIME HTML Archive file
    Dim ok As Boolean
    Dim tempFilePath, tempFileName As String
    tempFilePath = Environ("temp") & "\Scotia\Calculations\"
    
    'create directory if not exists
    If Not FSO.FolderExists(tempFilePath) Then
        FSOCreateFolder2 CStr(tempFilePath)
    End If
    
    'Dim ExApp As Excel.Application
    'Dim ExWbk As Workbook
    
    'Set ExApp = New Excel.Application
    
    'DisplayWindowsNotification "Response Email", "sending"
    'Set ExWbk = ExApp.Workbooks.Open("C:\wamp64\www\~Consult Anubhav Projects\scotia-bank-automation\SendBulkEmail.xlsm")
    
    'ExWbk.Application.Run "Email.SendResponse1", Item.Body, Item.SenderEmailAddress
    
    'ExWbk.Close SaveChanges:=True
End Sub


'MessageInfo = "" & _
        "Sender : " & Item.SenderEmailAddress & vbCrLf & _
        "Sent : " & Item.SentOn & vbCrLf & _
        "Received : " & Item.ReceivedTime & vbCrLf & _
        "Subject : " & Item.Subject & vbCrLf & _
        "Size : " & Item.Size & vbCrLf & _
        "Message Body : " & vbCrLf & Item.Body
    'Result = MsgBox(MessageInfo, vbOKOnly, "New Message Received")
    'Debug.Print "Hello"
    'Debug.Print Item.Subject
    
    'Start - Send Test email response
    'DisplayWindowsNotification "LATAM", "Saving File"
    'Set ExWbk1 = ExApp.Workbooks.Open("C:\wamp64\www\~Consult Anubhav Projects\scotia-bank-automation\SendBulkEmail.xlsm")
    'MsgBox "opening SendEmail"
    'ExWbk1.Application.Run "Module1.SendResponse1", Item.Body, Item.SenderEmailAddress
    'ExWbk1.Close SaveChanges:=True
    'End
