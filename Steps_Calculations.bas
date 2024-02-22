Attribute VB_Name = "Steps_Calculations"


'--- Calculations ---


Public Sub EmailToPDF(Item As Object, outputPath As String)
    On Error GoTo ErrorHandler
    
    If Item.Body = "" Then
    
        DisplayWindowsNotification "Error - Email", "Body is empty"
    
    Else
    
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        
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
    
    End If
    
ExitSub:
        Exit Sub

ErrorHandler:
    DisplayWindowsNotification "Error - " & Err.Number, "Email to PDF failed - " & Err.Description
    Resume ExitSub
End Sub


' Zip Calculations Folder
Public Sub ZipAllFilesInFolder(zippedFileFullName As String, folderToZipPath As String)
            
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objFiles = fso.GetFolder(folderToZipPath & "\Calculations").Files
        
    If objFiles.Count > 0 Then
    
        Dim ShellApp As Object
        Set ShellApp = CreateObject("Shell.Application")
        
        DisplayWindowsNotification "Zip", "Creating"
        'Create an empty zip file
        Open zippedFileFullName For Output As #1
        Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
        Close #1
        
        'Copy the files & folders into the zip file
        'ShellApp.NameSpace(zippedFileFullName).CopyHere ShellApp.NameSpace(folderToZipPath).Items 'copies only items within folder
        ShellApp.NameSpace(zippedFileFullName).CopyHere folderToZipPath & "\Calculations" 'copies folder and its contents
        
        DisplayWindowsNotification "Zip", "Completed"
        Set ShellApp = Nothing
    
    Else
    
        DisplayWindowsNotification "Error - Zip", "Calculations is empty"
    
    End If

End Sub
