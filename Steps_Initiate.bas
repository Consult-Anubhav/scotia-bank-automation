Attribute VB_Name = "Steps_Initiate"
Public Sub DownloadEmailAttachments(Item As Object, attachPath As String)
    Dim dirPath As String
    Dim olAttachment As Attachment
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim i As Integer
    i = 1
    
    DisplayWindowsNotification "Attachments", "Downloading"
    
    If Not fso.FolderExists(attachPath) Then
        FSOCreateFolder2 (attachPath)
    End If
    
    If Not fso.FolderExists(attachPath & "\Calculations") Then
        FSOCreateFolder2 (attachPath & "\Calculations")
    End If
    
    If Not fso.FolderExists(attachPath & "\" & fakeK2Path) Then
        FSOCreateFolder2 (attachPath & "\" & fakeK2Path)
    End If
    
    If Not fso.FolderExists(attachPath & "\" & fakeOPICSPath) Then
        FSOCreateFolder2 (attachPath & "\" & fakeOPICSPath)
    End If
    
    If Not fso.FolderExists(attachPath & "\" & fakeLATAMCFTCPath) Then
        FSOCreateFolder2 (attachPath & "\" & fakeLATAMCFTCPath)
    End If
    
    If Not fso.FolderExists(attachPath & "\" & fakeLATAMUSPPath) Then
        FSOCreateFolder2 (attachPath & "\" & fakeLATAMUSPPath)
    End If
    
    DisplayWindowsNotification Item.Attachments.Count & " Attachments", "Downloading"
    
    For Each olAttachment In Item.Attachments
        i = i + 1

       Select Case True 'UCase(fso.GetExtensionName(olAttachment.FileName))
       
            'Case "XLSX", "XLSM", "CSV", "XLS"
                
            'K2
            Case olAttachment.FileName Like "bookingpoint*"
                dirPath = attachPath & "\" & fakeK2Path
              
            Case olAttachment.FileName Like "CCD Extract*"
                dirPath = attachPath & "\" & fakeK2Path

            Case olAttachment.FileName Like "CFTCExtract*"
                dirPath = attachPath & "\" & fakeK2Path
                
            'OPICS
            Case olAttachment.FileName Like "FX (FORWARDS*"
                dirPath = attachPath & "\" & fakeOPICSPath
                
            'LATAM
            Case olAttachment.FileName Like "Cartera Fwd*"
                dirPath = attachPath & "\" & fakeLATAMCFTCPath
                
            Case olAttachment.FileName Like "DeMinimisReport_Colombia*"
                dirPath = attachPath & "\" & fakeLATAMCFTCPath
                
            Case olAttachment.FileName Like "DERIVATIVES MEXICO Ene*"
                dirPath = attachPath & "\" & fakeLATAMCFTCPath
                
            Case olAttachment.FileName Like "Dodd-Frank CCS*"
                dirPath = attachPath & "\" & fakeLATAMCFTCPath
                
            Case olAttachment.FileName Like "Dodd-Frank IRS*"
                dirPath = attachPath & "\" & fakeLATAMCFTCPath
                
            Case olAttachment.FileName Like "MINIMIS Calculation Template (Chile)*"
                dirPath = attachPath & "\" & fakeLATAMCFTCPath
                
            Case olAttachment.FileName Like "CHILE US Person*"
                dirPath = attachPath & "\" & fakeLATAMUSPPath
                
            Case olAttachment.FileName Like "Colombia_US Person identified_Dic*"
                dirPath = attachPath & "\" & fakeLATAMUSPPath
                
            Case olAttachment.FileName Like "US PERSON LIST*"
                dirPath = attachPath & "\" & fakeLATAMUSPPath
                
       End Select
       
       olAttachment.SaveAsFile dirPath & "\" & fso.GetBaseName(olAttachment.FileName) & GetTimeStamp & "." & fso.GetExtensionName(olAttachment.FileName)

    Next olAttachment
End Sub

Public Sub TestCopy()
    CopyPreviousReports "C:\wamp64\www\~Consult Anubhav Projects\Scotia Bank\~Scotia-Bank-Root\2024\Jan", "C:\wamp64\www\~Consult Anubhav Projects\Scotia Bank\~Scotia-Bank-Root\2024\Feb"
End Sub


Public Sub CopyPreviousReports(inputDir As String, outputDir As String)
    
    Dim fso As Object
    Dim inLink, outLink As String
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    
    DisplayWindowsNotification "Previous Reports", "Copying"
    
    'LATAM
    inLink = "\Latam De Minimis Calculation\SupportdataforMINIMIS Report Jan 1, 2023 to Dec 31, 2023.xlsx"
    outLink = "\Latam De Minimis Calculation\"
    
    If Not fso.FolderExists(outputDir & outLink) Then
        FSOCreateFolder2 (outputDir & outLink)
    End If
    Call fso.CopyFile(inputDir & inLink, outputDir & outLink)  'outputDir[, overwrite]
    
    'OPICS
    inLink = "\OPICS Scotia Investments Jamaica Limited\FX (FORWARDS).prn.xlsx"
    outLink = "\OPICS Scotia Investments Jamaica Limited\"
    
    If Not fso.FolderExists(outputDir & outLink) Then
        FSOCreateFolder2 (outputDir & outLink)
    End If
    Call fso.CopyFile(inputDir & inLink, outputDir & outLink)
    
    'SCOTS
    inLink = ""
    outLink = ""
    'Call fso.CopyFile(inputDir & inLink, outputDir & outLink)
    
    'K2
    inLink = "\Supporting Files K2 and Murex\K2\K2 and Portal Data Summary_Jan 1 2022 - Dec 31 2023.xlsx"
    outLink = "\Supporting Files K2 and Murex\K2\"
    
    If Not fso.FolderExists(outputDir & outLink) Then
        FSOCreateFolder2 (outputDir & outLink)
    End If
    Call fso.CopyFile(inputDir & inLink, outputDir & outLink)
    
    'Murex
    inLink = "\Supporting Files K2 and Murex\Murex\DF_DeMinimis_Extract (01012023-12312023).xlsx"
    outLink = "\Supporting Files K2 and Murex\Murex\"
    
    If Not fso.FolderExists(outputDir & outLink) Then
        FSOCreateFolder2 (outputDir & outLink)
    End If
    Call fso.CopyFile(inputDir & inLink, outputDir & outLink)
    
End Sub
