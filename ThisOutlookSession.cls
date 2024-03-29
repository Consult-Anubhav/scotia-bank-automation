VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private WithEvents inboxItems As Outlook.Items
Attribute inboxItems.VB_VarHelpID = -1

Private Sub Application_Startup()

    DisplayWindowsNotification "Initiating", "New Email Event Litsener"
    
    Dim outlookApp As Outlook.Application
    Dim objectNS As Outlook.NameSpace
  
    Set outlookApp = Outlook.Application
    Set objectNS = outlookApp.GetNamespace("MAPI")
    Set inboxItems = objectNS.GetDefaultFolder(olFolderInbox).Items
End Sub

'Public Sub Application_ItemSend()

'End Sub


Private Sub inboxItems_ItemAdd(ByVal Item As Object)
    On Error GoTo ErrorHandler
    
    DisplayWindowsNotification "New Email", "Checking Email"
    
    If TypeName(Item) = "MailItem" And Item.Subject Like fakeEmailSubject Then
    
        'Declare Variables
        
        Dim inputDir, outputDir, _
            emailMonthYear, previousMonthYear, _
            emailYear, previousYear, _
            emailMonth, previousMonth As String
            
        'Assign Variables
        
        emailMonthYear = GetEmailMonthYear(Item.Subject)
        previousMonthYear = GetPreviousMonthYear(emailMonthYear & "")
        emailYear = GetEmailYear(emailMonthYear & "")
        previousYear = GetPreviousYear(emailMonthYear & "")
        emailMonth = GetEmailMonth(emailMonthYear & "")
        previousMonth = GetPreviousMonth(emailMonthYear & "")
        outputDir = fakeRootPath & "\" & emailYear & "\" & emailMonth
        inputDir = fakeRootPath & "\" & previousYear & "\" & previousMonth
        
        '--- Initiate directory template ---
        
            ' Download all attachments to output directory
            DownloadEmailAttachments Item, outputDir & ""
            
            ' copy report templates from input directory
            CopyPreviousReport inputDir & "", outputDir & ""
            
        '--- Latam ---
        
        '--- OPICS ---
        
        '--- SCOTS ---
        
        '--- K2 ---
        
            'Start - Generate K2 Extract
                GenerateK2Extract outputDir & ""
            'End
        
        '--- Mutex ---
        
            'Start - Generate CCD Extract
                GenerateMutexExtract outputDir & ""
            'End
        
        '--- Calculations ---
        
            'Start - Store email as pdf
                EmailToPDF Item, outputDir & ""
            'End
            
            'Start - Compress final reports
                ZipAllFilesInFolder outputDir & "\Calculations " & Format(CStr(Now), "yyyy-mm-dd_hh.mm.ss") & ".zip", outputDir & ""
            'End
            
    Else
        DisplayWindowsNotification "Ignored", "irrelative email"
    End If
    
ExitSub:
        Exit Sub
    
ErrorHandler:
        DisplayWindowsNotification Err.Number, Err.Description
        Resume ExitSub
End Sub

