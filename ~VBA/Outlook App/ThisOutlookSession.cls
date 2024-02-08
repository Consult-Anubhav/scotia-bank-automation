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

Private Sub inboxItems_ItemAdd(ByVal Item As Object)
    On Error GoTo ErrorHandler
    
    Dim Msg As Outlook.MailItem
    Dim MessageInfo
    Dim Result
    
    DisplayWindowsNotification "New Email", "Checking Email"
    If TypeName(Item) = "MailItem" And Item.Subject Like "*Scotia Report - *" Then
        
        'Start - Generate CCD Extract
            GenerateCCDExtract
        'End
        
        'Start - Generate K2 Extract
            GenerateK2Extract
        'End
        
        'Start - Send response email
            TestEmail
        'End
    Else
        DisplayWindowsNotification "Ignored", "irrelative email"
    End If
    
ExitNewItem:
        Exit Sub
    
ErrorHandler:
        DisplayWindowsNotification Err.Number, Err.Description
        Resume ExitNewItem
End Sub

