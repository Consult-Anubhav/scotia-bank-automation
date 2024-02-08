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
  
    Dim ExApp As Excel.Application
    Dim ExWbk1, ExWbk2, ExWbk3 As Workbook
    Set ExApp = New Excel.Application
    ExApp.Visible = False
    
    'Start - Generate CCD extract
    DisplayWindowsNotification "CCD Extract", "Opening excel"
    Set ExWbk2 = ExApp.Workbooks.Open("C:\wamp64\www\~Consult Anubhav Projects\scotia-bank-automation\DF_DeMinimis_Extract (01012023-12312023).xlsm")
    
    DisplayWindowsNotification "CCD Extract", "CopyAndTrimSpecialEntity"
    ExWbk2.Application.Run "CopyAndTrimSpecialEntity.CopyAndTrimSpecialEntity"
    
    DisplayWindowsNotification "CCD Extract", "Saving File"
    ExWbk2.Close SaveChanges:=True
    'End
    
    'Start - Generate K2
    DisplayWindowsNotification "K2 Extract", "Opening excel"
    Set ExWbk3 = ExApp.Workbooks.Open("C:\wamp64\www\~Consult Anubhav Projects\scotia-bank-automation\K2 and Portal Data Summary_Jan 1 2022 - Dec 31 2023.xlsm")
    
    DisplayWindowsNotification "K2 Extract", "CCDExtractCSV"
    ExWbk3.Application.Run "Module1.CCDExtractCSV"
    
    DisplayWindowsNotification "K2 Extract", "CFCTE"
    ExWbk3.Application.Run "Module2.CFCTE"
    
    DisplayWindowsNotification "K2 Extract", "Saving File"
    ExWbk3.Close SaveChanges:=True
    'End
    
    'Start - Send Test email response
    'DisplayWindowsNotification "LATAM", "Saving File"
    'Set ExWbk1 = ExApp.Workbooks.Open("C:\wamp64\www\~Consult Anubhav Projects\scotia-bank-automation\SendBulkEmail.xlsm")
    'MsgBox "opening SendEmail"
    'ExWbk1.Application.Run "Module1.SendResponse1", Item.Body, Item.SenderEmailAddress
    'ExWbk1.Close SaveChanges:=True
    'End
    
    'Start - Send response email
    DisplayWindowsNotification "Response Email", "sending"
    Set ExWbk1 = ExApp.Workbooks.Open("C:\wamp64\www\~Consult Anubhav Projects\scotia-bank-automation\SendBulkEmail.xlsm")
    ExWbk1.Application.Run "Module1.SendResponse1", Item.Body, Item.SenderEmailAddress
    ExWbk1.Close SaveChanges:=True
    'End
Else
    DisplayWindowsNotification "Ignored", "irrelative email"
End If

ExitNewItem:
    Exit Sub

ErrorHandler:
    MsgBox Err.Number & " - " & Err.Description
    Resume ExitNewItem
End Sub

