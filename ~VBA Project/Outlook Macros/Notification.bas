Attribute VB_Name = "Notification"
Public Sub DisplayWindowsNotification(Subject As String, Comment As String)

Dim WsShell     As Object: Set WsShell = CreateObject("WScript.Shell")
Dim strCommand  As String

strCommand = "powershell.exe -Command " & Chr(34) & "& { "
strCommand = strCommand & "[reflection.assembly]::loadwithpartialname('System.Windows.Forms')"
strCommand = strCommand & "; [reflection.assembly]::loadwithpartialname('System.Drawing')"
strCommand = strCommand & "; $notify = new-object system.windows.forms.notifyicon"
strCommand = strCommand & "; $notify.icon = [System.Drawing.SystemIcons]::Information"
strCommand = strCommand & "; $notify.visible = $true"
strCommand = strCommand & "; $notify.showballoontip(10,'" & Subject & "','" & Comment & "',[system.windows.forms.tooltipicon]::None)"
strCommand = strCommand & " }" & Chr(34)
WsShell.Run strCommand, 0, False

End Sub

Public Sub RunMe()
Dim Msg As Outlook.MailItem
Dim MessageInfo
Dim Result
DisplayWindowsNotification "New Email", "Checking Email"
'If TypeName(Item) = "MailItem" And Item.Subject Like "*Scotia Report - *" Then
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
    'ExApp.AskToUpdateLinks = False
    'ExApp.DisplayAlerts = False
    ExApp.Visible = True
    
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
    ExWbk1.Application.Run "Email.SendResponse1", Item.Body, Item.SenderEmailAddress
    ExWbk1.Close SaveChanges:=True
    'End
'End If
End Sub


