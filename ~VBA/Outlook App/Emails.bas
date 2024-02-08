Attribute VB_Name = "Emails"
Public Sub TestEmail()
    Dim ExApp As Excel.Application
    Dim ExWbk As Workbook
    
    Set ExApp = New Excel.Application
    
    DisplayWindowsNotification "Response Email", "sending"
    Set ExWbk = ExApp.Workbooks.Open("C:\wamp64\www\~Consult Anubhav Projects\scotia-bank-automation\SendBulkEmail.xlsm")
    
    ExWbk.Application.Run "Email.SendResponse1", Item.Body, Item.SenderEmailAddress
    
    ExWbk.Close SaveChanges:=True
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
