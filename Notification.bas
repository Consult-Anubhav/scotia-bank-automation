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

