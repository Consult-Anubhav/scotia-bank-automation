Attribute VB_Name = "Helper_Functions"

' Show notifications
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

' Create nested directories
Public Function FSOCreateFolder2(strPath As String) As Boolean
    
    Static fso As New FileSystemObject
    
    If Not fso.FolderExists(fso.GetParentFolderName(strPath)) Then
        'walk back up until you find one that exists
        FSOCreateFolder2 fso.GetParentFolderName(strPath)
    End If
    
    fso.CreateFolder strPath
    
End Function

' GetTimeStamp
Public Function GetTimeStamp() As String
    GetTimeStamp = Format(CStr(Now), "yyyy-mm-dd_hh.mm.ss")
End Function

' GetEmailMonthYear
Public Function GetEmailMonthYear(ItemSubject As String) As String
    Dim ItemStr As String
    ItemStr = Mid(ItemSubject, Len(fakeEmailSubject), Len(ItemSubject))
    GetEmailMonthYear = ItemStr
End Function

' GetEmailYear
Public Function GetEmailYear(emailMonthYear As String) As String
    Dim ItemStr As String
    ItemStr = CDate("1 " & emailMonthYear)
    GetEmailYear = Format(ItemStr, "YYYY")
End Function

' GetEmailMonth
Public Function GetEmailMonth(emailMonthYear As String) As String
    Dim ItemStr As String
    ItemStr = CDate("1 " & emailMonthYear)
    GetEmailMonth = Format(ItemStr, "MMM")
End Function

' GetPreviousMonthYear
Public Function GetPreviousMonthYear(emailMonthYear As String) As String
    Dim ItemStr As String
    ItemStr = CDate("1 " & emailMonthYear)
    GetPreviousMonthYear = Format(DateAdd("M", -1, ItemStr), "MMM YY")
End Function

' GetPreviousYear
Public Function GetPreviousYear(emailMonthYear As String) As String
    Dim ItemStr As String
    ItemStr = CDate("1 " & emailMonthYear)
    GetPreviousYear = Format(DateAdd("M", -1, ItemStr), "YYYY")
End Function

' GetPreviousMonth
Public Function GetPreviousMonth(emailMonthYear As String) As String
    Dim ItemStr As String
    ItemStr = CDate("1 " & emailMonthYear)
    GetPreviousMonth = Format(DateAdd("M", -1, ItemStr), "MMM")
End Function

' Zip Calculations
Sub ZipAllFilesInFolder(zippedFileFullName, folderToZipPath)
    Dim ShellApp As Object
    Set ShellApp = CreateObject("Shell.Application")
  
    'Create an empty zip file
    Open zippedFileFullName For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
    
    'Copy the files & folders into the zip file
    'ShellApp.NameSpace(zippedFileFullName).CopyHere ShellApp.NameSpace(folderToZipPath).Items 'copies only items within folder
    ShellApp.NameSpace(zippedFileFullName).CopyHere folderToZipPath 'copies folder and its contents
     
    Set ShellApp = Nothing

End Sub
