Attribute VB_Name = "Helper"
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

Public Function FSOCreateFolder2(strPath As String) As Boolean
    
    Static fso As New FileSystemObject
    
    If Not fso.FolderExists(fso.GetParentFolderName(strPath)) Then
        'walk back up until you find one that exists
        FSOCreateFolder2 fso.GetParentFolderName(strPath)
    End If
    
    fso.CreateFolder strPath
    
End Function

Public Function GetTimeStamp() As String
    GetTimeStamp = Format(CStr(Now), "yyyy-mm-dd_hh.mm.ss")
End Function

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