Attribute VB_Name = "Helper"
Public Function FSOCreateFolder2(strPath As String) As Boolean
    Static FSO As New FileSystemObject
    If Not FSO.FolderExists(FSO.GetParentFolderName(strPath)) Then
        'walk back up until you find one that exists
        FSOCreateFolder2 FSO.GetParentFolderName(strPath)
    End If
    FSO.CreateFolder strPath
End Function
