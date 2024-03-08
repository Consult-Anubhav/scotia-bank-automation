Imports System.IO
Imports System.Windows.Forms

Module Helper_Functions
    ' Show notifications
    Public Sub DisplayWindowsNotification(Subject As String, Comment As String)
        Dim notify As New NotifyIcon()

        notify.Icon = SystemIcons.Information
        notify.Visible = True
        notify.ShowBalloonTip(10, Subject, Comment, ToolTipIcon.None)
    End Sub

    Public Sub UpdateLabel(subject As String, comment As String)
        ' Assuming you have a form called Form1 with a label named lblNotification
        If Form1.rTxtStatus.InvokeRequired Then
            Form1.rTxtStatus.Invoke(Sub() UpdateLabel(subject, comment))
        Else
            ' Update the label text
            Form1.rTxtStatus.Text = $"Subject: {subject}{Environment.NewLine}Comment: {comment}"
        End If
    End Sub

    ' Create nested directories
    Public Sub FSOCreateFolder2(strPath As String)
        Dim fso As New DirectoryInfo(strPath)
        If Not fso.Parent.Exists Then
            ' Walk back up until you find one that exists
            FSOCreateFolder2(fso.Parent.FullName)
        End If
        Directory.CreateDirectory(strPath)
    End Sub

    ' GetTimeStamp
    Public Function GetTimeStamp() As String
        Return DateTime.Now.ToString("yyyy-MM-dd_hh.mm.ss")
    End Function

    ' GetEmailMonthYear
    Public Function GetEmailMonthYear(ItemSubject As String) As String
        Return ItemSubject.Substring(Len(FakeEmailSubject()))
    End Function

    ' GetEmailYear
    Public Function GetEmailYear(emailMonthYear As String) As String
        Dim ItemStr As String = "1 " & emailMonthYear
        Return DateTime.Parse(ItemStr).ToString("yyyy")
    End Function

    ' GetEmailMonth
    Public Function GetEmailMonth(emailMonthYear As String) As String
        Dim ItemStr As String = "1 " & emailMonthYear
        Return DateTime.Parse(ItemStr).ToString("MMM")
    End Function

    ' GetPreviousMonthYear
    Public Function GetPreviousMonthYear(emailMonthYear As String) As String
        Dim ItemStr As String = "1 " & emailMonthYear
        Return DateTime.Parse(ItemStr).AddMonths(-1).ToString("MMM yy")
    End Function

    ' GetPreviousYear
    Public Function GetPreviousYear(emailMonthYear As String) As String
        Dim ItemStr As String = "1 " & emailMonthYear
        Return DateTime.Parse(ItemStr).AddMonths(-1).ToString("yyyy")
    End Function

    ' GetPreviousMonth
    Public Function GetPreviousMonth(emailMonthYear As String) As String
        Dim ItemStr As String = "1 " & emailMonthYear
        Return DateTime.Parse(ItemStr).AddMonths(-1).ToString("MMM")
    End Function
    Function SelectSingleFile() As String
        ' Show a file dialog to select a single file
        Dim fileDialog As New OpenFileDialog()

        ' Set the dialog title and filter if needed
        fileDialog.Title = "Select File"
        fileDialog.Filter = "All Files|*.*" ' You can adjust the filter as per your requirements

        If fileDialog.ShowDialog() = DialogResult.OK Then
            ' Return the selected file path
            Return fileDialog.FileName
        Else
            ' If the user cancels the dialog, return an empty string
            Return ""
        End If
    End Function
    Function SelectMultipleFiles() As String()
        ' Show a file dialog to select multiple files
        Dim fileDialog As New OpenFileDialog()
        fileDialog.Multiselect = True

        ' Set the dialog title and filter if needed
        fileDialog.Title = "Select Files"
        fileDialog.Filter = "All Files|*.*" ' You can adjust the filter as per your requirements

        If fileDialog.ShowDialog() = DialogResult.OK Then
            ' Return the selected file paths
            Return fileDialog.FileNames
        Else
            ' If the user cancels the dialog, return an empty array
            Return New String() {}
        End If
    End Function
    Function SelectFolder() As String
        ' Show a folder browser dialog to select the folder location
        Dim folderBrowserDialog As New FolderBrowserDialog()

        If folderBrowserDialog.ShowDialog() = DialogResult.OK Then
            ' Return the selected folder path
            Return folderBrowserDialog.SelectedPath
        Else
            ' If the user cancels the dialog, return an empty string
            Return ""
        End If
    End Function
    Function SelectFileWithMessage(message As String, filters As String) As String
        ' Display a message to the user
        MessageBox.Show(message, "Select File", MessageBoxButtons.OK, MessageBoxIcon.Information)

        ' Show a file dialog to select a file
        Dim fileDialog As New OpenFileDialog()

        ' Set the dialog title
        fileDialog.Title = "Select File - " & message

        ' Add the specified filters for file types
        fileDialog.Filter = filters 'e.g.: "Word Documents (*.doc, *.docx)|*.doc;*.docx|PDF Files (*.pdf)|*.pdf|Excel Files (*.xlsx)|*.xlsx|CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"


        If fileDialog.ShowDialog() = DialogResult.OK Then
            ' Return the selected file path
            Return fileDialog.FileName
        Else
            ' If the user cancels the dialog, return an empty string
            Return ""
        End If
    End Function
End Module
