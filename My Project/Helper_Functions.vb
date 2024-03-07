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
End Module
