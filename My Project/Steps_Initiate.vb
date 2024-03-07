Imports Microsoft.VisualBasic.FileIO
Imports Microsoft.VisualBasic.CompilerServices
Imports System.IO
Imports Microsoft.Office.Interop.Outlook

Module Steps_Initiate
    Public Sub DownloadEmailAttachments(Item As Object, attachPath As String)
        Try
            Dim fso As New FileSystem()

            DisplayWindowsNotification("Attachments", "Downloading")

            If Not fso.DirectoryExists(attachPath) Then
                FSOCreateFolder2(attachPath)
            End If

            Dim attachmentsCount As Integer = Item.Attachments.Count
            DisplayWindowsNotification($"{attachmentsCount} Attachments", "Downloading")

            For Each attachment As Attachment In Item.Attachments
                Dim dirPath As String = GetAttachmentFolderPath(attachment.FileName, attachPath)
                attachment.SaveAsFile(Path.Combine(dirPath, $"{Path.GetFileNameWithoutExtension(attachment.FileName)}_{GetTimeStamp()}.{Path.GetExtension(attachment.FileName)}"))
            Next
        Catch ex As System.Exception
            DisplayWindowsNotification($"Error - {ex.HResult}", $"Downloading attachments failed - {ex.Message}")
        End Try
    End Sub

    Private Function GetAttachmentFolderPath(fileName As String, attachPath As String) As String
        Select Case True
            Case fileName.StartsWith("bookingpoint") OrElse fileName.StartsWith("CCD Extract") OrElse fileName.StartsWith("CFTCExtract")
                Return Path.Combine(attachPath, "Calculations")
            Case fileName.StartsWith("FX (FORWARDS)")
                Return Path.Combine(attachPath, FakeOPICSPath)
            Case fileName.StartsWith("Cartera Fwd") OrElse fileName.StartsWith("DeMinimisReport_Colombia") OrElse fileName.StartsWith("DERIVATIVES MEXICO Ene") OrElse fileName.StartsWith("Dodd-Frank CCS") OrElse fileName.StartsWith("Dodd-Frank IRS") OrElse fileName.StartsWith("MINIMIS Calculation Template (Chile)")
                Return Path.Combine(attachPath, FakeLATAMCFTCPath)
            Case fileName.StartsWith("CHILE US Person") OrElse fileName.StartsWith("Colombia_US Person identified_Dic") OrElse fileName.StartsWith("US PERSON LIST")
                Return Path.Combine(attachPath, FakeLATAMUSPPath)
            Case Else
                Return Path.Combine(attachPath, FakeK2Path)
        End Select
    End Function

    Public Sub TestCopy()
        CopyPreviousReports("C:\wamp64\www\~Consult Anubhav Projects\Scotia Bank\~Scotia-Bank-Root\2024\Jan", "C:\wamp64\www\~Consult Anubhav Projects\Scotia Bank\~Scotia-Bank-Root\2024\Feb")
    End Sub

    Public Sub CopyPreviousReports(inputDir As String, outputDir As String)
        Try
            Dim fso As New FileSystem()

            DisplayWindowsNotification("Previous Reports", "Copying")

            ' LATAM
            CopyFileIfNotExists(inputDir & "\Latam De Minimis Calculation\SupportdataforMINIMIS Report Jan 1, 2023 to Dec 31, 2023.xlsx", outputDir & "\Latam De Minimis Calculation\")

            ' OPICS
            CopyFileIfNotExists(inputDir & "\OPICS Scotia Investments Jamaica Limited\FX (FORWARDS).prn.xlsx", outputDir & "\OPICS Scotia Investments Jamaica Limited\")

            ' K2
            CopyFileIfNotExists(inputDir & "\Supporting Files K2 and Murex\K2\K2 and Portal Data Summary_Jan 1 2022 - Dec 31 2023.xlsx", outputDir & "\Supporting Files K2 and Murex\K2\")

            ' Murex
            CopyFileIfNotExists(inputDir & "\Supporting Files K2 and Murex\Murex\DF_DeMinimis_Extract (01012023-12312023).xlsx", outputDir & "\Supporting Files K2 and Murex\Murex\")
        Catch ex As System.Exception
            DisplayWindowsNotification($"Error - {ex.HResult}", $"Copying previous reports failed - {ex.Message}")
        End Try
    End Sub

    Private Sub CopyFileIfNotExists(sourceFilePath As String, destinationFilePath As String)
        Dim fso As New FileSystem()

        If Not fso.DirectoryExists(Path.GetDirectoryName(destinationFilePath)) Then
            FSOCreateFolder2(Path.GetDirectoryName(destinationFilePath))
        End If

        If Not fso.FileExists(destinationFilePath) Then
            fso.CopyFile(sourceFilePath, destinationFilePath)
        End If
    End Sub
End Module
