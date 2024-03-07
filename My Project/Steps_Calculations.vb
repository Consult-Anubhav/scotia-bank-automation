Imports System.IO
Imports System.IO.Compression
Imports Microsoft.Office.Interop.Word

Module Steps_Calculations
    ' Convert Email to PDF
    Public Sub EmailToPDF(Item As Object, outputPath As String)
        Try
            If String.IsNullOrEmpty(Item.Body) Then
                DisplayWindowsNotification("Error - Email", "Body is empty")
            Else
                Dim tempFilePath As String = Path.Combine(Environment.GetEnvironmentVariable("temp"), "Scotia\Calculations\")
                Dim tempFileName As String = Path.Combine(tempFilePath, $"{DateTime.Now:yyyy-MM-dd_hh.mm.ss}.mht")

                ' Create directory if not exists
                If Not Directory.Exists(tempFilePath) Then
                    Directory.CreateDirectory(tempFilePath)
                End If

                DisplayWindowsNotification("Saving", "email")
                Item.SaveAs(tempFileName, Microsoft.Office.Interop.Outlook.OlSaveAsType.olMHTML)

                ' Convert MHT to PDF
                Dim wordApp As New Application()
                wordApp.Visible = False

                Dim doc As Document = wordApp.Documents.Open(tempFileName, ConfirmConversions:=False)

                Dim outputFilePath As String = Path.Combine(outputPath, "Calculations\")
                ' Create directory if not exists
                If Not Directory.Exists(outputFilePath) Then
                    Directory.CreateDirectory(outputFilePath)
                End If

                Dim outputFileName As String = Path.Combine(outputFilePath, $"ThisEmail_{DateTime.Now:yyyy-MM-dd_hh.mm.ss}.pdf")

                doc.ExportAsFixedFormat(outputFileName, WdExportFormat.wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:=WdExportOptimizeFor.wdExportOptimizeForPrint)

                doc.Close(False)
                wordApp.Quit(False)

                ' Delete MHT file
                ' File.Delete(tempFileName)

                ' ZipAllFilesInFolder(Path.Combine(outputPath, $"Calculations {DateTime.Now:yyyy-MM-dd_hh.mm.ss}.zip"), outputPath)

            End If
        Catch ex As Exception
            DisplayWindowsNotification($"Error - {ex.HResult}", $"Email to PDF failed - {ex.Message}")
        End Try
    End Sub

    ' Zip Calculations Folder
    Public Sub ZipAllFilesInFolder(zippedFileFullName As String, folderToZipPath As String)
        Try
            Dim calculationsFolderPath As String = Path.Combine(folderToZipPath, "Calculations")

            If Directory.Exists(calculationsFolderPath) Then
                Dim tempZipPath As String = Path.Combine(folderToZipPath, $"Calculations_{DateTime.Now:yyyy-MM-dd_hh.mm.ss}.zip")

                DisplayWindowsNotification("Zip", "Creating")

                ZipFile.CreateFromDirectory(calculationsFolderPath, tempZipPath)

                File.Move(tempZipPath, zippedFileFullName)

                DisplayWindowsNotification("Zip", "Completed")
            Else
                DisplayWindowsNotification("Error - Zip", "Calculations folder is empty")
            End If
        Catch ex As Exception
            DisplayWindowsNotification($"Error - {ex.HResult}", $"Error zipping files - {ex.Message}")
        End Try
    End Sub


End Module
