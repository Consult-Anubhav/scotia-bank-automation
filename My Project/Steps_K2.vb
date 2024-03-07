Imports Microsoft.Office.Interop.Excel
Imports Range = Microsoft.Office.Interop.Excel.Range

Module Steps_K2

    Public Sub TestK2()
        Dim inputDir, outputDir, emailMonthYear, previousMonthYear, emailYear, previousYear, emailMonth, previousMonth As String

        'Assign Variables
        emailMonthYear = GetEmailMonthYear("Re: Scotia Report - Dec 2023")
        previousMonthYear = GetPreviousMonthYear(emailMonthYear & "")
        emailYear = GetEmailYear(emailMonthYear & "")
        previousYear = GetPreviousYear(emailMonthYear & "")
        emailMonth = GetEmailMonth(emailMonthYear & "")
        previousMonth = GetPreviousMonth(emailMonthYear & "")
        outputDir = GetFakeRootPath() '& "\" & emailYear & "\" & emailMonth
        inputDir = GetFakeRootPath() '& "\" & previousYear & "\" & previousMonth
        'Test K2
        GenerateK2Extract(outputDir & "")
        'Test Murex
        'GenerateMutexExtract(outputDir & "")
    End Sub

    '--- K2 ---

    Public Sub GenerateK2Extract(dirPath As String)
        Try
            Dim RootPath As String
            Dim ExApp As New Microsoft.Office.Interop.Excel.Application
            Dim ExWbkReport, ExWbkCSV As Workbook
            Dim FileName As String
            Dim FilePath As String
            Dim wsCCD As Worksheet
            Dim csvDataRange As Range

            RootPath = dirPath & "\Supporting Files K2 and Murex\K2\"

            ExApp.AskToUpdateLinks = False
            ExApp.DisplayAlerts = False
            ExApp.Visible = True

            DisplayWindowsNotification("K2 Extract", "Opening Report")

            FileName = "K2 and Portal Data Summary_Jan 1 2022 - Dec 31 2023.xlsx"
            FilePath = RootPath & FileName
            ExWbkReport = ExApp.Workbooks.Open(FilePath)

            '--- CCDExtractCSV ---

            ' Change the file name and path accordingly
            FileName = "CCD Extract.csv"
            FilePath = RootPath & FileName

            DisplayWindowsNotification("K2 CCD Extract", "Opening CSV")
            ExWbkCSV = ExApp.Workbooks.Open(FilePath)

            ' Reference to CCD Extract sheet
            wsCCD = ExWbkCSV.Sheets("CCD Extract")

            ' Set the data range in the CSV file
            csvDataRange = ExWbkReport.Sheets("CCD Extract").UsedRange

            DisplayWindowsNotification("K2 CCD Extract", "Copying data")
            ' Copy data from CSV to CCD Extract sheet
            csvDataRange.Copy(wsCCD.Range("A1"))

            ' Close the CSV file without saving changes
            DisplayWindowsNotification("K2 CCDExtract", "Closing CSV")
            ExWbkCSV.Close()

            '--- --- CFCTExtractCSV ---

            '--- CFCTExtractCSV ---
            Dim wsK2 As Worksheet
            Dim lastRow As Long

            ' Change the file name and path accordingly
            FileName = "CFTCExtract_2023_12_28.csv"
            FilePath = RootPath & FileName

            ' Open the CSV file
            ExWbkCSV = ExApp.Workbooks.Open(FilePath)

            ' Reference to K2 Extract sheet
            wsK2 = ExWbkCSV.Sheets("CFTCExtract_2023_12_28")

            ' Copy data from CSV to K2 Extract sheet
            DisplayWindowsNotification("K2 CFCT Extract", "Copying data")
            lastRow = wsK2.Cells(wsK2.Rows.Count, "A").End(XlDirection.xlUp).Row

            ' Find the last row in column A of CSV file
            With ExWbkReport.Sheets("K2 Extract")
                .Range("A1:AN" & lastRow).Value = wsK2.Range("A1:AN" & lastRow).Value
            End With

            ' Close the CSV file without saving changes
            DisplayWindowsNotification("K2 CFCTE Extract", "Closing CSV")
            ExWbkCSV.Close(False)

            ' Close the Excel application
            DisplayWindowsNotification("K2 Extract", "Saving Report")
            ExWbkReport.Close(True)
            ExApp.Quit()


            ' Close the Excel application
            DisplayWindowsNotification("K2 Extract", "Saving Report")
            ExWbkReport.Close(True)
            ExApp.Quit()

        Catch ex As Exception
            DisplayWindowsNotification("Error", "GenerateK2Extract failed")
            DisplayWindowsNotification(ex.HResult.ToString(), ex.Message)
        End Try
    End Sub

End Module
