Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel
Imports Range = Microsoft.Office.Interop.Excel.Range

Module Steps_K2

    Public Sub TestK2()
        Dim inputDir, outputDir, emailMonthYear, previousMonthYear, emailYear, previousYear, emailMonth, previousMonth As String

        'Assign Variables
        'emailMonthYear = GetEmailMonthYear("Re: Scotia Report - Dec 2023")
        'previousMonthYear = GetPreviousMonthYear(emailMonthYear & "")
        'emailYear = GetEmailYear(emailMonthYear & "")
        'previousYear = GetPreviousYear(emailMonthYear & "")
        'emailMonth = GetEmailMonth(emailMonthYear & "")
        'previousMonth = GetPreviousMonth(emailMonthYear & "")
        'outputDir = GetFakeRootPath() & "\" & emailYear & "\" & emailMonth
        'outputDir = SelectFolder()
        'inputDir = GetFakeRootPath() '& "\" & previousYear & "\" & previousMonth
        'Test K2
        GenerateK2Extract()
        'Test Murex
        'GenerateMutexExtract(outputDir & "")
    End Sub

    '--- K2 ---
    Public Sub GenerateK2Extract()
        Try
            Dim ExApp As New Microsoft.Office.Interop.Excel.Application
            Dim ExWbkReport, ExWbkCSV As Workbook
            Dim FileName As String
            Dim FilePath As String
            Dim wsCCD, wsK2 As Worksheet
            Dim csvDataRange As Range

            ' Suppress alerts and make Excel visible
            ExApp.DisplayAlerts = False
            ExApp.Visible = True

            ' Open the report workbook
            FileName = SelectFileWithMessage("Open a Report File for K2", "Excel Files (*.xlsx)|*.xlsx")
            FilePath = FileName
            ExWbkReport = ExApp.Workbooks.Open(FilePath)

            ' Open the CCDExtract CSV file
            FileName = SelectFileWithMessage("Open a CCDExtract file", "CSV Files (*.csv)|*.csv")
            FilePath = FileName
            ExWbkCSV = ExApp.Workbooks.Open(FilePath)
            wsCCD = ExWbkCSV.Sheets("CCD Extract")
            csvDataRange = ExWbkReport.Sheets("CCD Extract").UsedRange
            csvDataRange.Copy(wsCCD.Range("A1"))
            ExWbkCSV.Close()

            ' Open the CFCTExtract CSV file
            FileName = SelectFileWithMessage("Open a CFCTExtract file", "CSV Files (*.csv)|*.csv")
            FilePath = FileName
            ExWbkCSV = ExApp.Workbooks.Open(FilePath)
            wsK2 = ExWbkCSV.Sheets("CFTCExtract_2023_12_28")
            wsK2.UsedRange.Copy(ExWbkReport.Sheets("K2 Extract").Range("A1"))
            ExWbkCSV.Close(False)

            ' Display message before saving the report
            UpdateLabel("K2 Extract", "Saving Report")

            ' Save and close the report workbook
            ExWbkReport.Close(True)
            ExApp.Quit()

            ' Release Excel objects
            Marshal.ReleaseComObject(wsCCD)
            Marshal.ReleaseComObject(wsK2)
            Marshal.ReleaseComObject(ExWbkReport)
            Marshal.ReleaseComObject(ExWbkCSV)
            Marshal.ReleaseComObject(ExApp)
            UpdateLabel("K2 Extract", "Complete")
        Catch ex As Exception
            ' Handle errors
            UpdateLabel("Error", "GenerateK2Extract failed")
            UpdateLabel(ex.HResult.ToString(), ex.Message)
        End Try
    End Sub

End Module
