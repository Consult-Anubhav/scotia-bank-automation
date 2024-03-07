Imports System.IO
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Module Steps_Murex

    Public Sub TestMurex()
        Dim inputDir, outputDir, emailMonthYear, previousMonthYear, emailYear, previousYear, emailMonth, previousMonth As String

        'Assign Variables
        emailMonthYear = GetEmailMonthYear("Re: Scotia Report - Dec 2023")
        previousMonthYear = GetPreviousMonthYear(emailMonthYear & "")
        emailYear = GetEmailYear(emailMonthYear & "")
        previousYear = GetPreviousYear(emailMonthYear & "")
        emailMonth = GetEmailMonth(emailMonthYear & "")
        previousMonth = GetPreviousMonth(emailMonthYear & "")
        outputDir = GetFakeRootPath() & "\" & emailYear & "\" & emailMonth
        inputDir = GetFakeRootPath() & "\" & previousYear & "\" & previousMonth
        'Test K2
        'GenerateK2Extract outputDir & ""
        'Test Murex
        GenerateMutexExtract(outputDir & "")
    End Sub

    '--- Murex ---

    Public Sub GenerateMutexExtract(dirPath As String)
        Dim ExApp As New Excel.Application
        Dim ExWbkReport As Excel.Workbook
        Dim ExWbkCSV As Excel.Workbook

        Dim ReportPath, CSVPath As String
        Dim FileName As String
        Dim FilePath As String

        ReportPath = Path.Combine(dirPath, "Supporting Files K2 and Murex\Murex\")
        CSVPath = Path.Combine(dirPath, "Supporting Files K2 and Murex\K2\")

        ExApp.AskToUpdateLinks = False
        ExApp.DisplayAlerts = False
        ExApp.Visible = True

        'DisplayWindowsNotification "CCD Extract", "Opening Report"
        ' Set ExWbkReport = ExApp.Workbooks.Open("C:\wamp64\www\~Consult Anubhav Projects\scotia-bank-automation\DF_DeMinimis_Extract (01012023-12312023).xlsm")

        'DisplayWindowsNotification "CCD Extract", "CopyAndTrimSpecialEntity"
        'ExWbk.Application.Run "CopyAndTrimSpecialEntity.CopyAndTrimSpecialEntity"

        '--- --- CCD Extract ---

        Dim ccdWs As Excel.Worksheet
        Dim dfWs As Excel.Worksheet
        Dim ccdLastRow As Long
        Dim dfLastRow As Long
        Dim i As Long

        DisplayWindowsNotification("Murex", "Opening Report")
        FileName = "DF_DeMinimis_Extract (01012023-12312023).xlsx"
        FilePath = Path.Combine(ReportPath, FileName)
        ExWbkReport = ExApp.Workbooks.Open(FilePath)

        ' Disable screen updating
        ExApp.Application.ScreenUpdating = False

        ' Set worksheets
        DisplayWindowsNotification("Murex CCD Extract", "Opening CSV")

        FileName = "CCD Extract.csv"
        FilePath = Path.Combine(CSVPath, FileName)
        ExWbkCSV = ExApp.Workbooks.Open(FilePath)
        ccdWs = ExWbkCSV.Sheets("CCD Extract")
        dfWs = ExWbkReport.Sheets("Murex_EM_DF_attributes")

        ' Find the last row in CCD Extract.csv
        ccdLastRow = ccdWs.Cells(ccdWs.Rows.Count, "Y").End(XlDirection.xlUp).Row

        ' Find the last row in DF_DeMinimis_Extract
        dfLastRow = dfWs.Cells(dfWs.Rows.Count, "Q").End(XlDirection.xlUp).Row

        ' Force the entire column to be in the desired format
        dfWs.Columns("Q:Q").NumberFormat = "@"

        DisplayWindowsNotification("Murex CCD Extract", "Copying data")

        ' Copy "Special Entity" from CCD Extract.csv to DF_DeMinimis_Extract
        For i = 2 To ccdLastRow ' Assuming the headers are in the first row
            ' Copy the value
            dfWs.Cells(i, "Q").Value = ccdWs.Cells(i, "Y").Value

            ' Trim the column Q after 15 spaces and store in R, T, and U columns
            Dim specialEntity As String = dfWs.Cells(i, "Q").Value

            ' Trim after 15 spaces
            Dim trimmedValue As String = specialEntity.Substring(0, Math.Min(15, specialEntity.Length))

            ' Store in R, T, and U columns
            dfWs.Cells(i, "R").Value = trimmedValue
            dfWs.Cells(i, "T").Value = trimmedValue
            dfWs.Cells(i, "U").Value = trimmedValue
        Next i

        DisplayWindowsNotification("Murex CCD Extract", "Special Entity copied")

        ' Enable screen updating
        ExApp.Application.ScreenUpdating = True
        'Application.Calculation = xlCalculationAutomatic
        DisplayWindowsNotification("Murex CCD Extract", "Enable screen updating")
        ExWbkCSV.Close(SaveChanges:=False)
        ExWbkReport.Close(SaveChanges:=True)
        ExApp.Quit()

        '--- --- CCD Extract ---
    End Sub

End Module
