Attribute VB_Name = "Steps_Murex"

Public Sub TestMurex()
    Dim inputDir, outputDir, emailMonthYear, previousMonthYear, emailYear, previousYear, emailMonth, previousMonth As String
        
    'Assign Variables
    emailMonthYear = GetEmailMonthYear("Re: Scotia Report - Dec 2023")
    previousMonthYear = GetPreviousMonthYear(emailMonthYear & "")
    emailYear = GetEmailYear(emailMonthYear & "")
    previousYear = GetPreviousYear(emailMonthYear & "")
    emailMonth = GetEmailMonth(emailMonthYear & "")
    previousMonth = GetPreviousMonth(emailMonthYear & "")
    outputDir = fakeRootPath & "\" & emailYear & "\" & emailMonth
    inputDir = fakeRootPath & "\" & previousYear & "\" & previousMonth
    'Test K2
    'GenerateK2Extract outputDir & ""
    'Test Murex
    GenerateMutexExtract outputDir & ""
End Sub

'--- Murex ---

Public Sub GenerateMutexExtract(dirPath As String)
    Dim ExApp As Excel.Application
    'Dim ExWbk As Workbook
    Dim ReportPath, CSVPath As String
    
    Set ExApp = New Excel.Application
    Dim ExWbkReport, ExWbkCSV As Workbook
    
    Dim FileName As String
    Dim FilePath As String
    
    ReportPath = dirPath & "\Supporting Files K2 and Murex\Murex\"
    CSVPath = dirPath & "\Supporting Files K2 and Murex\K2\"
    
    ExApp.AskToUpdateLinks = False
    ExApp.DisplayAlerts = False
    ExApp.Visible = True
    
    'DisplayWindowsNotification "CCD Extract", "Opening excel"
    'Set ExWbk = ExApp.Workbooks.Open("C:\wamp64\www\~Consult Anubhav Projects\scotia-bank-automation\DF_DeMinimis_Extract (01012023-12312023).xlsm")
    
    'DisplayWindowsNotification "CCD Extract", "CopyAndTrimSpecialEntity"
    'ExWbk.Application.Run "CopyAndTrimSpecialEntity.CopyAndTrimSpecialEntity"
    
    
    '--- --- CCD Extract ---
    
    
    Dim ccdWs As Worksheet
    Dim dfWs As Worksheet
    Dim ccdLastRow As Long
    Dim dfLastRow As Long
    Dim i As Long
    
    DisplayWindowsNotification "Murex", "Opening Report"
    FileName = "DF_DeMinimis_Extract (01012023-12312023).xlsx"
    FilePath = ReportPath & FileName
    Set ExWbkReport = ExApp.Workbooks.Open(FilePath)
    
    ' Disable screen updating and automatic calculations
    'DisplayWindowsNotification "DeMinimis", "Disable screen updating"
    ExApp.Application.ScreenUpdating = False
    'ExApp.Application.Calculation = xlCalculationManual
    
    ' Set worksheets
    DisplayWindowsNotification "Murex CCD Extract", "Opening CSV"
    
    FileName = "CCD Extract.csv"
    FilePath = CSVPath & FileName
    Set ExWbkCSV = Workbooks.Open(FilePath)
    Set ccdWs = ExWbkCSV.Sheets("CCD Extract")
    Set dfWs = ExWbkReport.Sheets("Murex_EM_DF_attributes")
    
    ' Find the last row in CCD Extract.csv
    ccdLastRow = ccdWs.Cells(ccdWs.Rows.Count, "Y").End(xlUp).Row
    
    ' Find the last row in DF_DeMinimis_Extract
    dfLastRow = dfWs.Cells(dfWs.Rows.Count, "Q").End(xlUp).Row
    
    ' Force the entire column to be in the desired format
    dfWs.Columns("Q:Q").NumberFormat = "@"
    
    DisplayWindowsNotification "Murex CCD Extract", "Copying data"
    ' Copy "Special Entity" from CCD Extract.csv to DF_DeMinimis_Extract
    For i = 2 To ccdLastRow ' Assuming the headers are in the first row
        ' Copy the value
        dfWs.Cells(i, "Q").Value = ccdWs.Cells(i, "Y").Value
        
        ' Trim the column Q after 15 spaces and store in R, T, and U columns
        Dim specialEntity As String
        specialEntity = dfWs.Cells(i, "Q").Value
        
        ' Trim after 15 spaces
        Dim trimmedValue As String
        trimmedValue = Trim(Mid(specialEntity, 1, 15))
        
        ' Store in R, T, and U columns
        dfWs.Cells(i, "R").Value = trimmedValue
        dfWs.Cells(i, "T").Value = trimmedValue
        dfWs.Cells(i, "U").Value = trimmedValue
    Next i
    
    DisplayWindowsNotification "Murex CCD Extract", "Special Entity copied"
    
    ' Enable screen updating and automatic calculations
    ExApp.Application.ScreenUpdating = True
    'Application.Calculation = xlCalculationAutomatic
    DisplayWindowsNotification "Murex CCD Extract", "Enable screen updating"
    ExWbkCSV.Close 'Workbooks.Close
    'MsgBox "Special Entity copied and trimmed successfully!", vbInformation
    ExWbkReport.Close SaveChanges:=True
    ExApp.Quit
    
    '--- --- CCD Extract ---
    
    
    'DisplayWindowsNotification "CCD Extract", "Saving File"
    'ExWbk.Close SaveChanges:=True
End Sub

