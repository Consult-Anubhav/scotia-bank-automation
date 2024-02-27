Attribute VB_Name = "Steps_K2"

Public Sub TestK2()
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
    GenerateK2Extract outputDir & ""
    'Test Murex
    'GenerateMutexExtract outputDir & ""
End Sub

'--- K2 ---

Public Sub GenerateK2Extract(dirPath As String)
    'On Error GoTo ErrorHandler
    
    Dim RootPath As String
    Dim ExApp As Excel.Application
    Dim ExWbkReport, ExWbkCSV As Workbook
    
    Dim FileName As String
    Dim FilePath As String
    Dim wsCCD As Worksheet
    Dim csvDataRange As Range
    
    RootPath = dirPath & "\Supporting Files K2 and Murex\K2\"
    Set ExApp = New Excel.Application
    
    ExApp.AskToUpdateLinks = False
    ExApp.DisplayAlerts = False
    ExApp.Visible = True
    
    DisplayWindowsNotification "K2 Extract", "Opening Report"
    
    FileName = "K2 and Portal Data Summary_Jan 1 2022 - Dec 31 2023.xlsx"
    FilePath = RootPath & FileName
    Set ExWbkReport = ExApp.Workbooks.Open(FilePath)
    'ExWbk.Application.Run "Module1.CCDExtractCSV"
    
    
    '--- CCDExtractCSV ---
    
    
    ' Change the file name and path accordingly
    FileName = "CCD Extract.csv"
    FilePath = RootPath & FileName
    ' Open the CSV file
    'Workbook.OpenText FileName:=csvFilePath, DataType:=xlDelimited, comma:=True
    
    DisplayWindowsNotification "K2 CCD Extract", "Opening CSV"
    Set ExWbkCSV = ExApp.Workbooks.Open(FilePath)
    
    ' Reference to CCD Extract sheet
    Set wsCCD = ExWbkCSV.Sheets("CCD Extract")
    
    ' Set the data range in the CSV file
    With ExWbkReport.Sheets("CCD Extract")
        Set csvDataRange = .UsedRange
    End With
    
    DisplayWindowsNotification "K2 CCD Extract", "copying data"
    ' Copy data from CSV to CCD Extract sheet
    csvDataRange.Copy wsCCD.Range("A1")
    
    ' Close the CSV file without saving changes
    'Workbooks(csvFileName).Close SaveChanges:=False
    DisplayWindowsNotification "K2 CCDExtract", "Closing CSV"
    ExWbkCSV.Close 'SaveChanges:=True
    
    
    '--- CCDExtractCSV ---
    
    'ExWbk.Application.Run "Module2.CFCTE"
    
    
    '--- CFCTExtractCSV ---
    
    
    'Dim csvFileName As String
    'Dim csvFilePath As String
    Dim wsK2 As Worksheet
    Dim lastRow As Long
    
    ' Change the file name and path accordingly
    FileName = "CFTCExtract_2023_12_28.csv"
    FilePath = RootPath & FileName
    
    ' Open the CSV file
    'Workbooks.OpenText FileName:=csvFilePath, DataType:=xlDelimited, comma:=True
    
    DisplayWindowsNotification "K2 CFCTE Extract", "Opening CSV"
    Set ExWbkCSV = ExApp.Workbooks.Open(FilePath)
    
    ' Reference to K2 Extract sheet
    Set wsK2 = ExWbkCSV.Sheets("CFTCExtract_2023_12_28")
    
    ' Copy data from CSV to K2 Extract sheet
    DisplayWindowsNotification "K2 CFCT Extract", "Copying data"
    With ExWbkReport.Sheets("K2 Extract")
        ' Find the last row in column A of CSV file
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        
        ' Copy data from CSV to K2 Extract sheet based on the mapping
        .Range("A1:A" & lastRow).Copy wsK2.Range("A1")
        .Range("B1:B" & lastRow).Copy wsK2.Range("B1")
        .Range("C1:C" & lastRow).Copy wsK2.Range("C1")
        .Range("D1:D" & lastRow).Copy wsK2.Range("D1")
        .Range("E1:E" & lastRow).Copy wsK2.Range("E1")
        .Range("F1:F" & lastRow).Copy wsK2.Range("F1")
        .Range("G1:G" & lastRow).Copy wsK2.Range("G1")
        .Range("H1:H" & lastRow).Copy wsK2.Range("H1")
        .Range("I1:I" & lastRow).Copy wsK2.Range("I1")
        .Range("J1:J" & lastRow).Copy wsK2.Range("K1")
        .Range("K1:K" & lastRow).Copy wsK2.Range("L1")
        .Range("L1:L" & lastRow).Copy wsK2.Range("M1")
        .Range("M1:M" & lastRow).Copy wsK2.Range("N1")
        .Range("N1:N" & lastRow).Copy wsK2.Range("O1")
        .Range("O1:O" & lastRow).Copy wsK2.Range("P1")
        .Range("P1:P" & lastRow).Copy wsK2.Range("Q1")
        .Range("Q1:Q" & lastRow).Copy wsK2.Range("S1")
        .Range("R1:R" & lastRow).Copy wsK2.Range("V1")
        .Range("S1:S" & lastRow).Copy wsK2.Range("W1")
        .Range("T1:T" & lastRow).Copy wsK2.Range("X1")
        .Range("U1:U" & lastRow).Copy wsK2.Range("Y1")
        .Range("V1:V" & lastRow).Copy wsK2.Range("Z1")
        .Range("W1:W" & lastRow).Copy wsK2.Range("AA1")
        .Range("X1:X" & lastRow).Copy wsK2.Range("AB1")
        .Range("Y1:Y" & lastRow).Copy wsK2.Range("AC1")
        .Range("Z1:Z" & lastRow).Copy wsK2.Range("AD1")
        .Range("AA1:AA" & lastRow).Copy wsK2.Range("AE1")
        .Range("AB1:AB" & lastRow).Copy wsK2.Range("AF1")
        .Range("AC1:AC" & lastRow).Copy wsK2.Range("AG1")
        .Range("AD1:AD" & lastRow).Copy wsK2.Range("AH1")
        .Range("AE1:AE" & lastRow).Copy wsK2.Range("AI1")
        .Range("AF1:AF" & lastRow).Copy wsK2.Range("AJ1")
        .Range("AG1:AG" & lastRow).Copy wsK2.Range("AK1")
        .Range("AH1:AH" & lastRow).Copy wsK2.Range("AL1")
        .Range("AI1:AI" & lastRow).Copy wsK2.Range("AM1")
        .Range("AJ1:AJ" & lastRow).Copy wsK2.Range("AN1")
    End With
    
    ' Close the CSV file without saving changes
    'Workbooks(csvFileName).Close SaveChanges:=False
    DisplayWindowsNotification "K2 CFCTE Extract", "Closing CSV"
    ExWbkCSV.Close 'SaveChanges:=True
    
    DisplayWindowsNotification "K2 Extract", "Saving Report"
    ExWbkReport.Close SaveChanges:=True
    ExApp.Quit
    
    
    '--- --- CFCTExtractCSV ---
    
    
ExitSub:
    Exit Sub
    
ErrorHandler:
    DisplayWindowsNotification "Error", "GenerateK2Extract failed"
    DisplayWindowsNotification Err.Number, Err.Description
    Resume ExitSub
End Sub
