Attribute VB_Name = "Excels"

'--- Latam ---
        
'--- OPICS ---

'--- SCOTS ---

'--- K2 ---

Public Sub Test()
    Dim inputDir, outputDir, emailMonthYear, previousMonthYear, emailYear, previousYear, emailMonth, previousMonth As String
        
    'Assign Variables
    emailMonthYear = GetEmailMonthYear("Re: Scotia Report - Dec 2023")
    previousMonthYear = GetPreviousMonthYear(emailMonthYear & "")
    emailYear = GetEmailYear(emailMonthYear & "")
    previousYear = GetPreviousYear(emailMonthYear & "")
    emailMonth = GetEmailMonth(emailMonthYear & "")
    previousMonth = GetPreviousMonth(emailMonthYear & "")
    outputDir = getRootPath & "\" & emailYear & "\" & emailMonth
    inputDir = getRootPath & "\" & previousYear & "\" & previousMonth
    GenerateK2Extract outputDir & ""
End Sub

Public Sub GenerateK2Extract(dirPath As String)
    'On Error GoTo ErrorHandler
    
    'Dim ExApp As Excel.Application
    'Dim ExWbk As Workbook
    
    'Set ExApp = New Excel.Application
    
    'ExApp.AskToUpdateLinks = False
    'ExApp.DisplayAlerts = False
    'ExApp.Visible = False
    
    'DisplayWindowsNotification "K2 Extract", "Opening excel"
    'Set ExWbk = ExApp.Workbooks.Open("C:\wamp64\www\~Consult Anubhav Projects\scotia-bank-automation\K2 and Portal Data Summary_Jan 1 2022 - Dec 31 2023.xlsm")
    
    'DisplayWindowsNotification "K2 Extract", "CCDExtractCSV"
    'ExWbk.Application.Run "Module1.CCDExtractCSV"
    
    '--- CCDExtractCSV ---
    
    Dim csvFileName As String
    Dim csvFilePath As String
    Dim wsCCD As Worksheet
    Dim csvDataRange As Range
    
    ' Change the file name and path accordingly
    csvFileName = "CCD Extract.csv"
    MsgBox dirPath
    csvFilePath = dirPath & "\Supporting Files K2 and Murex\K2\" & csvFileName
    MsgBox csvFilePath
    ' Open the CSV file
    Workbook.OpenText FileName:=csvFilePath, DataType:=xlDelimited, comma:=True
    
    ' Reference to CCD Extract sheet
    Set wsCCD = ThisWorkbook.Sheets("CCD Extract")
    
    ' Set the data range in the CSV file
    With Workbooks(csvFileName).Sheets(1)
        Set csvDataRange = .UsedRange
    End With
    
    ' Copy data from CSV to CCD Extract sheet
    csvDataRange.Copy wsCCD.Range("A1")
    
    ' Close the CSV file without saving changes
    Workbooks(csvFileName).Close SaveChanges:=False
    
    '--- CCDExtractCSV ---
    
    'DisplayWindowsNotification "K2 Extract", "CFCTE"
    'ExWbk.Application.Run "Module2.CFCTE"
    
    '--- CFCTExtractCSV ---
    
    
    'Dim csvFileName As String
    'Dim csvFilePath As String
    Dim wsK2 As Worksheet
    Dim lastRow As Long
    
    ' Change the file name and path accordingly
    csvFileName = "CFTCExtract_2023_12_28.csv"
    csvFilePath = ThisWorkbook.Path & "\" & csvFileName
    
    ' Open the CSV file
    Workbooks.OpenText FileName:=csvFilePath, DataType:=xlDelimited, comma:=True
    
    ' Reference to K2 Extract sheet
    Set wsK2 = ThisWorkbook.Sheets("K2 Extract")
    
    ' Copy data from CSV to K2 Extract sheet
    With Workbooks(csvFileName).Sheets(1)
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
    Workbooks(csvFileName).Close SaveChanges:=False
    
    
    '--- --- CFCTExtractCSV ---
    
    'DisplayWindowsNotification "K2 Extract", "Saving File"
    'ExWbk.Close SaveChanges:=True
    
ExitSub:
    Exit Sub
    
ErrorHandler:
    DisplayWindowsNotification "Error", "GenerateK2Extract failed"
    DisplayWindowsNotification Err.Number, Err.Description
    Resume ExitSub
End Sub

'--- Mutex ---

Public Sub GenerateCCDExtract()
    'Dim ExApp As Excel.Application
    'Dim ExWbk As Workbook
    
    'Set ExApp = New Excel.Application
    
    'ExApp.AskToUpdateLinks = False
    'ExApp.DisplayAlerts = False
    'ExApp.Visible = False
    
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
    
    ' Disable screen updating and automatic calculations
    DisplayWindowsNotification "DeMinimis", "Disable screen updating"
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Set worksheets
    DisplayWindowsNotification "DeMinimis", "opening CCD Extract"
    Set ccdWs = Workbooks.Open(ActiveWorkbook.Path & "\Docs\Supporting Files K2 and Murex\K2\CCD Extract.csv").Sheets(1)
    Set dfWs = ThisWorkbook.Sheets("Murex_EM_DF_attributes")
    
    ' Find the last row in CCD Extract.csv
    ccdLastRow = ccdWs.Cells(ccdWs.Rows.Count, "Y").End(xlUp).Row
    
    ' Find the last row in DF_DeMinimis_Extract
    dfLastRow = dfWs.Cells(dfWs.Rows.Count, "Q").End(xlUp).Row
    
    ' Force the entire column to be in the desired format
    dfWs.Columns("Q:Q").NumberFormat = "@"
    
    DisplayWindowsNotification "DeMinimis", "copying CCD Extract to DF_DeMinimis_Extract"
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
    
    DisplayWindowsNotification "DeMinimis", "Special Entity copied"
    
    ' Enable screen updating and automatic calculations
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    DisplayWindowsNotification "DeMinimis", "Enable screen updating"
    Workbooks.Close
    'MsgBox "Special Entity copied and trimmed successfully!", vbInformation
    
    
    '--- --- CCD Extract ---
    
    
    'DisplayWindowsNotification "CCD Extract", "Saving File"
    'ExWbk.Close SaveChanges:=True
End Sub

'--- Calculations ---

    
