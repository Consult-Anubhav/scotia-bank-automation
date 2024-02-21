Attribute VB_Name = "Module2"
Sub CFCTE()

    Dim csvFileName As String
    Dim csvFilePath As String
    Dim wsK2 As Worksheet
    Dim lastRow As Long
    
    ' Change the file name and path accordingly
    csvFileName = "CFTCExtract_2023_12_28.csv"
    csvFilePath = ThisWorkbook.Path & "\" & csvFileName
    
    ' Open the CSV file
    Workbooks.OpenText Filename:=csvFilePath, DataType:=xlDelimited, comma:=True
    
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

End Sub

