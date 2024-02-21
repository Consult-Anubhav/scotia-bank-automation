Attribute VB_Name = "Module1"
Sub CCDExtractCSV()

    Dim csvFileName As String
    Dim csvFilePath As String
    Dim wsCCD As Worksheet
    Dim csvDataRange As Range
    
    ' Change the file name and path accordingly
    csvFileName = "CCD Extract.csv"
    csvFilePath = ThisWorkbook.Path & "\" & csvFileName
    
    ' Open the CSV file
    Workbooks.OpenText Filename:=csvFilePath, DataType:=xlDelimited, comma:=True
    
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

End Sub

