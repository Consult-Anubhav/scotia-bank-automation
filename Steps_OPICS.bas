Attribute VB_Name = "Steps_OPICS"


        
'--- OPICS ---


Public Sub FXCalc()
    On Error GoTo ErrorHandler
    
    Dim localFilePath As String
    Dim fileData As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim grandSum As Double
    
    ' Define the path to the local CSV file
    localFilePath = "C:\Users\yash\Documents\___WORK\scotia-bank-automation\fx_rates_local.csv"
    
    ' Check if the local file exists
    If Dir(localFilePath) <> "" Then
        ' Read data from the local file
        Dim fileNum As Integer
        fileNum = FreeFile()
        Open localFilePath For Input As #fileNum
        
        ' Read the entire file data
        fileData = Input$(LOF(fileNum), fileNum)
        
        ' Close the file
        Close #fileNum
        
        ' Show message indicating the use of the local CSV file
        MsgBox "Using the local CSV file."
        
        ' Open the Excel file to which data will be appended
        Set wb = Workbooks.Open("C:\Users\yash\Downloads\Scotia\Was resume\OPICS Scotia Investments Jamaica Limited\FX (FORWARDS).prn.xlsx")
        Set ws = wb.Sheets(1) ' Assuming data will be appended to the first sheet
        
        ' Find the last row in column A of the Excel sheet
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
        
        ' Parse the CSV data and append it to the Excel sheet
        Dim csvRows() As String
        csvRows = Split(fileData, vbCrLf) ' Split CSV data into rows
        
        ' Skip the first row (headers) of the CSV file
        For i = LBound(csvRows) + 1 To UBound(csvRows)
            Dim csvColumns() As String
            csvColumns = Split(csvRows(i), ",") ' Split CSV row into columns
            
            ' Assuming columns A to S contain the required data
            ws.Cells(lastRow + i - LBound(csvRows), 1).Resize(1, UBound(csvColumns) + 1).Value = csvColumns
            
            ' Convert columns R and T to numeric values
            ws.Cells(lastRow + i - LBound(csvRows), 18).Value = CDbl(ws.Cells(lastRow + i - LBound(csvRows), 18).Value) ' Column R
            ws.Cells(lastRow + i - LBound(csvRows), 20).Value = Abs(CDbl(ws.Cells(lastRow + i - LBound(csvRows), 18).Value)) ' Column T
        Next i
        
        ' Calculate grand sum in cell V2
        ' grandSum = Application.WorksheetFunction.Sum(ws.Range("T" & lastRow & ":T" & lastRow + UBound(csvRows) - LBound(csvRows)))
        ' ws.Range("V2").Value = grandSum
        
        ' Save and close the workbook
        wb.Save
        wb.Close
        
        ' Display a message indicating successful processing
        MsgBox "CSV file processed successfully!"
    Else
        MsgBox "Local CSV file not found. Process cannot continue."
    End If
    
    ' Clean up
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
    ' Close the file if it was opened
    If fileNum > 0 Then Close #fileNum
End Sub

