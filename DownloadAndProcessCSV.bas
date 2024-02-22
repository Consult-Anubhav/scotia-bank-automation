Attribute VB_Name = "DownloadAndProcessCSV"
Sub DownloadAndProcessCSV()
    On Error GoTo ErrorHandler
    
    Dim url As String
    Dim httpRequest As Object
    Dim fileData As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim grandSum As Double
    
    ' Define the URL of the CSV file
    url = "http://Gcm.navigator.bns/doddfrank/index.asp?page=statusreport"
    
    ' Create a new instance of the XML HTTP Request object
    Set httpRequest = CreateObject("MSXML2.ServerXMLHTTP")
    
    ' Send a request to the URL
    httpRequest.Open "GET", url, False
    httpRequest.setRequestHeader "Content-Type", "text/csv"
    httpRequest.send
    
    ' Check if an error occurred during the request or if the host cannot be resolved
    If Err.Number <> 0 Or httpRequest.Status <> 200 Or httpRequest.Status = 12002 Then
        ' Show message indicating the use of the local CSV file
        MsgBox "Using the local CSV file."
        
        ' Use local file if an error occurred or if the file does not exist on the remote location
        Dim localFilePath As String
        localFilePath = "C:\Users\yash\Documents\___WORK\scotia-bank-automation\fx_rates_local.csv"
        
        ' Check if the local file exists
        If Dir(localFilePath) <> "" Then
            ' Read data from the local file
            Dim fileNum As Integer
            fileNum = FreeFile()
            Open localFilePath For Input As #fileNum
            fileData = Input$(LOF(fileNum), fileNum)
            Close #fileNum
        Else
            MsgBox "Local file not found. Process cannot continue."
            Exit Sub
        End If
    Else
        ' File successfully downloaded
        fileData = httpRequest.responseText
    End If
    
    ' Display the data from the local CSV file for debugging
    MsgBox fileData
    
    ' Open the Excel file to which data will be appended
    Set wb = Workbooks.Open("C:\Users\yash\Downloads\Scotia\Was resume\OPICS Scotia Investments Jamaica Limited\FX (FORWARDS).prn.xlsx")
    Set ws = wb.Sheets(1) ' Assuming data will be appended to the first sheet
    
    ' Find the last row in column A of the Excel sheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Parse the CSV data and append it to the Excel sheet
    Dim csvRows() As String
    csvRows = Split(fileData, vbCrLf) ' Split CSV data into rows
    
    For i = LBound(csvRows) To UBound(csvRows)
        Dim csvColumns() As String
        csvColumns = Split(csvRows(i), ",") ' Split CSV row into columns
        
        ' Assuming columns A to S contain the required data
        ws.Cells(lastRow + i, 1).Resize(1, UBound(csvColumns) + 1).Value = csvColumns
    Next i
    
    ' Calculate absolute values in column T
    ws.Range("T2:T" & lastRow + UBound(csvRows)).Formula = "=ABS(R2)"
    
    ' Calculate grand sum in cell V2
    grandSum = Application.WorksheetFunction.Sum(ws.Range("T2:T" & lastRow + UBound(csvRows)))
    ws.Range("V2").Value = grandSum
    
    ' Save and close the workbook
    wb.Save
    wb.Close
    
    ' Display a message indicating successful download and processing
    MsgBox "CSV file downloaded and processed successfully!"
    
    ' Clean up
    Set httpRequest = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
    ' Close the file if it was opened
    If fileNum > 0 Then Close #fileNum
    Set httpRequest = Nothing
End Sub

