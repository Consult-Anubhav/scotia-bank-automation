Attribute VB_Name = "Steps_LATAM"

'--- Latam ---



Public Sub DownloadAndProcessForexData()
    Dim url As String
    Dim httpRequest As Object
    Dim fileStream As Object
    Dim fileData As String
    Dim csvFileName As String
    Dim excelApp As Object
    Dim wb As Object
    Dim ws As Object
    Dim lastRow As Long
    Dim i As Long
    Dim totalAbsolute As Double
    
    ' Intranet URL to download CSV file
    url = "https://www.facebook.com/LICENCE" '"Gcm.navigator.bns/doddfrank/index.asp?page=statusreport"
    
    ' Create HTTP request object
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    
    ' Send request to download CSV file
    httpRequest.Open "GET", url, False
    httpRequest.Send
    
    ' Check if request was successful
    If httpRequest.Status <> 200 Then
        MsgBox "Error downloading file: " & httpRequest.StatusText
        Exit Sub
    End If
    
    ' Get CSV data
    fileData = httpRequest.responseText
    
    ' Check if CSV data is empty
    If Len(fileData) = 0 Then
        MsgBox "Error: Empty file data."
        Exit Sub
    End If
    
    ' Save CSV data to a file
    csvFileName = ThisWorkbook.Path & "\forex_data.csv"
    Set fileStream = CreateObject("ADODB.Stream")
    fileStream.Open
    fileStream.Type = 1 ' Binary
    fileStream.Write httpRequest.responseBody
    fileStream.SaveToFile csvFileName, 2 ' Overwrite existing file
    fileStream.Close
    
    ' Open Excel application and the target workbook
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False ' Make Excel application invisible
    Set wb = excelApp.Workbooks.Open(ThisWorkbook.Path & "\FX (Forwards).prn.xlsx")
    Set ws = wb.Sheets(1) ' Assuming data will be appended to the first sheet
    
    ' Find the last row in the target worksheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(-4162).Row ' -4162 represents xlUp
    
    ' Append CSV data to the target worksheet
    With ws.QueryTables.Add(Connection:="TEXT;" & csvFileName, Destination:=ws.Cells(lastRow + 1, 1))
        .TextFileParseType = 1 ' Delimited
        .TextFileCommaDelimiter = True ' Use comma delimiter
        .Refresh ' Refresh query table to load CSV data
    End With
    
    ' Calculate absolute values in column T
    For i = lastRow + 1 To ws.Cells(ws.Rows.Count, "A").End(-4162).Row
        ws.Cells(i, "T").Formula = "=ABS(R" & i & ")"
        totalAbsolute = totalAbsolute + Abs(ws.Cells(i, "R").Value)
    Next i
    
    ' Show the grand sum of absolute values in cell V2
    ws.Cells(2, "V").Value = totalAbsolute
    
    ' Save and close the workbook
    wb.Save
    wb.Close
    
    ' Clean up
    excelApp.Quit
    Set excelApp = Nothing
    Set httpRequest = Nothing
    Set fileStream = Nothing
    Set wb = Nothing
    Set ws = Nothing
    
    ' Delete the temporary CSV file
    Kill csvFileName
    
    MsgBox "Forex data processed successfully."
    
End Sub


