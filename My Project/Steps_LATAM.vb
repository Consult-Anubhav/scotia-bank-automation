Imports System.IO
Imports System.Net
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Module Steps_LATAM

    Public Sub DownloadAndProcessForexData()
        Dim url As String = "https://www.facebook.com/LICENCE" '"Gcm.navigator.bns/doddfrank/index.asp?page=statusreport"
        Dim httpRequest As HttpWebRequest = WebRequest.Create(url)
        Dim httpResponse As HttpWebResponse
        Dim fileStream As FileStream
        Dim fileData As String
        Dim csvFileName As String
        Dim excelApp As New Excel.Application
        Dim wb As Excel.Workbook
        Dim ws As Excel.Worksheet
        Dim lastRow As Long
        Dim i As Long
        Dim totalAbsolute As Double

        ' Send request to download CSV file
        Try
            httpResponse = httpRequest.GetResponse()

            ' Get CSV data
            Using reader As New StreamReader(httpResponse.GetResponseStream())
                fileData = reader.ReadToEnd()
            End Using

            ' Save CSV data to a file
            csvFileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "forex_data.csv")
            fileStream = New FileStream(csvFileName, FileMode.Create)
            Using writer As New StreamWriter(fileStream)
                writer.Write(fileData)
            End Using

            ' Open Excel application and the target workbook
            excelApp.Visible = False ' Make Excel application invisible
            wb = excelApp.Workbooks.Open(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "FX (Forwards).prn.xlsx"))
            ws = wb.Sheets(1) ' Assuming data will be appended to the first sheet

            ' Find the last row in the target worksheet
            lastRow = ws.Cells(ws.Rows.Count, "A").End(XlDirection.xlUp).Row ' -4162 represents xlUp

            ' Append CSV data to the target worksheet
            ws.QueryTables.Add(Connection:="TEXT;" & csvFileName, Destination:=ws.Cells(lastRow + 1, 1)).TextFileParseType = XlTextParsingType.xlDelimited
            ws.QueryTables(1).TextFileCommaDelimiter = True ' Use comma delimiter
            ws.QueryTables(1).Refresh() ' Refresh query table to load CSV data

            ' Calculate absolute values in column T
            For i = lastRow + 1 To ws.Cells(ws.Rows.Count, "A").End(XlDirection.xlUp).Row
                ws.Cells(i, "T").Formula = "=ABS(R" & i & ")"
                totalAbsolute += Math.Abs(ws.Cells(i, "R").Value)
            Next i

            ' Show the grand sum of absolute values in cell V2
            ws.Cells(2, "V").Value = totalAbsolute

            ' Save and close the workbook
            wb.Save()
            wb.Close()

            ' Clean up
            excelApp.Quit()
            File.Delete(csvFileName)

            MsgBox("Forex data processed successfully.")
        Catch ex As Exception
            MsgBox("Error downloading file: " & ex.Message)
        End Try
    End Sub

End Module
