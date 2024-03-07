Imports System.IO
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Module Steps_OPICS

    '--- OPICS ---

    Public Sub FXCalc()
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
        If File.Exists(localFilePath) Then
            ' Read data from the local file
            fileData = File.ReadAllText(localFilePath)

            ' Show message indicating the use of the local CSV file
            UpdateLabel("Using the local CSV file.", "")

            ' Open the Excel file to which data will be appended
            Dim excelApp As New Excel.Application
            excelApp.Visible = False
            wb = excelApp.Workbooks.Open("C:\Users\yash\Downloads\Scotia\Was resume\OPICS Scotia Investments Jamaica Limited\FX (FORWARDS).prn.xlsx")
            ws = wb.Sheets(1) ' Assuming data will be appended to the first sheet

            ' Find the last row in column A of the Excel sheet
            lastRow = ws.Cells(ws.Rows.Count, "A").End(XlDirection.xlUp).Row + 1

            ' Parse the CSV data and append it to the Excel sheet
            Dim csvRows() As String = fileData.Split(vbCrLf)

            ' Skip the first row (headers) of the CSV file
            For i = 1 To csvRows.Length - 1
                Dim csvColumns() As String = csvRows(i).Split(",")

                ' Assuming columns A to S contain the required data
                ws.Cells(lastRow + i - 1, 1).Resize(1, csvColumns.Length).Value = csvColumns

                ' Convert columns R and T to numeric values
                ws.Cells(lastRow + i - 1, 18).Value = CDbl(ws.Cells(lastRow + i - 1, 18).Value) ' Column R
                ws.Cells(lastRow + i - 1, 20).Value = Math.Abs(CDbl(ws.Cells(lastRow + i - 1, 18).Value)) ' Column T
            Next i

            ' Calculate grand sum in cell V2
            ' grandSum = Application.WorksheetFunction.Sum(ws.Range("T" & lastRow & ":T" & lastRow + UBound(csvRows) - LBound(csvRows)))
            ' ws.Range("V2").Value = grandSum

            ' Save and close the workbook
            wb.Save()
            wb.Close()

            ' Display a message indicating successful processing
            UpdateLabel("CSV file processed successfully!", "")
        Else
            UpdateLabel("Local CSV file not found. Process cannot continue.", "")
        End If
    End Sub

End Module
