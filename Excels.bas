Attribute VB_Name = "Excels"

'--- Latam ---
        
'--- OPICS ---

'--- SCOTS ---

'--- K2 ---

Public Sub GenerateK2Extract()
    Dim ExApp As Excel.Application
    Dim ExWbk As Workbook
    
    Set ExApp = New Excel.Application
    
    ExApp.AskToUpdateLinks = False
    ExApp.DisplayAlerts = False
    ExApp.Visible = True
    
    DisplayWindowsNotification "K2 Extract", "Opening excel"
    Set ExWbk = ExApp.Workbooks.Open("C:\wamp64\www\~Consult Anubhav Projects\scotia-bank-automation\K2 and Portal Data Summary_Jan 1 2022 - Dec 31 2023.xlsm")
    
    DisplayWindowsNotification "K2 Extract", "CCDExtractCSV"
    ExWbk.Application.Run "Module1.CCDExtractCSV"
    
    DisplayWindowsNotification "K2 Extract", "CFCTE"
    ExWbk.Application.Run "Module2.CFCTE"
    
    DisplayWindowsNotification "K2 Extract", "Saving File"
    ExWbk.Close SaveChanges:=True
End Sub

'--- Mutex ---

Public Sub GenerateCCDExtract()
    Dim ExApp As Excel.Application
    Dim ExWbk As Workbook
    
    Set ExApp = New Excel.Application
    
    ExApp.AskToUpdateLinks = False
    ExApp.DisplayAlerts = False
    ExApp.Visible = True
    
    DisplayWindowsNotification "CCD Extract", "Opening excel"
    Set ExWbk = ExApp.Workbooks.Open("C:\wamp64\www\~Consult Anubhav Projects\scotia-bank-automation\DF_DeMinimis_Extract (01012023-12312023).xlsm")
    
    DisplayWindowsNotification "CCD Extract", "CopyAndTrimSpecialEntity"
    ExWbk.Application.Run "CopyAndTrimSpecialEntity.CopyAndTrimSpecialEntity"
    
    DisplayWindowsNotification "CCD Extract", "Saving File"
    ExWbk.Close SaveChanges:=True
End Sub

'--- Calculations ---

    
