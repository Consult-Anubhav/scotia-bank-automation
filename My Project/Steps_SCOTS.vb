Module Steps_SCOTS

    '--- SCOTS ---

    Public Sub TestSCOTS()
        Dim inputDir, outputDir, emailMonthYear, previousMonthYear, emailYear, previousYear, emailMonth, previousMonth As String

        'Assign Variables
        emailMonthYear = GetEmailMonthYear("Re: Scotia Report - Dec 2023")
        previousMonthYear = GetPreviousMonthYear(emailMonthYear & "")
        emailYear = GetEmailYear(emailMonthYear & "")
        previousYear = GetPreviousYear(emailMonthYear & "")
        emailMonth = GetEmailMonth(emailMonthYear & "")
        previousMonth = GetPreviousMonth(emailMonthYear & "")
        outputDir = GetFakeRootPath() & "\" & emailYear & "\" & emailMonth
        inputDir = GetFakeRootPath() & "\" & previousYear & "\" & previousMonth
        'Test K2
        'GenerateK2Extract outputDir & ""
        'Test Murex
        'GenerateMutexExtract outputDir & ""
    End Sub

    '--- SCOTS ---

End Module
