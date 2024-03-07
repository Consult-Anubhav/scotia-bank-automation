Module Faker
    Private FakeRootPath
    Public Function FakeEmailSubject() As String
        Return "* Scotia Report - *"
    End Function

    Public Function GetFakeRootPath() As String
        Return FakeRootPath

    End Function

    Public Function SetFakeRootPath(path As String)
        FakeRootPath = path

    End Function

    Public Function FakeK2Path() As String
        Return "Supporting Files K2 and Murex\K2"
    End Function

    Public Function FakeOPICSPath() As String
        Return "OPICS Scotia Investments Jamaica Limited"
    End Function

    Public Function FakeSCOTSPath() As String
        Return "SCOTS"
    End Function

    Public Function FakeLATAMPath() As String
        Return "Latam De Minimis Calculation"
    End Function

    Public Function FakeLATAMCFTCPath() As String
        Return "Latam De Minimis Calculation\CFTC Deminimis LatAm Extracts"
    End Function

    Public Function FakeLATAMUSPPath() As String
        Return "Latam De Minimis Calculation\CFTC Deminimis LatAm Extracts\US Person List"
    End Function

End Module
