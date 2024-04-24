Public Class clsParameters
    Friend ReportType As String
    Friend EventName As String
    Friend FirstName As String
    Friend LastName As String
    Friend SourceName As String

    Friend CompareSourceName As String
    Friend CompareTimeControl As String
    Friend CompareRatingID As Short = -1
    Friend CompareScoreName As String

    Friend Sub ClearVariables()
        ReportType = ""
        EventName = ""
        FirstName = ""
        LastName = ""
        SourceName = ""
        CompareSourceName = ""
        CompareTimeControl = ""
        CompareRatingID = -1
        CompareScoreName = ""
    End Sub

    Friend Sub ClearCompareVariables()
        CompareSourceName = ""
        CompareTimeControl = ""
        CompareRatingID = -1
        CompareScoreName = ""
    End Sub
End Class
