Imports Microsoft.Data.SqlClient

Public Class clsParameters
    Friend ReportType As String
    Friend EventName As String
    Friend FirstName As String 'TODO: Standardize the first and last names to proper casing at some point
    Friend LastName As String
    Friend StartDate As Date?  'need these dates to be nullable since they are coming from a DatePicker
    Friend EndDate As Date?
    Friend SourceName As String

    'set the Compare* variables to defaults
    Friend CompareSourceName As String = "Control"
    Friend CompareTimeControl As String = "Classical"
    Friend CompareRatingID As Short = -1
    Friend CompareScoreName As String = "WinProbabilityLost"

    Friend EventID As Long = 0
    Friend PlayerID As Long = 0
    Friend SourceID As Short = 0
    Friend CompareSourceID As Short = 0
    Friend CompareTimeControlID As Short = 0
    Friend CompareScoreID As Short = 0

    Friend Sub ClearVariables()
        'reset everything to the default values
        ReportType = ""
        EventName = ""
        FirstName = ""
        LastName = ""
        StartDate = Nothing
        EndDate = Nothing
        SourceName = ""

        EventID = 0
        PlayerID = 0
        SourceID = 0

        ClearCompareVariables()
    End Sub

    Friend Sub ClearCompareVariables()
        CompareSourceName = "Control"
        CompareTimeControl = "Classical"
        CompareRatingID = -1
        CompareScoreName = "WinProbabilityLost"
        CompareSourceID = 0
        CompareTimeControlID = 0
        CompareScoreID = 0
    End Sub

    Friend Sub PopulateIDVariables()
        Select Case ReportType
            Case "Event"
                EventID = GetEventID(EventName)
            Case "Player"
                PlayerID = GetPlayerID(FirstName, LastName, SourceName)
        End Select

        SourceID = GetSourceID(SourceName)

        If CompareRatingID = -1 Then
            CompareRatingID = GetDefaultRatingID()
        Else
            CompareSourceID = GetSourceID(CompareSourceName)
            CompareTimeControlID = GetTimeControlID(CompareTimeControl)
            CompareScoreID = GetScoreID(CompareScoreName)
        End If
    End Sub

    Private Function GetDefaultRatingID() As Short
        Dim rtnval As Long
        Select Case ReportType
            Case "Event"
                Using objl_CMD As New SqlCommand(modQueries.EventAvgRating(), MainWindow.db_Connection)
                    objl_CMD.Parameters.AddWithValue("@EventID", EventID)
                    rtnval = Convert.ToInt64(objl_CMD.ExecuteScalar())
                End Using
            Case "Player"
                Using objl_CMD As New SqlCommand(modQueries.PlayerAvgRating(), MainWindow.db_Connection)
                    objl_CMD.Parameters.AddWithValue("@PlayerID", PlayerID)
                    objl_CMD.Parameters.AddWithValue("@StartDate", StartDate)
                    objl_CMD.Parameters.AddWithValue("@EndDate", EndDate)
                    rtnval = Convert.ToInt64(objl_CMD.ExecuteScalar())
                End Using
        End Select

        Return Math.Floor(rtnval / 100) * 100
    End Function

#Region "ID Functions"
    Private Function GetEventID(pi_EventName As String) As Long
        Dim rtn_EventID As Long = 0

        Using objl_CMD As New SqlCommand(modQueries.EventID(), MainWindow.db_Connection)
            objl_CMD.Parameters.AddWithValue("@EventName", pi_EventName)
            rtn_EventID = Convert.ToInt64(objl_CMD.ExecuteScalar())
        End Using

        Return rtn_EventID
    End Function

    Private Function GetPlayerID(pi_FirstName As String, pi_LastName As String, pi_SourceName As String) As Long
        Dim rtn_PlayerID As Long = 0

        Using objl_CMD As New SqlCommand(modQueries.PlayerID(), MainWindow.db_Connection)
            objl_CMD.Parameters.AddWithValue("@FirstName", pi_FirstName)
            objl_CMD.Parameters.AddWithValue("@LastName", pi_LastName)
            objl_CMD.Parameters.AddWithValue("@SourceName", pi_SourceName)
            rtn_PlayerID = Convert.ToInt64(objl_CMD.ExecuteScalar())
        End Using

        Return rtn_PlayerID
    End Function

    Private Function GetSourceID(pi_SourceName As String) As Short
        Dim rtn_SourceID As Short = 0

        Using objl_CMD As New SqlCommand(modQueries.SourceID(), MainWindow.db_Connection)
            objl_CMD.Parameters.AddWithValue("@SourceName", pi_SourceName)
            rtn_SourceID = Convert.ToInt16(objl_CMD.ExecuteScalar())
        End Using

        Return rtn_SourceID
    End Function

    Private Function GetTimeControlID(pi_TimeControlName As String) As Short
        Dim rtn_TimeControlID As Short = 0

        Using objl_CMD As New SqlCommand(modQueries.TimeControlID(), MainWindow.db_Connection)
            objl_CMD.Parameters.AddWithValue("@TimeControlName", pi_TimeControlName)
            rtn_TimeControlID = Convert.ToInt16(objl_CMD.ExecuteScalar())
        End Using

        Return rtn_TimeControlID
    End Function

    Private Function GetScoreID(pi_ScoreName As String) As Short
        Dim rtn_ScoreID As Short = 0

        Using objl_CMD As New SqlCommand(modQueries.ScoreID(), MainWindow.db_Connection)
            objl_CMD.Parameters.AddWithValue("@ScoreName", pi_ScoreName)
            rtn_ScoreID = Convert.ToInt16(objl_CMD.ExecuteScalar())
        End Using

        Return rtn_ScoreID
    End Function
#End Region
End Class
