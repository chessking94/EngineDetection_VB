Imports Microsoft.Data.SqlClient

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

    Friend EventID As Long = 0
    Friend PlayerID As Long = 0
    Friend SourceID As Short = 0
    Friend CompareSourceID As Short = 0
    Friend CompareTimeControlID As Short = 0
    Friend CompareScoreID As Short = 0

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

        EventID = 0
        PlayerID = 0
        SourceID = 0
        CompareSourceID = 0
        CompareTimeControl = 0
        CompareScoreID = 0
    End Sub

    Friend Sub ClearCompareVariables()
        CompareSourceName = ""
        CompareTimeControl = ""
        CompareRatingID = -1
        CompareScoreName = ""
        CompareSourceID = 0
        CompareTimeControl = 0
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

        If CompareSourceName <> "" Then
            CompareSourceID = GetSourceID(CompareSourceName)
            CompareTimeControlID = GetTimeControlID(CompareTimeControl)
            CompareScoreID = GetScoreID(CompareScoreName)
        End If
    End Sub

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
