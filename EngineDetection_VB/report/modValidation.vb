Imports Microsoft.Data.SqlClient

Friend Module modValidation
    Friend Function EventName(pi_EventName As String) As List(Of String)
        '''Return possible source values for a given event
        Dim rtn_SourceNames As New List(Of String)

        Using objl_CMD As New SqlCommand(modQueries.EventSources(), MainWindow.db_Connection)
            objl_CMD.Parameters.AddWithValue("@EventName", pi_EventName)
            With objl_CMD.ExecuteReader
                While .Read
                    rtn_SourceNames.Add(.Item("SourceName"))
                End While
            End With
        End Using

        Return rtn_SourceNames
    End Function

    Friend Function PlayerName(pi_FirstName As String, pi_LastName As String) As List(Of String)
        '''Return possible source values for a given player
        Dim rtn_SourceNames As New List(Of String)

        Using objl_CMD As New SqlCommand(modQueries.NameSources(), MainWindow.db_Connection)
            objl_CMD.Parameters.AddWithValue("@FirstName", pi_FirstName)
            objl_CMD.Parameters.AddWithValue("@LastName", pi_LastName)
            With objl_CMD.ExecuteReader
                While .Read
                    rtn_SourceNames.Add(.Item("SourceName"))
                End While
            End With
        End Using

        Return rtn_SourceNames
    End Function

    Friend Function CompareSources() As List(Of String)
        '''Return possible sources for comparision statistics
        Dim rtn_SourceNames As New List(Of String)

        Using objl_CMD As New SqlCommand(modQueries.CompareSources(), MainWindow.db_Connection)
            With objl_CMD.ExecuteReader
                While .Read
                    rtn_SourceNames.Add(.Item("SourceName"))
                End While
            End With
        End Using

        Return rtn_SourceNames
    End Function

    Friend Function CompareTimeControls(pi_SourceName As String) As List(Of String)
        '''Return possible time controls for comparison statistics
        Dim rtn_TimeControls As New List(Of String)

        Using objl_CMD As New SqlCommand(modQueries.CompareTimeControls(), MainWindow.db_Connection)
            objl_CMD.Parameters.AddWithValue("@SourceName", pi_SourceName)
            With objl_CMD.ExecuteReader
                While .Read
                    rtn_TimeControls.Add(.Item("TimeControlName"))
                End While
            End With
        End Using

        Return rtn_TimeControls
    End Function

    Friend Function CompareRatingIDs(pi_SourceName As String, pi_TimeControl As String) As List(Of Short)
        '''Return possible rating ID's for comparison statistics
        Dim rtn_RatingIDs As New List(Of Short)

        Using objl_CMD As New SqlCommand(modQueries.CompareRatingIDs(), MainWindow.db_Connection)
            objl_CMD.Parameters.AddWithValue("@SourceName", pi_SourceName)
            objl_CMD.Parameters.AddWithValue("@TimeControlName", pi_TimeControl)
            With objl_CMD.ExecuteReader
                While .Read
                    rtn_RatingIDs.Add(.Item("RatingID"))
                End While
            End With
        End Using

        Return rtn_RatingIDs
    End Function
End Module
