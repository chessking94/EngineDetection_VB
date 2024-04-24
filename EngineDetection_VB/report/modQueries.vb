Friend Module modQueries
#Region "Validation"
    Public Function EventSources() As String
        Dim rtn_Query As String =
            "
                SELECT
                e.EventName,
                s.SourceName

                FROM dim.Events e
                JOIN dim.Sources s ON e.SourceID = s.SourceID

                WHERE e.EventName = @EventName
            "
        Return rtn_Query
    End Function

    Public Function NameSources() As String
        Dim rtn_Query As String =
            "
                SELECT
                p.FirstName,
                p.LastName,
                s.SourceName

                FROM dim.Players p
                JOIN dim.Sources s ON p.SourceID = s.SourceID

                WHERE p.FirstName = @FirstName
                AND p.LastName = @LastName
            "
        Return rtn_Query
    End Function

    Public Function CompareSources() As String
        Dim rtn_Query As String =
            "
                SELECT DISTINCT
                s.SourceName

                FROM stat.StatisticsSummary ss
                JOIN dim.Sources s ON ss.SourceID = s.SourceID
            "
        Return rtn_Query
    End Function

    Public Function CompareTimeControls() As String
        Dim rtn_Query As String =
            "
                SELECT DISTINCT
                tc.TimeControlName

                FROM stat.StatisticsSummary ss
                JOIN dim.Sources s ON ss.SourceID = s.SourceID                
                JOIN dim.TimeControls tc ON ss.TimeControlID = tc.TimeControlID                

                WHERE s.SourceName = @SourceName
            "
        Return rtn_Query
    End Function

    Public Function CompareRatingIDs() As String
        Dim rtn_Query As String =
            "
                SELECT DISTINCT
                ss.RatingID

                FROM stat.StatisticsSummary ss
                JOIN dim.Sources s ON ss.SourceID = s.SourceID
                JOIN dim.TimeControls tc ON ss.TimeControlID = tc.TimeControlID

                WHERE s.SourceName = @SourceName
                AND tc.TimeControlName = @TimeControlName

                ORDER BY ss.RatingID
            "
        Return rtn_Query
    End Function
#End Region
End Module
