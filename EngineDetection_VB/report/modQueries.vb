Friend Module modQueries
#Region "Validation"
    Public Function EventSources() As String
        Return _
            "
                SELECT
                e.EventName,
                s.SourceName

                FROM dim.Events e
                JOIN dim.Sources s
                    ON e.SourceID = s.SourceID

                WHERE e.EventName = @EventName
            "
    End Function

    Public Function NameSources() As String
        Return _
            "
                SELECT
                p.FirstName,
                p.LastName,
                s.SourceName

                FROM dim.Players p
                JOIN dim.Sources s
                    ON p.SourceID = s.SourceID

                WHERE p.FirstName = @FirstName
                AND p.LastName = @LastName
            "
    End Function

    Public Function CompareSources() As String
        Return _
            "
                SELECT DISTINCT
                s.SourceName

                FROM stat.StatisticsSummary ss
                JOIN dim.Sources s
                    ON ss.SourceID = s.SourceID
            "
    End Function

    Public Function CompareTimeControls() As String
        Return _
            "
                SELECT DISTINCT
                tc.TimeControlName

                FROM stat.StatisticsSummary ss
                JOIN dim.Sources s
                    ON ss.SourceID = s.SourceID                
                JOIN dim.TimeControls tc
                    ON ss.TimeControlID = tc.TimeControlID                

                WHERE s.SourceName = @SourceName
            "
    End Function

    Public Function CompareRatingIDs() As String
        Return _
            "
                SELECT DISTINCT
                ss.RatingID

                FROM stat.StatisticsSummary ss
                JOIN dim.Sources s
                    ON ss.SourceID = s.SourceID
                JOIN dim.TimeControls tc
                    ON ss.TimeControlID = tc.TimeControlID

                WHERE s.SourceName = @SourceName
                AND tc.TimeControlName = @TimeControlName

                ORDER BY ss.RatingID
            "
    End Function
#End Region

#Region "ID's"
    Public Function SourceID() As String
        Return "SELECT SourceID FROM dim.Sources WHERE SourceName = @SourceName"
    End Function

    Public Function EventID() As String
        Return "SELECT EventID FROM dim.Events WHERE EventName = @EventName"
    End Function

    Public Function TimeControlID() As String
        Return "SELECT TimeControlID FROM dim.TimeControls WHERE TimeControlName = @TimeControlName"
    End Function

    Public Function ScoreID() As String
        Return "SELECT ScoreID FROM dim.Scores WHERE ScoreName = @ScoreName"
    End Function

    Public Function PlayerID() As String
        Return _
            "
                SELECT
                p.PlayerID

                FROM dim.Players p
                JOIN dim.Sources s ON p.SourceID = s.SourceID

                WHERE p.FirstName = @FirstName
                AND p.LastName = @LastName
                AND s.SourceName = @SourceName
            "
    End Function
#End Region

#Region "Info"
    Public Function EventEngine() As String
        Return _
            "
                SELECT TOP(1)
                eng.EngineName

                FROM lake.Moves m
                JOIN lake.Games g
                    ON m.GameID = g.GameID
                JOIN dim.Engines eng
                    ON m.EngineID = eng.EngineID

                WHERE g.EventID = @EventID

                GROUP BY
                eng.EngineName

                ORDER BY
                COUNT(m.MoveNumber) DESC
            "
    End Function

    Public Function PlayerEngine() As String
        Return _
            "
                SELECT TOP(1)
                eng.EngineName

                FROM lake.Moves m
                JOIN lake.Games g
	                ON m.GameID = g.GameID
                JOIN dim.Colors c
	                ON m.ColorID = c.ColorID
                JOIN dim.Players wp
	                ON g.WhitePlayerID = wp.PlayerID
                JOIN dim.Players bp
	                ON g.BlackPlayerID = bp.PlayerID
                JOIN dim.Engines eng
	                ON m.EngineID = eng.EngineID

                WHERE (CASE WHEN c.Color = 'White' THEN g.WhitePlayerID ELSE g.BlackPlayerID END) = @PlayerID
                AND g.GameDate BETWEEN @StartDate AND @EndDate

                GROUP BY
                eng.EngineName

                ORDER BY
                COUNT(m.MoveNumber) DESC
            "
    End Function

    Public Function EventDepth() As String
        Return _
            "
                SELECT TOP(1)
                m.Depth

                FROM lake.Moves m
                JOIN lake.Games g
                    ON m.GameID = g.GameID

                WHERE g.EventID = @EventID

                GROUP BY
                m.Depth

                ORDER BY
                COUNT(m.MoveNumber) DESC
            "
    End Function

    Public Function PlayerDepth() As String
        Return _
            "
                SELECT TOP(1)
                m.Depth

                FROM lake.Moves m
                JOIN lake.Games g
	                ON m.GameID = g.GameID
                JOIN dim.Colors c
	                ON m.ColorID = c.ColorID
                JOIN dim.Players wp
	                ON g.WhitePlayerID = wp.PlayerID
                JOIN dim.Players bp
	                ON g.BlackPlayerID = bp.PlayerID

                WHERE (CASE WHEN c.Color = 'White' THEN g.WhitePlayerID ELSE g.BlackPlayerID END) = @PlayerID
                AND m.IsTablebase = 0
                AND g.GameDate BETWEEN @StartDate AND @EndDate

                GROUP BY
                m.Depth

                ORDER BY
                COUNT(m.MoveNumber) DESC
            "
    End Function

    Public Function MaxEval() As String
        Return "SELECT Value FROM dbo.Settings WHERE Name = 'Max Eval'"
    End Function

    Public Function EventSummary() As String
        Return _
            "
                SELECT
                CONVERT(varchar(10), MIN(GameDate), 101) + ' - ' + CONVERT(varchar(10), MAX(GameDate), 101) AS EventDates,
                MAX(RoundNum) AS Rounds,
                (COUNT(DISTINCT WhitePlayerID) + COUNT(DISTINCT BlackPlayerID))/2 AS Players

                FROM lake.Games

                WHERE EventID = @EventID
            "
    End Function
#End Region
End Module
