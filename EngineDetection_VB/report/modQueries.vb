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

#Region "Parameters"
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

    Public Function EventAvgRating() As String
        Return "SELECT ROUND(AVG((WhiteElo + BlackElo)/2), 0) AS AvgGameRating FROM lake.Games WHERE EventID = @EventID"
    End Function

    Public Function PlayerAvgRating() As String
        Return _
            "
                SELECT
                AVG(r.Elo) AS Elo

                FROM (
                    SELECT
                    NULLIF(NULLIF(WhiteElo, ''), 0) AS Elo

                    FROM lake.Games

                    WHERE WhitePlayerID = @PlayerID
                    AND GameDate BETWEEN @StartDate AND @EndDate

                    UNION ALL

                    SELECT
                    NULLIF(NULLIF(BlackElo, ''), 0) AS Elo

                    FROM lake.Games

                    WHERE BlackPlayerID = @PlayerID
                    AND GameDate BETWEEN @StartDate AND @EndDate
                ) r
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

#Region "Detail"
#Region "Event Detail"
    Public Function EventRatings() As String
        Return _
            "
                SELECT
                ROUND(AVG((WhiteElo + BlackElo)/2), 0) AS AvgRating,
                ROUND(MIN((WhiteElo + BlackElo)/2), 0) AS MinRating,
                ROUND(MAX((WhiteElo + BlackElo)/2), 0) AS MaxRating

                FROM lake.Games

                WHERE EventID = @EventID
            "
    End Function

    Public Function EventMoveCounts() As String
        Return _
            "
                SELECT
                COUNT(m.MoveNumber) AS TotalMoves,
                SUM(CAST(m.MoveScored AS tinyint)) AS ScoredMoves

                FROM lake.Moves m
                JOIN lake.Games g
                    ON m.GameID = g.GameID

                WHERE g.EventID = @EventID
            "
    End Function

    Public Function EventTraces() As String
        Return _
            "
                SELECT
                t.TraceKey,
                t.TraceDescription,
                COUNT(m.MoveNumber) AS MoveCount

                FROM lake.Moves m
                JOIN lake.Games g
                    ON m.GameID = g.GameID
                JOIN dim.Traces t
                    ON m.TraceKey = t.TraceKey

                WHERE g.EventID = @EventID
                AND t.TraceKey NOT IN ('0', 'M')

                GROUP BY
                t.TraceKey,
                t.TraceDescription

                ORDER BY
                MoveCount DESC
            "
    End Function

    Public Function EventBaseStats() As String
        Return _
            "
                SELECT
                SUM(CASE WHEN m.Move_Rank <= 1 THEN 1 ELSE 0 END) AS T1,
                SUM(CASE WHEN m.Move_Rank <= 2 THEN 1 ELSE 0 END) AS T2,
                SUM(CASE WHEN m.Move_Rank <= 3 THEN 1 ELSE 0 END) AS T3,
                SUM(CASE WHEN m.Move_Rank <= 4 THEN 1 ELSE 0 END) AS T4,
                SUM(CASE WHEN m.Move_Rank <= 5 THEN 1 ELSE 0 END) AS T5,
                AVG(m.ScACPL) AS ACPL,
                ISNULL(STDEV(m.ScACPL), 0) AS SDCPL,
                SUM(CASE WHEN m.CP_Loss > 2 THEN 1 ELSE 0 END) AS Blunders

                FROM lake.Moves m
                JOIN lake.Games g
                    ON m.GameID = g.GameID
                JOIN dim.Colors c
                    ON m.ColorID = c.ColorID

                WHERE g.EventID = @EventID
                AND m.MoveScored = 1
            "
    End Function

    Public Function EventTotalScore() As String
        Return _
            "
                SELECT
                CASE
                    WHEN ISNULL(100*SUM(ms.ScoreValue)/NULLIF(SUM(ms.MaxScoreValue), 0), 100) > 100 THEN 100
                    ELSE ISNULL(100*SUM(ms.ScoreValue)/NULLIF(SUM(ms.MaxScoreValue), 0), 100)
                END AS Score

                FROM lake.Moves m
                JOIN stat.MoveScores ms
                    ON m.GameID = ms.GameID
                    AND m.MoveNumber = ms.MoveNumber
                    AND m.ColorID = ms.ColorID
                JOIN lake.Games g
                    ON m.GameID = g.GameID

                WHERE g.EventID = @EventID
                AND ms.ScoreID = @ScoreID
                AND m.MoveScored = 1
            "
    End Function
#End Region

#Region "Player Detail"
    Public Function PlayerRatings() As String
        Return _
            "
                SELECT
                AVG(r.OppElo) AS AvgRating,
                MIN(r.OppElo) AS MinRating,
                MAX(r.OppElo) AS MaxRating

                FROM (
                    SELECT
                    NULLIF(NULLIF(WhiteElo, ''), 0) AS Elo,
                    NULLIF(NULLIF(BlackElo, ''), 0) AS OppElo

                    FROM lake.Games

                    WHERE WhitePlayerID = @PlayerID
                    AND GameDate BETWEEN @StartDate AND @EndDate

                    UNION ALL

                    SELECT
                    NULLIF(NULLIF(BlackElo, ''), 0) AS Elo,
                    NULLIF(NULLIF(WhiteElo, ''), 0) AS OppElo

                    FROM lake.Games

                    WHERE BlackPlayerID = @PlayerID
                    AND GameDate BETWEEN @StartDate AND @EndDate
                ) r
            "
    End Function

    Public Function PlayerMoveCounts() As String
        Return _
            "
                SELECT
                COUNT(m.MoveNumber) AS TotalMoves,
                SUM(CAST(m.MoveScored AS tinyint)) AS ScoredMoves

                FROM lake.Moves m
                JOIN lake.Games g
                    ON m.GameID = g.GameID
                JOIN dim.Colors c
                    ON m.ColorID = c.ColorID

                WHERE (CASE WHEN c.Color = 'White' THEN g.WhitePlayerID ELSE g.BlackPlayerID END) = @PlayerID
                AND g.GameDate BETWEEN @StartDate AND @EndDate
            "
    End Function

    Public Function PlayerTraces() As String
        Return _
            "
                SELECT
                t.TraceDescription,
                COUNT(m.MoveNumber) AS MoveCount

                FROM lake.Moves m
                JOIN lake.Games g
                    ON m.GameID = g.GameID
                JOIN dim.Traces t
                    ON m.TraceKey = t.TraceKey
                JOIN dim.Colors c
                    ON m.ColorID = c.ColorID

                WHERE (CASE WHEN c.Color = 'White' THEN g.WhitePlayerID ELSE g.BlackPlayerID END) = @PlayerID
                AND g.GameDate BETWEEN @StartDate AND @EndDate
                AND t.TraceKey NOT IN ('0', 'M')

                GROUP BY
                t.TraceDescription

                ORDER BY
                COUNT(m.MoveNumber) DESC
            "
    End Function

    Public Function PlayerBaseStats() As String
        Return _
            "
                SELECT
                SUM(CASE WHEN m.Move_Rank <= 1 THEN 1 ELSE 0 END) AS T1,
                SUM(CASE WHEN m.Move_Rank <= 2 THEN 1 ELSE 0 END) AS T2,
                SUM(CASE WHEN m.Move_Rank <= 3 THEN 1 ELSE 0 END) AS T3,
                SUM(CASE WHEN m.Move_Rank <= 4 THEN 1 ELSE 0 END) AS T4,
                SUM(CASE WHEN m.Move_Rank <= 5 THEN 1 ELSE 0 END) AS T5,
                AVG(m.ScACPL) AS ACPL,
                ISNULL(STDEV(m.ScACPL), 0) AS SDCPL,
                SUM(CASE WHEN m.CP_Loss > 2 THEN 1 ELSE 0 END) AS Blunders

                FROM lake.Moves m
                JOIN lake.Games g
                    ON m.GameID = g.GameID
                JOIN dim.Colors c
                    ON m.ColorID = c.ColorID

                WHERE (CASE WHEN c.Color = 'White' THEN g.WhitePlayerID ELSE g.BlackPlayerID END) = @PlayerID
                AND g.GameDate BETWEEN @StartDate AND @EndDate
                AND m.MoveScored = 1
            "
    End Function

    Public Function PlayerTotalScore() As String
        Return _
            "
                SELECT
                CASE
                    WHEN ISNULL(100*SUM(ms.ScoreValue)/NULLIF(SUM(ms.MaxScoreValue), 0), 100) > 100 THEN 100
                    ELSE ISNULL(100*SUM(ms.ScoreValue)/NULLIF(SUM(ms.MaxScoreValue), 0), 100)
                END AS Score

                FROM lake.Moves m
                JOIN stat.MoveScores ms
                    ON m.GameID = ms.GameID
                    AND m.MoveNumber = ms.MoveNumber
                    AND m.ColorID = ms.ColorID
                JOIN lake.Games g
                    ON m.GameID = g.GameID
                JOIN dim.Colors c
                    ON m.ColorID = c.ColorID

                WHERE (CASE WHEN c.Color = 'White' THEN g.WhitePlayerID ELSE g.BlackPlayerID END) = @PlayerID
                AND g.GameDate BETWEEN @StartDate AND @EndDate
                AND ms.ScoreID = @ScoreID
                AND m.MoveScored = 1
            "
    End Function
#End Region

#Region "Other Detail"
    Public Function ZScoreData(Optional ColorID As Short = 0) As String
        Dim qry As String =
            "
SELECT
ms.MeasurementName,
ss.Average,
ss.StandardDeviation

FROM stat.StatisticsSummary ss
JOIN dim.Aggregations agg
    ON ss.AggregationID = agg.AggregationID
JOIN dim.Measurements ms
    ON ss.MeasurementID = ms.MeasurementID
LEFT JOIN dim.Colors c
    ON ss.ColorID = c.ColorID

WHERE agg.AggregationName = @AggregationName
AND ms.MeasurementName IN ('T1', 'ScACPL', @ScoreName)
AND ss.SourceID = @SourceID
AND ss.TimeControlID = @TimeControlID
AND ss.RatingID = @RatingID
            "

        If ColorID > 0 Then qry += $"AND c.ColorID = {ColorID}"

        Return qry
    End Function
#End Region
#End Region
End Module
