Imports Microsoft.Data.SqlClient
Imports MathNet.Numerics.Distributions

Public Class clsDetail
    Private ReadOnly params As clsParameters = MainWindow.objl_Parameters
    Private Const EventLength As Short = 35

#Region "Key Stats"
    Friend Sub KeyStats()
        Select Case params.ReportType
            Case "Event"
                objm_Lines.Add("WHOLE-EVENT STATISTICS:")
            Case "Player"
                objm_Lines.Add("WHOLE-SAMPLE STATISTICS:")
        End Select

        objm_Lines.Add(New String("-", 25))

        Dim stats As New _keystats With {
            .objl_Parameters = params
        }
        stats.Build()

        'rating summary
        Dim tempText As String = ""
        Select Case params.ReportType
            Case "Event"
                tempText = "Average rating by game:"
            Case "Player"
                tempText = "Average opponent rating:"
        End Select
        objm_Lines.Add(tempText.PadRight(EventLength, " "c) & $"{stats.AvgRating}; min {stats.MinRating}, max {stats.MaxRating}")

        'trace summary
        objm_Lines.Add("Scored Moves:".PadRight(EventLength, " "c) & $"{stats.ScoredMoves} / {stats.TotalMoves} = {Convert.ToDouble(100 * stats.ScoredMoves / stats.TotalMoves):0.00}%")
        For Each kvp As KeyValuePair(Of String, Short) In stats.TraceCounts
            objm_Lines.Add($"{kvp.Key}:".PadRight(EventLength, " "c) & $"{kvp.Value} / {stats.TotalMoves} = {Convert.ToDouble(100 * kvp.Value / stats.TotalMoves):0.00}%")
        Next
        objm_Lines.Add("")

        'base stats
        For Each kvp As KeyValuePair(Of String, Short) In stats.TCounts
            objm_Lines.Add($"{kvp.Key}:".PadRight(EventLength, " "c) & $"{kvp.Value} / {stats.ScoredMoves} = {Convert.ToDouble(100 * kvp.Value / stats.ScoredMoves):0.00}%")
        Next
        objm_Lines.Add("Blunders:".PadRight(EventLength, " "c) & $"{stats.Blunders} / {stats.ScoredMoves} = {Convert.ToDouble(100 * stats.Blunders / stats.ScoredMoves):0.00}%")
        objm_Lines.Add("ScACPL:".PadRight(EventLength, " "c) & stats.ACPL.ToString("0.0000"))
        objm_Lines.Add("ScSDCPL:".PadRight(EventLength, " "c) & stats.SDCPL.ToString("0.0000"))

        'advanced stats
        'TODO: Add the asterisks to each of these when needed
        objm_Lines.Add("Score:".PadRight(EventLength, " "c) & stats.Score.ToString("0.00"))
        objm_Lines.Add("ROI:".PadRight(EventLength, " "c) & stats.ROI.ToString("0.0"))
        objm_Lines.Add("PValue:".PadRight(EventLength, " "c) & $"{100 * stats.PValue:0.00}%")
        objm_Lines.Add("")
        objm_Lines.Add("")
    End Sub

    Private Class _keystats
        Friend Property objl_Parameters As clsParameters
        Friend Property AvgRating As Short
        Friend Property MinRating As Short
        Friend Property MaxRating As Short
        Friend Property ScoredMoves As Short
        Friend Property TotalMoves As Short
        Friend Property TraceCounts As New Dictionary(Of String, Short)
        Friend Property TCounts As New Dictionary(Of String, Short)
        Friend Property ACPL As Double
        Friend Property SDCPL As Double
        Friend Property Blunders As Short
        Friend Property Score As Double
        Friend Property ROI As Double
        Friend Property PValue As Double

        Friend Sub Build()
            Ratings()
            MoveCounts()
            Traces()
            BaseStats()
            AdvancedStats()
        End Sub

        Private Sub Ratings()
            Dim objl_CMD As New SqlCommand
            Select Case objl_Parameters.ReportType
                Case "Event"
                    With objl_CMD
                        .Connection = MainWindow.db_Connection
                        .CommandText = modQueries.EventRatings()
                        .Parameters.AddWithValue("@EventID", objl_Parameters.EventID)
                    End With
                Case "Player"
                    With objl_CMD
                        .Connection = MainWindow.db_Connection
                        .CommandText = modQueries.PlayerRatings()
                        .Parameters.AddWithValue("@PlayerID", objl_Parameters.PlayerID)
                        .Parameters.AddWithValue("@StartDate", objl_Parameters.StartDate)
                        .Parameters.AddWithValue("@EndDate", objl_Parameters.EndDate)
                    End With
            End Select

            With objl_CMD.ExecuteReader
                While .Read
                    AvgRating = Convert.ToInt16(.Item("AvgRating"))
                    MinRating = Convert.ToInt16(.Item("MinRating"))
                    MaxRating = Convert.ToInt16(.Item("MaxRating"))
                End While
                .Close()
            End With
        End Sub

        Private Sub MoveCounts()
            Dim objl_CMD As New SqlCommand
            Select Case objl_Parameters.ReportType
                Case "Event"
                    With objl_CMD
                        .Connection = MainWindow.db_Connection
                        .CommandText = modQueries.EventMoveCounts()
                        .Parameters.AddWithValue("@EventID", objl_Parameters.EventID)
                    End With
                Case "Player"
                    With objl_CMD
                        .Connection = MainWindow.db_Connection
                        .CommandText = modQueries.PlayerMoveCounts()
                        .Parameters.AddWithValue("@PlayerID", objl_Parameters.PlayerID)
                        .Parameters.AddWithValue("@StartDate", objl_Parameters.StartDate)
                        .Parameters.AddWithValue("@EndDate", objl_Parameters.EndDate)
                    End With
            End Select

            With objl_CMD.ExecuteReader
                While .Read
                    TotalMoves = Convert.ToInt16(.Item("TotalMoves"))
                    ScoredMoves = Convert.ToInt16(.Item("ScoredMoves"))
                End While
                .Close()
            End With
        End Sub

        Private Sub Traces()
            Dim objl_CMD As New SqlCommand
            Select Case objl_Parameters.ReportType
                Case "Event"
                    With objl_CMD
                        .Connection = MainWindow.db_Connection
                        .CommandText = modQueries.EventTraces()
                        .Parameters.AddWithValue("@EventID", objl_Parameters.EventID)
                    End With
                Case "Player"
                    With objl_CMD
                        .Connection = MainWindow.db_Connection
                        .CommandText = modQueries.PlayerTraces()
                        .Parameters.AddWithValue("@PlayerID", objl_Parameters.PlayerID)
                        .Parameters.AddWithValue("@StartDate", objl_Parameters.StartDate)
                        .Parameters.AddWithValue("@EndDate", objl_Parameters.EndDate)
                    End With
            End Select

            With objl_CMD.ExecuteReader
                While .Read
                    TraceCounts.Add(.Item("TraceDescription"), .Item("MoveCount"))
                End While
                .Close()
            End With
        End Sub

        Private Sub BaseStats()
            Dim objl_CMD As New SqlCommand
            Select Case objl_Parameters.ReportType
                Case "Event"
                    With objl_CMD
                        .Connection = MainWindow.db_Connection
                        .CommandText = modQueries.EventBaseStats()
                        .Parameters.AddWithValue("@EventID", objl_Parameters.EventID)
                    End With
                Case "Player"
                    With objl_CMD
                        .Connection = MainWindow.db_Connection
                        .CommandText = modQueries.PlayerBaseStats()
                        .Parameters.AddWithValue("@PlayerID", objl_Parameters.PlayerID)
                        .Parameters.AddWithValue("@StartDate", objl_Parameters.StartDate)
                        .Parameters.AddWithValue("@EndDate", objl_Parameters.EndDate)
                    End With
            End Select

            With objl_CMD.ExecuteReader
                While .Read
                    TCounts.Add("T1", .Item("T1"))
                    TCounts.Add("T2", .Item("T2"))
                    TCounts.Add("T3", .Item("T3"))
                    TCounts.Add("T4", .Item("T4"))
                    TCounts.Add("T5", .Item("T5"))
                    ACPL = .Item("ACPL")
                    SDCPL = .Item("SDCPL")
                    Blunders = .Item("Blunders")
                End While
                .Close()
            End With
        End Sub

        Private Sub AdvancedStats()
            'score
            Dim objl_CMD As New SqlCommand
            Select Case objl_Parameters.ReportType
                Case "Event"
                    With objl_CMD
                        .Connection = MainWindow.db_Connection
                        .CommandText = modQueries.EventTotalScore()
                        .Parameters.AddWithValue("@EventID", objl_Parameters.EventID)
                        .Parameters.AddWithValue("@ScoreID", objl_Parameters.CompareScoreID)
                    End With
                Case "Player"
                    With objl_CMD
                        .Connection = MainWindow.db_Connection
                        .CommandText = modQueries.PlayerTotalScore()
                        .Parameters.AddWithValue("@PlayerID", objl_Parameters.PlayerID)
                        .Parameters.AddWithValue("@StartDate", objl_Parameters.StartDate)
                        .Parameters.AddWithValue("@EndDate", objl_Parameters.EndDate)
                        .Parameters.AddWithValue("@ScoreID", objl_Parameters.CompareScoreID)
                    End With
            End Select

            With objl_CMD.ExecuteReader
                While .Read
                    Score = .Item("Score")
                End While
                .Close()
            End With

            'ROI
            With objl_CMD
                .Parameters.Clear()
                .CommandText = modQueries.ZScoreData()
                .Parameters.AddWithValue("AggregationName", "Event")  'since this value is inclusive of multiple games, it should always be compared against Event
                .Parameters.AddWithValue("@ScoreName", objl_Parameters.CompareScoreName)
                .Parameters.AddWithValue("@SourceID", 3)  'hard-coding sourceID since Lichess doesn't have event stats
                .Parameters.AddWithValue("@TimeControlID", objl_Parameters.CompareTimeControlID)
                .Parameters.AddWithValue("@RatingID", objl_Parameters.CompareRatingID)
            End With

            Dim objm_z As New Dictionary(Of String, Double)
            With objl_CMD.ExecuteReader
                While .Read
                    Dim z_score As New Double
                    Select Case .Item("MeasurementName")
                        Case "T1"
                            z_score = ((Convert.ToDouble(TCounts("T1")) / ScoredMoves) - Convert.ToDouble(.Item("Average"))) / Convert.ToDouble(.Item("StandardDeviation"))
                            objm_z.Add("T1", z_score)
                        Case "ScACPL"
                            z_score = -1 * (ACPL - Convert.ToDouble(.Item("Average"))) / Convert.ToDouble(.Item("StandardDeviation"))
                            objm_z.Add("ScACPL", z_score)
                        Case Else
                            'all possible score measurement names
                            z_score = (Score - Convert.ToDouble(.Item("Average"))) / Convert.ToDouble(.Item("StandardDeviation"))
                            objm_z.Add("Score", z_score)
                    End Select
                    ROI = CalculateROI(objm_z)
                End While
                .Close()
            End With

            'PValue
            Dim testStatistic As Double() = {Convert.ToDouble(TCounts("T1")) / TotalMoves, ACPL, Score}

            Dim t1Average As Double = 0
            With objl_CMD
                .Parameters.Clear()
                .CommandText = modQueries.GetStatAverage()
                .Parameters.AddWithValue("@SourceID", 3)  'hard-coding sourceID since Lichess doesn't have event stats
                .Parameters.AddWithValue("AggregationName", "Event")  'since this value is inclusive of multiple games, it should always be compared against Event
                .Parameters.AddWithValue("@RatingID", objl_Parameters.CompareRatingID)
                .Parameters.AddWithValue("@TimeControlID", objl_Parameters.CompareTimeControlID)
                .Parameters.AddWithValue("@ColorID", 0)
                .Parameters.AddWithValue("@EvaluationGroupID", 0)
                .Parameters.AddWithValue("@MeasurementName", "T1")
                t1Average = Convert.ToDouble(objl_CMD.ExecuteScalar())
            End With

            Dim t1Variance As Double = 0
            With objl_CMD
                .Parameters.Clear()
                .CommandText = modQueries.GetStatCovar()
                .Parameters.AddWithValue("@SourceID", 3)
                .Parameters.AddWithValue("AggregationName", "Event")
                .Parameters.AddWithValue("@RatingID", objl_Parameters.CompareRatingID)
                .Parameters.AddWithValue("@TimeControlID", objl_Parameters.CompareTimeControlID)
                .Parameters.AddWithValue("@ColorID", 0)
                .Parameters.AddWithValue("@EvaluationGroupID", 0)
                .Parameters.AddWithValue("@MeasurementName1", "T1")
                .Parameters.AddWithValue("@MeasurementName2", "T1")
                t1Variance = Convert.ToDouble(objl_CMD.ExecuteScalar())
            End With

            Dim cplAverage As Double = 0
            With objl_CMD
                .Parameters.Clear()
                .CommandText = modQueries.GetStatAverage()
                .Parameters.AddWithValue("@SourceID", 3)
                .Parameters.AddWithValue("AggregationName", "Event")
                .Parameters.AddWithValue("@RatingID", objl_Parameters.CompareRatingID)
                .Parameters.AddWithValue("@TimeControlID", objl_Parameters.CompareTimeControlID)
                .Parameters.AddWithValue("@ColorID", 0)
                .Parameters.AddWithValue("@EvaluationGroupID", 0)
                .Parameters.AddWithValue("@MeasurementName", "ScACPL")
                cplAverage = Convert.ToDouble(objl_CMD.ExecuteScalar())
            End With

            Dim cplVariance As Double = 0
            With objl_CMD
                .Parameters.Clear()
                .CommandText = modQueries.GetStatCovar()
                .Parameters.AddWithValue("@SourceID", 3)
                .Parameters.AddWithValue("AggregationName", "Event")
                .Parameters.AddWithValue("@RatingID", objl_Parameters.CompareRatingID)
                .Parameters.AddWithValue("@TimeControlID", objl_Parameters.CompareTimeControlID)
                .Parameters.AddWithValue("@ColorID", 0)
                .Parameters.AddWithValue("@EvaluationGroupID", 0)
                .Parameters.AddWithValue("@MeasurementName1", "ScACPL")
                .Parameters.AddWithValue("@MeasurementName2", "ScACPL")
                cplVariance = Convert.ToDouble(objl_CMD.ExecuteScalar())
            End With

            Dim scoreAverage As Double = 0
            With objl_CMD
                .Parameters.Clear()
                .CommandText = modQueries.GetStatAverage()
                .Parameters.AddWithValue("@SourceID", 3)
                .Parameters.AddWithValue("AggregationName", "Event")
                .Parameters.AddWithValue("@RatingID", objl_Parameters.CompareRatingID)
                .Parameters.AddWithValue("@TimeControlID", objl_Parameters.CompareTimeControlID)
                .Parameters.AddWithValue("@ColorID", 0)
                .Parameters.AddWithValue("@EvaluationGroupID", 0)
                .Parameters.AddWithValue("@MeasurementName", objl_Parameters.CompareScoreName)
                scoreAverage = Convert.ToDouble(objl_CMD.ExecuteScalar())
            End With

            Dim scoreVariance As Double = 0
            With objl_CMD
                .Parameters.Clear()
                .CommandText = modQueries.GetStatCovar()
                .Parameters.AddWithValue("@SourceID", 3)
                .Parameters.AddWithValue("AggregationName", "Event")
                .Parameters.AddWithValue("@RatingID", objl_Parameters.CompareRatingID)
                .Parameters.AddWithValue("@TimeControlID", objl_Parameters.CompareTimeControlID)
                .Parameters.AddWithValue("@ColorID", 0)
                .Parameters.AddWithValue("@EvaluationGroupID", 0)
                .Parameters.AddWithValue("@MeasurementName1", objl_Parameters.CompareScoreName)
                .Parameters.AddWithValue("@MeasurementName2", objl_Parameters.CompareScoreName)
                scoreVariance = Convert.ToDouble(objl_CMD.ExecuteScalar())
            End With

            Dim t1cplCovariance As Double = 0
            With objl_CMD
                .Parameters.Clear()
                .CommandText = modQueries.GetStatCovar()
                .Parameters.AddWithValue("@SourceID", 3)
                .Parameters.AddWithValue("AggregationName", "Event")
                .Parameters.AddWithValue("@RatingID", objl_Parameters.CompareRatingID)
                .Parameters.AddWithValue("@TimeControlID", objl_Parameters.CompareTimeControlID)
                .Parameters.AddWithValue("@ColorID", 0)
                .Parameters.AddWithValue("@EvaluationGroupID", 0)
                .Parameters.AddWithValue("@MeasurementName1", "T1")
                .Parameters.AddWithValue("@MeasurementName2", "ScACPL")
                t1cplCovariance = Convert.ToDouble(objl_CMD.ExecuteScalar())
            End With

            Dim t1scoreCovariance As Double = 0
            With objl_CMD
                .Parameters.Clear()
                .CommandText = modQueries.GetStatCovar()
                .Parameters.AddWithValue("@SourceID", 3)
                .Parameters.AddWithValue("AggregationName", "Event")
                .Parameters.AddWithValue("@RatingID", objl_Parameters.CompareRatingID)
                .Parameters.AddWithValue("@TimeControlID", objl_Parameters.CompareTimeControlID)
                .Parameters.AddWithValue("@ColorID", 0)
                .Parameters.AddWithValue("@EvaluationGroupID", 0)
                .Parameters.AddWithValue("@MeasurementName1", "T1")
                .Parameters.AddWithValue("@MeasurementName2", "Score")
                t1scoreCovariance = Convert.ToDouble(objl_CMD.ExecuteScalar())
            End With

            Dim cplscoreCovariance As Double = 0
            With objl_CMD
                .Parameters.Clear()
                .CommandText = modQueries.GetStatCovar()
                .Parameters.AddWithValue("@SourceID", 3)
                .Parameters.AddWithValue("AggregationName", "Event")
                .Parameters.AddWithValue("@RatingID", objl_Parameters.CompareRatingID)
                .Parameters.AddWithValue("@TimeControlID", objl_Parameters.CompareTimeControlID)
                .Parameters.AddWithValue("@ColorID", 0)
                .Parameters.AddWithValue("@EvaluationGroupID", 0)
                .Parameters.AddWithValue("@MeasurementName1", "ScACPL")
                .Parameters.AddWithValue("@MeasurementName2", "Score")
                cplscoreCovariance = Convert.ToDouble(objl_CMD.ExecuteScalar())
            End With

            Dim means As Double() = {t1Average, cplAverage, scoreAverage}

            Dim covarianceMatrix As Double(,) = {
                {t1Variance, t1cplCovariance, t1scoreCovariance},
                {t1cplCovariance, cplVariance, cplscoreCovariance},
                {t1scoreCovariance, cplscoreCovariance, scoreVariance}
            }

            Dim mahalanobis As Double = MahalanobisDistance(testStatistic, means, covarianceMatrix)
            Dim chiSquareDist As ChiSquared = New ChiSquared(3)
            PValue = chiSquareDist.CumulativeDistribution(Math.Pow(mahalanobis, 2))
        End Sub
    End Class
#End Region

#Region "Player Summary"
    Friend Sub PlayerSummary()
        Dim PlayerLength As Short = 30
        Dim EloLength As Short = 7
        Dim RecordLength As Short = 13
        Dim PerformanceLength As Short = 7
        Dim EvmLength As Short = 24
        Dim BlunderLength As Short = 24
        Dim AcplLength As Short = 11
        Dim SdcplLength As Short = 11
        Dim ScoreLength As Short = 10
        Dim RoiLength As Short = 9
        Dim PvalLength As Short = 11

        Dim tempText As String = ""
        tempText += "Player Name".PadRight(PlayerLength, " "c)
        tempText += "Elo".PadRight(EloLength, " "c)
        tempText += "Record".PadRight(RecordLength, " "c)
        tempText += "Perf".PadRight(PerformanceLength, " "c)
        tempText += "EVM / Turns = Pcnt".PadRight(EvmLength, " "c)
        tempText += "Blund / Turns = Pcnt".PadRight(BlunderLength, " "c)
        tempText += "ScACPL".PadRight(AcplLength, " "c)
        tempText += "ScSDCPL".PadRight(SdcplLength, " "c)
        tempText += "Score".PadRight(ScoreLength, " "c)
        tempText += "ROI".PadRight(RoiLength, " "c)
        tempText += "PValue".PadRight(PvalLength, " "c)
        tempText += "Opp EVM Pcnt".PadRight(EvmLength, " "c)
        tempText += "Opp Blund Pcnt".PadRight(BlunderLength, " "c)
        tempText += "OppScACPL".PadRight(AcplLength, " "c)
        tempText += "OppScSDCPL".PadRight(SdcplLength, " "c)
        tempText += "OppScore".PadRight(ScoreLength, " "c)
        tempText += "OppROI".PadRight(RoiLength, " "c)
        tempText += "OppPValue".PadRight(RoiLength, " "c)
        objm_Lines.Add(tempText)
        objm_Lines.Add(New String("-", 257))

        Dim objl_CMD As New SqlCommand
        Select Case params.ReportType
            Case "Event"
                With objl_CMD
                    .Connection = MainWindow.db_Connection
                    .CommandText = modQueries.EventPlayerSummary()
                    .Parameters.AddWithValue("@EventID", params.EventID)
                    .Parameters.AddWithValue("@ScoreID", params.CompareScoreID)
                End With
            Case "Player"
                With objl_CMD
                    .Connection = MainWindow.db_Connection
                    .CommandText = modQueries.PlayerPlayerSummary()
                    .Parameters.AddWithValue("@PlayerID", params.PlayerID)
                    .Parameters.AddWithValue("@StartDate", params.StartDate)
                    .Parameters.AddWithValue("@EndDate", params.EndDate)
                    .Parameters.AddWithValue("@ScoreID", params.CompareScoreID)
                End With
        End Select

        Dim objm_PlayerSummaries As New List(Of _playersummary)
        With objl_CMD.ExecuteReader
            While .Read
                Dim objl_Player As New _playersummary

                objl_Player.Name = .Item("Name")
                objl_Player.Rating = .Item("Rating")
                objl_Player.Record = .Item("Record")
                objl_Player.GamesPlayed = .Item("GamesPlayed")
                objl_Player.PerfRating = .Item("Perf")
                objl_Player.EVM = .Item("EVM")
                objl_Player.Blunders = .Item("Blunders")
                objl_Player.ScoredMoves = .Item("ScoredMoves")
                objl_Player.ACPL = .Item("ACPL")
                objl_Player.SDCPL = .Item("SDCPL")
                objl_Player.Score = .Item("Score")
                objl_Player.OppEVM = .Item("OppEVM")
                objl_Player.OppBlunders = .Item("OppBlunders")
                objl_Player.OppScoredMoves = .Item("OppScoredMoves")
                objl_Player.OppACPL = .Item("OppACPL")
                objl_Player.OppSDCPL = .Item("OppSDCPL")
                objl_Player.OppScore = .Item("OppScore")

                objm_PlayerSummaries.Add(objl_Player)
            End While
            .Close()
        End With

        'Due to the query itself, objm_PlayerSummaries will already be sorted alphabetically by name; if needed, could sort here

        For Each player As _playersummary In objm_PlayerSummaries
            tempText = ""  'reset this from prior use
            Dim tmp As String = ""

            tempText += Left(player.Name, PlayerLength).PadRight(PlayerLength, " "c)
            tempText += player.Rating.ToString().PadRight(EloLength, " "c)

            Dim tempLength As Short = player.GamesPlayed.ToString().Length + 2
            tmp = $"{player.Record.ToString().PadRight(tempLength, " "c)} / {player.GamesPlayed}"
            tempText += tmp.PadRight(RecordLength, " "c)

            If player.PerfRating > 0 Then
                tempText += $"+{player.PerfRating}".PadRight(PerformanceLength, " "c)
            Else
                tempText += player.PerfRating.ToString().PadRight(PerformanceLength, " "c)
            End If

            tmp = player.EVM.ToString().PadRight(4, " "c) & " / " & player.ScoredMoves.ToString().PadRight(4, " "c) & $" = {Convert.ToDouble(100 * player.EVM / player.ScoredMoves):0.00}%"
            tempText += tmp.PadRight(EvmLength, " "c)

            tmp = player.Blunders.ToString().PadRight(4, " "c) & " / " & player.ScoredMoves.ToString().PadRight(4, " "c) & $" = {Convert.ToDouble(100 * player.Blunders / player.ScoredMoves):0.00}%"
            tempText += tmp.PadRight(BlunderLength, " "c)

            tempText += player.ACPL.ToString("0.0000").PadRight(AcplLength, " "c)
            tempText += player.SDCPL.ToString("0.0000").PadRight(SdcplLength, " "c)
            tempText += player.Score.ToString("0.00").PadRight(ScoreLength, " "c)

            With objl_CMD
                .Parameters.Clear()
                .CommandText = modQueries.ZScoreData()
                .Parameters.AddWithValue("AggregationName", "Event")  'since this value is inclusive of multiple games, it should always be compared against Event
                .Parameters.AddWithValue("@ScoreName", params.CompareScoreName)
                .Parameters.AddWithValue("@SourceID", 3)  'hard-coding sourceID since Lichess doesn't have event stats
                .Parameters.AddWithValue("@TimeControlID", params.CompareTimeControlID)
                .Parameters.AddWithValue("@RatingID", Math.Floor(player.Rating / 100) * 100)
            End With

            Dim objm_z As New Dictionary(Of String, Double)
            Dim objm_zopp As New Dictionary(Of String, Double)
            With objl_CMD.ExecuteReader
                While .Read
                    Dim z_score As New Double
                    Dim z_score_opp As New Double
                    Select Case .Item("MeasurementName")
                        Case "T1"
                            z_score = ((Convert.ToDouble(player.EVM) / player.ScoredMoves) - Convert.ToDouble(.Item("Average"))) / Convert.ToDouble(.Item("StandardDeviation"))
                            objm_z.Add("T1", z_score)

                            z_score_opp = ((Convert.ToDouble(player.OppEVM) / player.OppScoredMoves) - Convert.ToDouble(.Item("Average"))) / Convert.ToDouble(.Item("StandardDeviation"))
                            objm_zopp.Add("T1", z_score_opp)
                        Case "ScACPL"
                            z_score = -1 * (player.ACPL - Convert.ToDouble(.Item("Average"))) / Convert.ToDouble(.Item("StandardDeviation"))
                            objm_z.Add("ScACPL", z_score)

                            z_score_opp = -1 * (player.OppACPL - Convert.ToDouble(.Item("Average"))) / Convert.ToDouble(.Item("StandardDeviation"))
                            objm_zopp.Add("ScACPL", z_score_opp)
                        Case Else
                            'all possible score measurement names
                            z_score = (player.Score - Convert.ToDouble(.Item("Average"))) / Convert.ToDouble(.Item("StandardDeviation"))
                            objm_z.Add("Score", z_score)

                            z_score_opp = (player.OppScore - Convert.ToDouble(.Item("Average"))) / Convert.ToDouble(.Item("StandardDeviation"))
                            objm_zopp.Add("Score", z_score_opp)
                    End Select
                    tmp = CalculateROI(objm_z).ToString("0.0")
                End While
                .Close()
            End With
            tempText += tmp.PadRight(RoiLength, " "c)

            tempText += "".PadRight(PvalLength, " "c)  'TODO: PValue values

            tmp = player.OppEVM.ToString().PadRight(4, " "c) & " / " & player.OppScoredMoves.ToString().PadRight(4, " "c) & $" = {Convert.ToDouble(100 * player.OppEVM / player.OppScoredMoves):0.00}%"
            tempText += tmp.PadRight(EvmLength, " "c)

            tmp = player.OppBlunders.ToString().PadRight(4, " "c) & " / " & player.OppScoredMoves.ToString().PadRight(4, " "c) & $" = {Convert.ToDouble(100 * player.OppBlunders / player.OppScoredMoves):0.00}%"
            tempText += tmp.PadRight(BlunderLength, " "c)

            tempText += player.OppACPL.ToString("0.0000").PadRight(AcplLength, " "c)
            tempText += player.OppSDCPL.ToString("0.0000").PadRight(SdcplLength, " "c)
            tempText += player.OppScore.ToString("0.00").PadRight(ScoreLength, " "c)

            tmp = CalculateROI(objm_zopp).ToString("0.0")
            tempText += tmp.PadRight(RoiLength, " "c)

            tempText += "".PadRight(PvalLength, " "c)  'TODO: OppPValue values

            objm_Lines.Add(tempText)
        Next

        objm_Lines.Add("")
        objm_Lines.Add("")
    End Sub

    Private Class _playersummary
        Friend Property Name As String
        Friend Property Rating As Short
        Friend Property Record As Double
        Friend Property GamesPlayed As Short
        Friend Property PerfRating As Short
        Friend Property EVM As Short
        Friend Property Blunders As Short
        Friend Property ScoredMoves As Short
        Friend Property ACPL As Double
        Friend Property SDCPL As Double
        Friend Property Score As Double
        Friend Property OppEVM As Short
        Friend Property OppBlunders As Short
        Friend Property OppScoredMoves As Short
        Friend Property OppACPL As Double
        Friend Property OppSDCPL As Double
        Friend Property OppScore As Double
    End Class
#End Region

#Region "Game Traces"
    Friend Sub GameTraces()

    End Sub

    Private Class _gametraces
        'do I want this to be an individual game or all games from a player? Do I need one class for the entire set of games, and a separate class for a single game?
    End Class
#End Region
End Class
