Imports Microsoft.Data.SqlClient

Public Class clsDetail
    Private Shared params As clsParameters = MainWindow.objl_Parameters

#Region "Key Stats"
    Friend Sub KeyStats()
        Const EventLength As Short = 35

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

            Dim objm_ROI As New clsAdvancedStats.ROI
            With objm_ROI
                .AggregationName = "Event"  'since this value is inclusive of multiple games, it should always be compared against Event
                .ScoreName = objl_Parameters.CompareScoreName
                .SourceName = "Control"  'hard-coding sourceID since Lichess doesn't have event stats
                .TimeControlName = objl_Parameters.CompareTimeControl
                .RatingID = objl_Parameters.CompareRatingID
                .T1_Pcnt = Convert.ToDouble(TCounts("T1")) / ScoredMoves
                .ACPL = ACPL
                .Score = Score
            End With
            ROI = objm_ROI.GetROI()

            Dim objm_PValue As New clsAdvancedStats.PValue
            With objm_PValue
                .T1_Pcnt = Convert.ToDouble(TCounts("T1")) / ScoredMoves
                .ACPL = ACPL
                .Score = Score
                .SourceName = "Control"  'hard-coding since Lichess doesn't have event stats
                .AggregationName = "Event"  'since this value is inclusive of multiple games, it should always be compared against Event
                .RatingID = objl_Parameters.CompareRatingID
                .TimeControlName = objl_Parameters.CompareTimeControl
                .Color = ""
                .ScoreName = objl_Parameters.CompareScoreName
            End With
            PValue = objm_PValue.GetPValue()
        End Sub
    End Class
#End Region

#Region "Player Summary"
    Friend Sub PlayerSummary()
        Const PlayerLength As Short = 30
        Const EloLength As Short = 7
        Const RecordLength As Short = 13
        Const PerformanceLength As Short = 7
        Const EvmLength As Short = 24
        Const BlunderLength As Short = 24
        Const AcplLength As Short = 11
        Const SdcplLength As Short = 11
        Const ScoreLength As Short = 10
        Const RoiLength As Short = 9
        Const PvalLength As Short = 11

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

            tmp = player.EVM.ToString().PadRight(4, " "c) & " / " & player.ScoredMoves.ToString().PadRight(4, " "c) & $" = {Convert.ToDouble(100 * player.EVM / player.ScoredMoves):0.0}%"
            tempText += tmp.PadRight(EvmLength, " "c)

            tmp = player.Blunders.ToString().PadRight(4, " "c) & " / " & player.ScoredMoves.ToString().PadRight(4, " "c) & $" = {Convert.ToDouble(100 * player.Blunders / player.ScoredMoves):0.00}%"
            tempText += tmp.PadRight(BlunderLength, " "c)

            tempText += player.ACPL.ToString("0.0000").PadRight(AcplLength, " "c)
            tempText += player.SDCPL.ToString("0.0000").PadRight(SdcplLength, " "c)
            tempText += player.Score.ToString("0.00").PadRight(ScoreLength, " "c)

            Dim objm_ROI As New clsAdvancedStats.ROI
            With objm_ROI
                .AggregationName = "Event"  'since this value is inclusive of multiple games, it should always be compared against Event
                .ScoreName = params.CompareScoreName
                .SourceName = "Control"  'hard-coding since Lichess doesn't have event stats
                .TimeControlName = params.CompareTimeControl
                .RatingID = params.CompareRatingID
                .T1_Pcnt = Convert.ToDouble(player.EVM) / player.ScoredMoves
                .ACPL = player.ACPL
                .Score = player.Score
            End With
            tmp = objm_ROI.GetROI().ToString("0.0")
            tempText += tmp.PadRight(RoiLength, " "c)

            Dim objm_PValue As New clsAdvancedStats.PValue
            With objm_PValue
                .T1_Pcnt = Convert.ToDouble(player.EVM) / player.ScoredMoves
                .ACPL = player.ACPL
                .Score = player.Score
                .SourceName = "Control"  'hard-coding since Lichess doesn't have event stats
                .AggregationName = "Event"  'since this value is inclusive of multiple games, it should always be compared against Event
                .RatingID = params.CompareRatingID
                .TimeControlName = params.CompareTimeControl
                .Color = ""
                .ScoreName = params.CompareScoreName
            End With
            Dim PValue As Double = objm_PValue.GetPValue()
            Dim strPValue As String = (100 * PValue).ToString("0.00") & "%"
            tempText += strPValue.PadRight(PvalLength, " "c)

            tmp = player.OppEVM.ToString().PadRight(4, " "c) & " / " & player.OppScoredMoves.ToString().PadRight(4, " "c) & $" = {Convert.ToDouble(100 * player.OppEVM / player.OppScoredMoves):0.0}%"
            tempText += tmp.PadRight(EvmLength, " "c)

            tmp = player.OppBlunders.ToString().PadRight(4, " "c) & " / " & player.OppScoredMoves.ToString().PadRight(4, " "c) & $" = {Convert.ToDouble(100 * player.OppBlunders / player.OppScoredMoves):0.00}%"
            tempText += tmp.PadRight(BlunderLength, " "c)

            tempText += player.OppACPL.ToString("0.0000").PadRight(AcplLength, " "c)
            tempText += player.OppSDCPL.ToString("0.0000").PadRight(SdcplLength, " "c)
            tempText += player.OppScore.ToString("0.00").PadRight(ScoreLength, " "c)

            objm_ROI = New clsAdvancedStats.ROI
            With objm_ROI
                .AggregationName = "Event"  'since this value is inclusive of multiple games, it should always be compared against Event
                .ScoreName = params.CompareScoreName
                .SourceName = "Control"  'hard-coding sourceID since Lichess doesn't have event stats
                .TimeControlName = params.CompareTimeControl
                .RatingID = params.CompareRatingID
                .T1_Pcnt = Convert.ToDouble(player.OppEVM) / player.OppScoredMoves
                .ACPL = player.OppACPL
                .Score = player.OppScore
            End With
            tmp = objm_ROI.GetROI().ToString("0.0")
            tempText += tmp.PadRight(RoiLength, " "c)

            objm_PValue = New clsAdvancedStats.PValue
            With objm_PValue
                .T1_Pcnt = Convert.ToDouble(player.OppEVM) / player.OppScoredMoves
                .ACPL = player.OppACPL
                .Score = player.OppScore
                .SourceName = "Control"  'hard-coding since Lichess doesn't have event stats
                .AggregationName = "Event"  'since this value is inclusive of multiple games, it should always be compared against Event
                .RatingID = params.CompareRatingID
                .TimeControlName = params.CompareTimeControl
                .Color = ""
                .ScoreName = params.CompareScoreName
            End With
            PValue = objm_PValue.GetPValue()
            strPValue = (100 * PValue).ToString("0.00") & "%"
            tempText += strPValue.PadRight(PvalLength, " "c)

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
        objm_Lines.Add(New String("-", 31))

        Dim objl_CMD As New SqlCommand
        Select Case params.ReportType
            Case "Event"
                With objl_CMD
                    .Connection = MainWindow.db_Connection
                    .CommandText = modQueries.EventPlayers()
                    .Parameters.AddWithValue("@EventID", params.EventID)
                End With
            Case "Player"
                With objl_CMD
                    .Connection = MainWindow.db_Connection
                    .CommandText = modQueries.PlayerPlayers()
                    .Parameters.AddWithValue("@PlayerID", params.PlayerID)
                    .Parameters.AddWithValue("@StartDate", params.StartDate)
                    .Parameters.AddWithValue("@EndDate", params.EndDate)
                End With
        End Select

        Dim objm_Players As New List(Of _player)
        With objl_CMD.ExecuteReader
            While .Read
                Dim objl_Player As New _player
                objl_Player.PlayerID = Convert.ToInt64(.Item("PlayerID"))
                objl_Player.Name = .Item("Name")
                objl_Player.Rating = Convert.ToInt16(.Item("Rating"))
                objl_Player.ScoredMoves = Convert.ToInt16(.Item("ScoredMoves"))
                objl_Player.BuildGames()

                objm_Players.Add(objl_Player)
            End While
            .Close()
        End With

        For Each player As _player In objm_Players
            objm_Lines.Add($"{player.Name} {player.Rating} (Moves={player.ScoredMoves})")

            For Each game As _game In player.Games
                Dim tempText As String = ""
                tempText += Right($" {game.Round}", 2)
                tempText += $"{game.ReportColor} "
                tempText += $"{game.Result} "
                tempText += game.OppName.PadRight(25, " "c)
                tempText += game.OppElo.ToString().PadRight(4, " "c) & ":  "

                Dim tmp2 As String = ""
                tmp2 += game.EVM.ToString().PadRight(3, " "c) & " / " & game.ScoredMoves.ToString().PadRight(3, " "c) & $" = {Convert.ToDouble(100 * game.EVM / game.ScoredMoves):0}%"
                tempText += tmp2.PadRight(18, " "c)

                tempText += $"{game.ACPL:0.0000}".PadRight(8, " "c)
                tempText += $"{game.SDCPL:0.0000}".PadRight(8, " "c)
                tempText += $"{game.Score:0.00}".PadRight(7, " "c)
                tempText += $"{game.ROI:0.0}".PadRight(6, " "c)
                tempText += $"{(100 * game.PValue):0.00}%".PadRight(8, " "c)

                For i As Short = 0 To game.Trace.Length - 1
                    If i > 0 Then
                        If i Mod 60 = 0 Then
                            objm_Lines.Add(tempText)
                            tempText = New String(" ", 93)
                        Else
                            If i Mod 10 = 0 Then tempText += " "
                        End If
                    End If
                    tempText += game.Trace(i)
                Next

                objm_Lines.Add(tempText)
            Next

            objm_Lines.Add(New String("-", 25))
        Next
    End Sub

    Private Class _player
        Friend Property PlayerID As Long
        Friend Property Name As String
        Friend Property Rating As Short
        Friend Property ScoredMoves As Short

        Friend Property Games As New List(Of _game)

        Friend Sub BuildGames()
            Dim objl_CMD As New SqlCommand
            Select Case params.ReportType
                Case "Event"
                    With objl_CMD
                        .Connection = MainWindow.db_Connection
                        .CommandText = modQueries.EventPlayerOpponents()
                        .Parameters.AddWithValue("@PlayerID", PlayerID)
                        .Parameters.AddWithValue("@EventID", params.EventID)
                        .Parameters.AddWithValue("@ScoreID", params.CompareScoreID)
                    End With
                Case "Player"
                    With objl_CMD
                        .Connection = MainWindow.db_Connection
                        .CommandText = modQueries.PlayerPlayerOpponents()
                        .Parameters.AddWithValue("@PlayerID", PlayerID)
                        .Parameters.AddWithValue("@StartDate", params.StartDate)
                        .Parameters.AddWithValue("@EndDate", params.EndDate)
                        .Parameters.AddWithValue("@ScoreID", params.CompareScoreID)
                    End With
            End Select

            With objl_CMD.ExecuteReader
                While .Read
                    Dim objl_Game As New _game
                    objl_Game.GameID = Convert.ToInt64(.Item("GameID"))
                    objl_Game.Round = Convert.ToInt16(.Item("RoundNum"))
                    objl_Game.Color = .Item("Color")
                    objl_Game.ReportColor = objl_Game.Color.Substring(0, 1).ToLower()
                    objl_Game.Result = .Item("Result")
                    objl_Game.OppName = .Item("OppName")
                    objl_Game.OppElo = Convert.ToInt16(.Item("OppRating"))
                    objl_Game.EVM = Convert.ToInt16(.Item("EVM"))
                    objl_Game.ScoredMoves = Convert.ToInt16(.Item("ScoredMoves"))
                    objl_Game.ACPL = Convert.ToDouble(.Item("ACPL"))
                    objl_Game.SDCPL = Convert.ToDouble(.Item("SDCPL"))
                    objl_Game.Score = Convert.ToDouble(.Item("Score"))
                    objl_Game.PopulateTrace()
                    objl_Game.PopulateAdvancedStats()

                    Games.Add(objl_Game)
                End While
                .Close()
            End With
        End Sub
    End Class

    Private Class _game
        Friend Property GameID As Long
        Friend Property Round As Short
        Friend Property Color As String
        Friend Property Result As Char
        Friend Property OppName As String
        Friend Property OppElo As Short
        Friend Property EVM As Short
        Friend Property ScoredMoves As Short
        Friend Property ACPL As Double
        Friend Property SDCPL As Double
        Friend Property Score As Double

        Friend Property ReportColor As Char
        Friend Property Trace As String

        Friend Property ROI As Double
        Friend Property PValue As Double

        Friend Sub PopulateTrace()
            Dim objl_CMD As New SqlCommand
            With objl_CMD
                .Connection = MainWindow.db_Connection
                .CommandText = modQueries.GameTrace()
                .Parameters.AddWithValue("@GameID", GameID)
                .Parameters.AddWithValue("@Color", Color)
            End With

            With objl_CMD.ExecuteReader
                While .Read
                    Trace += .Item("MoveTrace")
                End While
                .Close()
            End With
        End Sub

        Friend Sub PopulateAdvancedStats()
            Dim objm_ROI As New clsAdvancedStats.ROI
            With objm_ROI
                .AggregationName = "Game"
                .ScoreName = params.CompareScoreName
                .SourceName = params.CompareSourceName
                .TimeControlName = params.CompareTimeControl
                .RatingID = params.CompareRatingID
                .T1_Pcnt = Convert.ToDouble(EVM) / ScoredMoves
                .ACPL = ACPL
                .Score = Score
            End With
            ROI = objm_ROI.GetROI(Color)

            Dim objm_PValue As New clsAdvancedStats.PValue
            With objm_PValue
                .T1_Pcnt = Convert.ToDouble(EVM) / ScoredMoves
                .ACPL = ACPL
                .Score = Score
                .SourceName = params.CompareSourceName
                .AggregationName = "Game"
                .RatingID = params.CompareRatingID
                .TimeControlName = params.CompareTimeControl
                .Color = Color
                .ScoreName = params.CompareScoreName
            End With
            PValue = objm_PValue.GetPValue()
        End Sub
    End Class
#End Region
End Class
