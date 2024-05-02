Imports Microsoft.Data.SqlClient

Public Class clsDetail
    Private ReadOnly params As clsParameters = MainWindow.objl_Parameters
    Private Const EventLength As Short = 35

    Friend Sub KeyStats()
        'should this be turned into a sub-class or something?
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
        '.ReportType = params.ReportType,
        '.EventID = params.EventID,
        '.PlayerID = params.PlayerID,
        '.StartDate = params.StartDate,
        '.EndDate = params.EndDate,
        '.ScoreID = params.CompareScoreID
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
        objm_Lines.Add("Score:".PadRight(EventLength, " "c) & stats.Score.ToString("0.00"))  'TODO: Add the asterisk when needed
        objm_Lines.Add("ROI:".PadRight(EventLength, " "c) & stats.ROI.ToString("0.0"))
        objm_Lines.Add("PValue:".PadRight(EventLength, " "c) & $"{stats.PValue:0.00}%")
        objm_Lines.Add("")
        objm_Lines.Add("")
    End Sub

    Private Class _keystats
        Friend Property objl_Parameters As clsParameters
        'Friend Property ReportType As String
        'Friend Property EventID As Long
        'Friend Property PlayerID As Long
        'Friend Property StartDate As Date
        'Friend Property EndDate As Date
        'Friend Property ScoreID As Short

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
                        .Parameters.AddWithValue("@EventID", EventID)
                    End With
                Case "Player"
                    With objl_CMD
                        .Connection = MainWindow.db_Connection
                        .CommandText = modQueries.PlayerRatings()
                        .Parameters.AddWithValue("@PlayerID", PlayerID)
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
            ROI = 0
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
                            z_score = 0  'TODO
                            objm_z.Add("T1", z_score)
                        Case "ScACPL"
                            z_score = 0  'TODO
                            objm_z.Add("ScACPL", z_score)
                        Case Else
                            'all possible score measurement names
                            z_score = 0  'TODO
                            objm_z.Add("Score", z_score)
                    End Select
                    ROI = .Item("Score")
                End While
                .Close()
            End With

            'PValue
            PValue = 0
        End Sub
    End Class
End Class
