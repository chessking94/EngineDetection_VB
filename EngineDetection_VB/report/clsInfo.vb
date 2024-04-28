Imports Microsoft.Data.SqlClient

Public Class clsInfo
    Private ReadOnly params As clsParameters = MainWindow.objl_Parameters
    Friend ReadOnly Engine As String
    Friend ReadOnly Depth As Short
    Friend ReadOnly MaxEval As Double
    Private Const HeaderLength As Short = 30
    Private Const PlayerKeyLength As Short = 10

#Region "Initialize"
    Public Sub New()
        Engine = GetEngineName(params)
        Depth = GetDepth(params)
        MaxEval = GetMaxEval()
    End Sub

    Private Function GetEngineName(pi_params As clsParameters) As String
        Dim rtnval As String

        If params.EventID > 0 Then
            Using objl_CMD As New SqlCommand(modQueries.EventEngine(), MainWindow.db_Connection)
                objl_CMD.Parameters.AddWithValue("@EventID", pi_params.EventID)
                rtnval = Convert.ToString(objl_CMD.ExecuteScalar())
            End Using
        Else
            Using objl_CMD As New SqlCommand(modQueries.PlayerEngine(), MainWindow.db_Connection)
                objl_CMD.Parameters.AddWithValue("@PlayerID", pi_params.PlayerID)
                'TODO: Add date parameters
                rtnval = Convert.ToString(objl_CMD.ExecuteScalar())
            End Using
        End If

        Return rtnval
    End Function

    Private Function GetDepth(pi_params As clsParameters) As String
        Dim rtnval As Short

        If params.EventID > 0 Then
            Using objl_CMD As New SqlCommand(modQueries.EventDepth(), MainWindow.db_Connection)
                objl_CMD.Parameters.AddWithValue("@EventID", pi_params.EventID)
                rtnval = Convert.ToInt16(objl_CMD.ExecuteScalar())
            End Using
        Else
            Using objl_CMD As New SqlCommand(modQueries.PlayerDepth(), MainWindow.db_Connection)
                objl_CMD.Parameters.AddWithValue("@PlayerID", pi_params.PlayerID)
                'TODO: Add date parameters
                rtnval = Convert.ToInt16(objl_CMD.ExecuteScalar())
            End Using
        End If

        Return rtnval
    End Function

    Private Function GetMaxEval() As String
        Dim rtnval As Double
        Using objl_CMD As New SqlCommand(modQueries.MaxEval(), MainWindow.db_Connection)
            rtnval = Convert.ToDouble(objl_CMD.ExecuteScalar())
        End Using

        Return rtnval
    End Function
#End Region

#Region "Header"
    Friend Sub Header()
        ComparisonVariables()

        Select Case params.ReportType
            Case "Event"
                EventDetails()
            Case "Player"
                PlayerDetails()
        End Select

        EngineDetails()
    End Sub

    Private Sub ComparisonVariables()
        objm_Lines.Add(New String("-", 100))
        objm_Lines.Add("Analysis Type:".PadRight(HeaderLength, " "c) & params.ReportType)
        objm_Lines.Add("Compared Source:".PadRight(HeaderLength, " "c) & params.CompareSourceName)  'TODO: how do Compare* variables populate if no comparison data is chosen?
        objm_Lines.Add("Compared Time Control:".PadRight(HeaderLength, " "c) & params.CompareTimeControl)
        objm_Lines.Add("Compared Rating:".PadRight(HeaderLength, " "c) & params.CompareRatingID)
        objm_Lines.Add("Scoring Method Used:".PadRight(HeaderLength, " "c) & params.CompareScoreName)
        objm_Lines.Add("")
    End Sub

    Private Sub EventDetails()
        Dim eventDates As String = ""
        Dim eventRounds As Short = 0
        Dim eventPlayers As Short = 0

        Using objl_CMD As New SqlCommand(modQueries.EventSummary(), MainWindow.db_Connection)
            objl_CMD.Parameters.AddWithValue("@EventID", params.EventID)
            With objl_CMD.ExecuteReader
                While .Read
                    eventDates = .Item("EventDates")
                    eventRounds = .Item("Rounds")
                    eventPlayers = .Item("Players")
                End While
            End With
        End Using

        objm_Lines.Add("Event Name:".PadRight(HeaderLength, " "c) & params.EventName)
        objm_Lines.Add("Event Date:".PadRight(HeaderLength, " "c) & eventDates)
        objm_Lines.Add("Rounds:".PadRight(HeaderLength, " "c) & eventRounds.ToString())
        objm_Lines.Add("Players:".PadRight(HeaderLength, " "c) & eventPlayers.ToString())
        objm_Lines.Add("")
    End Sub

    Private Sub PlayerDetails()
        objm_Lines.Add("Player Name:".PadRight(HeaderLength, " "c) & $"{params.FirstName} {params.LastName}")
        objm_Lines.Add("Games Between:".PadRight(HeaderLength, " "c) & "PENDING DEVELOPMENT")  'TODO: Integrate date range for ReportType = Player
        objm_Lines.Add("")
    End Sub

    Private Sub EngineDetails()
        objm_Lines.Add("Engine Name:".PadRight(HeaderLength, " "c) & Engine)
        objm_Lines.Add("Depth:".PadRight(HeaderLength, " "c) & Depth)
        objm_Lines.Add("Report Date:".PadRight(HeaderLength, " "c) & Date.Today.ToString("MM/dd/yyyy"))
        objm_Lines.Add("")
        objm_Lines.Add("")
    End Sub
#End Region

#Region "Keys"
    Friend Sub ScoringKey()
        objm_Lines.Add("MOVE SCORING")
        objm_Lines.Add(New String("-", 25))
        objm_Lines.Add("A move is Scored if it does not meet any of the following:")
        objm_Lines.Add(Space(4) & "Is theoretical opening move")
        objm_Lines.Add(Space(4) & "Is a tablebase hit")
        objm_Lines.Add(Space(4) & $"The best engine evaluation is greater than {MaxEval} centipawns or a mate in N")
        objm_Lines.Add(Space(4) & "The engine evaluation of the move played is a mate in N")
        objm_Lines.Add(Space(4) & "Only one legal move exists or the difference in evaluation between the top 2 engine moves is greater than 200 centipawns")
        objm_Lines.Add(Space(4) & "Is the second or third occurance of the position")
        objm_Lines.Add("")
        objm_Lines.Add("")
    End Sub

    Friend Sub PlayerKey()
        objm_Lines.Add("PLAYER KEY")
        objm_Lines.Add(New String("-", 100))
        objm_Lines.Add("EVM:".PadRight(PlayerKeyLength, " "c) & "Equal Value Match; moves with an evaluation that matches the best engine evaluation")
        objm_Lines.Add("Blund:".PadRight(PlayerKeyLength, " "c) & "Blunders; moves that lost 200 centipawns or more")
        objm_Lines.Add("ScACPL:".PadRight(PlayerKeyLength, " "c) & "Scaled Average Centipawn Loss; sum of total centipawn loss divided by the number of moves, scaled by position evaluation")
        objm_Lines.Add("ScSDCPL:".PadRight(PlayerKeyLength, " "c) & "Scaled Standard Deviation Centipawn Loss; standard deviation of centipawn loss values from each move played, scaled by position evaluation")
        objm_Lines.Add("Score:".PadRight(PlayerKeyLength, " "c) & "Game Score; measurement of how accurately the game was played, ranges from 0 to 100")
        objm_Lines.Add("ROI:".PadRight(PlayerKeyLength, " "c) & "Raw Outlier Index; standardized value where 50 represents the mean for that rating level and each increment of 5 is one standard deviation")
        objm_Lines.Add("PValue:".PadRight(PlayerKeyLength, " "c) & "Chi-square statistic associated with the Mahalanobis distance of the test point (T1, ScACPL, Score)")
        objm_Lines.Add(Space(PlayerKeyLength) & "An asterisk (*) following any statistic indicates an outlier that should be reviewed more closely")
        objm_Lines.Add("")
    End Sub

    Friend Sub GameKey()
        objm_Lines.Add("GAME KEY")
        objm_Lines.Add(New String("-", 100))
        objm_Lines.Add("(Player Name) (Elo) (Scored Moves)")
        objm_Lines.Add(Space(1) & "(Round)(Color) (Result) (Opp) (Opp Rating): (EVM/Turns = EVM%) (ScACPL) (ScSDCPL) (Score) (ROI) (PValue) (game trace)")
        objm_Lines.Add("Game trace key: b = Book move; M = EV match; 0 = Inferior move; e = Eliminated because one side far ahead; t = Tablebase hit; f = Forced move; r = Repeated move")
        objm_Lines.Add("")
    End Sub
#End Region
End Class
