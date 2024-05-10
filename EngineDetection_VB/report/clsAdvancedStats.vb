Imports MathNet.Numerics.Distributions
Imports Microsoft.Data.SqlClient

Friend Class clsAdvancedStats
    Friend Class ROI
        Friend Property AggregationName As String
        Friend Property ScoreName As String
        Friend Property SourceName As String
        Friend Property TimeControlName As String
        Friend Property RatingID As Short

        Friend Property T1_Pcnt As Double
        Friend Property ACPL As Double
        Friend Property Score As Double

        Private ReadOnly Weights As New Dictionary(Of String, Double)

        Public Sub New()
            PopulateWeights()
        End Sub

        Private Sub PopulateWeights()
            Weights.Add("T1", 0.2)
            Weights.Add("ScACPL", 0.35)
            Weights.Add("Score", 0.45)
        End Sub

        Friend Function GetROI(Optional Color As String = "") As Double
            Dim objl_CMD As New SqlCommand
            With objl_CMD
                .Connection = MainWindow.db_Connection
                If Color = "" Then
                    .CommandText = modQueries.ZScoreData()
                Else
                    .CommandText = modQueries.ZScoreData(Color)
                End If
                .Parameters.AddWithValue("AggregationName", AggregationName)
                .Parameters.AddWithValue("@ScoreName", ScoreName)
                .Parameters.AddWithValue("@SourceName", SourceName)
                .Parameters.AddWithValue("@TimeControlName", TimeControlName)
                .Parameters.AddWithValue("@RatingID", RatingID)
            End With

            Dim objm_z As New Dictionary(Of String, Double)
            With objl_CMD.ExecuteReader
                While .Read
                    Dim z_score As New Double
                    Select Case .Item("MeasurementName")
                        Case "T1"
                            z_score = (T1_Pcnt - Convert.ToDouble(.Item("Average"))) / Convert.ToDouble(.Item("StandardDeviation"))
                            objm_z.Add("T1", z_score)
                        Case "ScACPL"
                            z_score = -1 * (ACPL - Convert.ToDouble(.Item("Average"))) / Convert.ToDouble(.Item("StandardDeviation"))
                            objm_z.Add("ScACPL", z_score)
                        Case Else
                            'all possible score measurement names
                            z_score = (Score - Convert.ToDouble(.Item("Average"))) / Convert.ToDouble(.Item("StandardDeviation"))
                            objm_z.Add("Score", z_score)
                    End Select
                End While
                .Close()
            End With

            Dim dotproduct As Double = 0
            For Each kvp As KeyValuePair(Of String, Double) In objm_z
                dotproduct += kvp.Value * Weights(kvp.Key)
            Next

            Dim sumsquaresroot As Double = 0
            For Each kvp As KeyValuePair(Of String, Double) In Weights
                sumsquaresroot += Math.Pow(kvp.Value, 2)
            Next
            sumsquaresroot = Math.Sqrt(sumsquaresroot)

            Dim z As Double = dotproduct / sumsquaresroot
            Dim roi As Double = 5 * z + 50

            Return roi
        End Function
    End Class

    Friend Class PValue
        'for the comparison statistic
        Friend Property T1_Pcnt As Double
        Friend Property ACPL As Double
        Friend Property Score As Double

        Friend Property SourceName As String
        Friend Property AggregationName As String
        Friend Property RatingID As Short
        Friend Property TimeControlName As String
        Friend Property Color As String
        Friend Property EvaluationGroupID As Short = 0  'this isn't currently used anywhere except to further restrict results returned in the query
        Friend Property ScoreName As String

        Friend Function GetPValue() As Double
            If Color = "" Then Color = "N/A"  'since the name of the unspecified color in the database is this
            Dim testStatistic As Double() = {T1_Pcnt, ACPL, Score}
            Dim meanVector As Double() = GetMeanVector()
            Dim covarianceMatrix As Double(,) = GetCovarianceMatrix()
            Dim mahalanobis As Double = MahalanobisDistance(testStatistic, meanVector, covarianceMatrix)
            Dim chiSquareDist As ChiSquared = New ChiSquared(3)
            Dim PValue As Double = 1 - chiSquareDist.CumulativeDistribution(Math.Pow(mahalanobis, 2))
            Return PValue
        End Function

        Private Function GetMeanVector() As Double()
            Dim t1Average As Double = 0
            Dim cplAverage As Double = 0
            Dim scoreAverage As Double = 0

            Dim objl_CMD As New SqlCommand
            With objl_CMD
                .Connection = MainWindow.db_Connection
                .CommandText = modQueries.GetStatAverage()
                .Parameters.AddWithValue("@SourceName", SourceName)
                .Parameters.AddWithValue("AggregationName", AggregationName)
                .Parameters.AddWithValue("@RatingID", RatingID)
                .Parameters.AddWithValue("@TimeControlName", TimeControlName)
                .Parameters.AddWithValue("@Color", Color)
                .Parameters.AddWithValue("@EvaluationGroupID", EvaluationGroupID)
                .Parameters.AddWithValue("@MeasurementName", "T1")
                t1Average = Convert.ToDouble(objl_CMD.ExecuteScalar())
            End With

            With objl_CMD
                .Parameters("@MeasurementName").Value = "ScACPL"
                cplAverage = Convert.ToDouble(objl_CMD.ExecuteScalar())
            End With

            With objl_CMD
                .Parameters("@MeasurementName").Value = ScoreName
                scoreAverage = Convert.ToDouble(objl_CMD.ExecuteScalar())
            End With

            Dim meanVector As Double() = {t1Average, cplAverage, scoreAverage}
            Return meanVector
        End Function

        Private Function GetCovarianceMatrix() As Double(,)
            Dim t1Variance As Double = 0
            Dim cplVariance As Double = 0
            Dim scoreVariance As Double = 0

            Dim t1cplCovariance As Double = 0
            Dim t1scoreCovariance As Double = 0
            Dim cplscoreCovariance As Double = 0

            Dim objl_CMD As New SqlCommand
            With objl_CMD
                .Connection = MainWindow.db_Connection
                .CommandText = modQueries.GetStatCovar()
                .Parameters.AddWithValue("@SourceName", SourceName)
                .Parameters.AddWithValue("AggregationName", AggregationName)
                .Parameters.AddWithValue("@RatingID", RatingID)
                .Parameters.AddWithValue("@TimeControlName", TimeControlName)
                .Parameters.AddWithValue("@Color", Color)
                .Parameters.AddWithValue("@EvaluationGroupID", EvaluationGroupID)
                .Parameters.AddWithValue("@MeasurementName1", "T1")
                .Parameters.AddWithValue("@MeasurementName2", "T1")
                t1Variance = Convert.ToDouble(objl_CMD.ExecuteScalar())
            End With

            With objl_CMD
                .Parameters("@MeasurementName1").Value = "ScACPL"
                .Parameters("@MeasurementName2").Value = "ScACPL"
                cplVariance = Convert.ToDouble(objl_CMD.ExecuteScalar())
            End With

            With objl_CMD
                .Parameters("@MeasurementName1").Value = ScoreName
                .Parameters("@MeasurementName2").Value = ScoreName
                scoreVariance = Convert.ToDouble(objl_CMD.ExecuteScalar())
            End With

            With objl_CMD
                .Parameters("@MeasurementName1").Value = "T1"
                .Parameters("@MeasurementName2").Value = "ScACPL"
                t1cplCovariance = Convert.ToDouble(objl_CMD.ExecuteScalar())
            End With

            With objl_CMD
                .Parameters("@MeasurementName1").Value = "T1"
                .Parameters("@MeasurementName2").Value = ScoreName
                t1scoreCovariance = Convert.ToDouble(objl_CMD.ExecuteScalar())
            End With

            With objl_CMD
                .Parameters("@MeasurementName1").Value = "ScACPL"
                .Parameters("@MeasurementName2").Value = ScoreName
                cplscoreCovariance = Convert.ToDouble(objl_CMD.ExecuteScalar())
            End With

            Dim covarianceMatrix As Double(,) = {
                {t1Variance, t1cplCovariance, t1scoreCovariance},
                {t1cplCovariance, cplVariance, cplscoreCovariance},
                {t1scoreCovariance, cplscoreCovariance, scoreVariance}
            }

            Return covarianceMatrix
        End Function

        Private Function MahalanobisDistance(point As Double(), mean As Double(), covarianceMatrix As Double(,)) As Double
            '''Return the Mahalanobis distance of a point in n-dimensional space; ChatGPT did pretty good! The output matches my original Python calculation
            If point.Length <> mean.Length Then
                Throw New ArgumentException("Point and mean vector must have the same length.")
            End If

            Dim dimension As Integer = point.Length
            Dim deviationVector(dimension - 1) As Double
            For i As Integer = 0 To dimension - 1
                deviationVector(i) = point(i) - mean(i)
            Next

            Dim invCovMatrix(,) As Double = MatrixInverse(covarianceMatrix)
            Dim dotProduct As Double = 0
            For i As Integer = 0 To dimension - 1
                For j As Integer = 0 To dimension - 1
                    dotProduct += deviationVector(i) * invCovMatrix(i, j) * deviationVector(j)
                Next
            Next

            Return Math.Sqrt(dotProduct)
        End Function

        Private Function MatrixInverse(matrix As Double(,)) As Double(,)
            Dim n As Integer = matrix.GetLength(0)
            If matrix.GetLength(0) <> matrix.GetLength(1) Then
                Throw New ArgumentException("Matrix must be square.")
            End If

            Dim augmentedMatrix(n - 1, 2 * n - 1) As Double
            For i As Integer = 0 To n - 1
                For j As Integer = 0 To n - 1
                    augmentedMatrix(i, j) = matrix(i, j)
                Next
                augmentedMatrix(i, i + n) = 1
            Next

            For i As Integer = 0 To n - 1
                Dim divisor As Double = augmentedMatrix(i, i)
                For j As Integer = 0 To 2 * n - 1
                    augmentedMatrix(i, j) /= divisor
                Next

                For k As Integer = 0 To n - 1
                    If k <> i Then
                        Dim factor As Double = augmentedMatrix(k, i)
                        For j As Integer = 0 To 2 * n - 1
                            augmentedMatrix(k, j) -= factor * augmentedMatrix(i, j)
                        Next
                    End If
                Next
            Next

            Dim inverseMatrix(n - 1, n - 1) As Double
            For i As Integer = 0 To n - 1
                For j As Integer = 0 To n - 1
                    inverseMatrix(i, j) = augmentedMatrix(i, j + n)
                Next
            Next

            Return inverseMatrix
        End Function
    End Class
End Class