Friend Module modStatistics
    Friend Function MahalanobisDistance(point As Double(), mean As Double(), covarianceMatrix As Double(,)) As Double
        'generated originally by ChatGPT
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
        'generated originally by ChatGPT
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

    Friend Function CalculateROI(pi_zscores As Dictionary(Of String, Double)) As Double
        Dim weights As New Dictionary(Of String, Double) From {
            {"T1", 0.2},
            {"ScACPL", 0.35},
            {"Score", 0.45}
        }

        Dim dotproduct As Double = 0
        For Each kvp As KeyValuePair(Of String, Double) In pi_zscores
            dotproduct += kvp.Value * weights(kvp.Key)
        Next

        Dim sumsquaresroot As Double = 0
        For Each kvp As KeyValuePair(Of String, Double) In weights
            sumsquaresroot += Math.Pow(kvp.Value, 2)
        Next
        sumsquaresroot = Math.Sqrt(sumsquaresroot)

        Dim z As Double = dotproduct / sumsquaresroot
        Dim roi As Double = 5 * z + 50

        Return roi
    End Function
End Module
