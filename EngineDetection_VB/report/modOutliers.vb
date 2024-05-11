Imports System.Drawing
Imports Microsoft.Data.SqlClient

Friend Module modOutliers
    Friend Function FlagEVM(piEVM As Double, piSource As String, piTimeControl As String, piAggregation As String, piRating As Short, Optional piColor As String = "") As Char
        'Return an asterisk flag character if the provided EVM value is an extreme value
        Dim objl_CMD As New SqlCommand
        With objl_CMD
            .Connection = MainWindow.db_Connection
            If piColor = "" Then
                .CommandText = modQueries.EVM_Outlier()
            Else
                .CommandText = modQueries.EVM_Outlier(piColor)
            End If
            .Parameters.AddWithValue("@SourceName", piSource)
            .Parameters.AddWithValue("@TimeControlName", piTimeControl)
            .Parameters.AddWithValue("AggregationName", piAggregation)
            .Parameters.AddWithValue("@RatingID", piRating)
        End With

        Dim flg As Char = ""
        With objl_CMD.ExecuteReader
            While .Read
                Dim z_score As Double = (piEVM - Convert.ToDouble(.Item("Average"))) / Convert.ToDouble(.Item("StandardDeviation"))
                If z_score >= 4 OrElse piEVM >= Convert.ToDouble(.Item("MaxValue")) Then flg = "*"
            End While
            .Close()
        End With

        Return flg
    End Function

    Friend Function FlagCPL(piCPL As Double, piMeasurementName As String, piSource As String, piTimeControl As String, piAggregation As String, piRating As Short, Optional piColor As String = "") As Char
        'Return an asterisk flag character if the provided ACPL/SDCPL value is an extreme value
        Dim objl_CMD As New SqlCommand
        With objl_CMD
            .Connection = MainWindow.db_Connection
            If piColor = "" Then
                .CommandText = modQueries.CPL_Outlier()
            Else
                .CommandText = modQueries.CPL_Outlier(piColor)
            End If
            .Parameters.AddWithValue("@MeasurementName", piMeasurementName)
            .Parameters.AddWithValue("@SourceName", piSource)
            .Parameters.AddWithValue("@TimeControlName", piTimeControl)
            .Parameters.AddWithValue("AggregationName", piAggregation)
            .Parameters.AddWithValue("@RatingID", piRating)
        End With

        Dim flg As Char = ""
        With objl_CMD.ExecuteReader
            While .Read
                Dim z_score As Double = (piCPL - Convert.ToDouble(.Item("Average"))) / Convert.ToDouble(.Item("StandardDeviation"))
                If z_score <= -4 OrElse piCPL <= Convert.ToDouble(.Item("MinValue")) Then flg = "*"
            End While
            .Close()
        End With

        Return flg
    End Function
End Module
