Imports Microsoft.Data.SqlClient

Public Class clsInfo
    Friend ReadOnly CompareSourceName As String
    Friend ReadOnly CompareTimeControl As String
    Friend ReadOnly CompareRatingID As Short
    Friend ReadOnly CompareScoreName As String
    Friend ReadOnly Engine As String
    Friend ReadOnly Depth As Short
    Friend ReadOnly MaxEval As Double

    Public Sub New()
        Dim params As clsParameters = MainWindow.objl_Parameters

        CompareSourceName = params.CompareSourceName
        CompareTimeControl = params.CompareTimeControl
        CompareRatingID = params.CompareRatingID
        CompareScoreName = params.CompareScoreName

        Engine = GetEngineName(params)
        Depth = GetDepth(params)
        MaxEval = GetMaxEval()
    End Sub

    Private Function GetEngineName(params As clsParameters) As String
        Dim rtnval As String

        If params.EventID > 0 Then
            Using objl_CMD As New SqlCommand(modQueries.EventEngine(), MainWindow.db_Connection)
                objl_CMD.Parameters.AddWithValue("@EventID", params.EventID)
                rtnval = Convert.ToString(objl_CMD.ExecuteScalar())
            End Using
        Else
            Using objl_CMD As New SqlCommand(modQueries.PlayerEngine(), MainWindow.db_Connection)
                objl_CMD.Parameters.AddWithValue("@PlayerID", params.PlayerID)
                rtnval = Convert.ToString(objl_CMD.ExecuteScalar())
            End Using
        End If

        Return rtnval
    End Function

    Private Function GetDepth(params As clsParameters) As String
        Dim rtnval As Short

        If params.EventID > 0 Then
            Using objl_CMD As New SqlCommand(modQueries.EventDepth(), MainWindow.db_Connection)
                objl_CMD.Parameters.AddWithValue("@EventID", params.EventID)
                rtnval = Convert.ToInt16(objl_CMD.ExecuteScalar())
            End Using
        Else
            Using objl_CMD As New SqlCommand(modQueries.PlayerDepth(), MainWindow.db_Connection)
                objl_CMD.Parameters.AddWithValue("@PlayerID", params.PlayerID)
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
End Class
