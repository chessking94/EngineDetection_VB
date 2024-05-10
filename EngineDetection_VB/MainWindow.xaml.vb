Imports Microsoft.Data.SqlClient
Imports System.Reflection

' TODO: Mark outliers with an asterisk (EVM, ACPL, SDCPL, Score, ROI, PValue)

Class MainWindow
    Private ErrorReason As String = ""

    Public Shared db_Connection As SqlConnection = Utilities_NetCore.ConnectionLocal("ChessWarehouse", Assembly.GetCallingAssembly().GetName().Name)
    Public Shared objl_Parameters As New clsParameters

    Private Sub WindowLoaded() Handles Me.Loaded
        'build selection options for report types
        For Each report As enum_ReportType In [Enum].GetValues(GetType(enum_ReportType))
            cb_ReportType.Items.Add(report.ToString())
        Next

        'hide all other elements unless ready
        ToggleEvent(Visibility.Hidden)
        ToggleName(Visibility.Hidden)
        ToggleSource(Visibility.Hidden)
        TogglePreCompareStats(Visibility.Hidden)
        ToggleCompareStats(Visibility.Hidden)
        btn_Generate.IsEnabled = False

        Try
            If db_Connection.State <> System.Data.ConnectionState.Open Then
                db_Connection.Open()
            End If
        Catch ex As Exception
            cb_ReportType.IsEnabled = False
            Throw New Exception($"Unable to establish database connection: {ex.Message}")
        End Try
    End Sub

    Private Sub WindowClosed() Handles Me.Closed
        If db_Connection.State = System.Data.ConnectionState.Open Then
            db_Connection.Close()
        End If
        db_Connection.Dispose()
    End Sub

#Region "Buttons"
    Private Sub ValidateParameters() Handles btn_ValidateParameters.Click
        btn_ValidateParameters.IsEnabled = False
        tb_EventName.IsEnabled = False
        tb_FirstName.IsEnabled = False
        tb_LastName.IsEnabled = False
        dp_StartDate.IsEnabled = False
        dp_EndDate.IsEnabled = False

        ErrorReason = ""
        If dp_StartDate.SelectedDate > dp_EndDate.SelectedDate Then
            ErrorReason = "Start Date is after End Date"
        End If

        If ErrorReason = "" Then
            Dim objm_Sources As List(Of String)
            If objl_Parameters.EventName <> "" Then
                objm_Sources = EventName(objl_Parameters.EventName)
            Else
                objm_Sources = PlayerName(objl_Parameters.FirstName, objl_Parameters.LastName)
            End If

            If objm_Sources.Count = 0 Then
                ErrorReason = "No sources for provided name found"
            Else
                For Each source As String In objm_Sources
                    cb_SourceName.Items.Add(source)
                Next
            End If
        End If

        If ErrorReason = "" Then
            ToggleSource(Visibility.Visible)
            cb_SourceName.IsEnabled = True
            If cb_SourceName.Items.Count = 1 Then
                cb_SourceName.SelectedIndex = 0
                cb_SourceName.IsEnabled = False
            End If
        Else
            ToggleSource(Visibility.Hidden)
            tb_EventName.IsEnabled = True
            tb_FirstName.IsEnabled = True
            tb_LastName.IsEnabled = True
            dp_StartDate.IsEnabled = True
            dp_EndDate.IsEnabled = True
            MessageBox.Show(ErrorReason, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            If objl_Parameters.EventName <> "" Then
                tb_EventName.Text = ""
            Else
                tb_LastName.Text = ""
                tb_FirstName.Text = ""
                dp_StartDate.SelectedDate = Nothing
                dp_EndDate.SelectedDate = Nothing
            End If
        End If
    End Sub

    Private Sub Generate() Handles btn_Generate.Click
        btn_Generate.IsEnabled = False
        objl_Parameters.PopulateIDVariables()
        BuildReport()
    End Sub
#End Region

#Region "GUI Updates"
    Private Sub ReportTypeChanged() Handles cb_ReportType.SelectionChanged
        'ensure these are reset if the user went back and changed
        objl_Parameters.ClearVariables()
        tb_EventName.Text = ""
        tb_LastName.Text = ""
        tb_FirstName.Text = ""
        dp_StartDate.SelectedDate = Nothing
        dp_EndDate.SelectedDate = Nothing
        cb_SourceName.Items.Clear()
        chk_UseCompareStats.IsChecked = False
        cb_CompareSource.Items.Clear()
        cb_CompareTimeControl.Items.Clear()
        cb_CompareRatingID.Items.Clear()
        cb_CompareScoreName.Items.Clear()
        tb_EventName.IsEnabled = True
        tb_FirstName.IsEnabled = True
        tb_LastName.IsEnabled = True
        dp_StartDate.IsEnabled = True
        dp_EndDate.IsEnabled = True

        objl_Parameters.ReportType = cb_ReportType.SelectedValue
        If cb_ReportType.SelectedIndex >= 0 Then
            Select Case cb_ReportType.SelectedValue
                Case "Event"
                    ToggleEvent(Visibility.Visible)
                    ToggleName(Visibility.Hidden)
                    tb_EventName.Focus()
                Case "Player"
                    ToggleEvent(Visibility.Hidden)
                    ToggleName(Visibility.Visible)
                    tb_FirstName.Focus()
            End Select

            ToggleSource(Visibility.Hidden)
            TogglePreCompareStats(Visibility.Hidden)
            ToggleCompareStats(Visibility.Hidden)
        End If
    End Sub

    Private Sub CompareStats_Clicked() Handles chk_UseCompareStats.Click
        If chk_UseCompareStats.IsChecked Then
            ToggleCompareStats(Visibility.Visible)
            btn_Generate.IsEnabled = False

            cb_CompareSource.IsEnabled = True
            cb_CompareTimeControl.IsEnabled = False
            cb_CompareRatingID.IsEnabled = False
            cb_CompareScoreName.IsEnabled = False

            Dim objm_Sources As List(Of String) = modValidation.CompareSources()
            For Each source As String In objm_Sources
                cb_CompareSource.Items.Add(source)
            Next

            If cb_CompareSource.Items.Count = 1 Then
                cb_CompareSource.SelectedIndex = 0
                cb_CompareSource.IsEnabled = False
            End If
        Else
            ToggleCompareStats(Visibility.Hidden)
            btn_Generate.IsEnabled = True
            objl_Parameters.ClearCompareVariables()
            cb_CompareSource.Items.Clear()
            cb_CompareTimeControl.Items.Clear()
            cb_CompareRatingID.Items.Clear()
            cb_CompareScoreName.Items.Clear()
        End If
    End Sub

    Private Sub EventChanged() Handles tb_EventName.TextChanged
        objl_Parameters.EventName = tb_EventName.Text

        If tb_EventName.Text = "" Then
            btn_ValidateParameters.Visibility = Visibility.Hidden
        Else
            btn_ValidateParameters.IsEnabled = True
            btn_ValidateParameters.Visibility = Visibility.Visible
        End If
    End Sub

    Private Sub NameChanged() Handles tb_FirstName.TextChanged, tb_LastName.TextChanged
        objl_Parameters.LastName = tb_LastName.Text
        objl_Parameters.FirstName = tb_FirstName.Text

        If tb_LastName.Text = "" AndAlso tb_FirstName.Text = "" Then
            tb_LastName.Background = Nothing
            tb_FirstName.Background = Nothing
        ElseIf tb_LastName.Text = "" OrElse tb_FirstName.Text = "" Then
            If tb_LastName.Text = "" Then
                tb_LastName.Background = WarningColor()
            Else
                tb_FirstName.Background = WarningColor()
            End If
        Else
            tb_LastName.Background = Nothing
            tb_FirstName.Background = Nothing
        End If
    End Sub

    Private Sub DatesChanged() Handles dp_StartDate.SelectedDateChanged, dp_EndDate.SelectedDateChanged
        objl_Parameters.StartDate = dp_StartDate.SelectedDate
        objl_Parameters.EndDate = dp_EndDate.SelectedDate

        If dp_StartDate.SelectedDate Is Nothing AndAlso dp_EndDate.SelectedDate Is Nothing Then
            dp_StartDate.Background = Nothing
            dp_EndDate.Background = Nothing
            btn_ValidateParameters.Visibility = Visibility.Hidden
        ElseIf dp_StartDate.SelectedDate Is Nothing OrElse dp_EndDate.SelectedDate Is Nothing Then
            If dp_StartDate.SelectedDate Is Nothing Then
                dp_StartDate.Background = WarningColor()
            Else
                dp_EndDate.Background = WarningColor()
            End If
            btn_ValidateParameters.Visibility = Visibility.Hidden
        Else
            dp_StartDate.Background = Nothing
            dp_EndDate.Background = Nothing
            btn_ValidateParameters.IsEnabled = True
            btn_ValidateParameters.Visibility = Visibility.Visible
        End If
    End Sub

    Private Function WarningColor() As SolidColorBrush
        Dim redValue As Byte = 255
        Dim greenValue As Byte = 192
        Dim blueValue As Byte = 192
        Dim opacityValue As Double = 0.4
        Dim myColor As Color = Color.FromArgb(opacityValue * 255, redValue, greenValue, blueValue)
        Dim brush As New SolidColorBrush(myColor)

        Return brush
    End Function

    Private Sub SourceChanged() Handles cb_SourceName.SelectionChanged
        objl_Parameters.SourceName = cb_SourceName.SelectedValue
        TogglePreCompareStats(Visibility.Visible)
        btn_Generate.IsEnabled = True
    End Sub

    Private Sub CompareSourceChanged() Handles cb_CompareSource.SelectionChanged
        If cb_CompareSource.SelectedIndex >= 0 Then
            objl_Parameters.CompareSourceName = cb_CompareSource.SelectedValue
            cb_CompareTimeControl.Items.Clear()
            cb_CompareRatingID.Items.Clear()
            cb_CompareScoreName.Items.Clear()

            cb_CompareTimeControl.IsEnabled = True

            Dim objm_TimeControls As List(Of String) = modValidation.CompareTimeControls(objl_Parameters.CompareSourceName)
            For Each source As String In objm_TimeControls
                cb_CompareTimeControl.Items.Add(source)
            Next

            objl_Parameters.CompareRatingID = -1
            cb_CompareRatingID.SelectedValue = Nothing

            objl_Parameters.CompareScoreName = ""
            cb_CompareScoreName.SelectedValue = Nothing

            cb_CompareRatingID.IsEnabled = False
            cb_CompareScoreName.IsEnabled = False

            If cb_CompareTimeControl.Items.Count = 1 Then
                cb_CompareTimeControl.SelectedIndex = 0
                cb_CompareTimeControl.IsEnabled = False
            End If
        End If
    End Sub

    Private Sub CompareTimeControlChanged() Handles cb_CompareTimeControl.SelectionChanged
        If cb_CompareTimeControl.SelectedIndex >= 0 Then
            objl_Parameters.CompareTimeControl = cb_CompareTimeControl.SelectedValue
            cb_CompareRatingID.Items.Clear()
            cb_CompareScoreName.Items.Clear()

            cb_CompareRatingID.IsEnabled = True

            Dim objm_RatingIDs As List(Of Short) = modValidation.CompareRatingIDs(objl_Parameters.CompareSourceName, objl_Parameters.CompareTimeControl)
            For Each source As String In objm_RatingIDs
                cb_CompareRatingID.Items.Add(source)
            Next

            objl_Parameters.CompareScoreName = ""
            cb_CompareScoreName.SelectedValue = Nothing
            cb_CompareScoreName.IsEnabled = False

            If cb_CompareRatingID.Items.Count = 1 Then
                cb_CompareRatingID.SelectedIndex = 0
                cb_CompareRatingID.IsEnabled = False
            End If
        End If
    End Sub

    Private Sub CompareRatingIDChanged() Handles cb_CompareRatingID.SelectionChanged
        If cb_CompareRatingID.SelectedIndex >= 0 Then
            objl_Parameters.CompareRatingID = cb_CompareRatingID.SelectedValue
            cb_CompareScoreName.Items.Clear()

            cb_CompareScoreName.IsEnabled = True

            For Each score As enum_ScoreName In [Enum].GetValues(GetType(enum_ScoreName))
                cb_CompareScoreName.Items.Add(score.ToString())
            Next

            If cb_CompareScoreName.Items.Count = 1 Then
                cb_CompareScoreName.SelectedIndex = 0
                cb_CompareScoreName.IsEnabled = False
            End If
        End If
    End Sub

    Private Sub CompareScoreNameChanged() Handles cb_CompareScoreName.SelectionChanged
        If cb_CompareScoreName.SelectedIndex >= 0 Then
            objl_Parameters.CompareScoreName = cb_CompareScoreName.SelectedValue
        End If
        btn_Generate.IsEnabled = True
    End Sub
#End Region

#Region "Visibility Toggles"
    Private Sub ToggleEvent(pi_Visibility As Visibility)
        lab_EventName.Visibility = pi_Visibility
        tb_EventName.Visibility = pi_Visibility
    End Sub

    Private Sub ToggleName(pi_Visibility As Visibility)
        lab_FirstName.Visibility = pi_Visibility
        tb_FirstName.Visibility = pi_Visibility
        lab_LastName.Visibility = pi_Visibility
        tb_LastName.Visibility = pi_Visibility
        lab_StartDate.Visibility = pi_Visibility
        dp_StartDate.Visibility = pi_Visibility
        lab_EndDate.Visibility = pi_Visibility
        dp_EndDate.Visibility = pi_Visibility
    End Sub

    Private Sub ToggleSource(pi_Visibility As Visibility)
        lab_SourceName.Visibility = pi_Visibility
        cb_SourceName.Visibility = pi_Visibility
        btn_ValidateParameters.Visibility = pi_Visibility
    End Sub

    Private Sub TogglePreCompareStats(pi_Visibility As Visibility)
        sep_CompareStats.Visibility = pi_Visibility
        chk_UseCompareStats.Visibility = pi_Visibility
    End Sub

    Private Sub ToggleCompareStats(pi_Visibility As Visibility)
        lab_CompareSourceName.Visibility = pi_Visibility
        cb_CompareSource.Visibility = pi_Visibility
        lab_CompareTimeControl.Visibility = pi_Visibility
        cb_CompareTimeControl.Visibility = pi_Visibility
        lab_CompareRatingID.Visibility = pi_Visibility
        cb_CompareRatingID.Visibility = pi_Visibility
        lab_CompareScoreName.Visibility = pi_Visibility
        cb_CompareScoreName.Visibility = pi_Visibility
    End Sub
#End Region
End Class
