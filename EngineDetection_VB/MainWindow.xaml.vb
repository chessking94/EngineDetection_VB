Imports Microsoft.Data.SqlClient
Imports System.Reflection

Class MainWindow
    'window opens: only option is to choose report type (Event or Player)
    'going to be ugly, but am having trouble dynamically generating xaml elements. may need to create/position them all, then update the visibility at each step
    '''Event path
    '''expose input box for event name
    '''validate input against database. if no events of that name are found, error out
    '''expose input box for game source (values from dim.Source.SourceName). if the event entered only has one possible source, pre-populate it and do not allow field value to change
    '''if there's multiple options, give user option to choose which one

    '''Player path
    '''expose input boxes for player first and last names
    '''validate inputs against database. if no players are found, error out
    '''expose input box for game source (values from dim.Source). if the event entered only has one possible source, pre-populate it and do not allow field value to change
    '''if there's multiple options, give user option to choose which one
    '''expose date entries for start and end dates and allow user to choose the dates. validate to ensure start date is on or before the end date

    'for both paths, guide user in selecting the comparison dataset. All options should come from the DB and/or be predefined
    '''1. choose a source
    '''2. choose a time control
    '''3. choose a ratingID
    '''4. choose a score name

    'other variables:
    '''engine - get from DB instead of an input parameter
    '''depth - get from DB instead of an input parameter
    '''max eval - seems like getting from the DB would be better

    Private bool_Error As Boolean
    Private lst_Errors As New List(Of String)

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
        Dim objm_Sources As List(Of String)
        If objl_Parameters.EventName <> "" Then
            objm_Sources = EventName(objl_Parameters.EventName)
            'validate event entered exists in the database
            'if yes, populate source if only one source for the event exists, otherwise populate the options in the sel_Source ComboBox
            'if no, clear out the event text box and tell the user it was a bad name
        Else
            objm_Sources = PlayerName(objl_Parameters.FirstName, objl_Parameters.LastName)
            'validate name entered exists in the database
            'if yes, populate source if only one source for the name exists, otherwise populate the options in the sel_Source ComboBox
            'if no, clear out the first and last name text boxes and tell the user it was a bad name
        End If

        bool_Error = False
        If objm_Sources.Count = 0 Then
            bool_Error = True
        Else
            For Each source As String In objm_Sources
                cb_SourceName.Items.Add(source)
            Next
        End If

        If Not bool_Error Then
            ToggleSource(Visibility.Visible)
        Else
            ToggleSource(Visibility.Hidden)
            MessageBox.Show("No sources for provided name found", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            If objl_Parameters.EventName <> "" Then
                tb_EventName.Text = ""
            Else
                tb_LastName.Text = ""
                tb_FirstName.Text = ""
            End If
        End If

        btn_ValidateParameters.IsEnabled = False
    End Sub

    Private Sub Generate() Handles btn_Generate.Click
        MessageBox.Show("Stub", "Stub", MessageBoxButton.OK, MessageBoxImage.Information)
    End Sub
#End Region

#Region "GUI Updates"
    Private Sub ReportTypeChanged() Handles cb_ReportType.SelectionChanged
        'ensure these are reset if the user went back and changed
        objl_Parameters.ClearVariables()
        tb_EventName.Text = ""
        tb_LastName.Text = ""
        tb_FirstName.Text = ""
        cb_SourceName.Items.Clear()
        chk_UseCompareStats.IsChecked = False
        cb_CompareSource.Items.Clear()
        cb_CompareTimeControl.Items.Clear()
        cb_CompareRatingID.Items.Clear()
        cb_CompareScoreName.Items.Clear()

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
            btn_ValidateParameters.Visibility = Visibility.Hidden
        ElseIf tb_LastName.Text = "" OrElse tb_FirstName.Text = "" Then
            If tb_LastName.Text = "" Then
                tb_LastName.Background = NameWarningColor()
            Else
                tb_FirstName.Background = NameWarningColor()
            End If
            btn_ValidateParameters.Visibility = Visibility.Hidden
        Else
            tb_LastName.Background = Nothing
            tb_FirstName.Background = Nothing
            btn_ValidateParameters.IsEnabled = True
            btn_ValidateParameters.Visibility = Visibility.Visible
        End If
    End Sub

    Private Function NameWarningColor() As SolidColorBrush
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
        End If
    End Sub

    Private Sub CompareTimeControlChanged() Handles cb_CompareTimeControl.SelectionChanged
        If cb_CompareTimeControl.SelectedIndex >= 0 Then
            objl_Parameters.CompareTimeControl = cb_CompareTimeControl.SelectedValue

            cb_CompareRatingID.IsEnabled = True

            Dim objm_RatingIDs As List(Of Short) = modValidation.CompareRatingIDs(objl_Parameters.CompareSourceName, objl_Parameters.CompareTimeControl)
            For Each source As String In objm_RatingIDs
                cb_CompareRatingID.Items.Add(source)
            Next

            objl_Parameters.CompareScoreName = ""
            cb_CompareScoreName.SelectedValue = Nothing
            cb_CompareScoreName.IsEnabled = False
        End If
    End Sub

    Private Sub CompareRatingIDChanged() Handles cb_CompareRatingID.SelectionChanged
        If cb_CompareRatingID.SelectedIndex >= 0 Then
            objl_Parameters.CompareRatingID = cb_CompareRatingID.SelectedValue

            cb_CompareScoreName.IsEnabled = True

            For Each score As enum_ScoreName In [Enum].GetValues(GetType(enum_ScoreName))
                cb_CompareScoreName.Items.Add(score.ToString())
            Next
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
