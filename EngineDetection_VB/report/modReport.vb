Imports Microsoft.VisualBasic.FileIO
Imports System.IO
Imports System.Reflection

Friend Module modReport
    Public objm_Lines As New List(Of String)

    Friend Sub BuildReport()
        Dim objl_Info As New clsInfo
        'Dim objl_Detail As New clsDetail
        'call methods in order to append sections to objm_Lines

        '''Python sections:
        'g.header_type(rpt, engine, ev, full_name, start_date, end_date)
        'g.header_info(engine_name, depth)
        'g.scoring_desc(engine)
        'r.key_stats()
        'g.player_key()
        'r.player_summary()
        'g.game_key()
        'r.game_traces()

        WriteReport()
    End Sub

    Private Sub WriteReport()
        Dim params As clsParameters = MainWindow.objl_Parameters
        Dim outputDir As String = Path.Combine(SpecialDirectories.Desktop, "Local_Applications", Assembly.GetCallingAssembly().GetName().Name)
        Dim reportName As String = "ReportType_Name_StartDate_EndDate.txt"

        Select Case params.ReportType
            Case "Event"
                reportName = $"{params.ReportType}_{params.EventName}.txt"
            Case "Player"
                reportName = $"{params.ReportType}_{params.FirstName} {params.LastName}.txt"
        End Select

        If Not Directory.Exists(outputDir) Then
            Directory.CreateDirectory(outputDir)
        End If

        Dim fileName As String = Path.Combine(outputDir, reportName)
        Dim abortReason As String = ""
        If File.Exists(fileName) Then
            Dim result As Forms.DialogResult = MessageBox.Show("This report already exists, do you want to overwrite it?", "Report exists", MessageBoxButton.YesNo, MessageBoxImage.Question)
            Select Case result
                Case Forms.DialogResult.Yes
                    Try
                        File.Delete(fileName)
                    Catch ex As Exception
                        abortReason = "Unable to delete file"
                    End Try
                Case Forms.DialogResult.No
                    abortReason = "User cancelled report write"
            End Select
        End If

        If abortReason <> "" Then
            MessageBox.Show(abortReason, "Report Creation Aborted", MessageBoxButton.OK, MessageBoxImage.Exclamation)
        Else
            'File.WriteAllLines(fileName, objm_Lines)  'TODO: Investigate if this can work, didn't in initial test
            Using writer As New StreamWriter(fileName, False, Encoding.UTF8)
                For Each line In objm_Lines
                    writer.WriteLine(line)
                Next
            End Using

            Try
                Process.Start("explorer.exe", outputDir)
            Catch ex As Exception
                MessageBox.Show($"Process complete! File located at {fileName}", "Process Complete", MessageBoxButton.OK, MessageBoxImage.Information)
            End Try
        End If
    End Sub
End Module
