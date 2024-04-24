Public Module modEnums
    Public Enum enum_ReportType
        'if a new Enum is added, it also needs to be added in MainWindow.ReportTypeChanged()
        [Event]
        Player
    End Enum

    Public Enum enum_ScoreName
        WinProbabilityLost
        EvaluationGroupComparison
    End Enum
End Module
