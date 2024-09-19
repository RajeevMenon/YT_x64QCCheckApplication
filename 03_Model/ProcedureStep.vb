Public Class ProcedureStep

    Public Property ProcedureMessage As String
    Public Property ImagePath_ProcedureStep As String
    Public Property SetSizeMode As SizeMode

    Public Enum SizeMode
        Normal
        StretchImage
        AutoSize
        CenterImage
        Zoom
    End Enum

End Class
