Public Class CheckSheetStep
    Enum MethodOption
        UserIput = 1
        SinglePointInstruction = 2
        ProcedureStep = 3
        QrCodeCheck = 4
        DocumentCheck = 5
    End Enum
    Public Property ProcessNo As String
    Public Property ProcessStep As String
    Public Property ActivityToCheck As String
    Public Property Method As MethodOption
    Public Property Initial As String
    Public Property UserInputAction As New UserInput
    Public Property ProcedureStepAction As New ProcedureStep
    Public Property SinglePointAction As New SinglePointInst

End Class
