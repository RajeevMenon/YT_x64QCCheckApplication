Public Class CheckSheetStep
    Enum MethodOption
        UserIput = 1
        SinglePntInst = 2
        ProcedureStep = 3
        QrCodeCheck = 4
        DocumentCheck = 5
    End Enum
    Public Property StepNo As String
    Public Property ProcessNo As String
    Public Property ProcessStep As String
    Public Property ActivityToCheck As String
    Public Property Method As MethodOption
    Public Property Initial As String
    Public Property UserInputAction As New UserInput
    Public Property ProcedureStepAction As New ProcedureStep
    Public Property SinglePointAction As New SinglePointInstCheck
    Public Property ViewDocAction As New DocumentCheck

End Class
