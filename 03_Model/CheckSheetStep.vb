Public Class CheckSheetStep
    Enum MethodOption
        UserIput = 1
        SinglePntInst = 2
        ProcedureStep = 3
        QrCodeCheck = 4
        DocumentCheck = 5
    End Enum
    Public Property ProcessNo As String 'As in EJA/YTA QC Checksheet
    Public Property ProcessStep As String 'As in YTA QC Checksheet
    Public Property Activity As String 'As in EJA/YTA QC Checksheet
    Public Property ToCheck As String 'As in EJA/YTA QC Checksheet
    Public Property Method As MethodOption 'Not As in EJA/YTA QC Checksheet
    Public Property Initial As String 'As in EJA/YTA QC Checksheet
    Public Property Result As String 'To Write to EJA/YTA QC Checksheet with Action comma X,Y Coordinates
    Public Property StepNo As String 'Some processno needs substeps. So software need increment by 1 stepno. This is for that purpose.
    Public Property StepNo_Group As String 'If there are more than one substep. Then those needs to be grouped here. If all grouped steps pass, main processno pass
    Public Property ActivityToCheck As String 'Text to display to inspector to perform inspectioon
    Public Property UserInputAction As New UserInput 'Selection from a list of 4 items
    Public Property ProcedureStepAction As New ProcedureStep 'Picture view to do some action
    Public Property SinglePointAction As New SinglePointInstCheck 'Selection from 2 different GO/NG pictures
    Public Property ViewDocAction As New DocumentCheck 'Load a document to check and confirm result
    Public Property CheckResult As New Boolean 'Final check result for the StepNo selected. This is set from subform
End Class
