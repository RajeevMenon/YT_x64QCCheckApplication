Public Class YTA_CheckSheet

#Region "Common"
    Public Shared Function ProcessStepNo(ByVal StepNo As Integer, ByVal Initial As String) As CheckSheetStep
        Try
            Dim ProcessStepReturn As New CheckSheetStep

            If StepNo = 20 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo20(Initial)

            Return ProcessStepReturn
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
#End Region

#Region "Version 1.2"
    Public Function ProcessStepNo20(ByVal Initial As String) As CheckSheetStep
        Try
            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "20"
            ProcessStepReturn.ProcessStep = "Parts correctness"
            ProcessStepReturn.ActivityToCheck = "Complete unit and KD Part correctness check. Part numbers and Quantity as per Bill of Material"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.ProcedureStep
            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.ProcedureStepAction.ImagePath_ProcedureStep = MainForm.Setting.Var_51_ProStepView_ImagePath & "YTA\20\" & "B_J_Bracket.jpg"
            Return ProcessStepReturn
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Function ProcessStepNo30(ByVal Initial As String, ByVal IndexNo As String, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try
            Dim TmlEntity As New MFG_ENTITY.Op(MainForm.Setting.Var_03_MySql_YGSP)
            Dim FieldName As String = ""
            If IndexNo.Length = 10 Then FieldName = "BAR" Else FieldName = "INDEX_NO"
            Dim CustOrd = TmlEntity.GetDatabaseTableAs_Object(Of POCO_YGSP.cust_ord)(FieldName, IndexNo, FieldName, IndexNo, ErrMsg)
            If ErrMsg.Length > 0 Then
                Return Nothing
                Exit Function
            End If

            Dim PlatePartNo As String = ""
            Dim Inst_Lib As New TML_Library.Instrument
            PlatePartNo = Inst_Lib.GetNamePlatePartNumber(CustOrd.MS_CODE, MainForm.Setting.Var_05_Factory, IndexNo)
            Dim allParts() As String
            allParts = Inst_Lib.GetYTA_AllPlatePartNumberList(CustOrd.MS_CODE, MainForm.Setting.Var_05_Factory, ErrMsg)
            If Len(ErrMsg) > 0 Then
                ErrMsg = "Error in reading all plate part numbers:" & ErrMsg
                Exit Function
            End If
            If allParts.Length > 0 Then
                Dim SelParts(3) As String
                Dim rl As New TML_Library.RandomArray
                SelParts = rl.MakeRandomControlledlist(PlatePartNo, allParts, 4)
                If CustOrd.SERIAL_NO Like "S5*" Then
                    For i As Integer = 0 To SelParts.Length - 1
                        If SelParts(i) = "F9220MV" Then
                            SelParts(i) = "F9220LV"
                        End If
                        If SelParts(i) = "F9220MW" Then
                            SelParts(i) = "F9220LW"
                        End If
                    Next
                    If PlatePartNo = "F9220MW" Then
                        PlatePartNo = "F9220LW"
                    End If
                    If PlatePartNo = "F9220LV" Then
                        PlatePartNo = "F9220LW"
                    End If
                End If

                Dim ProcessStepReturn As New CheckSheetStep
                ProcessStepReturn.ProcessNo = "30"
                ProcessStepReturn.ProcessStep = "Data Plate Marking"
                ProcessStepReturn.ActivityToCheck = "Data Plates with correct contents."
                ProcessStepReturn.Method = CheckSheetStep.MethodOption.UserIput
                ProcessStepReturn.Initial = Initial
                ProcessStepReturn.UserInputAction.UserActionMessage = "Choose the Plate Part number from below list and Click SELECT button."
                ProcessStepReturn.UserInputAction.UserInputList = SelParts
                ProcessStepReturn.UserInputAction.UserInputCorrect = PlatePartNo
                Return ProcessStepReturn

            End If

        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Function ProcessStepNo31(ByVal Initial As String) As CheckSheetStep
        Try
            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "30"
            ProcessStepReturn.ProcessStep = "Data Plate/ Tag Plate Marking"
            ProcessStepReturn.ActivityToCheck = "Tag Plate Marking"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.ProcedureStep
            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.ProcedureStepAction.ImagePath_ProcedureStep = MainForm.Setting.Var_51_ProStepView_ImagePath & "YTA\20\" & "B_J_Bracket.jpg"
            Return ProcessStepReturn
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

#End Region
End Class
