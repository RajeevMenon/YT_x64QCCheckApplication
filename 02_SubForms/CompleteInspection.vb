Public Class CompleteInspection

    Dim TmlEntityQA As New MFG_ENTITY.Op(MainForm.Setting.Var_04_MySql_QA)
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim Tbl As New POCO_QA.yta_qcc_v1p2
        Dim ProcessNos = MainForm.AllCheckResult.Select(Function(x) x.ProcessNo).Distinct.ToArray
        Dim Sql(1000) As String
        Dim Count As Integer = 0
        Dim ErrMsg As String = ""

        For Each process In ProcessNos

            Dim GroupStep = MainForm.AllCheckResult.Where(Function(x) x.ProcessNo = process).Select(Function(y) y.StepNo).ToArray
            Dim AllStepPass As Boolean = False
            Dim ProcessSave As New CheckSheetStep
            Dim StepCount As Integer = 0
            Dim ConsolidatedResultString As String = ""
            For Each IndividualStep In GroupStep
                For Each item In MainForm.AllCheckResult
                    If item.ProcessNo = process And item.StepNo Like IndividualStep Then
                        StepCount += 1
                        ProcessSave = item
                        If item.CheckResult = True Then
                            AllStepPass = True
                            If ConsolidatedResultString.Length = 0 Then
                                ConsolidatedResultString = item.Result
                            Else
                                If item.Result.Length > 0 Then ConsolidatedResultString &= "$" & item.Result
                            End If
                        Else
                            AllStepPass = False
                            GoTo UpdateResult
                        End If
                    End If
                Next
            Next

UpdateResult:
            If AllStepPass And StepCount = GroupStep.Length Then
                Tbl.INDEX_NO = MainForm.CustOrd.INDEX_NO
                Tbl.PROCESS_NO = ProcessSave.ProcessNo
                Tbl.PROCESS_STEP = ProcessSave.ProcessStep
                Tbl.ACTIVITY = ProcessSave.Activity
                Tbl.TO_CHECK = ProcessSave.ToCheck
                Tbl.CHECK_RESULT = ConsolidatedResultString 'ProcessSave.Result
                Tbl.METHOD = ProcessSave.Method
                Tbl.INITIAL = ProcessSave.Initial
                Tbl.REMARK = ProcessSave.StepNo_Group & "-GO"
            Else
                Tbl.INDEX_NO = MainForm.CustOrd.INDEX_NO
                Tbl.PROCESS_NO = ProcessSave.ProcessNo
                Tbl.PROCESS_STEP = ProcessSave.ProcessStep
                Tbl.ACTIVITY = ProcessSave.Activity
                Tbl.TO_CHECK = ProcessSave.ToCheck
                Tbl.CHECK_RESULT = ConsolidatedResultString 'ProcessSave.Result
                Tbl.METHOD = ProcessSave.Method
                Tbl.INITIAL = ProcessSave.Initial
                Tbl.REMARK = ProcessSave.StepNo_Group & "-NG"
            End If

            Dim SavedLine = TmlEntityQA.GetDatabaseTableAs_Object(Of POCO_QA.yta_qcc_v1p2)("INDEX_NO", MainForm.CustOrd.INDEX_NO, "PROCESS_NO", Decimal.Parse(ProcessSave.ProcessNo).ToString("N1"), ErrMsg)
            If ErrMsg.Length > 0 Then
                MsgBox(ErrMsg, MsgBoxStyle.Critical, "Error")
                Exit Sub
            End If
            If Not IsNothing(SavedLine.INDEX_NO) Then
                Sql(Count) = TmlEntityQA.SetDatabaseModel_Sql(Tbl, "INDEX_NO", MainForm.CustOrd.INDEX_NO, "PROCESS_NO", Decimal.Parse(ProcessSave.ProcessNo).ToString("N1"), ErrMsg)
                Count += 1
            Else
                Sql(Count) = TmlEntityQA.SaveDatabaseModel_Sql(Tbl, ErrMsg)
                Count += 1
            End If
        Next
        ReDim Preserve Sql(Count - 1)
        TmlEntityQA.ExecuteTransactionQuery(Sql, ErrMsg)
        If ErrMsg.Length > 0 Then
            MsgBox(ErrMsg, MsgBoxStyle.Critical, "Error")
        End If
    End Sub

    Private Sub Button_ViewQCC_Click(sender As Object, e As EventArgs) Handles Button_ViewQCC.Click
        MainForm.PrintQcc_Rev1()
    End Sub


End Class