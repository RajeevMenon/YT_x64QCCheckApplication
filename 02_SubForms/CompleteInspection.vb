Imports Newtonsoft.Json

Public Class CompleteInspection

    Dim TmlEntityQA As New MFG_ENTITY.Op(MainForm.Setting.Var_04_MySql_QA)
    Dim TmlEntityYGS As New MFG_ENTITY.Op(MainForm.Setting.Var_03_MySql_YGSP)
    Dim WMsg As New WarningForm
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If MainForm.CurrentQCC_Version = "1.4" Then
            Call CompleteInspection_1p4()
        Else
            Call CompleteInspection_1p2()
        End If
    End Sub

    Private Sub CompleteInspection_1p4()
        Try
            Dim ErrMsg As String = ""

            Dim Tbl As New POCO_QA.yta_qcc_v1p4
            Dim ProcessNos = MainForm.AllCheckResult.Select(Function(x) x.ProcessNo).Distinct.ToArray
            Dim Sql(1000) As String
            Dim Count As Integer = 0

            For Each process In ProcessNos
                Dim GroupStep = MainForm.AllCheckResult.Where(Function(x) x.ProcessNo = process).Select(Function(y) y.StepNo).ToArray
                Dim AllStepPass As Boolean = False
                Dim CheckedResults As New List(Of String)
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
                                CheckedResults.Add(item.Result)
                            Else
                                AllStepPass = False
                                GoTo UpdateResult
                            End If
                        End If
                    Next
                Next

UpdateResult:
                If CheckedResults.Count > 0 Then
                    Dim ResultFields As New List(Of OpenPdfOperation_x64.Field)
                    For Each CheckedResult As String In CheckedResults
                        If CheckedResult.Length > 0 Then
                            Dim ReportFields = JsonConvert.DeserializeObject(Of List(Of OpenPdfOperation_x64.Field))(CheckedResult)
                            ResultFields.AddRange(ReportFields)
                        End If
                    Next
                    Dim EjaCheckSheet As New YTA_CheckSheet_v1p4
                    Dim FinalTemplate = EjaCheckSheet.MakeProcessReturnTemplate(ResultFields)
                    ConsolidatedResultString = Newtonsoft.Json.JsonConvert.SerializeObject(FinalTemplate, Newtonsoft.Json.Formatting.Indented)
                Else
                    ConsolidatedResultString = ""
                End If

                If AllStepPass And StepCount = GroupStep.Length Then
                    Tbl.INDEX_NO = MainForm.CustOrd.INDEX_NO
                    Tbl.PROCESS_NO = ProcessSave.ProcessNo
                    Tbl.PROCESS_STEP = "All Steps of " & ProcessSave.ProcessNo 'ProcessSave.ProcessStep
                    Tbl.ACTIVITY = "Inprocess Inspection of Step:" & ProcessSave.ProcessNo 'ProcessSave.Activity
                    Tbl.TO_CHECK = "As per QC Checksheet " & ProcessSave.ProcessNo 'ProcessSave.ToCheck
                    Tbl.CHECK_RESULT = ConsolidatedResultString 'ProcessSave.Result
                    Tbl.METHOD = "Visual/Tools as per QC Process Chart" 'ProcessSave.Method
                    Tbl.INITIAL = ProcessSave.Initial
                    Tbl.REMARK = String.Join(",", GroupStep) & "-GO" 'ProcessSave.StepNo_Group & "-GO"
                Else
                    Tbl.INDEX_NO = MainForm.CustOrd.INDEX_NO
                    Tbl.PROCESS_NO = ProcessSave.ProcessNo
                    Tbl.PROCESS_STEP = "All Steps of " & ProcessSave.ProcessNo 'ProcessSave.ProcessStep
                    Tbl.ACTIVITY = "Inprocess Inspection of Step:" & ProcessSave.ProcessNo 'ProcessSave.Activity
                    Tbl.TO_CHECK = "As per QC Checksheet " & ProcessSave.ProcessNo 'ProcessSave.ToCheck
                    Tbl.CHECK_RESULT = ConsolidatedResultString 'ProcessSave.Result
                    Tbl.METHOD = "Visual/Tools as per QC Process Chart" 'ProcessSave.Method
                    Tbl.INITIAL = ProcessSave.Initial
                    Tbl.REMARK = String.Join(",", GroupStep) & "-GO" 'ProcessSave.StepNo_Group & "-NG"
                End If

                Dim SavedLine = TmlEntityQA.GetDatabaseTableAs_Object(Of POCO_QA.yta_qcc_v1p4)("INDEX_NO", MainForm.CustOrd.INDEX_NO, "PROCESS_NO", ProcessSave.ProcessNo, ErrMsg)
                If ErrMsg.Length > 0 Then
                    WMsg.Message = ErrMsg
                    WMsg.ShowDialog()
                    Exit Sub
                End If
                If Not IsNothing(SavedLine.INDEX_NO) Then
                    Sql(Count) = TmlEntityQA.SetDatabaseModel_Sql(Tbl, "INDEX_NO", MainForm.CustOrd.INDEX_NO, "PROCESS_NO", ProcessSave.ProcessNo, ErrMsg)
                    Count += 1
                Else
                    Sql(Count) = TmlEntityQA.SaveDatabaseModel_Sql(Tbl, ErrMsg)
                    Count += 1
                End If

                If My.Settings.Station = "PACKING" Then
                    If Not IsDate(MainForm.CustOrd.ACTUAL_FINISH_DATE) Then
                        'Check if Images are taken for evidence
                        If System.IO.Directory.GetFiles(Link.EvidenceImagePath, MainForm.CustOrd.INDEX_NO & "*").Count = 0 Then
                            WMsg.Message = "Take pictures of Packing Box contents!"
                            WMsg.ShowDialog()
                        End If

                        'Set Actual finish date
                        Dim FinishDate As String = Date.Today.ToString("yyyy-MM-dd")
                        Dim FinishTime As String = DateAndTime.Now.ToString("hh:mm:ss tt")
                        Dim SqlQry(0) As String
                        SqlQry(0) = "UPDATE cust_ord SET ACTUAL_FINISH_DATE='" & FinishDate & "', ACTUAL_FINISH_TIME='" & FinishTime & "' WHERE INDEX_NO='" & MainForm.CustOrd.INDEX_NO & "';"
                        TmlEntityYGS.ExecuteTransactionQuery(SqlQry, ErrMsg)
                        If ErrMsg.Length > 0 Then
                            WMsg.Message = "Actual Finish Date Update Error:" & ErrMsg
                            WMsg.ShowDialog()
                        Else
                            LoggerHelper.LogInfo($"Actual Finish Date Updated for Index No: {MainForm.CustOrd.INDEX_NO} on {FinishDate} at {FinishTime}")
                        End If
                    End If
                End If

            Next

            ReDim Preserve Sql(Count - 1)
            TmlEntityQA.ExecuteTransactionQuery(Sql, ErrMsg)
            If ErrMsg.Length > 0 Then
                WMsg.Message = ErrMsg
                WMsg.ShowDialog()
            Else
                LoggerHelper.LogInfo($"Completed {MainForm.StationName} Station Inspection steps for Index No: {MainForm.CustOrd.INDEX_NO}")
                MainForm.PrintQcc_Rev2()
                MainForm.LoadIndexScan()
            End If
        Catch ex As Exception
            WMsg.Message = $"CompleteInspection_1p6() Exception: {ex.Message}"
            WMsg.ShowDialog()
        End Try

    End Sub

    Private Sub CompleteInspection_1p2()

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
                Tbl.REMARK = String.Join(",", GroupStep) & "-GO" 'ProcessSave.StepNo_Group & "-GO"
            Else
                Tbl.INDEX_NO = MainForm.CustOrd.INDEX_NO
                Tbl.PROCESS_NO = ProcessSave.ProcessNo
                Tbl.PROCESS_STEP = ProcessSave.ProcessStep
                Tbl.ACTIVITY = ProcessSave.Activity
                Tbl.TO_CHECK = ProcessSave.ToCheck
                Tbl.CHECK_RESULT = ConsolidatedResultString 'ProcessSave.Result
                Tbl.METHOD = ProcessSave.Method
                Tbl.INITIAL = ProcessSave.Initial
                Tbl.REMARK = String.Join(",", GroupStep) & "-GO" 'ProcessSave.StepNo_Group & "-NG"
            End If

            'Dim SavedLine = TmlEntityQA.GetDatabaseTableAs_Object(Of POCO_QA.yta_qcc_v1p2)("INDEX_NO", MainForm.CustOrd.INDEX_NO, "PROCESS_NO", Decimal.Parse(ProcessSave.ProcessNo).ToString("N1"), ErrMsg)
            Dim SavedLine = TmlEntityQA.GetDatabaseTableAs_Object(Of POCO_QA.yta_qcc_v1p2)("INDEX_NO", MainForm.CustOrd.INDEX_NO, "PROCESS_NO", ProcessSave.ProcessNo, ErrMsg)
            If ErrMsg.Length > 0 Then
                MsgBox(ErrMsg, MsgBoxStyle.Critical, "Error")
                Exit Sub
            End If
            If Not IsNothing(SavedLine.INDEX_NO) Then
                'Sql(Count) = TmlEntityQA.SetDatabaseModel_Sql(Tbl, "INDEX_NO", MainForm.CustOrd.INDEX_NO, "PROCESS_NO", Decimal.Parse(ProcessSave.ProcessNo).ToString("N1"), ErrMsg)
                Sql(Count) = TmlEntityQA.SetDatabaseModel_Sql(Tbl, "INDEX_NO", MainForm.CustOrd.INDEX_NO, "PROCESS_NO", ProcessSave.ProcessNo, ErrMsg)
                Count += 1
            Else
                Sql(Count) = TmlEntityQA.SaveDatabaseModel_Sql(Tbl, ErrMsg)
                Count += 1
            End If

            If My.Settings.Station = "PACKING" Then
                If Not IsDate(MainForm.CustOrd.ACTUAL_FINISH_DATE) Then
                    'Check if Images are taken for evidence
                    If System.IO.Directory.GetFiles(Link.EvidenceImagePath, MainForm.CustOrd.INDEX_NO & "*").Count = 0 Then
                        WMsg.Message = "Take pictures of Packing Box contents!"
                        WMsg.ShowDialog()
                    End If

                    'Set Actual finish date
                    Dim FinishDate As String = Date.Today.ToString("yyyy-MM-dd")
                    Dim FinishTime As String = DateAndTime.Now.ToString("hh:mm:ss tt")
                    Dim SqlQry(0) As String
                    SqlQry(0) = "UPDATE cust_ord SET ACTUAL_FINISH_DATE='" & FinishDate & "', ACTUAL_FINISH_TIME='" & FinishTime & "' WHERE INDEX_NO='" & MainForm.CustOrd.INDEX_NO & "';"
                    TmlEntityYGS.ExecuteTransactionQuery(SqlQry, ErrMsg)
                    If ErrMsg.Length > 0 Then
                        WMsg.Message = "Actual Finish Date Update Error:" & ErrMsg
                        WMsg.ShowDialog()
                    End If
                End If
            End If

        Next
        ReDim Preserve Sql(Count - 1)
        TmlEntityQA.ExecuteTransactionQuery(Sql, ErrMsg)
        If ErrMsg.Length > 0 Then
            MsgBox(ErrMsg, MsgBoxStyle.Critical, "Error")
        Else
            MainForm.PrintQcc_Rev1()
            MainForm.LoadIndexScan()
        End If
    End Sub

    Private Sub Button_ViewQCC_Click(sender As Object, e As EventArgs) Handles Button_ViewQCC.Click
        If MainForm.CurrentQCC_Version = "1.4" Then
            MainForm.PrintQcc_Rev2()
        Else
            MainForm.PrintQcc_Rev1()
        End If
    End Sub


End Class