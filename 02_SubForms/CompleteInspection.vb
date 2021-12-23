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
            Dim SavedLine = TmlEntityQA.GetDatabaseTableAs_Object(Of POCO_QA.yta_qcc_v1p2)("INDEX_NO", MainForm.CustOrd.INDEX_NO, "PROCESS_NO", ProcessSave.ProcessNo, ErrMsg)
            If ErrMsg.Length > 0 Then
                MsgBox(ErrMsg, MsgBoxStyle.Critical, "Error")
                Exit Sub
            End If
            If Not IsNothing(SavedLine.INDEX_NO) Then
                Sql(Count) = TmlEntityQA.SetDatabaseModel_Sql(Tbl, "INDEX_NO", MainForm.CustOrd.INDEX_NO, "PROCESS_NO", ProcessSave.ProcessNo, ErrMsg)
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
        Try
            Dim ErrMsg As String = ""
            MainForm.QcData = TmlEntityQA.GetDatabaseTableAs_List(Of POCO_QA.yta_qcc_v1p2)("INDEX_NO", MainForm.CustOrd.INDEX_NO, "INDEX_NO", MainForm.CustOrd.INDEX_NO, ErrMsg)
            If MainForm.QcData.Count > 0 Then
                Dim BlankDoc As String = MainForm.Setting.Var_06_DocsStore & "Production Release Documents\QC Check Sheets\" & MainForm.CustOrd.PROD_NO & "\Line-" & MainForm.CustOrd.LINE_NO & "-(Qty " & MainForm.CustOrd.TOT_QTY & " Pcs)\" & MainForm.CustOrd.INDEX_NO & "-QCSHEET.pdf"
                Dim FinalDoc As String = MainForm.Setting.Var_06_DocsStore & "Production Complete Documents\Signed_QCC\" & MainForm.CustOrd.PROD_NO & "\Line-" & MainForm.CustOrd.LINE_NO & "\" & MainForm.CustOrd.INDEX_NO & "-QCS-Signed.pdf"
                Dim UseDoc As New PdfSharp.Pdf.PdfDocument
                'If System.IO.File.Exists(FinalDoc) Then
                '    For Each QcWrite In MainForm.QcData
                '        Dim WriteParams = QcWrite.CHECK_RESULT.Split("$")
                '        For Each WriteParam In WriteParams
                '            Dim WriteInput = WriteParam.Split("-")(0)
                '            Dim WriteSXY = WriteParam.Split("-")(1)
                '            Dim WriteS = WriteSXY.Split(",")(0)
                '            Dim WriteX = WriteSXY.Split(",")(1)
                '            Dim WriteY = WriteSXY.Split(",")(2)
                '            Dim TextToWrite As String = WriteInput
                '            Dim FontToWrite As OpenPdfOperation.FileOp.FontName = OpenPdfOperation.FileOp.FontName.Arial
                '            If WriteInput = "Tick" Then
                '                TextToWrite = "ü"
                '                FontToWrite = OpenPdfOperation.FileOp.FontName.Wingdings
                '            End If
                '            Dim WP As New OpenPdfOperation.WriteTextParameters With {
                '                        .Name = FontToWrite,
                '                        .Colour = OpenPdfOperation.FileOp.FontColor.Black,
                '                        .Style = OpenPdfOperation.FileOp.FontStyle.Regular,
                '                        .TextSize = Integer.Parse(WriteS),
                '                        .TextValue = TextToWrite,
                '                        .X_Position = CDbl(WriteX / 100),
                '                        .Y_Position = CDbl(WriteY / 100)}
                '            Dim WPS As New List(Of OpenPdfOperation.WriteTextParameters)
                '            WPS.Add(WP)
                '            OpenPdfOperation.FileOp.PDF_WriteText(FinalDoc, FinalDoc, WPS, ErrMsg)
                '            If ErrMsg.Length > 0 Then
                '                MsgBox(ErrMsg)
                '                Exit Sub
                '            End If
                '        Next
                '        Process.Start(FinalDoc)
                '    Next
                'ElseIf System.IO.File.Exists(BlankDoc) Then
                Dim P_Doc = OpenPdfOperation.FileOp.GetDocument(BlankDoc, ErrMsg)
                    For Each QcWrite In MainForm.QcData
                        Dim WriteParams = QcWrite.CHECK_RESULT.Split("$")
                        For Each WriteParam In WriteParams
                            Dim WriteInput = WriteParam.Split("-")(0)
                            Dim WriteSXY = WriteParam.Split("-")(1)
                            Dim WriteS = WriteSXY.Split(",")(0)
                            Dim WriteX = WriteSXY.Split(",")(1)
                            Dim WriteY = WriteSXY.Split(",")(2)
                            Dim TextToWrite As String = WriteInput
                            Dim FontToWrite As OpenPdfOperation.FileOp.FontName = OpenPdfOperation.FileOp.FontName.Arial
                        If WriteInput = "Tick" Then
                            TextToWrite = "ü"
                            FontToWrite = OpenPdfOperation.FileOp.FontName.Wingdings
                        ElseIf WriteInput = "Circle" Then
                            TextToWrite = "¡"
                            FontToWrite = OpenPdfOperation.FileOp.FontName.Wingdings
                        ElseIf WriteInput = "Cross" Then
                            TextToWrite = "û"
                            FontToWrite = OpenPdfOperation.FileOp.FontName.Wingdings
                        End If
                        Dim WP As New OpenPdfOperation.WriteTextParameters With {
                                        .Name = FontToWrite,
                                        .Colour = OpenPdfOperation.FileOp.FontColor.Blue,
                                        .Style = OpenPdfOperation.FileOp.FontStyle.Regular,
                                        .TextSize = Integer.Parse(WriteS),
                                        .TextValue = TextToWrite,
                                        .X_Position = CDbl(WriteX / 100),
                                        .Y_Position = CDbl(WriteY / 100)}
                        Dim WPS As New List(Of OpenPdfOperation.WriteTextParameters)
                            WPS.Add(WP)
                            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(FinalDoc))
                            OpenPdfOperation.FileOp.PDF_WriteText(P_Doc, WPS, ErrMsg)
                            If ErrMsg.Length > 0 Then
                                MsgBox(ErrMsg)
                                Exit Sub
                            End If
                        Next
                        P_Doc.Save(FinalDoc)
                        Process.Start(FinalDoc)
                    Next
                'Else
                '    MsgBox("QC Template Not created in TML System. Please create it first.")
                '    Exit Sub
                'End If



            End If

        Catch ex As Exception

        End Try
    End Sub
End Class