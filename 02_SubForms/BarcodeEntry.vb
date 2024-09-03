Public Class BarcodeEntry

    Dim WMsg As New WarningForm
    Private Sub TextBox_Scan_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox_Scan.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                Dim ErrMsg As String = ""
                Dim TmlEntityYGS As New MFG_ENTITY.Op(MainForm.Setting.Var_03_MySql_YGSP)
                Dim TmlEntityQA As New MFG_ENTITY.Op(MainForm.Setting.Var_04_MySql_QA)
                Dim FieldName As String = ""
                Dim IndexNo = ""

                'TextBox_Scan.Text = "100003771640"

#Region "Scan Code Read"
                If TextBox_Scan.Text.ToUpper Like "*MFR:*S/N:*" Then
                    IndexNo = TextBox_Scan.Text.Substring(TextBox_Scan.Text.Length - 9, 9)
                Else
                    IndexNo = TextBox_Scan.Text
                End If
                If IndexNo.Length = 10 Then
                    FieldName = "BAR"
                ElseIf IndexNo.Length = 9 Then
                    FieldName = "SERIAL_NO"
                ElseIf IndexNo.Length = 12 Then
                    FieldName = "INDEX_NO"
                Else
                    Label_Message.Text = "Please scan SerialNo./IndexNo./BarNo. or QR Code"
                    Exit Sub
                End If
#End Region

#Region "CustOrd Read"
                MainForm.CustOrd = TmlEntityYGS.GetDatabaseTableAs_Object(Of POCO_YGSP.cust_ord)(FieldName, IndexNo, FieldName, IndexNo, ErrMsg)
                If ErrMsg.Length > 0 Then
                    Label_Message.Text = ErrMsg
                    Exit Sub
                End If
                If IsDate(MainForm.CustOrd.ACTUAL_FINISH_DATE) Then
                    If MsgBox("Production already completed. Do you want to see QC Checksheet?", MsgBoxStyle.YesNoCancel) = MsgBoxResult.Yes Then
                        Dim ActFinDate As String = MainForm.CustOrd.ACTUAL_FINISH_DATE
                        Dim OneWeekBefore As String = Date.Today.AddDays(-7).ToString("yyyy-MM-dd")
                        If ActFinDate > OneWeekBefore Then
                            MainForm.PrintQcc_Rev1()
                            Exit Sub
                        Else
                            WMsg.Message = "Production Already completed. Please check QC-Checksheet from Intranter Portal."
                            WMsg.ShowDialog()
                            Exit Sub
                        End If
                    Else
                        WMsg.Message = "Production Already completed. Please check QC-Checksheet from Intranter Portal."
                        WMsg.ShowDialog()
                        Exit Sub
                    End If
                End If
                If Not (MainForm.CustOrd.MS_CODE Like "YTA[67]10-???????*") Then
                    WMsg.Message = "Selected Index_No does not belongs to YTA Model. Model:" & MainForm.CustOrd.MS_CODE
                    WMsg.ShowDialog()
                    Exit Sub
                End If
                MainForm.CoHeader = TmlEntityYGS.GetDatabaseTableAs_Object(Of POCO_YGSP.co_register)("ORDER_NO", MainForm.CustOrd.PROD_NO, "LINE_NO", MainForm.CustOrd.LINE_NO, ErrMsg:=ErrMsg)
                If ErrMsg.Length > 0 Then
                    Label_Message.Text = ErrMsg
                    Exit Sub
                End If
#End Region

#Region "Validate if All Inspection or Previous Step Inspections Already Done"

                Dim TotalInspectionSteps = TmlEntityQA.GetDatabaseTableAs_List(Of POCO_QA.yta_qcc_steps)("STEP_NO", "%", ErrMsg).OrderBy(Function(y) y.SLNO).ToList
                If ErrMsg.Length > 0 Then
                    Label_Message.Text = ErrMsg
                    Exit Sub
                End If
                If Not (MainForm.CustOrd.MS_CODE Like "YTA[67]10-???????*") Then
                    'QR Code scan restricted only to final insp. to avoid use of 2D Scanner at First station
                    TotalInspectionSteps.Remove(TotalInspectionSteps.Where(Function(x) x.STEP_NO = "20_00_01").ToList.Item(0))
                End If
                MainForm.QcSteps = TotalInspectionSteps
                MainForm.AllowedSteps = MainForm.QcSteps.OrderBy(Function(x) x.SLNO).Where(Function(x) x.STATION = My.Settings.Station).Select(Of String)(Function(x) x.STEP_NO).ToArray
                ReDim MainForm.AllCheckResult(MainForm.AllowedSteps.Length - 1)

                'Inspection Results till now in DB
                Dim InspRes = TmlEntityQA.GetDatabaseTableAs_List(Of POCO_QA.yta_qcc_v1p2)("INDEX_NO", MainForm.CustOrd.INDEX_NO, "INDEX_NO", MainForm.CustOrd.INDEX_NO, ErrMsg)
                If ErrMsg.Length > 0 Then
                    Label_Message.Text = ErrMsg
                    Exit Sub
                End If
                Dim TotalStepsInspected(0) As String
                If InspRes.Count > 0 Then
                    For Each process In InspRes
                        Dim StepNo = process.REMARK.Split("-")(0).Split(",")
                        For Each Steps In StepNo
                            TotalStepsInspected(TotalStepsInspected.Length - 1) = Steps
                            ReDim Preserve TotalStepsInspected(TotalStepsInspected.Length)
                        Next
                    Next
                    ReDim Preserve TotalStepsInspected(TotalStepsInspected.Length - 2)
                End If

                'Check if all Previous Inspections done, else stop new inspections
                Dim PreviousInspectionsDone As Boolean = True
                Dim NotDoneStep As String = ""
                Dim CurrentSteps = TotalInspectionSteps.Where(Function(x) x.STATION = My.Settings.Station).ToList.OrderBy(Function(y) y.SLNO).ToList
                Dim CurrentMin = CurrentSteps.Min(Function(x) x.SLNO)
                Dim PreviousSteps = TotalInspectionSteps.Where(Function(x) x.SLNO < CurrentMin).OrderBy(Function(y) y.SLNO).ToList
                For Each ToDoStep In PreviousSteps 'MainForm.Setting.Var_08_StepsPrevous.Split(",")
                    If Array.Find(TotalStepsInspected, Function(x) x = ToDoStep.STEP_NO) <> ToDoStep.STEP_NO Then
                        PreviousInspectionsDone = False
                        NotDoneStep = ToDoStep.STEP_NO
                        Exit For
                    End If
                Next
                If PreviousInspectionsDone = False Then
                    MsgBox("Some of the Previous Inspection steps are not completed. eg." & NotDoneStep, MsgBoxStyle.OkCancel)
                    TextBox_Scan.Text = ""
                    Exit Sub
                End If


                'All Station Inspection Step Check
                If TotalStepsInspected.Length = TotalInspectionSteps.Count Then 'If all Inspection Done
                    If MsgBox("Inspection already completed. Do you want to see QC Checksheet?", MsgBoxStyle.YesNoCancel) = MsgBoxResult.Yes Then
                        MainForm.PrintQcc_Rev1()
                        Exit Sub
                    Else
                        If MsgBox("Do you want to re-inspect the unit?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                            Exit Sub
                        Else
                            MainForm.RichTextBox_ActivityToCheck.Text = "Loading inspection check points. Wait.."
                            MainForm.Refresh()
                        End If
                    End If
                End If

                'Current Station Inspection Step Check
                Dim CurrentInspComplete As Boolean = True
                For Each InspStep In CurrentSteps
                    If Array.Find(TotalStepsInspected, Function(x) x = InspStep.STEP_NO) <> InspStep.STEP_NO Then
                        PreviousInspectionsDone = False
                        NotDoneStep = InspStep.STEP_NO
                        CurrentInspComplete = False
                        Exit For
                    End If
                Next
                If CurrentInspComplete = True Then
                    If MsgBox("Current Station Inspections are complete. Do you want to re-inspect the unit?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                        Exit Sub
                    Else
                        MainForm.RichTextBox_ActivityToCheck.Text = "Loading inspection check points. Wait.."
                        MainForm.Refresh()
                    End If
                End If

#End Region


                MainForm.QcData = TmlEntityQA.GetDatabaseTableAs_List(Of POCO_QA.yta_qcc_v1p2)("INDEX_NO", MainForm.CustOrd.INDEX_NO, "INDEX_NO", MainForm.CustOrd.INDEX_NO, ErrMsg)
                    If ErrMsg.Length > 0 Then
                        Label_Message.Text = ErrMsg
                        Exit Sub
                    End If

                    'Once reaches upto here, then the application is allowing user to do remaining inspections.
                    'So before proceeding, it will fill HiPOT, CRC & Plate status for further use.
                    'So get that and allocate to variable.

                    'Get HIPOT Result
                    If Array.Find(MainForm.AllowedSteps, Function(x) x.StartsWith("130_01_00")) = "130_01_00" Or
                        Array.Find(MainForm.AllowedSteps, Function(x) x.StartsWith("140_01_00")) = "140_01_00" Or
                        Array.Find(MainForm.AllowedSteps, Function(x) x.StartsWith("180_01_00")) = "180_01_00" Then

                        Dim Hipots = TmlEntityYGS.GetDatabaseTableAs_List(Of POCO_YGSP.hipot_tb)("index_no", MainForm.CustOrd.INDEX_NO, "index_no", MainForm.CustOrd.INDEX_NO, ErrMsg)
                        If Hipots.Count > 0 Then
                            MainForm.Hipot = Hipots.Where(Function(x) x.rec_no = (Hipots.Max(Function(y) y.rec_no))).FirstOrDefault
                            If ErrMsg.Length > 0 Then
                                Label_Message.Text = ErrMsg
                                Exit Sub
                            End If
                        Else
                            MainForm.Hipot = New POCO_YGSP.hipot_tb
                            MainForm.Hipot.acw_test_result = "NA"
                            MainForm.Hipot.ir_test_result = "NA"
                        End If
                    End If

                    'Get CRC Result
                    If Array.Find(MainForm.AllowedSteps, Function(x) x.StartsWith("150_01_00")) = "150_01_00" Or
                        Array.Find(MainForm.AllowedSteps, Function(x) x.StartsWith("180_01_00")) = "180_01_00" Then
                        Dim YTA_CrcList = TmlEntityYGS.GetDatabaseTableAs_List(Of POCO_YGSP.yta710_inspection_tb)("SERIAL", MainForm.CustOrd.SERIAL_NO, "SERIAL", MainForm.CustOrd.SERIAL_NO, ErrMsg)
                        If ErrMsg.Length > 0 Then
                            Label_Message.Text = ErrMsg
                            Exit Sub
                        Else
                            If YTA_CrcList.Count > 0 Then
                                MainForm.YTA_Crc = YTA_CrcList.Where(Function(x) x.ID = (YTA_CrcList.Max(Function(y) y.ID))).FirstOrDefault
                            Else
                                MainForm.YTA_Crc = New POCO_YGSP.yta710_inspection_tb
                                MainForm.YTA_Crc.RESULT = "NA"
                            End If
                        End If
                    End If

                    'Get PLATE PART NUMBER
                    If Array.Find(MainForm.AllowedSteps, Function(x) x.StartsWith("40_01_00")) = "40_01_00" Or Array.Find(MainForm.AllowedSteps, Function(x) x.StartsWith("190_11_00")) = "190_11_00" Then
                        Call SelectYtaPlates(MainForm.CustOrd, ErrMsg)
                        If ErrMsg.Length > 0 Then
                            Label_Message.Text = ErrMsg
                            Exit Sub
                        End If
                    End If

                    MainForm.Text = MainForm.Text.Split("-")(0) & "-" & " User: " & MainForm.Initial & ", Current Unit: " & MainForm.CustOrd.SERIAL_NO & " [ " & MainForm.CustOrd.MS_CODE & " ]"
                    MainForm.TextBox_Step.Text = MainForm.Setting.Var_08_StepsCurrent.Split(",")(0)
                    MainForm.FirstCheckPoint()
                    TextBox_Scan.Text = ""
                    Me.Close()

                End If
        Catch ex As Exception
            Label_Message.Text = "Runtime Error:" & ex.Message.Substring(0, 50) & ".."
        End Try
    End Sub
    Private Sub SelectYtaPlates(ByVal CustOrd As POCO_YGSP.cust_ord, ByRef ErrMsg As String)
        Try
            ErrMsg = ""
            Dim PlatePartNo As String = ""
            Dim Inst_Lib As New TML_Library.Instrument
            Dim Factory = TML_Library.Instrument.FACTORY._YKSA_F_NET
            If MainForm.Setting.Var_05_Factory = "1" Then Factory = TML_Library.Instrument.FACTORY._YUAE_M_NET
            If MainForm.Setting.Var_05_Factory = "2" Then Factory = TML_Library.Instrument.FACTORY._YKSA_F_NET
            If MainForm.Setting.Var_05_Factory = "3" Then Factory = TML_Library.Instrument.FACTORY._YUAE_F_NET
            If MainForm.Setting.Var_05_Factory = "4" Then Factory = TML_Library.Instrument.FACTORY._YKSA_F_NET
            If MainForm.Setting.Var_05_Factory = "5" Then Factory = TML_Library.Instrument.FACTORY._YUAE_F_NET

            PlatePartNo = Inst_Lib.GetNamePlatePartNumber(CustOrd.MS_CODE, MainForm.Setting.Var_05_Factory, CustOrd.INDEX_NO)
            Dim allParts() As String
            allParts = Inst_Lib.GetYTA_AllPlatePartNumberList(CustOrd.MS_CODE, Factory, ErrMsg)
            If Len(ErrMsg) > 0 Then
                ErrMsg = "Error in reading all plate part numbers:" & ErrMsg
                Exit Sub
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
                MainForm.DataPlateCheck = SelParts
                MainForm.DataPlateCorrect = PlatePartNo
            End If
        Catch ex As Exception
            ErrMsg = ex.Message
        End Try
    End Sub
End Class