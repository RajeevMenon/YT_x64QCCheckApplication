Public Class YTA_CheckSheet

#Region "Common"

    Public Shared Function ProcessStepNo(ByVal StepNo As Integer, ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep

            'If StepNo = 11 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo11(Initial, CustOrd, ErrMsg)

            If StepNo = 21 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo21(Initial, CustOrd, ErrMsg)

            If StepNo = 31 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo31(Initial, CustOrd, ErrMsg)
            If StepNo = 32 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo32(Initial, CustOrd, ErrMsg)
            If StepNo = 33 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo33(Initial, CustOrd, ErrMsg)
            If StepNo = 34 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo34(Initial, CustOrd, ErrMsg)

            If StepNo = 41 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo41(Initial, CustOrd, ErrMsg)
            If StepNo = 42 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo42(Initial, CustOrd, ErrMsg)

            If StepNo = 51 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo51(Initial, CustOrd, ErrMsg)

            If StepNo = 61 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo61(Initial, CustOrd, ErrMsg)
            If StepNo = 62 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo62(Initial, CustOrd, ErrMsg)

            If StepNo = 71 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo71(Initial, CustOrd, ErrMsg)
            If StepNo = 72 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo72(Initial, CustOrd, ErrMsg)

            If StepNo = 81 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo81(Initial, CustOrd, ErrMsg)
            If StepNo = 91 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo91(Initial, CustOrd, ErrMsg)
            If StepNo = 101 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo101(Initial, CustOrd, ErrMsg)

            If StepNo = 111 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo111(Initial, CustOrd, ErrMsg)
            If StepNo = 112 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo112(Initial, CustOrd, ErrMsg)

            If StepNo = 121 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo121(Initial, CustOrd, ErrMsg)
            If StepNo = 122 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo122(Initial, CustOrd, ErrMsg)
            If StepNo = 123 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo123(Initial, CustOrd, ErrMsg)
            If StepNo = 124 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo124(Initial, CustOrd, ErrMsg)
            If StepNo = 125 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo125(Initial, CustOrd, ErrMsg)

            If StepNo = 161 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo161(Initial, CustOrd, ErrMsg)
            If StepNo = 171 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo171(Initial, CustOrd, ErrMsg)
            If StepNo = 172 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo172(Initial, CustOrd, ErrMsg)

            If StepNo = 181 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo181(Initial, CustOrd, ErrMsg)

            MainForm.SetInspectionColor(StepNo.ToString, ProcessStepReturn.ProcessNo)

            Return ProcessStepReturn
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
#End Region

#Region "Version 1.2"

    'Public Function ProcessStepNo11(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
    '    Try
    '        Dim ProcessStepReturn As New CheckSheetStep
    '        ProcessStepReturn.ProcessNo = "10"
    '        ProcessStepReturn.ProcessStep = "Line Receiving Inspection "
    '        ProcessStepReturn.Activity = ""
    '        ProcessStepReturn.ToCheck = ""
    '        ProcessStepReturn.Method = CheckSheetStep.MethodOption.DocumentCheck
    '        ProcessStepReturn.Initial = Initial

    '        ProcessStepReturn.StepNo = "11"
    '        ProcessStepReturn.StepNo_Group = "11"
    '        ProcessStepReturn.ActivityToCheck = "Check QC Checksheet correctness."
    '        ProcessStepReturn.ViewDocAction.DocumentCheckMessage = "QCC Correct?."
    '        'ProcessStepReturn.ViewDocAction.PdfPath_DocumentCheck = MainForm.Setting.Var_06_DocsStore & "Production Release Documents\OrderTag\" &
    '        '    "" & CustOrd.PROD_NO & "\" & "Line-" & CustOrd.LINE_NO & "-(Qty " & CustOrd.TOT_QTY & " Pcs)\" & CustOrd.INDEX_NO & "-OrderTag.pdf"
    '        ProcessStepReturn.ViewDocAction.PdfPath_DocumentCheck = MainForm.Setting.Var_06_DocsStore & "Production Release Documents\QC Check Sheets\" &
    '            "" & CustOrd.PROD_NO & "\" & "Line-" & CustOrd.LINE_NO & "-(Qty " & CustOrd.TOT_QTY & " Pcs)\" & CustOrd.INDEX_NO & "-QCSHEET.pdf"
    '        Return ProcessStepReturn

    '    Catch ex As Exception
    '        Return Nothing
    '    End Try
    'End Function
    Public Function ProcessStepNo21(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try
            Dim ProcessStepReturn As New CheckSheetStep
            Dim RL_Tag As New TML_Library.RandomArray
            Dim SelPartsSNO = RL_Tag.RandomStringArraySERIAL(CustOrd.SERIAL_NO_BEFORE, 4)

            ProcessStepReturn.ProcessNo = "20"
            ProcessStepReturn.ProcessStep = "Parts correctness"
            ProcessStepReturn.Activity = "Complete unit and KD Part correctness check"
            ProcessStepReturn.ToCheck = "Part numbers and Quantity as per Bill of Material"
            ProcessStepReturn.ActivityToCheck = "Check Complete Unit Correctly Picked."
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.UserIput
            ProcessStepReturn.Initial = Initial
            'ProcessStepReturn.Result = "Tick-13,70,18$" & Initial & "-11,84,18$" & CustOrd.SERIAL_NO_BEFORE & "-10,83,9.7"
            ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_020_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            ProcessStepReturn.Result &= "$" & MainForm.Setting.Var_60_020_Position_Initial.Replace("Initial", MainForm.Initial)
            ProcessStepReturn.Result &= "$" & MainForm.Setting.Var_60_020_Position_SerialBefore.Replace("SERIAL_NO_BEFORE", MainForm.CustOrd.SERIAL_NO_BEFORE)

            ProcessStepReturn.StepNo = "21"
            ProcessStepReturn.StepNo_Group = "21"
            ProcessStepReturn.UserInputAction.UserActionMessage = "Choose the Serial number printed In the plate Of before Modificaiton unit."
            ProcessStepReturn.UserInputAction.UserInputList = SelPartsSNO
            ProcessStepReturn.UserInputAction.UserInputCorrect = CustOrd.SERIAL_NO_BEFORE
            Return ProcessStepReturn

        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo31(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try
            Dim RL_Tag As New TML_Library.RandomArray
            Dim SelPartsTAG = RL_Tag.RandomStringArrayYTA(CustOrd.TAG_NO_525, 4)

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "30"
            ProcessStepReturn.ProcessStep = "Data Plate/Tag Plate Marking"
            ProcessStepReturn.Activity = "Prepare Data and Tag Plates with correct contents"
            ProcessStepReturn.ToCheck = "Country of Origin"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.UserIput
            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.StepNo_Group = "31,32,33,34"

            ProcessStepReturn.StepNo = "31"
            ProcessStepReturn.ActivityToCheck = "Prepare Data Plates with correct COO."
            ProcessStepReturn.UserInputAction.UserActionMessage = "Choose the COO printed In the plate."
            ProcessStepReturn.UserInputAction.UserInputList = {"Made In China", "Made In Japan", "Made In KSA", "Made In Singapore"}
            If CustOrd.SERIAL_NO Like "Y3*" And CustOrd.EU_COUNTRY = "SA" Then
                ProcessStepReturn.UserInputAction.UserInputCorrect = "Made In KSA"
                ProcessStepReturn.Result = "Made In KSA-8,55,19.6$Tick-13,70,20$" & Initial & "-11,84,20"
            Else
                ProcessStepReturn.UserInputAction.UserInputCorrect = "Made In China"
                ProcessStepReturn.Result = "Made In China-8,55,19.6$Tick-15,69,21$" & Initial & "-11,84,20"
            End If
            Return ProcessStepReturn
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo32(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try
            Dim RL_Tag As New TML_Library.RandomArray
            Dim SelPartsSNO = RL_Tag.RandomStringArraySERIAL(CustOrd.SERIAL_NO, 4)

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "30"
            ProcessStepReturn.ProcessStep = "Data Plate/Tag Plate Marking"
            ProcessStepReturn.Activity = "Prepare Data and Tag Plates with correct contents"
            ProcessStepReturn.ToCheck = "Plate data correct?  YES NO"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.UserIput
            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.StepNo_Group = "31,32,33,34"

            ProcessStepReturn.StepNo = "32"
            ProcessStepReturn.ActivityToCheck = "Date Plates With correct contents."
            ProcessStepReturn.UserInputAction.UserActionMessage = "Choose the Serial number printed In the plate And Click Select button."
            ProcessStepReturn.UserInputAction.UserInputList = SelPartsSNO
            ProcessStepReturn.UserInputAction.UserInputCorrect = CustOrd.SERIAL_NO
            ProcessStepReturn.Result = ""
            Return ProcessStepReturn
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo33(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try
            Dim RL_Tag As New TML_Library.RandomArray
            Dim SelPartsTAG = RL_Tag.RandomStringArrayYTA(CustOrd.TAG_NO_525, 4)

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "30"
            ProcessStepReturn.ProcessStep = "Tag Plate Marking"
            ProcessStepReturn.Activity = "Prepare Data and Tag Plates with correct contents"
            ProcessStepReturn.ToCheck = "Plate data correct?  YES NO"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.UserIput
            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.StepNo_Group = "31,32,33,34"

            ProcessStepReturn.StepNo = "33"
            ProcessStepReturn.ActivityToCheck = "Tag Plates With correct contents."
            ProcessStepReturn.UserInputAction.UserActionMessage = "Choose the Tag number printed In the plate And Click Select button."
            ProcessStepReturn.UserInputAction.UserInputList = SelPartsTAG
            ProcessStepReturn.UserInputAction.UserInputCorrect = CustOrd.TAG_NO_525
            ProcessStepReturn.Result = ""
            Return ProcessStepReturn
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo34(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "30"
            ProcessStepReturn.ProcessStep = "Data Plate Marking"
            ProcessStepReturn.Activity = "Prepare Data and Tag Plates with correct contents"
            ProcessStepReturn.ToCheck = "Plate data correct?  YES NO"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.StepNo_Group = "31,32,33,34"

            ProcessStepReturn.StepNo = "34"
            ProcessStepReturn.ActivityToCheck = "RoHS Confirmatatory Marking"
            ProcessStepReturn.SinglePointAction.SPI_Message = "The Plate has which Of the following [with CE Or without CE]. This Is sample picture, so don't check other contents."
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\30\" & "NamePlate_with_CE.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\30\" & "NamePlate_without_CE.jpg"
            If CustOrd.SERIAL_NO_BEFORE Like "S5WC*" Or CustOrd.SERIAL_NO_BEFORE Like "S5X1*" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "NamePlate_without_CE.jpg"
                ProcessStepReturn.Result = "Yes without CE-8,57,20.6"
            Else
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "NamePlate_with_CE.jpg"
                ProcessStepReturn.Result = "Yes with CE-8,57,20.6"
            End If

            Return ProcessStepReturn
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo41(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim PlatePartNo As String = ""
            Dim Inst_Lib As New TML_Library.Instrument
            PlatePartNo = Inst_Lib.GetNamePlatePartNumber(CustOrd.MS_CODE, MainForm.Setting.Var_05_Factory, CustOrd.INDEX_NO)
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
                ProcessStepReturn.ProcessNo = "40"
                ProcessStepReturn.ProcessStep = "Data Plate/Tag Plate Mounting"
                ProcessStepReturn.Activity = "Mount the marked Data and Tag Plates to the unit"
                ProcessStepReturn.ToCheck = "Approval type Data plate Part No. and No gap between plate and housing"
                ProcessStepReturn.Method = CheckSheetStep.MethodOption.UserIput
                ProcessStepReturn.Initial = Initial
                ProcessStepReturn.StepNo_Group = "41,42"

                ProcessStepReturn.StepNo = "41"
                ProcessStepReturn.ActivityToCheck = "Mount the marked Data and Tag Plates to the unit"
                ProcessStepReturn.UserInputAction.UserActionMessage = "Choose the Plate Part number from below list and Click SELECT button."
                ProcessStepReturn.UserInputAction.UserInputList = SelParts
                ProcessStepReturn.UserInputAction.UserInputCorrect = PlatePartNo
                ProcessStepReturn.Result = PlatePartNo & "-8,62,21.8"
                Return ProcessStepReturn

            End If

        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo42(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "40"
            ProcessStepReturn.ProcessStep = "Data Plate/Tag Plate Mounting"
            ProcessStepReturn.Activity = "Mount the marked Data and Tag Plates to the unit"
            ProcessStepReturn.ToCheck = "Approval type Data plate Part No. and No gap between plate and housing"
            ProcessStepReturn.StepNo_Group = "41,42"

            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.StepNo = "42"
            ProcessStepReturn.ActivityToCheck = "Mount the marked Data and Tag Plates to the unit"
            ProcessStepReturn.SinglePointAction.SPI_Message = "[Screw Tightness/Appearance]Is the Plate fixed correctly with No space or raised edges."
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\40\" & "Nameplate_with_gaps.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\40\" & "NamePlate_Edge_Correct.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "NamePlate_Edge_Correct.jpg"
            ProcessStepReturn.Result = "OK-8,64,22.8$Tick-13,70,22$" & Initial & "-11,84,22"
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo51(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "50"
            ProcessStepReturn.ProcessStep = "Modification / Assembly Checks"
            ProcessStepReturn.Activity = "Display, Terminal Cover Open"
            ProcessStepReturn.ToCheck = "No Thread damage"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.StepNo_Group = "51"

            ProcessStepReturn.StepNo = "51"
            ProcessStepReturn.ActivityToCheck = "Display, Terminal Cover Open"
            ProcessStepReturn.SinglePointAction.SPI_Message = "[Cover open/Appearance] No Thread damage."
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\50\" & "DamagedCovers.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\50\" & "UnDamagedCovers.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "UnDamagedCovers.jpg"
            ProcessStepReturn.Result = "[GO]-8,65,24.6$Tick-13,70,24$" & Initial & "-11,84,30"
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo61(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "60"
            ProcessStepReturn.ProcessStep = "Modification / Assembly Checks"
            ProcessStepReturn.Activity = "BOUT Switch setting/Plug for electrical connection"
            ProcessStepReturn.ToCheck = "/C1  /C2  /C3  NA / Plug Installed"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.StepNo_Group = "61,62"
            ProcessStepReturn.StepNo = "61"

            If ((CustOrd.MS_CODE Like "YTA*/C[123]*") And (Not CustOrd.MS_CODE_BEFORE Like "YTA*/C[123]*")) _
                Or ((Not CustOrd.MS_CODE Like "YTA*/C[123]*") And (CustOrd.MS_CODE_BEFORE Like "YTA*/C[123]*")) Then
                ProcessStepReturn.ActivityToCheck = "BOUT Switch setting"
                ProcessStepReturn.SinglePointAction.SPI_Message = "[Burn-Out Switch] Confirm Switch [1] Position on Switch Block - SW1"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\50\" & "BOUT_C1_C2.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\50\" & "BOUT_C3_NORMAL.jpg"
                If CustOrd.MS_CODE Like "YTA*/C[12]*" Then
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "BOUT_C1_C2.jpg"
                    If CustOrd.MS_CODE Like "YTA*/C1*" Then ProcessStepReturn.Result = "Circle-16,50,25.8"
                    If CustOrd.MS_CODE Like "YTA*/C2*" Then ProcessStepReturn.Result = "Circle-16,54.1,25.8"
                Else
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "BOUT_C3_NORMAL.jpg"
                    ProcessStepReturn.Result = "Circle-16,58.5,25.8"
                End If
            Else
                If CustOrd.MS_CODE Like "YTA*/C[12]*" Then
                    ProcessStepReturn.ActivityToCheck = "BOUT Switch setting"
                    ProcessStepReturn.SinglePointAction.SPI_Message = "[Burn-Out Switch] Confirm Switch [1] Position on Switch Block - SW1"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\50\" & "BOUT_C1_C2.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\50\" & "BOUT_C3_NORMAL.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "BOUT_C1_C2.jpg"
                    If CustOrd.MS_CODE Like "YTA*/C1*" Then ProcessStepReturn.Result = "Circle-16,50,25.8"
                    If CustOrd.MS_CODE Like "YTA*/C2*" Then ProcessStepReturn.Result = "Circle-16,54.1,25.8"
                ElseIf CustOrd.MS_CODE Like "YTA*/C3*" Then
                    ProcessStepReturn.ActivityToCheck = "BOUT Switch setting"
                    ProcessStepReturn.SinglePointAction.SPI_Message = "[Burn-Out Switch] Confirm Switch [1] Position on Switch Block - SW1"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\50\" & "BOUT_C1_C2.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\50\" & "BOUT_C3_NORMAL.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "BOUT_C3_NORMAL.jpg"
                    ProcessStepReturn.Result = "Circle-16,58.5,25.8"
                Else
                    ProcessStepReturn.ActivityToCheck = "BOUT Switch setting"
                    ProcessStepReturn.SinglePointAction.SPI_Message = "[Burn-Out Switch] Confirm Switch [1] Position on Switch Block - SW1"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\50\" & "BOUT_C1_C2.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "No Change.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "No Change.jpg"
                    ProcessStepReturn.Result = "Circle-16,62.5,25.8"
                End If
            End If

            ProcessStepReturn.Result &= "$Tick-13,70,26"
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo62(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            If CustOrd.MS_CODE Like "YTA*" Then
                ProcessStepReturn.ProcessNo = "60"
                ProcessStepReturn.ProcessStep = "Modification / Assembly Checks"
                ProcessStepReturn.Activity = "BOUT Switch setting/Plug for electrical connection"
                ProcessStepReturn.ToCheck = "/C1  /C2  /C3  NA / Plug Installed"
                ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
                ProcessStepReturn.Initial = Initial
                ProcessStepReturn.StepNo_Group = "61,62"

                ProcessStepReturn.StepNo = "62"
                ProcessStepReturn.ActivityToCheck = "Plug for Electrical connection"
                ProcessStepReturn.SinglePointAction.SPI_Message = "[Plug] Confirm Red Plug on Electrical Connection"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Installed.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Not Installed.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Installed.jpg"
                ProcessStepReturn.Result &= "Tick-13,70,27.6" '"OK-8,65,26.6$Tick-13,70,26"
            End If
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo71(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            If CustOrd.MS_CODE Like "YTA*" Then
                ProcessStepReturn.ProcessNo = "70"
                ProcessStepReturn.ProcessStep = "Modification / Assembly Checks"
                ProcessStepReturn.Activity = "Display Indicator"
                ProcessStepReturn.ToCheck = "ADDED   REMOVED    NO-CHANGE"
                ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
                ProcessStepReturn.Initial = Initial
                ProcessStepReturn.StepNo_Group = "71,72"

                ProcessStepReturn.StepNo = "71"
                ProcessStepReturn.ActivityToCheck = "Display Indicator"
                If CustOrd.MS_CODE_BEFORE Like "YTA[67]10-?????D?*" And CustOrd.MS_CODE Like "YTA[67]10-?????D?*" Then
                    ProcessStepReturn.SinglePointAction.SPI_Message = "[Indicator] Confirm Indicator Status?"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "No Change.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Not Installed.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "No Change.jpg"
                    ProcessStepReturn.Result = "Circle-16,62,28.8"
                ElseIf CustOrd.MS_CODE_BEFORE Like "YTA[67]10-?????D?*" And CustOrd.MS_CODE Like "YTA[67]10-?????N?*" Then
                    ProcessStepReturn.SinglePointAction.SPI_Message = "[Indicator] Confirm Indicator Added or Removed?"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Added.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Removed.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Removed.jpg"
                    ProcessStepReturn.Result = "Circle-16,54,28.8"
                ElseIf CustOrd.MS_CODE_BEFORE Like "YTA[67]10-?????N?*" And CustOrd.MS_CODE Like "YTA[67]10-?????D?*" Then
                    ProcessStepReturn.SinglePointAction.SPI_Message = "[Indicator] Confirm Indicator Added or Removed?"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Removed.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Added.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Installed.jpg"
                    ProcessStepReturn.Result = "Circle-16,47.5,28.8"
                End If

            End If
            ProcessStepReturn.Result &= "$Tick-13,70,29.1"
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo72(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            If CustOrd.MS_CODE Like "YTA*" Then
                ProcessStepReturn.ProcessNo = "70"
                ProcessStepReturn.ProcessStep = "Modification / Assembly Checks"
                ProcessStepReturn.Activity = "Display Indicator"
                ProcessStepReturn.ToCheck = "ADDED   REMOVED    NO-CHANGE"
                ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
                ProcessStepReturn.Initial = Initial
                ProcessStepReturn.StepNo_Group = "71,72"

                ProcessStepReturn.StepNo = "72"
                ProcessStepReturn.ActivityToCheck = "Display Indicator"
                If CustOrd.MS_CODE_BEFORE Like "YTA[67]10-?????D?*" And CustOrd.MS_CODE Like "YTA[67]10-?????D?*" Then
                    ProcessStepReturn.SinglePointAction.SPI_Message = "[Nut/Stud] Confirm Nut/Stud usage for Indicator/Board fixing to case"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\70\" & "Indicator_W_Nut.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\70\" & "NoneInd_w_Stud.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Indicator_W_Nut.jpg"
                ElseIf CustOrd.MS_CODE_BEFORE Like "YTA[67]10-?????D?*" And CustOrd.MS_CODE Like "YTA[67]10-?????N?*" Then
                    ProcessStepReturn.SinglePointAction.SPI_Message = "[Nut/Stud] Confirm Nut/Stud usage for Indicator/Board fixing to case"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\70\" & "Indicator_W_Nut.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\70\" & "NoneInd_w_Stud.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "NoneInd_w_Stud.jpg"
                ElseIf CustOrd.MS_CODE_BEFORE Like "YTA[67]10-?????N?*" And CustOrd.MS_CODE Like "YTA[67]10-?????D?*" Then
                    ProcessStepReturn.SinglePointAction.SPI_Message = "[Nut/Stud] Confirm Nut/Stud usage for Indicator/Board fixing to case"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\70\" & "Indicator_W_Nut.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\70\" & "NoneInd_w_Stud.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Indicator_W_Nut.jpg"
                End If

            End If
            ProcessStepReturn.Result = ""
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo81(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "80"
            ProcessStepReturn.ProcessStep = "Modification / Assembly Checks"
            ProcessStepReturn.Activity = "Lock Screw installation"
            ProcessStepReturn.ToCheck = "YES    NO"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.StepNo_Group = "81"

            ProcessStepReturn.StepNo = "81"
            ProcessStepReturn.ActivityToCheck = "Lock Screw installation"
            If CustOrd.MS_CODE Like "YTA[67]10-?????D?*/[KSF][UF][12]*" Then
                ProcessStepReturn.SinglePointAction.SPI_Message = "[Lock Screw] Confirm Installation of Lock Screw [F9900RP] ?"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\80\" & "Indicator_w_F9900RP.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\80\" & "Indicator_wo_F9900RP.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Indicator_w_F9900RP.jpg"
                ProcessStepReturn.Result = "Circle-16,51,31.8"
            ElseIf CustOrd.MS_CODE Like "YTA[67]10-?????D?*" Then
                ProcessStepReturn.SinglePointAction.SPI_Message = "[Lock Screw] Confirm Installation of Lock Screw [F9900RP] ?"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\80\" & "Indicator_w_F9900RP.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\80\" & "Indicator_wo_F9900RP.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Indicator_wo_F9900RP.jpg"
                ProcessStepReturn.Result = "Circle-16,61.5,31.8"
            ElseIf CustOrd.MS_CODE Like "YTA[67]10-?????N?*/[KSF][UF][12]*" Then
                ProcessStepReturn.SinglePointAction.SPI_Message = "[Lock Screw] Confirm Installation of Lock Screw [F9900RP] ?"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\80\" & "Blind_w_F9900RP.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\80\" & "Blind_wo_F9900RP.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Blind_w_F9900RP.jpg"
                ProcessStepReturn.Result = "Circle-16,51,31.8"
            Else
                ProcessStepReturn.SinglePointAction.SPI_Message = "[Lock Screw] Confirm Installation of Lock Screw [F9900RP] ?"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\80\" & "Blind_w_F9900RP.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\80\" & "Blind_wo_F9900RP.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Blind_wo_F9900RP.jpg"
                ProcessStepReturn.Result = "Circle-16,61.5,31.8"
            End If
            ProcessStepReturn.Result &= "$Tick-13,70,32.1"
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo91(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            If CustOrd.MS_CODE Like "YTA[67]10-?????D?*" Then
                ProcessStepReturn.ProcessNo = "90"
                ProcessStepReturn.ProcessStep = "Modification / Assembly Checks"
                ProcessStepReturn.Activity = "Display Cover Assembling"
                ProcessStepReturn.ToCheck = "Clean & No galling"
                ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
                ProcessStepReturn.Initial = Initial
                ProcessStepReturn.StepNo_Group = "91"

                ProcessStepReturn.StepNo = "91"
                ProcessStepReturn.ActivityToCheck = "Display Cover Assembling"
                ProcessStepReturn.SinglePointAction.SPI_Message = "[Indicator] Confirm Indicator Cover Quality after Assembling?"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\50\" & "No Galling.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\50\" & "Galling.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "No Galling.jpg"
            End If
            ProcessStepReturn.Result = "Tick-13,70,30.6"
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo101(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            If CustOrd.MS_CODE Like "YTA*" Then
                ProcessStepReturn.ProcessNo = "100"
                ProcessStepReturn.ProcessStep = "Modification / Assembly Checks"
                ProcessStepReturn.Activity = "Lightning Protector"
                ProcessStepReturn.ToCheck = "ADDED    REMOVED    NO-CHANGE"
                ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
                ProcessStepReturn.Initial = Initial
                ProcessStepReturn.StepNo_Group = "101"

                ProcessStepReturn.StepNo = "101"
                ProcessStepReturn.ActivityToCheck = "[Screw Tighten/Appearance] Lightning Arrestor Check"
                If Not CustOrd.MS_CODE_BEFORE Like "YTA[67]10-???????*/A*" And CustOrd.MS_CODE Like "YTA[67]10-???????*/A*" Then
                    ProcessStepReturn.SinglePointAction.SPI_Message = "[/A Option] Lightning Arrestor Installed?"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Installed.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Not Installed.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Installed.jpg"
                    ProcessStepReturn.Result = "Circle-16,47.4,33.3"
                ElseIf CustOrd.MS_CODE_BEFORE Like "YTA[67]10-???????*/A*" And Not CustOrd.MS_CODE Like "YTA[67]10-???????*/A*" Then
                    ProcessStepReturn.SinglePointAction.SPI_Message = "[/A Option] Confirm Lightning Arrestor Status?"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Added.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Removed.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Removed.jpg"
                    ProcessStepReturn.Result = "Circle-16,55,33.3"
                ElseIf Not CustOrd.MS_CODE_BEFORE Like "YTA[67]10-???????*/A*" And Not CustOrd.MS_CODE Like "YTA[67]10-???????*/A*" Then
                    ProcessStepReturn.SinglePointAction.SPI_Message = "[/A Option] Confirm Lightning Arrestor Status?"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "No Change.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Added.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "No Change.jpg"
                    ProcessStepReturn.Result = "Circle-16,63,33.3"
                End If

            End If
            ProcessStepReturn.Result &= "$Tick-13,70,33.6"
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo111(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "110"
            ProcessStepReturn.ProcessStep = "Modification / Assembly Checks"
            ProcessStepReturn.Activity = "Ground Screws Fixing/DLM QR Label Fixing"
            ProcessStepReturn.ToCheck = "FIXED INTERNAL & EXTERNAL/FIXED QR LABEL AT PROPER PLACE ON CASE"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.StepNo_Group = "111,112"

            ProcessStepReturn.StepNo = "111"
            ProcessStepReturn.ActivityToCheck = "External Ground Screw"
            If CustOrd.MS_CODE Like "YTA[67]10-???????*/[KS][UFS][12]*" Then
                ProcessStepReturn.SinglePointAction.SPI_Message = "[Screw tighten/Appearance] Confirm Type of Ground Screw Installed ?"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\110\" & "ExtEarth_w_Washer.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\110\" & "ExtEarth_wo_Washer.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "ExtEarth_w_Washer.jpg"
            Else
                ProcessStepReturn.SinglePointAction.SPI_Message = "[Screw tighten/Appearance] Confirm Type of Ground Screw Installed ?"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\110\" & "ExtEarth_w_Washer.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\110\" & "ExtEarth_wo_Washer.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "ExtEarth_wo_Washer.jpg"
            End If
            ProcessStepReturn.Result = "Tick-13,70,35"
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo112(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "110"
            ProcessStepReturn.ProcessStep = "Modification / Assembly Checks"
            ProcessStepReturn.Activity = "Ground Screws Fixing/DLM QR Label Fixing"
            ProcessStepReturn.ToCheck = "FIXED INTERNAL & EXTERNAL/FIXED QR LABEL AT PROPER PLACE ON CASE"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.StepNo_Group = "111,112"

            ProcessStepReturn.StepNo = "112"
            ProcessStepReturn.ActivityToCheck = "DLM QR Label Fixing"
            ProcessStepReturn.SinglePointAction.SPI_Message = "[Affix QR Label] Confirm QR Label Attached"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\110\" & "WithOutQR_Label.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\110\" & "WithQR_Label.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "WithQR_Label.jpg"

            ProcessStepReturn.Result = "Tick-13,70,36.5"
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo121(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "120"
            ProcessStepReturn.ProcessStep = "Basic Checks"
            ProcessStepReturn.Activity = "Protocol check,Sensor Input,MOC,Electrical Connection,Display Type"
            ProcessStepReturn.ToCheck = "J/F||1/2||A/C||0/2/4||D/N"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.StepNo_Group = "121,122,123,124,125"

            ProcessStepReturn.StepNo = "121"
            ProcessStepReturn.ActivityToCheck = "Output Singal of YTA Unit"
            ProcessStepReturn.SinglePointAction.SPI_Message = "[Output Singal] Confirm Communication Protocol"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\120\" & "HART.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\120\" & "Fieldbus.jpg"
            If MainForm.CustOrd.MS_CODE Like "YTA???-J*" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "HART.jpg"
            ElseIf MainForm.CustOrd.MS_CODE Like "YTA???-F*" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Fieldbus.jpg"
            Else
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "BRAIN.jpg"
            End If

            ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_120_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1) '"Tick-13,70,36.5"
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo122(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "120"
            ProcessStepReturn.ProcessStep = "Basic Checks"
            ProcessStepReturn.Activity = "Protocol check,Sensor Input,MOC,Electrical Connection,Display Type"
            ProcessStepReturn.ToCheck = "J/F||1/2||A/C||0/2/4||D/N"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.StepNo_Group = "121,122,123,124,125"

            ProcessStepReturn.StepNo = "122"
            ProcessStepReturn.ActivityToCheck = "Number of Sensor Inputs of YTA Unit"
            ProcessStepReturn.SinglePointAction.SPI_Message = "[Inputs] Confirm the number of Sensor Inputs"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\120\" & "SingleSensor.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\120\" & "DualSensor.jpg"
            If MainForm.CustOrd.MS_CODE Like "YTA???-??1*" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "SingleSensor.jpg"
            ElseIf MainForm.CustOrd.MS_CODE Like "YTA???-??2*" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "DualSensor.jpg"
            End If

            ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_120_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1) '"Tick-13,70,36.5"
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo123(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "120"
            ProcessStepReturn.ProcessStep = "Basic Checks"
            ProcessStepReturn.Activity = "Protocol check,Sensor Input,MOC,Electrical Connection,Display Type"
            ProcessStepReturn.ToCheck = "J/F||1/2||A/C||0/2/4||D/N"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.StepNo_Group = "121,122,123,124,125"

            ProcessStepReturn.StepNo = "123"
            ProcessStepReturn.ActivityToCheck = "Housing Material of YTA Unit"
            ProcessStepReturn.SinglePointAction.SPI_Message = "[MOC] Confirm the Material of Housing and Covers"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\120\" & "Aluminum.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\120\" & "SUS.jpg"
            If MainForm.CustOrd.MS_CODE Like "YTA???-???A*" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Aluminum.jpg"
            ElseIf MainForm.CustOrd.MS_CODE Like "YTA???-???C*" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "SUS.jpg"
            End If

            ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_120_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1) '"Tick-13,70,36.5"
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo124(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "120"
            ProcessStepReturn.ProcessStep = "Basic Checks"
            ProcessStepReturn.Activity = "Protocol check,Sensor Input,MOC,Electrical Connection,Display Type"
            ProcessStepReturn.ToCheck = "J/F||1/2||A/C||0/2/4||D/N"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.StepNo_Group = "121,122,123,124,125"

            ProcessStepReturn.StepNo = "124"
            ProcessStepReturn.ActivityToCheck = "Electrical Connection Type of YTA Unit"
            ProcessStepReturn.SinglePointAction.SPI_Message = "Confirm the Electrical Connection type of Housing/Case"

            If MainForm.CustOrd.MS_CODE Like "YTA???-????2*" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\120\" & "NPT.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\120\" & "M20.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "NPT.jpg"
            ElseIf MainForm.CustOrd.MS_CODE Like "YTA???-????4*" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\120\" & "NPT.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\120\" & "M20.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "M20.jpg"
            ElseIf MainForm.CustOrd.MS_CODE Like "YTA???-????0*" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\120\" & "G12.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\120\" & "NPT.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "G12.jpg"
            End If

            ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_120_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1) '"Tick-13,70,36.5"
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo125(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "120"
            ProcessStepReturn.ProcessStep = "Basic Checks"
            ProcessStepReturn.Activity = "Protocol check,Sensor Input,MOC,Electrical Connection,Display Type"
            ProcessStepReturn.ToCheck = "J/F||1/2||A/C||0/2/4||D/N"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.StepNo_Group = "121,122,123,124,125"

            ProcessStepReturn.StepNo = "125"
            ProcessStepReturn.ActivityToCheck = "Display Type of YTA Unit"
            ProcessStepReturn.SinglePointAction.SPI_Message = "Confirm the Display Type"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\120\" & "INDICATOR.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\120\" & "NO INDICATOR.jpg"
            If MainForm.CustOrd.MS_CODE Like "YTA???-?????D*" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "INDICATOR.jpg"
            ElseIf MainForm.CustOrd.MS_CODE Like "YTA???-?????N*" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "NO INDICATOR.jpg"
            End If

            ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_120_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1) '"Tick-13,70,46.5$" & Initial & "-11,84,40"
            ProcessStepReturn.Result &= "$" & MainForm.Setting.Var_60_120_Position_Initial.Replace("Initial", MainForm.Initial)
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo161(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "160"
            ProcessStepReturn.ProcessStep = "Final Assembly"
            ProcessStepReturn.Activity = "Mount Terminal cover"
            ProcessStepReturn.ToCheck = "Mounted Terminal cover YES || NO"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.StepNo_Group = "161"

            ProcessStepReturn.StepNo = "161"
            ProcessStepReturn.ActivityToCheck = "Mount Terminal cover"
            ProcessStepReturn.SinglePointAction.SPI_Message = "Cover Tighten and Loosen 1/2 turn, then tighten complete"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Installed.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Not Installed.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Installed.jpg"
            ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_160_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1) '"Tick-13,70,46.5$" & Initial & "-11,84,40"
            ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_160_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo171(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "170"
            ProcessStepReturn.ProcessStep = "Final Assembly"
            ProcessStepReturn.Activity = "Remove or Attach Mounting Bracket"
            ProcessStepReturn.ToCheck = "Remove or Attach Brackets YES || NO"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.StepNo_Group = "171,172"

            ProcessStepReturn.StepNo = "171"
            If MainForm.CustOrd.MS_CODE_BEFORE Like "YTA???-??????N" Or (MainForm.CustOrd.MS_CODE_BEFORE.Substring(13, 1) = MainForm.CustOrd.MS_CODE.Substring(13, 1)) Then
                ProcessStepReturn.ActivityToCheck = "Bracket Assy"
                ProcessStepReturn.SinglePointAction.SPI_Message = "Confirm status of Mounting Bracket of Before modification unit?"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Removed.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "No Change.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "No Change.jpg"
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_170_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_170_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(1)
            Else
                ProcessStepReturn.ActivityToCheck = "Bracket Assy"
                ProcessStepReturn.SinglePointAction.SPI_Message = "Confirm status of Mounting Bracket of Before modification unit?"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Removed.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "No Change.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Removed.jpg"
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_170_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_170_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)
            End If
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo172(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "170"
            ProcessStepReturn.ProcessStep = "Final Assembly"
            ProcessStepReturn.Activity = "Remove or Attach Mounting Bracket"
            ProcessStepReturn.ToCheck = "Remove or Attach Brackets YES || NO"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.StepNo_Group = "171,172"

            ProcessStepReturn.StepNo = "172"
            If MainForm.CustOrd.MS_CODE_BEFORE.Substring(13, 1) = MainForm.CustOrd.MS_CODE.Substring(13, 1) Then
                ProcessStepReturn.ActivityToCheck = "Bracket Assy"
                ProcessStepReturn.SinglePointAction.SPI_Message = "Confirm status of Mounting Bracket on Final Unit?"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "No Change.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Added.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "No Change.jpg"
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_170_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & MainForm.Setting.Var_60_170_Position_Initial.Replace("Initial", MainForm.Initial)
                If MainForm.CustOrd.MS_CODE_BEFORE.Substring(13, 1) = "N" Then
                    ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_170_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(1)
                Else
                    ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_170_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)
                End If
            Else
                ProcessStepReturn.ActivityToCheck = "Bracket Assy"
                ProcessStepReturn.SinglePointAction.SPI_Message = "Confirm status of Mounting Bracket on Final Unit?"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "No Change.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Added.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Added.jpg"
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_170_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & MainForm.Setting.Var_60_170_Position_Initial.Replace("Initial", MainForm.Initial)
                If MainForm.CustOrd.MS_CODE_BEFORE.Substring(13, 1) <> "N" Then
                    ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_170_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)
                Else
                    ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_170_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(1)
                End If
            End If
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo181(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "180"
            ProcessStepReturn.ProcessStep = "Print QIC"
            ProcessStepReturn.Activity = "Print Certificates"
            ProcessStepReturn.ToCheck = "Printed YES || NO"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.DocumentCheck
            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.StepNo_Group = "181"

            ProcessStepReturn.StepNo = "181"
            ProcessStepReturn.ActivityToCheck = "Print QIC"
            ProcessStepReturn.ViewDocAction.DocumentCheckMessage = "Printed correctly YES || NO ?"
            ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_180_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            ProcessStepReturn.Result &= "$" & MainForm.Setting.Var_60_180_Position_Initial.Replace("Initial", MainForm.Initial)
            ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_180_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)


            Dim QicFolerPath As String = MainForm.Setting.Var_06_DocsStore & "Production Complete Documents\QICDOC\"
            If Not (System.IO.Directory.Exists(QicFolerPath & MainForm.CustOrd.PROD_NO)) Then
                System.IO.Directory.CreateDirectory(QicFolerPath & MainForm.CustOrd.PROD_NO)
            End If
            QicFolerPath = QicFolerPath & MainForm.CustOrd.PROD_NO & "\"
            Dim lotNo As String = MainForm.CustOrd.SHIP_LOT
            If Not (System.IO.Directory.Exists(QicFolerPath & lotNo)) Then
                System.IO.Directory.CreateDirectory(QicFolerPath & lotNo)
            End If
            QicFolerPath = QicFolerPath & lotNo & "\"

            Dim QicFile As String = MainForm.Setting.Var_06_DocsStore & "Production Complete Documents\YTA_QIC\" & MainForm.CustOrd.SERIAL_NO
            Dim TargetFile As String = QicFolerPath & "LINE-" & MainForm.CustOrd.LINE_NO & "-" & MainForm.CustOrd.INDEX_NO & "(1)-CQIC"
            If System.IO.File.Exists(QicFile & ".pdf") Then
                If Not System.IO.File.Exists(TargetFile & ".pdf") Then
                    System.IO.File.Copy(QicFile & ".pdf", TargetFile & ".pdf", True)
                End If
            End If

            ProcessStepReturn.ViewDocAction.PdfPath_DocumentCheck = TargetFile & ".pdf"

            Return ProcessStepReturn

        Catch ex As Exception
            Return Nothing
        End Try
    End Function


#End Region
End Class
