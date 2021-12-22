Public Class YTA_CheckSheet

#Region "Common"
    Public Shared Function ProcessStepNo(ByVal StepNo As Integer, ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try
            Dim ProcessStepReturn As New CheckSheetStep

            If StepNo = 11 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo11(Initial, CustOrd, ErrMsg)

            If StepNo = 21 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo21(Initial, CustOrd, ErrMsg)

            If StepNo = 31 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo31(Initial, CustOrd, ErrMsg)
            If StepNo = 32 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo32(Initial, CustOrd, ErrMsg)
            If StepNo = 33 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo33(Initial, CustOrd, ErrMsg)
            If StepNo = 34 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo34(Initial, CustOrd, ErrMsg)

            If StepNo = 41 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo41(Initial, CustOrd, ErrMsg)
            If StepNo = 42 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo42(Initial, CustOrd, ErrMsg)

            If StepNo = 51 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo51(Initial, CustOrd, ErrMsg)
            If StepNo = 52 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo52(Initial, CustOrd, ErrMsg)
            If StepNo = 53 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo53(Initial, CustOrd, ErrMsg)
            If StepNo = 54 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo54(Initial, CustOrd, ErrMsg)
            If StepNo = 55 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo55(Initial, CustOrd, ErrMsg)

            Return ProcessStepReturn
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
#End Region

#Region "Version 1.2"

    Public Function ProcessStepNo11(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try
            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.StepNo = "11"
            ProcessStepReturn.ProcessNo = "10"
            ProcessStepReturn.ProcessStep = "Line Receiving Inspection "
            ProcessStepReturn.ActivityToCheck = "Check QC Checksheet correctness."
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.DocumentCheck
            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.ViewDocAction.DocumentCheckMessage = "QCC Correct?."
            'ProcessStepReturn.ViewDocAction.PdfPath_DocumentCheck = MainForm.Setting.Var_06_DocsStore & "Production Release Documents\OrderTag\" &
            '    "" & CustOrd.PROD_NO & "\" & "Line-" & CustOrd.LINE_NO & "-(Qty " & CustOrd.TOT_QTY & " Pcs)\" & CustOrd.INDEX_NO & "-OrderTag.pdf"
            ProcessStepReturn.ViewDocAction.PdfPath_DocumentCheck = MainForm.Setting.Var_06_DocsStore & "Production Release Documents\QC Check Sheets\" &
                "" & CustOrd.PROD_NO & "\" & "Line-" & CustOrd.LINE_NO & "-(Qty " & CustOrd.TOT_QTY & " Pcs)\" & CustOrd.INDEX_NO & "-QCSHEET.pdf"
            Return ProcessStepReturn

        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo21(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try
            Dim ProcessStepReturn As New CheckSheetStep
            Dim RL_Tag As New TML_Library.RandomArray
            Dim SelPartsSNO = RL_Tag.RandomStringArraySERIAL(CustOrd.SERIAL_NO_BEFORE, 4)
            ProcessStepReturn.StepNo = "21"
            ProcessStepReturn.ProcessNo = "20"
            ProcessStepReturn.ProcessStep = "Parts correctness"
            ProcessStepReturn.ActivityToCheck = "Check Complete Unit Correctly Picked."
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.UserIput
            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.UserInputAction.UserActionMessage = "Choose the Serial number printed In the plate Of before Modificaiton unit."
            ProcessStepReturn.UserInputAction.UserInputList = SelPartsSNO
            ProcessStepReturn.UserInputAction.UserInputCorrect = CustOrd.SERIAL_NO_BEFORE
            Return ProcessStepReturn

        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    'Public Function ProcessStepNo22(ByVal Initial As String) As CheckSheetStep
    '    Try
    '        Dim ProcessStepReturn As New CheckSheetStep
    '        ProcessStepReturn.ProcessNo = "22"
    '        ProcessStepReturn.ProcessStep = "Parts correctness"
    '        ProcessStepReturn.ActivityToCheck = "Complete unit And KD Part correctness check. Part numbers And Quantity As per Bill Of Material"
    '        ProcessStepReturn.Method = CheckSheetStep.MethodOption.ProcedureStep
    '        ProcessStepReturn.Initial = Initial
    '        ProcessStepReturn.ProcedureStepAction.ImagePath_ProcedureStep = MainForm.Setting.Var_51_ProStepView_ImagePath & "YTA\20\" & "B_J_Bracket.jpg"
    '        Return ProcessStepReturn
    '    Catch ex As Exception
    '        Return Nothing
    '    End Try
    'End Function
    Public Function ProcessStepNo31(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try
            Dim RL_Tag As New TML_Library.RandomArray
            Dim SelPartsTAG = RL_Tag.RandomStringArrayYTA(CustOrd.TAG_NO_525, 4)

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.StepNo = "31"
            ProcessStepReturn.ProcessNo = "30"
            ProcessStepReturn.ProcessStep = "Tag Plate Marking"
            ProcessStepReturn.ActivityToCheck = "Tag Plates With correct contents."
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.UserIput
            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.UserInputAction.UserActionMessage = "Choose the Tag number printed In the plate And Click Select button."
            ProcessStepReturn.UserInputAction.UserInputList = SelPartsTAG
            ProcessStepReturn.UserInputAction.UserInputCorrect = CustOrd.TAG_NO_525
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
            ProcessStepReturn.StepNo = "32"
            ProcessStepReturn.ProcessNo = "30"
            ProcessStepReturn.ProcessStep = "Date Plate Marking"
            ProcessStepReturn.ActivityToCheck = "Date Plates With correct contents."
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.UserIput
            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.UserInputAction.UserActionMessage = "Choose the Serial number printed In the plate And Click Select button."
            ProcessStepReturn.UserInputAction.UserInputList = SelPartsSNO
            ProcessStepReturn.UserInputAction.UserInputCorrect = CustOrd.SERIAL_NO
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
            ProcessStepReturn.StepNo = "33"
            ProcessStepReturn.ProcessNo = "30"
            ProcessStepReturn.ProcessStep = "Data Plate Marking"
            ProcessStepReturn.ActivityToCheck = "Data Plates With correct COO."
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.UserIput
            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.UserInputAction.UserActionMessage = "Choose the COO printed In the plate."
            ProcessStepReturn.UserInputAction.UserInputList = {"Made In China", "Made In Japan", "Made In KSA", "Made In Singapore"}
            If CustOrd.SERIAL_NO Like "Y3*" And CustOrd.EU_COUNTRY = "SA" Then
                ProcessStepReturn.UserInputAction.UserInputCorrect = "Made In KSA"
            Else
                ProcessStepReturn.UserInputAction.UserInputCorrect = "Made In China"
            End If
            Return ProcessStepReturn
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo34(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.StepNo = "34"
            ProcessStepReturn.ProcessNo = "30"
            ProcessStepReturn.ProcessStep = "Data Plate Marking"
            ProcessStepReturn.ActivityToCheck = "RoHS Confirmatatory Marking"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.SinglePointAction.SPI_Message = "The Plate has which Of the following [with CE Or without CE]. This Is sample picture, so don't check other contents."
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\30\" & "NamePlate_with_CE.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\30\" & "NamePlate_without_CE.jpg"
            If CustOrd.SERIAL_NO_BEFORE Like "S5WC*" Or CustOrd.SERIAL_NO_BEFORE Like "S5X1*" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "NamePlate_without_CE.jpg"
            Else
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "NamePlate_with_CE.jpg"
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
                ProcessStepReturn.StepNo = "41"
                ProcessStepReturn.ProcessNo = "40"
                ProcessStepReturn.ProcessStep = "Data Plate Mounting"
                ProcessStepReturn.ActivityToCheck = "Mount the marked Data and Tag Plates to the unit"
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
    Public Function ProcessStepNo42(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.StepNo = "42"
            ProcessStepReturn.ProcessNo = "40"
            ProcessStepReturn.ProcessStep = "Data Plate Mounting"
            ProcessStepReturn.ActivityToCheck = "Mount the marked Data and Tag Plates to the unit"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.SinglePointAction.SPI_Message = "[Screw Tightness/Appearance]Is the Plate fixed correctly with No space or raised edges."
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\40\" & "Nameplate_with_gaps.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\40\" & "NamePlate_Edge_Correct.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "NamePlate_Edge_Correct.jpg"
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo51(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.StepNo = "51"
            ProcessStepReturn.ProcessNo = "50"
            ProcessStepReturn.ProcessStep = "Modification / Assembly Checks"
            ProcessStepReturn.ActivityToCheck = "Display, Terminal Cover Open"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.SinglePointAction.SPI_Message = "[Cover open/Appearance] No Thread damage."
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\50\" & "DamagedCovers.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\50\" & "UnDamagedCovers.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "UnDamagedCovers.jpg"
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo52(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            If CustOrd.MS_CODE Like "YTA*/C[123]*" Then
                ProcessStepReturn.StepNo = "52"
                ProcessStepReturn.ProcessNo = "50"
                ProcessStepReturn.ProcessStep = "Modification / Assembly Checks"
                ProcessStepReturn.ActivityToCheck = "BOUT Switch setting"
                ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
                ProcessStepReturn.Initial = Initial
                ProcessStepReturn.SinglePointAction.SPI_Message = "[Burn-Out Switch] Confirm Switch [1] Position on Switch Block - SW1"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\50\" & "BOUT_C1_C2.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\50\" & "BOUT_C3_NORMAL.jpg"
                If CustOrd.MS_CODE Like "YTA*/C[12]*" Then
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "BOUT_C1_C2.jpg"
                Else
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "BOUT_C3_NORMAL.jpg"
                End If
            End If
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo53(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            If CustOrd.MS_CODE Like "YTA*" Then
                ProcessStepReturn.StepNo = "53"
                ProcessStepReturn.ProcessNo = "50"
                ProcessStepReturn.ProcessStep = "Modification / Assembly Checks"
                ProcessStepReturn.ActivityToCheck = "Plug for Electrical connection"
                ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
                ProcessStepReturn.Initial = Initial
                ProcessStepReturn.SinglePointAction.SPI_Message = "[Plug] Confirm Red Plug on Electrical Connection"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\50\" & "Installed.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\50\" & "Not Installed.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Installed.jpg"
            End If
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo54(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            If CustOrd.MS_CODE Like "YTA*" Then
                ProcessStepReturn.StepNo = "54"
                ProcessStepReturn.ProcessNo = "50"
                ProcessStepReturn.ProcessStep = "Modification / Assembly Checks"
                ProcessStepReturn.ActivityToCheck = "Display Indicator"
                ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
                ProcessStepReturn.Initial = Initial
                If CustOrd.MS_CODE_BEFORE Like "YTA[67]10-?????D?*" And CustOrd.MS_CODE Like "YTA[67]10-?????D?*" Then
                    ProcessStepReturn.SinglePointAction.SPI_Message = "[Indicator] Confirm Indicator Status?"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\50\" & "Installed.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\50\" & "Not Installed.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Installed.jpg"
                ElseIf CustOrd.MS_CODE_BEFORE Like "YTA[67]10-?????D?*" And CustOrd.MS_CODE Like "YTA[67]10-?????N?*" Then
                    ProcessStepReturn.SinglePointAction.SPI_Message = "[Indicator] Confirm Indicator Added or Removed?"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\50\" & "Added.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\50\" & "Removed.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Removed.jpg"
                ElseIf CustOrd.MS_CODE_BEFORE Like "YTA[67]10-?????N?*" And CustOrd.MS_CODE Like "YTA[67]10-?????D?*" Then
                    ProcessStepReturn.SinglePointAction.SPI_Message = "[Indicator] Confirm Indicator Added or Removed?"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\50\" & "Removed.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\50\" & "Added.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Installed.jpg"
                End If

            End If
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo55(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            If CustOrd.MS_CODE Like "YTA[67]10-?????D?*" Then
                ProcessStepReturn.StepNo = "55"
                ProcessStepReturn.ProcessNo = "50"
                ProcessStepReturn.ProcessStep = "Modification / Assembly Checks"
                ProcessStepReturn.ActivityToCheck = "Display Cover Assembling"
                ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
                ProcessStepReturn.Initial = Initial
                ProcessStepReturn.SinglePointAction.SPI_Message = "[Indicator] Confirm Indicator Cover Quality after Assembling?"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\50\" & "No Galling.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\50\" & "Galling.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "No Galling.jpg"
            End If
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function

#End Region
End Class
