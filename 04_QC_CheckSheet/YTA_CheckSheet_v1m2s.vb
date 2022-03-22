Public Class YTA_CheckSheet_v1m2s

#Region "Common"

    Public Shared Function ProcessStepNo(ByVal StepNo As String, ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep

            'If StepNo = 11 Then ProcessStepReturn = (New YTA_CheckSheet).ProcessStepNo11(Initial, CustOrd, ErrMsg)
            If StepNo = "20_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo20_01_00(Initial, CustOrd, ErrMsg)

            If StepNo = "30_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo30_01_00(Initial, CustOrd, ErrMsg)
            If StepNo = "30_02_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo30_02_00(Initial, CustOrd, ErrMsg)
            If StepNo = "30_03_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo30_03_00(Initial, CustOrd, ErrMsg)
            If StepNo = "30_04_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo30_04_00(Initial, CustOrd, ErrMsg)

            If StepNo = "40_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo40_01_00(Initial, CustOrd, ErrMsg)

            If StepNo = "50_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo50_01_00(Initial, CustOrd, ErrMsg)

            If StepNo = "60_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo60_01_00(Initial, CustOrd, ErrMsg)
            If StepNo = "60_02_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo60_02_00(Initial, CustOrd, ErrMsg)

            If StepNo = "70_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo70_01_00(Initial, CustOrd, ErrMsg)
            If StepNo = "70_02_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo70_02_00(Initial, CustOrd, ErrMsg)

            If StepNo = "80_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo80_01_00(Initial, CustOrd, ErrMsg)
            If StepNo = "90_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo90_01_00(Initial, CustOrd, ErrMsg)
            If StepNo = "100_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo100_01_00(Initial, CustOrd, ErrMsg)

            If StepNo = "110_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo110_01_00(Initial, CustOrd, ErrMsg)
            If StepNo = "110_02_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo110_02_00(Initial, CustOrd, ErrMsg)

            If StepNo = "130_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo130_01_00(Initial, CustOrd, ErrMsg)
            If StepNo = "140_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo140_01_00(Initial, CustOrd, ErrMsg)
            If StepNo = "150_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo150_01_00(Initial, CustOrd, ErrMsg)

            If StepNo = "120_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo120_01_00(Initial, CustOrd, ErrMsg)
            If StepNo = "120_02_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo120_02_00(Initial, CustOrd, ErrMsg)
            If StepNo = "120_03_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo120_03_00(Initial, CustOrd, ErrMsg)
            If StepNo = "120_04_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo120_04_00(Initial, CustOrd, ErrMsg)
            If StepNo = "120_05_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo120_05_00(Initial, CustOrd, ErrMsg)

            If StepNo = "160_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo160_01_00(Initial, CustOrd, ErrMsg)
            If StepNo = "170_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo170_01_00(Initial, CustOrd, ErrMsg)
            If StepNo = "170_02_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo170_02_00(Initial, CustOrd, ErrMsg)

            If StepNo = "180_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo180_01_00(Initial, CustOrd, ErrMsg)

            If StepNo = "190_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo190_01_00(Initial, CustOrd, ErrMsg)
            If StepNo = "190_02_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo190_02_00(Initial, CustOrd, ErrMsg)
            If StepNo = "190_03_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo190_03_00(Initial, CustOrd, ErrMsg)
            If StepNo = "190_04_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo190_04_00(Initial, CustOrd, ErrMsg)
            If StepNo = "190_05_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo190_05_00(Initial, CustOrd, ErrMsg)
            If StepNo = "190_06_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo190_06_00(Initial, CustOrd, ErrMsg)
            If StepNo = "190_07_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo190_07_00(Initial, CustOrd, ErrMsg)
            If StepNo = "190_08_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo190_08_00(Initial, CustOrd, ErrMsg)
            If StepNo = "190_09_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo190_09_00(Initial, CustOrd, ErrMsg)
            If StepNo = "190_10_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo190_10_00(Initial, CustOrd, ErrMsg)
            If StepNo = "190_11_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo190_11_00(Initial, CustOrd, ErrMsg)
            If StepNo = "190_12_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo190_12_00(Initial, CustOrd, ErrMsg)
            If StepNo = "190_13_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo190_13_00(Initial, CustOrd, ErrMsg)
            If StepNo = "190_14_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo190_14_00(Initial, CustOrd, ErrMsg)
            If StepNo = "190_15_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo190_15_00(Initial, CustOrd, ErrMsg)
            If StepNo = "190_16_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo190_16_00(Initial, CustOrd, ErrMsg)

            If StepNo = "200_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo200_01_00(Initial, CustOrd, ErrMsg)
            If StepNo = "200_02_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo200_02_00(Initial, CustOrd, ErrMsg)
            If StepNo = "200_03_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo200_03_00(Initial, CustOrd, ErrMsg)
            If StepNo = "200_04_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo200_04_00(Initial, CustOrd, ErrMsg)
            If StepNo = "200_05_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo200_05_00(Initial, CustOrd, ErrMsg)
            If StepNo = "200_06_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo200_06_00(Initial, CustOrd, ErrMsg)
            If StepNo = "200_07_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo200_07_00(Initial, CustOrd, ErrMsg)
            If StepNo = "200_08_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1m2s).ProcessStepNo200_08_00(Initial, CustOrd, ErrMsg)


            MainForm.SetInspectionColor(StepNo.ToString, ProcessStepReturn.ProcessNo)

            Return ProcessStepReturn
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
#End Region

#Region "Version 1.2"

    Public Function ProcessStepNo20_01_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try
            Dim ProcessStepReturn As New CheckSheetStep
            Dim RL_Tag As New TML_Library.RandomArray
            Dim SelPartsSNO = RL_Tag.RandomStringArraySERIAL(CustOrd.SERIAL_NO_BEFORE, 4)

            ProcessStepReturn.ProcessNo = "20"
            ProcessStepReturn.ProcessStep = "Parts correctness"
            ProcessStepReturn.Activity = "Complete unit and KD Part correctness check"
            ProcessStepReturn.ToCheck = "Part numbers and Quantity as per Bill of Material"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.UserIput
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "20_01_00"
            ProcessStepReturn.ActivityToCheck = "Check Complete Unit Correctly Picked."
            ProcessStepReturn.UserInputAction.UserActionMessage = "Choose the Serial number printed In the plate Of before Modificaiton unit."
            ProcessStepReturn.UserInputAction.UserInputList = SelPartsSNO
            ProcessStepReturn.UserInputAction.UserInputCorrect = CustOrd.SERIAL_NO_BEFORE
            ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_020_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            ProcessStepReturn.Result &= "$" & MainForm.Setting.Var_60_020_Position_Initial.Replace("Initial", MainForm.Initial)
            ProcessStepReturn.Result &= "$" & MainForm.Setting.Var_60_020_Position_DateManufacture.Replace("DATE", Date.Today.ToString("dd"))
            Return ProcessStepReturn

        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo30_01_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "30"
            ProcessStepReturn.ProcessStep = "Data Plate/Tag Plate Marking"
            ProcessStepReturn.Activity = "Prepare Data and Tag Plates with correct contents"
            ProcessStepReturn.ToCheck = "Country of Origin"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.UserIput
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "30_01_00"
            ProcessStepReturn.ActivityToCheck = "Prepare Data Plates with correct COO."
            ProcessStepReturn.UserInputAction.UserActionMessage = "Choose the COO printed In the plate."
            ProcessStepReturn.UserInputAction.UserInputList = {"Made In China", "Made In Japan", "Made In KSA", "Made In Singapore"}
            If CustOrd.SERIAL_NO Like "Y3*" And CustOrd.EU_COUNTRY = "SA" Then
                ProcessStepReturn.UserInputAction.UserInputCorrect = "Made In KSA"
                ProcessStepReturn.Result = "Made In KSA-8,55,19.6$Tick-13,70,20$" & Initial & "-11,84,20"
            Else
                ProcessStepReturn.UserInputAction.UserInputCorrect = "Made In China"
                ProcessStepReturn.Result = "Made In China-8,55,19.6$Tick-13,70,20$" & Initial & "-11,84,20"
            End If
            ProcessStepReturn.Result &= "$" & MainForm.Setting.Var_60_030_Position_SerialBefore.Replace("SERIAL_NO_BEFORE", MainForm.CustOrd.SERIAL_NO_BEFORE)
            Return ProcessStepReturn
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo30_02_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
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

            ProcessStepReturn.StepNo = "30_02_00"
            ProcessStepReturn.ActivityToCheck = "Date Plates With correct contents."
            ProcessStepReturn.UserInputAction.UserActionMessage = "Choose the Serial number printed In the plate."
            ProcessStepReturn.UserInputAction.UserInputList = SelPartsSNO
            ProcessStepReturn.UserInputAction.UserInputCorrect = CustOrd.SERIAL_NO
            ProcessStepReturn.Result = ""
            Return ProcessStepReturn
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo30_03_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
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

            ProcessStepReturn.StepNo = "30_03_00"
            ProcessStepReturn.ActivityToCheck = "Tag Plates With correct contents."
            ProcessStepReturn.UserInputAction.UserActionMessage = "Choose the Tag number printed In the plate."
            ProcessStepReturn.UserInputAction.UserInputList = SelPartsTAG
            If CustOrd.TAG_NO_525 = "" Then
                ProcessStepReturn.UserInputAction.UserInputCorrect = "BLANK"
            Else
                ProcessStepReturn.UserInputAction.UserInputCorrect = CustOrd.TAG_NO_525
            End If
            ProcessStepReturn.Result = ""
            Return ProcessStepReturn
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo30_04_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "30"
            ProcessStepReturn.ProcessStep = "Data Plate Marking"
            ProcessStepReturn.Activity = "Prepare Data and Tag Plates with correct contents"
            ProcessStepReturn.ToCheck = "Plate data correct?  YES NO"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "30_04_00"
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
    Public Function ProcessStepNo40_01_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "40"
            ProcessStepReturn.ProcessStep = "Data Plate/Tag Plate Mounting"
            ProcessStepReturn.Activity = "Mount the marked Data and Tag Plates to the unit"
            ProcessStepReturn.ToCheck = "Approval type Data plate Part No. and No gap between plate and housing"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.UserIput
            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.StepNo_Group = "40"

            ProcessStepReturn.StepNo = "40_01_00"
            ProcessStepReturn.ActivityToCheck = "Correctness of Dataplate"
            ProcessStepReturn.UserInputAction.UserActionMessage = "Choose the Plate Part number from below list."
            ProcessStepReturn.UserInputAction.UserInputList = MainForm.DataPlateCheck 'SelParts
            ProcessStepReturn.UserInputAction.UserInputCorrect = MainForm.DataPlateCorrect 'PlatePartNo
            ProcessStepReturn.Result = MainForm.DataPlateCorrect & "-8,62,21.8"
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo50_01_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "50"
            ProcessStepReturn.ProcessStep = "Modification / Assembly Checks"
            ProcessStepReturn.Activity = "Display, Terminal Cover Open"
            ProcessStepReturn.ToCheck = "No Thread damage"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "50_01_00"
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
    Public Function ProcessStepNo60_01_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "60"
            ProcessStepReturn.ProcessStep = "Modification / Assembly Checks"
            ProcessStepReturn.Activity = "BOUT Switch setting/Plug for electrical connection"
            ProcessStepReturn.ToCheck = "/C1  /C2  /C3  NA / Plug Installed"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.StepNo = "60_01_00"

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
    Public Function ProcessStepNo60_02_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            If CustOrd.MS_CODE Like "YTA*" Then
                ProcessStepReturn.ProcessNo = "60"
                ProcessStepReturn.ProcessStep = "Modification / Assembly Checks"
                ProcessStepReturn.Activity = "BOUT Switch setting/Plug for electrical connection"
                ProcessStepReturn.ToCheck = "/C1  /C2  /C3  NA / Plug Installed"
                ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
                ProcessStepReturn.Initial = Initial

                ProcessStepReturn.StepNo = "60_02_00"
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
    Public Function ProcessStepNo70_01_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            If CustOrd.MS_CODE Like "YTA*" Then
                ProcessStepReturn.ProcessNo = "70"
                ProcessStepReturn.ProcessStep = "Modification / Assembly Checks"
                ProcessStepReturn.Activity = "Display Indicator"
                ProcessStepReturn.ToCheck = "ADDED   REMOVED    NO-CHANGE"
                ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
                ProcessStepReturn.Initial = Initial

                ProcessStepReturn.StepNo = "70_01_00"
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
    Public Function ProcessStepNo70_02_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            If CustOrd.MS_CODE Like "YTA*" Then
                ProcessStepReturn.ProcessNo = "70"
                ProcessStepReturn.ProcessStep = "Modification / Assembly Checks"
                ProcessStepReturn.Activity = "Display Indicator"
                ProcessStepReturn.ToCheck = "ADDED   REMOVED    NO-CHANGE"
                ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
                ProcessStepReturn.Initial = Initial

                ProcessStepReturn.StepNo = "70_02_00"
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
    Public Function ProcessStepNo80_01_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "80"
            ProcessStepReturn.ProcessStep = "Modification / Assembly Checks"
            ProcessStepReturn.Activity = "Lock Screw installation"
            ProcessStepReturn.ToCheck = "YES    NO"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "80_01_00"
            ProcessStepReturn.ActivityToCheck = "Lock Screw installation"
            If CustOrd.MS_CODE Like "YTA[67]10-?????D?*/[KS][UF][12]*" Then
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
            ElseIf CustOrd.MS_CODE Like "YTA[67]10-?????N?*/[KS][UF][12]*" Then
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
    Public Function ProcessStepNo90_01_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            If CustOrd.MS_CODE Like "YTA[67]10-?????D?*" Then
                ProcessStepReturn.ProcessNo = "90"
                ProcessStepReturn.ProcessStep = "Modification / Assembly Checks"
                ProcessStepReturn.Activity = "Display Cover Assembling"
                ProcessStepReturn.ToCheck = "Clean & No galling"
                ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
                ProcessStepReturn.Initial = Initial

                ProcessStepReturn.StepNo = "90_01_00"
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
    Public Function ProcessStepNo100_01_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            If CustOrd.MS_CODE Like "YTA*" Then
                ProcessStepReturn.ProcessNo = "100"
                ProcessStepReturn.ProcessStep = "Modification / Assembly Checks"
                ProcessStepReturn.Activity = "Lightning Protector"
                ProcessStepReturn.ToCheck = "ADDED    REMOVED    NO-CHANGE"
                ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
                ProcessStepReturn.Initial = Initial

                ProcessStepReturn.StepNo = "100_01_00"
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
    Public Function ProcessStepNo110_01_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "110"
            ProcessStepReturn.ProcessStep = "Modification / Assembly Checks"
            ProcessStepReturn.Activity = "Ground Screws Fixing/DLM QR Label Fixing"
            ProcessStepReturn.ToCheck = "FIXED INTERNAL & EXTERNAL/FIXED QR LABEL AT PROPER PLACE ON CASE"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "110_01_00"
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
    Public Function ProcessStepNo110_02_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "110"
            ProcessStepReturn.ProcessStep = "Modification / Assembly Checks"
            ProcessStepReturn.Activity = "Ground Screws Fixing/DLM QR Label Fixing"
            ProcessStepReturn.ToCheck = "FIXED INTERNAL & EXTERNAL/FIXED QR LABEL AT PROPER PLACE ON CASE"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "110_02_00"
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
    Public Function ProcessStepNo120_01_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "120"
            ProcessStepReturn.ProcessStep = "Basic Checks"
            ProcessStepReturn.Activity = "Protocol check,Sensor Input,MOC,Electrical Connection,Display Type"
            ProcessStepReturn.ToCheck = "J/F||1/2||A/C||0/2/4||D/N"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "120_01_00"
            ProcessStepReturn.ActivityToCheck = "Output Singal of YTA Unit"
            ProcessStepReturn.SinglePointAction.SPI_Message = "[Output Singal] Confirm Communication Protocol"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\120\" & "HART.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\120\" & "Fieldbus.jpg"
            ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_120_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1) '"Tick-13,70,36.5"
            If MainForm.CustOrd.MS_CODE Like "YTA???-J*" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "HART.jpg"
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_120_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)
            ElseIf MainForm.CustOrd.MS_CODE Like "YTA???-F*" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Fieldbus.jpg"
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_120_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(1)
            Else
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "BRAIN.jpg"
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_120_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(2)
            End If
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo120_02_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "120"
            ProcessStepReturn.ProcessStep = "Basic Checks"
            ProcessStepReturn.Activity = "Protocol check,Sensor Input,MOC,Electrical Connection,Display Type"
            ProcessStepReturn.ToCheck = "J/F||1/2||A/C||0/2/4||D/N"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "120_02_00"
            ProcessStepReturn.ActivityToCheck = "Number of Sensor Inputs of YTA Unit"
            ProcessStepReturn.SinglePointAction.SPI_Message = "[Inputs] Confirm the number of Sensor Inputs"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\120\" & "SingleSensor.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\120\" & "DualSensor.jpg"

            ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_120_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            If MainForm.CustOrd.MS_CODE Like "YTA???-??1*" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "SingleSensor.jpg"
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_120_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)
            ElseIf MainForm.CustOrd.MS_CODE Like "YTA???-??2*" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "DualSensor.jpg"
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_120_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(1)
            End If
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo120_03_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "120"
            ProcessStepReturn.ProcessStep = "Basic Checks"
            ProcessStepReturn.Activity = "Protocol check,Sensor Input,MOC,Electrical Connection,Display Type"
            ProcessStepReturn.ToCheck = "J/F||1/2||A/C||0/2/4||D/N"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "120_03_00"
            ProcessStepReturn.ActivityToCheck = "Housing Material of YTA Unit"
            ProcessStepReturn.SinglePointAction.SPI_Message = "[MOC] Confirm the Material of Housing and Covers"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\120\" & "Aluminum.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\120\" & "SUS.jpg"

            ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_120_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            If MainForm.CustOrd.MS_CODE Like "YTA???-???A*" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Aluminum.jpg"
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_120_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)
            ElseIf MainForm.CustOrd.MS_CODE Like "YTA???-???C*" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "SUS.jpg"
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_120_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(1)
            End If
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo120_04_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "120"
            ProcessStepReturn.ProcessStep = "Basic Checks"
            ProcessStepReturn.Activity = "Protocol check,Sensor Input,MOC,Electrical Connection,Display Type"
            ProcessStepReturn.ToCheck = "J/F||1/2||A/C||0/2/4||D/N"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "120_04_00"
            ProcessStepReturn.ActivityToCheck = "Electrical Connection Type of YTA Unit"
            ProcessStepReturn.SinglePointAction.SPI_Message = "Confirm the Electrical Connection type of Housing/Case"

            ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_120_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            If MainForm.CustOrd.MS_CODE Like "YTA???-????2*" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\120\" & "NPT.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\120\" & "M20.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "NPT.jpg"
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_120_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(1)
            ElseIf MainForm.CustOrd.MS_CODE Like "YTA???-????4*" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\120\" & "NPT.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\120\" & "M20.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "M20.jpg"
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_120_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(2)
            ElseIf MainForm.CustOrd.MS_CODE Like "YTA???-????0*" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\120\" & "G12.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\120\" & "NPT.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "G12.jpg"
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_120_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)
            End If
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo120_05_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "120"
            ProcessStepReturn.ProcessStep = "Basic Checks"
            ProcessStepReturn.Activity = "Protocol check,Sensor Input,MOC,Electrical Connection,Display Type"
            ProcessStepReturn.ToCheck = "J/F||1/2||A/C||0/2/4||D/N"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "120_05_00"
            ProcessStepReturn.ActivityToCheck = "Display Type of YTA Unit"
            ProcessStepReturn.SinglePointAction.SPI_Message = "Confirm the Display Type"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\120\" & "INDICATOR.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\120\" & "NO INDICATOR.jpg"

            ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_120_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            If MainForm.CustOrd.MS_CODE Like "YTA???-?????D*" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "INDICATOR.jpg"
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_120_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(1)
            ElseIf MainForm.CustOrd.MS_CODE Like "YTA???-?????N*" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "NO INDICATOR.jpg"
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_120_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)
            End If
            ProcessStepReturn.Result &= "$" & MainForm.Setting.Var_60_120_Position_Initial.Replace("Initial", MainForm.Initial)
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo130_01_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "130"
            ProcessStepReturn.ProcessStep = "ACW Test"
            ProcessStepReturn.Activity = "AC Withstand Voltage test"
            ProcessStepReturn.ToCheck = "Test result is GO or No-GO"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "130_01_00"
            ProcessStepReturn.ActivityToCheck = "AC Withstand Voltage test. DB Result:" & MainForm.Hipot.acw_test_result
            ProcessStepReturn.SinglePointAction.SPI_Message = "Is the ACW Test completed successfully?"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Yes.jpg"

            If MainForm.Hipot.acw_test_result Like "*PASS*PASS*" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "No.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Yes.jpg"
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_130_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_130_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)
            ElseIf Not NeedsHiPot(CustOrd.MS_CODE, CustOrd.MS_CODE_BEFORE) Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Not Applicable.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Not Applicable.jpg"
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_130_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_130_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(2)
            Else
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "No.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "No.jpg"
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_130_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_130_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(1)
            End If

            Return ProcessStepReturn

        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo140_01_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "140"
            ProcessStepReturn.ProcessStep = "IR Test"
            ProcessStepReturn.Activity = "Insulation resistance test"
            ProcessStepReturn.ToCheck = "Test result is GO or No-GO"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "140_01_00"
            ProcessStepReturn.ActivityToCheck = "Insulation resistance test. DB Result:" & MainForm.Hipot.ir_test_result
            ProcessStepReturn.SinglePointAction.SPI_Message = "Is the IR Test completed successfully?"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Yes.jpg"
            If MainForm.Hipot.ir_test_result Like "*PASS*PASS*" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "No.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Yes.jpg"
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_140_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_140_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)
            ElseIf Not NeedsHiPot(CustOrd.MS_CODE, CustOrd.MS_CODE_BEFORE) Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Not Applicable.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Not Applicable.jpg"
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_140_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_140_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(2)
            Else
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "No.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "No.jpg"
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_140_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_140_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(1)
            End If
            ProcessStepReturn.Result &= "$" & MainForm.Setting.Var_60_140_Position_Initial.Replace("Initial", MainForm.Initial)
            Return ProcessStepReturn

        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo150_01_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "150"
            ProcessStepReturn.ProcessStep = "Programming and Calibration"
            ProcessStepReturn.Activity = "Calibration and Programming Procedure"
            ProcessStepReturn.ToCheck = "Test result is GO or No-GO"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "150_01_00"
            ProcessStepReturn.ActivityToCheck = "Calibration result Go or No-Go. DB Result:" & MainForm.YTA_Crc.RESULT
            ProcessStepReturn.SinglePointAction.SPI_Message = "Is the Calibration completed successfully?"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Yes.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "No.jpg"
            If MainForm.YTA_Crc.RESULT = "GO" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Yes.jpg"
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_150_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_150_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)
            Else
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "No.jpg"
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_150_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_150_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(1)
            End If
            ProcessStepReturn.Result &= "$" & MainForm.Setting.Var_60_150_Position_Initial.Replace("Initial", MainForm.Initial)
            Return ProcessStepReturn

        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo160_01_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "160"
            ProcessStepReturn.ProcessStep = "Final Assembly"
            ProcessStepReturn.Activity = "Mount Terminal cover"
            ProcessStepReturn.ToCheck = "Mounted Terminal cover YES || NO"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "160_01_00"
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
    Public Function ProcessStepNo170_01_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "170"
            ProcessStepReturn.ProcessStep = "Final Assembly"
            ProcessStepReturn.Activity = "Remove or Attach Mounting Bracket"
            ProcessStepReturn.ToCheck = "Remove or Attach Brackets YES || NO"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "170_01_00"
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
    Public Function ProcessStepNo170_02_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "170"
            ProcessStepReturn.ProcessStep = "Final Assembly"
            ProcessStepReturn.Activity = "Remove or Attach Mounting Bracket"
            ProcessStepReturn.ToCheck = "Remove or Attach Brackets YES || NO"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "170_02_00"
            If CustOrd.MS_CODE_BEFORE.Substring(13, 1) = CustOrd.MS_CODE.Substring(13, 1) Then
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
                If CustOrd.MS_CODE_BEFORE.Substring(13, 1) <> "N" Then
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
    Public Function ProcessStepNo180_01_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "180"
            ProcessStepReturn.ProcessStep = "Print QIC"
            ProcessStepReturn.Activity = "Print Certificates"
            ProcessStepReturn.ToCheck = "Printed YES || NO"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.DocumentCheck
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "180_01_00"
            ProcessStepReturn.ActivityToCheck = "Print QIC"
            ProcessStepReturn.ViewDocAction.DocumentCheckMessage = "Printed correctly YES || NO ?"
            ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_180_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            'ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_180_Position_QicTick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            ProcessStepReturn.Result &= "$" & MainForm.Setting.Var_60_180_Position_Initial.Replace("Initial", MainForm.Initial)
            ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_180_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)
            'ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_180_Position_QicCircle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)
            'ProcessStepReturn.Result &= "$" & MainForm.Setting.Var_60_180_Position_QicInitial.Replace("Initial", MainForm.Initial)


            Dim QicFolerPath As String = MainForm.Setting.Var_06_DocsStore & "Production Complete Documents\QICDOC\"
            If Not (System.IO.Directory.Exists(QicFolerPath & CustOrd.PROD_NO)) Then
                System.IO.Directory.CreateDirectory(QicFolerPath & CustOrd.PROD_NO)
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
    Public Function ProcessStepNo190_01_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "190"
            ProcessStepReturn.ProcessStep = "Compare Production Order with Test Reports and Unit"
            ProcessStepReturn.Activity = "Model Code||Serial Number||Calibration Range||Tag Number||Agency ||pproval||Check data of QIC"
            ProcessStepReturn.ToCheck = "Correct || Not Correct"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.UserIput
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "190_01_00"
            ProcessStepReturn.ActivityToCheck = "Correct MS-CODE"
            ProcessStepReturn.UserInputAction.UserActionMessage = "Confirm MS-CODE in YTA, QIC and OrderTag/SAP Sheet"
            Dim Var1, Var2, Var3 As New String("")
            Dim Mscode As String = CustOrd.MS_CODE
            If Mscode Like "YTA610*" Then Var1 = Mscode.Replace("YTA610", "YTA710")
            If Mscode Like "YTA710*" Then Var1 = Mscode.Replace("YTA710", "YTA610")
            If Mscode Like "YTA???-??1*" Then Var2 = Mscode.Remove(9, 1).Insert(9, "2")
            If Mscode Like "YTA???-??2*" Then Var2 = Mscode.Remove(9, 1).Insert(9, "1")
            If Mscode Like "YTA???-?????D*" Then Var3 = Mscode.Remove(12, 1).Insert(12, "N")
            If Mscode Like "YTA???-?????N*" Then Var3 = Mscode.Remove(12, 1).Insert(12, "D")
            ProcessStepReturn.UserInputAction.UserInputList = {Mscode, Var1, Var2, Var3}
            ProcessStepReturn.UserInputAction.UserInputCorrect = Mscode
            ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_190_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_190_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo190_02_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "190"
            ProcessStepReturn.ProcessStep = "Compare Production Order with Test Reports and Unit"
            ProcessStepReturn.Activity = "Model Code||Serial Number||Calibration Range||Tag Number||Agency ||pproval||Check data of QIC"
            ProcessStepReturn.ToCheck = "Correct || Not Correct"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.UserIput
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "190_02_00"
            ProcessStepReturn.ActivityToCheck = "Correct SERIAL Number"
            ProcessStepReturn.UserInputAction.UserActionMessage = "Confirm SERIAL NO. in YTA, QIC and OrderTag/SAP Sheet"
            Dim Var1, Var2, Var3 As New String("")
            Dim SerialNo As String = CustOrd.SERIAL_NO
            Dim Randm As Random = New Random

FixVar1:
            Var1 = SerialNo.Remove(8, 1).Insert(8, Randm.Next(0, 9))
            If SerialNo Like Var1 Then
                GoTo FixVar1
            End If

FixVar2:
            Var2 = SerialNo.Remove(7, 1).Insert(7, Randm.Next(0, 9))
            If SerialNo Like Var2 Then
                GoTo FixVar2
            End If

FixVar3:
            Var3 = SerialNo.Remove(6, 1).Insert(6, Randm.Next(0, 9))
            If SerialNo Like Var3 Then
                GoTo FixVar3
            End If

            ProcessStepReturn.UserInputAction.UserInputList = {SerialNo, Var1, Var2, Var3}
            ProcessStepReturn.UserInputAction.UserInputCorrect = SerialNo
            ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_190_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_190_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo190_03_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "190"
            ProcessStepReturn.ProcessStep = "Compare Production Order with Test Reports and Unit"
            ProcessStepReturn.Activity = "Model Code||Serial Number||Calibration Range||Tag Number||Agency ||pproval||Check data of QIC"
            ProcessStepReturn.ToCheck = "Correct || Not Correct"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.UserIput
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "190_03_00"
            ProcessStepReturn.ActivityToCheck = "Correct Calibration Range"
            ProcessStepReturn.UserInputAction.UserActionMessage = "Confirm Range Value. in YTA, QIC and OrderTag/SAP Sheet"
            Dim Var1, Var2, Var3 As New String("")
            Dim Range_1 As String = ""
            Dim Range_2 As String = ""
            Dim Range As String
            Dim Randm As Random = New Random
            Dim Mscode = CustOrd.MS_CODE
            If Mscode Like "YTA???-??2*" Then
                Range_1 = "S1: " & CustOrd.ORD_INST_MIN_T70 & " TO " & CustOrd.ORD_INST_MAX_T70 & " " & CustOrd.UNIT_T70
                Range_2 = "S2: " & CustOrd.ORD_INST_MIN_T71 & " TO " & CustOrd.ORD_INST_MAX_T71 & " " & CustOrd.UNIT_T71
                Var1 = "S1: " & CustOrd.ORD_INST_MIN_T70 & " TO " & (Integer.Parse(CustOrd.ORD_INST_MAX_T70) + 25).ToString & " " & CustOrd.UNIT_T70
                Var1 &= "|| S2: " & CustOrd.ORD_INST_MIN_T71 & " TO " & (Integer.Parse(CustOrd.ORD_INST_MAX_T71) - 25).ToString & " " & CustOrd.UNIT_T71
                Var2 = "S1: " & CustOrd.ORD_INST_MIN_T70 & " TO " & (Integer.Parse(CustOrd.ORD_INST_MAX_T70) - 25).ToString & " " & CustOrd.UNIT_T70
                Var2 &= "|| S2: " & CustOrd.ORD_INST_MIN_T71 & " TO " & (Integer.Parse(CustOrd.ORD_INST_MAX_T71) + 25).ToString & " " & CustOrd.UNIT_T71
                Var3 = "S1: " & CustOrd.ORD_INST_MIN_T70 & " TO " & (Integer.Parse(CustOrd.ORD_INST_MAX_T70) + 50).ToString & " " & CustOrd.UNIT_T70
                Var3 &= "|| S2: " & CustOrd.ORD_INST_MIN_T71 & " TO " & (Integer.Parse(CustOrd.ORD_INST_MAX_T71) - 50).ToString & " " & CustOrd.UNIT_T71
            Else
                Range_1 = "S1: " & CustOrd.ORD_INST_MIN_T70 & " TO " & CustOrd.ORD_INST_MAX_T70 & " " & CustOrd.UNIT_T70
                Range_2 = "S2: "
                Var1 = "S1: " & CustOrd.ORD_INST_MIN_T70 & " TO " & (Integer.Parse(CustOrd.ORD_INST_MAX_T70) + 25).ToString & " " & CustOrd.UNIT_T70
                Var1 &= "|| S2: "
                Var2 = "S1: " & CustOrd.ORD_INST_MIN_T70 & " TO " & (Integer.Parse(CustOrd.ORD_INST_MAX_T70) - 25).ToString & " " & CustOrd.UNIT_T70
                Var2 &= "|| S2: "
                Var3 = "S1: " & CustOrd.ORD_INST_MIN_T70 & " TO " & (Integer.Parse(CustOrd.ORD_INST_MAX_T70) + 50).ToString & " " & CustOrd.UNIT_T70
                Var3 &= "|| S2: "
            End If
            Range = Range_1 & " || " & Range_2
            ProcessStepReturn.UserInputAction.UserInputList = {Range, Var1, Var2, Var3}
            ProcessStepReturn.UserInputAction.UserInputCorrect = Range
            ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_190_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_190_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo190_04_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try
            Dim RL_Tag As New TML_Library.RandomArray
            Dim SelPartsTAG = RL_Tag.RandomStringArrayYTA(CustOrd.TAG_NO_525, 4)

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "190"
            ProcessStepReturn.ProcessStep = "Compare Production Order with Test Reports and Unit"
            ProcessStepReturn.Activity = "Model Code||Serial Number||Calibration Range||Tag Number||Agency ||pproval||Check data of QIC"
            ProcessStepReturn.ToCheck = "Correct || Not Correct"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.UserIput
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "190_04_00"
            ProcessStepReturn.ActivityToCheck = "Tag Number correctness"
            ProcessStepReturn.UserInputAction.UserActionMessage = "Choose the Tag number printed In the plate."
            ProcessStepReturn.UserInputAction.UserInputList = SelPartsTAG
            If CustOrd.TAG_NO_525 = "" Then
                ProcessStepReturn.UserInputAction.UserInputCorrect = "BLANK"
            Else
                ProcessStepReturn.UserInputAction.UserInputCorrect = CustOrd.TAG_NO_525
            End If
            ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_190_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_190_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            Return ProcessStepReturn
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo190_05_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "190"
            ProcessStepReturn.ProcessStep = "Compare Production Order with Test Reports and Unit"
            ProcessStepReturn.Activity = "Model Code||Serial Number||Calibration Range||Tag Number||Agency ||pproval||Check data of QIC"
            ProcessStepReturn.ToCheck = "Correct || Not Correct"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.UserIput
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "190_05_00"
            ProcessStepReturn.ActivityToCheck = "Agency Approvals"
            ProcessStepReturn.UserInputAction.UserActionMessage = "Choose the correct Approval of the YTA Unit."
            ProcessStepReturn.UserInputAction.UserInputList = {"ATEX", "FM", "IECEx", "NONE"}
            If CustOrd.MS_CODE Like "YTA???-*/K[FSU]*" Then
                ProcessStepReturn.UserInputAction.UserInputCorrect = "ATEX"
            ElseIf CustOrd.MS_CODE Like "YTA???-*/F[FSU]*" Then
                ProcessStepReturn.UserInputAction.UserInputCorrect = "FM"
            ElseIf CustOrd.MS_CODE Like "YTA???-*/S[FSU]*" Then
                ProcessStepReturn.UserInputAction.UserInputCorrect = "IECEx"
            Else
                ProcessStepReturn.UserInputAction.UserInputCorrect = "NONE"
            End If
            ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_190_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_190_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            Return ProcessStepReturn
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo190_06_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "190"
            ProcessStepReturn.ProcessStep = "Compare Production Order with Test Reports and Unit"
            ProcessStepReturn.Activity = "Model Code||Serial Number||Calibration Range||Tag Number||Agency ||pproval||Check data of QIC"
            ProcessStepReturn.ToCheck = "Correct || Not Correct"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "190_06_00"
            ProcessStepReturn.ActivityToCheck = "RoHS Confirmatatory Marking"
            ProcessStepReturn.SinglePointAction.SPI_Message = "The Plate has which Of the following [with CE Or without CE]. This Is sample picture, so don't check other contents."
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\30\" & "NamePlate_with_CE.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\30\" & "NamePlate_without_CE.jpg"
            If CustOrd.SERIAL_NO_BEFORE Like "S5WC*" Or CustOrd.SERIAL_NO_BEFORE Like "S5X1*" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "NamePlate_without_CE.jpg"
                ProcessStepReturn.Result = "" ' "Yes without CE-8,57,20.6"
            Else
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "NamePlate_with_CE.jpg"
                ProcessStepReturn.Result = "" '"Yes with CE-8,57,20.6"
            End If

            Return ProcessStepReturn
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo190_07_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "190"
            ProcessStepReturn.ProcessStep = "Compare Production Order with Test Reports and Unit"
            ProcessStepReturn.Activity = "Model Code||Serial Number||Calibration Range||Tag Number||Agency ||pproval||Check data of QIC"
            ProcessStepReturn.ToCheck = "Correct || Not Correct"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "190_07_00"
            ProcessStepReturn.ActivityToCheck = "Check QIC Data [1] Cal.Point/Error [2] Temp.Humidity [3] Overall"
            ProcessStepReturn.SinglePointAction.SPI_Message = "Check QIC Overall data including Cal. Points and Temp./Humidity"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Correct.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Wrong.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Correct.jpg"
            ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_190_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_190_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            ProcessStepReturn.Result &= "$" & MainForm.Setting.Var_60_190_Position_Initial.Replace("Initial", Initial)
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo190_08_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "190"
            ProcessStepReturn.ProcessStep = "Visually Inspect Unit"
            ProcessStepReturn.Activity = "Display||Clean||Lock Screw||Approval Plate||Tag Plate||N4 Plate||N4 Tagnumber||Bracket"
            ProcessStepReturn.ToCheck = "Correct || Not Correct"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "190_08_00"
            ProcessStepReturn.ActivityToCheck = "Does the unit have a Display?"
            ProcessStepReturn.SinglePointAction.SPI_Message = "Check presence of Display  YES  ||  NO"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Yes.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "No.jpg"
            If CustOrd.MS_CODE Like "YTA???-?????D?*" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Yes.jpg"
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_1900_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_1900_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)
            Else
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "No.jpg"
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_1900_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_1900_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(1)
            End If
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo190_09_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "190"
            ProcessStepReturn.ProcessStep = "Visually Inspect Unit"
            ProcessStepReturn.Activity = "Display||Clean||Lock Screw||Approval Plate||Tag Plate||N4 Plate||N4 Tagnumber||Bracket"
            ProcessStepReturn.ToCheck = "Correct || Not Correct"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "190_09_00"
            If CustOrd.MS_CODE Like "YTA???-?????D?*" Then
                ProcessStepReturn.ActivityToCheck = "Unit and Glass/Back Covers Clean?"
            Else
                ProcessStepReturn.ActivityToCheck = "Unit and Blind/Back Covers Clean?"
            End If
            ProcessStepReturn.SinglePointAction.SPI_Message = "Check the cleanliness of Unit  YES  ||  NO"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Yes.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "No.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Yes.jpg"
            ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_1900_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_1900_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo190_10_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "190"
            ProcessStepReturn.ProcessStep = "Visually Inspect Unit"
            ProcessStepReturn.Activity = "Display||Clean||Lock Screw||Approval Plate||Tag Plate||N4 Plate||N4 Tagnumber||Bracket"
            ProcessStepReturn.ToCheck = "Correct || Not Correct"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "190_10_00"
            ProcessStepReturn.ActivityToCheck = "Lock Screw tightened?"
            ProcessStepReturn.SinglePointAction.SPI_Message = "Check the Lock Screw tigtened  YES  ||  NO"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Yes.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "No.jpg"
            If CustOrd.MS_CODE Like "YTA???-???????*/[KS][UF]*" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Yes.jpg"
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_1900_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_1900_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)
            Else
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "No.jpg"
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_1900_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_1900_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(1)
            End If
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo190_11_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try
            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "190"
            ProcessStepReturn.ProcessStep = "Visually Inspect Unit"
            ProcessStepReturn.Activity = "Display||Clean||Lock Screw||Approval Plate||Tag Plate||N4 Plate||N4 Tagnumber||Bracket"
            ProcessStepReturn.ToCheck = "Correct || Not Correct"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.UserIput
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "190_11_00"
            ProcessStepReturn.ActivityToCheck = "Approval Name plate Fixed?"
            ProcessStepReturn.UserInputAction.UserActionMessage = "Choose the Plate Part number from below list."
            ProcessStepReturn.UserInputAction.UserInputList = MainForm.DataPlateCheck
            ProcessStepReturn.UserInputAction.UserInputCorrect = MainForm.DataPlateCorrect
            ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_1900_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_1900_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)
            Return ProcessStepReturn

        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo190_12_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "190"
            ProcessStepReturn.ProcessStep = "Data Plate/Tag Plate Mounting"
            ProcessStepReturn.Activity = "Mount the marked Data and Tag Plates to the unit"
            ProcessStepReturn.ToCheck = "Approval type Data plate Part No. and No gap between plate and housing"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst

            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.StepNo = "190_12_00"
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
    Public Function ProcessStepNo190_13_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "190"
            ProcessStepReturn.ProcessStep = "Visually Inspect Unit"
            ProcessStepReturn.Activity = "Display||Clean||Lock Screw||Approval Plate||Tag Plate||N4 Plate||N4 Tagnumber||Bracket"
            ProcessStepReturn.ToCheck = "Correct || Not Correct"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "190_13_00"
            ProcessStepReturn.ActivityToCheck = "Tag Plate Fixed?"
            ProcessStepReturn.SinglePointAction.SPI_Message = "Check the Tag plate fixed correctly  YES  ||  NO"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Yes.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "No.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Yes.jpg"
            ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_1900_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_1900_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo190_14_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "190"
            ProcessStepReturn.ProcessStep = "Visually Inspect Unit"
            ProcessStepReturn.Activity = "Display||Clean||Lock Screw||Approval Plate||Tag Plate||N4 Plate||N4 Tagnumber||Bracket"
            ProcessStepReturn.ToCheck = "Correct || Not Correct"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "190_14_00"
            ProcessStepReturn.ActivityToCheck = "/N4 Tag Plate Fixed?"
            ProcessStepReturn.SinglePointAction.SPI_Message = "Check /N4 Tag plate fixed correctly  YES  ||  NO"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Yes.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Not Applicable.jpg"
            If CustOrd.MS_CODE Like "YTA???-???????*/N4*" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Yes.jpg"
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_1900_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_1900_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)
            Else
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Not Applicable.jpg"
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_1900_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_1900_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(1)
            End If
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo190_15_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "190"
            ProcessStepReturn.ProcessStep = "Visually Inspect Unit"
            ProcessStepReturn.Activity = "Display||Clean||Lock Screw||Approval Plate||Tag Plate||N4 Plate||N4 Tagnumber||Bracket"
            ProcessStepReturn.ToCheck = "Correct || Not Correct"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "190_15_00"
            ProcessStepReturn.ActivityToCheck = "/N4 Tag Plate has correct Tag number?"
            ProcessStepReturn.SinglePointAction.SPI_Message = "Check /N4 Tag number  YES  ||  NO"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Yes.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Not Applicable.jpg"
            If CustOrd.MS_CODE Like "YTA???-???????*/N4*" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Yes.jpg"
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_1900_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_1900_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)
            Else
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Not Applicable.jpg"
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_1900_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_1900_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(1)
            End If
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo190_16_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "190"
            ProcessStepReturn.ProcessStep = "Visually Inspect Unit"
            ProcessStepReturn.Activity = "Display||Clean||Lock Screw||Approval Plate||Tag Plate||N4 Plate||N4 Tagnumber||Bracket"
            ProcessStepReturn.ToCheck = "Correct || Not Correct"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "190_16_00"
            ProcessStepReturn.ActivityToCheck = "Mounting Bracket Correct?"
            ProcessStepReturn.SinglePointAction.SPI_Message = "Is the Mounting Bracket correct  YES  ||  NO"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Yes.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Not Applicable.jpg"
            If Not (CustOrd.MS_CODE Like "YTA???-??????N*") Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Yes.jpg"
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_1900_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_1900_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)
            Else
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Not Applicable.jpg"
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_1900_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_1900_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(1)
            End If
            ProcessStepReturn.Result &= "$" & MainForm.Setting.Var_60_1900_Position_Initial.Replace("Initial", Initial)
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo200_01_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "200"
            ProcessStepReturn.ProcessStep = "Packing"
            ProcessStepReturn.Activity = "UnitKept||IM||EuDOC||QIC||Bracket||VisualCheck||Markings||OrderTag"
            ProcessStepReturn.ToCheck = "Correct||Not Correct||Details"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "200_01_00"
            ProcessStepReturn.ActivityToCheck = "Unit Kept Inside Box?"
            ProcessStepReturn.SinglePointAction.SPI_Message = "Is the YTA unit properly packed and kept inside Box?  YES  ||  NO"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Yes.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Not Applicable.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Yes.jpg"
            ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_2000_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_2000_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo200_02_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "200"
            ProcessStepReturn.ProcessStep = "Packing"
            ProcessStepReturn.Activity = "UnitKept||IM||EuDOC||QIC||Bracket||VisualCheck||Markings||OrderTag"
            ProcessStepReturn.ToCheck = "Correct||Not Correct||Details"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.AddedDocs
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "200_02_00"
            ProcessStepReturn.ActivityToCheck = "Scan QR-code of IM,SafetyIM,EU-Doc etc, if any."
            If Integer.Parse(CustOrd.QTY_NO) = 1 And CustOrd.MS_CODE Like "YTA*/K[UFS]*" Then
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_2000_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_2000_Position_IM.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_2000_Position_EUDoC.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            ElseIf Integer.Parse(CustOrd.QTY_NO) = 1 And Not (CustOrd.MS_CODE Like "YTA*/K[UFS]*") Then
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_2000_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_2000_Position_IM.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_2000_Position_NoEUDoC.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            ElseIf Integer.Parse(CustOrd.QTY_NO) <> 1 And (CustOrd.MS_CODE Like "YTA*/K[UFS]*") Then
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_2000_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_2000_Position_NoIM.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_2000_Position_EUDoC.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            Else
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_2000_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_2000_Position_NoIM.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_2000_Position_NoEUDoC.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            End If

            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo200_03_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "200"
            ProcessStepReturn.ProcessStep = "Packing"
            ProcessStepReturn.Activity = "UnitKept||IM||EuDOC||QIC||Bracket||VisualCheck||Markings||OrderTag"
            ProcessStepReturn.ToCheck = "Correct||Not Correct||Details"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "200_03_00"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Yes.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Not Applicable.jpg"

            If Integer.Parse(CustOrd.QTY_NO) = 1 And CustOrd.MS_CODE Like "YTA*/K[UFS]*" Then
                ProcessStepReturn.ActivityToCheck = "IM, SafetyIM and EU-Documents kept Inside Box?"
                ProcessStepReturn.SinglePointAction.SPI_Message = "IM, SafetyIM & EU-Documents kept Inside Box??  YES  ||  NO"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Yes.jpg"
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_2000_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            ElseIf Integer.Parse(CustOrd.QTY_NO) = 1 And Not (CustOrd.MS_CODE Like "YTA*/K[UFS]*") Then
                ProcessStepReturn.ActivityToCheck = "IM+SafetyIM only kept Inside Box with [NO] EuDocument?"
                ProcessStepReturn.SinglePointAction.SPI_Message = "Only IM and SafetyIM kept Inside Box??  YES  ||  NO"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Yes.jpg"
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_2000_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_2000_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)
            ElseIf Integer.Parse(CustOrd.QTY_NO) <> 1 And (CustOrd.MS_CODE Like "YTA*/K[UFS]*") Then
                ProcessStepReturn.ActivityToCheck = "Eu-Document only kept Inside Box with [NO] IM/SafetyIM?"
                ProcessStepReturn.SinglePointAction.SPI_Message = "Only Eu-Document kept Inside Box??  YES  ||  NO"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Yes.jpg"
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_2000_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            Else
                ProcessStepReturn.ActivityToCheck = "IM or SafetyIM or EuDocument kept Inside Box?"
                ProcessStepReturn.SinglePointAction.SPI_Message = "Any of the IMs or Eu-Documents kept Inside Box??  YES  ||  Not Applicable"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Not Applicable.jpg"
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_2000_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_2000_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)
            End If

            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo200_04_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "200"
            ProcessStepReturn.ProcessStep = "Packing"
            ProcessStepReturn.Activity = "UnitKept||IM||EuDOC||QIC||Bracket||VisualCheck||Markings||OrderTag"
            ProcessStepReturn.ToCheck = "Correct||Not Correct||Details"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "200_04_00"
            ProcessStepReturn.ActivityToCheck = "QIC Status?"
            ProcessStepReturn.SinglePointAction.SPI_Message = "How the QIC requested by customer?  Attached with Unit  ||  Seperate"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Attached.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Seperate.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Seperate.jpg"
            ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_2000_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_2000_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo200_05_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "200"
            ProcessStepReturn.ProcessStep = "Packing"
            ProcessStepReturn.Activity = "UnitKept||IM||EuDOC||QIC||Bracket||VisualCheck||Markings||OrderTag"
            ProcessStepReturn.ToCheck = "Correct||Not Correct||Details"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "200_05_00"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Yes.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Not Applicable.jpg"
            If CustOrd.MS_CODE Like "YTA???-??????N*" Then
                ProcessStepReturn.ActivityToCheck = "Mounting Brackets Kept Inside Box?"
                ProcessStepReturn.SinglePointAction.SPI_Message = "Is there any Mounting Brackets kept inside Box?  YES  ||  NO"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Not Applicable.jpg"
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_2000_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_2000_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(1)
            Else
                ProcessStepReturn.ActivityToCheck = "Mounting Brackets Kept Inside Box?"
                ProcessStepReturn.SinglePointAction.SPI_Message = "Is the Mounting Brackets kept inside Box?  YES  ||  NO"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Yes.jpg"
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_2000_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_2000_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)
            End If
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo200_06_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "200"
            ProcessStepReturn.ProcessStep = "Packing"
            ProcessStepReturn.Activity = "UnitKept||IM||EuDOC||QIC||Bracket||VisualCheck||Markings||OrderTag"
            ProcessStepReturn.ToCheck = "Correct||Not Correct||Details"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "200_06_00"
            ProcessStepReturn.ActivityToCheck = "Visually Check the Packing Box?"
            ProcessStepReturn.SinglePointAction.SPI_Message = "Is all items kept inside BOX and closed? Is it good to Ship?  YES  ||  NO"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Yes.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Not Applicable.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Yes.jpg"
            ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_2000_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_2000_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo200_07_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "200"
            ProcessStepReturn.ProcessStep = "Packing"
            ProcessStepReturn.Activity = "UnitKept||IM||EuDOC||QIC||Bracket||VisualCheck||Markings||OrderTag"
            ProcessStepReturn.ToCheck = "Correct||Not Correct||Details"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "200_07_00"
            ProcessStepReturn.ActivityToCheck = "Marking visible outside the Packing Box?"
            ProcessStepReturn.SinglePointAction.SPI_Message = "Is Only KC + Moroccan mark visible and CPL Hidden?  YES  ||  NO"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Yes.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Not Applicable.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Yes.jpg"
            ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_2000_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_2000_Position_Markings.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo200_08_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "200"
            ProcessStepReturn.ProcessStep = "Packing"
            ProcessStepReturn.Activity = "UnitKept||IM||EuDOC||QIC||Bracket||VisualCheck||Markings||OrderTag"
            ProcessStepReturn.ToCheck = "Correct||Not Correct||Details"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "200_08_00"
            ProcessStepReturn.ActivityToCheck = "Signed Order Tag"
            ProcessStepReturn.SinglePointAction.SPI_Message = "Is Signed Order Tag Attached to the Packing Box??  YES  ||  NO"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Attached.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Not Applicable.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Attached.jpg"
            ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_2000_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_2000_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)
            ProcessStepReturn.Result &= "$" & MainForm.Setting.Var_60_2000_Position_Initial.Replace("Initial", Initial)
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function

#End Region

#Region "Common"
    Private Function NeedsHiPot(ByVal NewModel As String, ByVal OldModel As String) As Boolean
        If NewModel.Substring(12, 1) <> OldModel.Substring(12, 1) Then
            Return True
        ElseIf Array.Find(NewModel.Split("/"), Function(x) x = "C1") <> Array.Find(OldModel.Split("/"), Function(x) x = "C1") Then
            Return True
        ElseIf Array.Find(NewModel.Split("/"), Function(x) x = "C2") <> Array.Find(OldModel.Split("/"), Function(x) x = "C2") Then
            Return True
        ElseIf Array.Find(NewModel.Split("/"), Function(x) x = "C3") <> Array.Find(OldModel.Split("/"), Function(x) x = "C3") Then
            Return True
        ElseIf Array.Find(NewModel.Split("/"), Function(x) x = "A") <> Array.Find(OldModel.Split("/"), Function(x) x = "A") Then
            Return True
        Else
            Return False
        End If
    End Function
#End Region

End Class
