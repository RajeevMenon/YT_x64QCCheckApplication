Imports OpenPdfOperation_x64

Public Class YTA_CheckSheet_v1p4

    Dim TmlEntityYGS As New MFG_ENTITY.Op(Link.Mysql_YGSP_ConStr)
    Dim WMsg As New WarningForm

#Region "Common"
    Public Shared Function ProcessStepNo(ByVal StepNo As String, ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try
            Dim ProcessStepReturn As New CheckSheetStep
            Dim stepMethodName As String = "ProcessStepNo" & StepNo
            Dim checker As New YTA_CheckSheet_v1p3
            Dim methodInfo = checker.GetType().GetMethod(stepMethodName)
            If methodInfo IsNot Nothing Then
                ProcessStepReturn = methodInfo.Invoke(checker, New Object() {Initial, CustOrd, ErrMsg})
            Else
                ' Handle unknown step
                ErrMsg = "Invalid StepNo: " & StepNo
            End If
            MainForm.SetInspectionColor(StepNo.ToString, ProcessStepReturn.ProcessNo)
            Return ProcessStepReturn
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    'Public Shared Function ProcessStepNo(ByVal StepNo As String, ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
    '    Try

    '        Dim ProcessStepReturn As New CheckSheetStep

    '        'BUILD
    '        If StepNo = "20_00_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo20_00_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "20_00_01" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo20_00_01(Initial, CustOrd, ErrMsg)
    '        If StepNo = "20_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo20_01_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "30_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo30_01_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "30_02_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo30_02_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "30_03_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo30_03_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "30_04_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo30_04_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "40_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo40_01_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "50_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo50_01_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "60_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo60_01_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "60_02_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo60_02_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "70_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo70_01_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "70_02_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo70_02_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "80_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo80_01_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "90_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo90_01_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "100_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo100_01_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "110_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo110_01_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "110_02_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo110_02_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "120_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo120_01_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "120_02_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo120_02_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "120_03_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo120_03_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "120_04_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo120_04_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "120_05_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo120_05_00(Initial, CustOrd, ErrMsg)
    '        'HIPOT
    '        If StepNo = "130_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo130_01_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "140_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo140_01_00(Initial, CustOrd, ErrMsg)
    '        'YTACRC
    '        If StepNo = "150_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo150_01_00(Initial, CustOrd, ErrMsg)
    '        'FINALASSY
    '        If StepNo = "160_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo160_01_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "170_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo170_01_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "170_02_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo170_02_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "180_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo180_01_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "180_02_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo180_02_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "180_03_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo180_03_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "190_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo190_01_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "190_02_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo190_02_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "190_03_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo190_03_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "190_04_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo190_04_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "190_05_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo190_05_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "190_06_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo190_06_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "190_07_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo190_07_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "190_08_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo190_08_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "190_09_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo190_09_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "190_10_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo190_10_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "190_11_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo190_11_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "190_12_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo190_12_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "190_12_01" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo190_12_01(Initial, CustOrd, ErrMsg)
    '        If StepNo = "190_13_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo190_13_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "190_14_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo190_14_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "190_15_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo190_15_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "190_16_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo190_16_00(Initial, CustOrd, ErrMsg)
    '        'PACKING
    '        If StepNo = "200_01_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo200_01_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "200_02_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo200_02_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "200_03_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo200_03_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "200_04_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo200_04_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "200_05_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo200_05_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "200_05_01" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo200_05_01(Initial, CustOrd, ErrMsg)
    '        If StepNo = "200_06_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo200_06_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "200_07_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo200_07_00(Initial, CustOrd, ErrMsg)
    '        If StepNo = "200_08_00" Then ProcessStepReturn = (New YTA_CheckSheet_v1p3).ProcessStepNo200_08_00(Initial, CustOrd, ErrMsg)


    '        MainForm.SetInspectionColor(StepNo.ToString, ProcessStepReturn.ProcessNo)

    '        Return ProcessStepReturn
    '    Catch ex As Exception
    '        Return Nothing
    '    End Try
    'End Function
#End Region

#Region "Version 1.2"

    'Process Step-20 Initial Inspection
    Public Function ProcessStepNo20_00_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try


            Dim ProcessStepReturn As New CheckSheetStep

            Dim RndAry As New TML_Library.RandomArray
            Dim NewUnitSrNo = RndAry.RandomStringArraySERIAL(CustOrd.SERIAL_NO, 4)

            ProcessStepReturn.ProcessNo = "20"
            ProcessStepReturn.ProcessStep = "Order Tag Correctness"
            ProcessStepReturn.Activity = "Serial No. correctness check"
            ProcessStepReturn.ToCheck = "Correct Serial Number of unit to be Prepared"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.UserIput

            ProcessStepReturn.StepNo = "20_00_00"
            ProcessStepReturn.ActivityToCheck = "Check correct Serial No. is used"
            ProcessStepReturn.UserInputAction.UserActionMessage = "Choose the Serial number printed In the Order Tag"
            ProcessStepReturn.UserInputAction.UserInputList = NewUnitSrNo
            ProcessStepReturn.UserInputAction.UserInputCorrect = CustOrd.SERIAL_NO

            Dim WriteFields As New List(Of OpenPdfOperation_x64.WriteField)
            AddWriteField("Index_No", CustOrd.INDEX_NO, WriteFields:=WriteFields)
            AddWriteField("Sr_No_After", CustOrd.SERIAL_NO, WriteFields:=WriteFields)
            AddWriteField("Sales_Order_No", $"SO-No.:{CustOrd.PROD_NO} Line-No.:{CustOrd.LINE_NO}", WriteFields:=WriteFields)
            AddWriteField("MS_Code_Before", CustOrd.MS_CODE_BEFORE, WriteFields:=WriteFields)
            AddWriteField("MS_Code_After", CustOrd.MS_CODE, WriteFields:=WriteFields)
            Dim ResultJson = AddResultTexts(WriteFields, ErrMsg:=ErrMsg)
            If ErrMsg.Length > 0 Then
                WMsg.Message = $"AddResultTexts() Error for Step:{ProcessStepReturn.StepNo}: {ErrMsg}"
                WMsg.ShowDialog()
                Return Nothing
            End If
            ProcessStepReturn.Result = Newtonsoft.Json.JsonConvert.SerializeObject(ResultJson, Newtonsoft.Json.Formatting.Indented)
            Return ProcessStepReturn

        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo20_00_01(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try


            Dim ProcessStepReturn As New CheckSheetStep

            ProcessStepReturn.ProcessNo = "20"
            ProcessStepReturn.ProcessStep = "Before Modification Unit"
            ProcessStepReturn.Activity = "Before Modification Serial Number?"
            ProcessStepReturn.ToCheck = "Correct Serial Number of unit to be Prepared"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.MakeUsrInpt

            ProcessStepReturn.StepNo = "20_00_01"
            ProcessStepReturn.ActivityToCheck = "Before Modification Unit"
            ProcessStepReturn.MakeUserInputAction.UserActionMessage = "Before Modification item Serial Number?"
            ProcessStepReturn.MakeUserInputAction.UserInputOld = CustOrd.SERIAL_NO_BEFORE
            ProcessStepReturn.MakeUserInputAction.UserInputSaveConnectionString = MainForm.Setting.Var_03_MySql_YGSP
            ProcessStepReturn.MakeUserInputAction.UserInputSaveTableName = "cust_ord"
            ProcessStepReturn.MakeUserInputAction.UserInputSaveTableField = "SERIAL_NO_BEFORE"
            ProcessStepReturn.MakeUserInputAction.UserInputWhereFieldName = "INDEX_NO"
            ProcessStepReturn.MakeUserInputAction.UserInputWhereFieldValue = CustOrd.INDEX_NO

            ProcessStepReturn.Result = "" 'This is blank as 'MakeUserInputAction' is for saving user input to DB only
            Return ProcessStepReturn

        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo20_01_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

#Region "CustOrd Read"

            If Not IsDate(MainForm.CustOrd.ACTUAL_START_DATE) Then
                Dim Sdate = Date.Today.ToString("yyyy-MM-dd")
                Dim Stime = Now.ToString("hh:mm:ss tt")

                Dim SqlU(0) As String
                SqlU(0) = "UPDATE cust_ord SET ACTUAL_START_DATE='" & Sdate & "',ACTUAL_START_TIME='" & Stime & "' WHERE INDEX_NO='" & MainForm.CustOrd.INDEX_NO & "';"
                If Not IsDate(MainForm.CoHeader.ACT_START_DATE) Then
                    ReDim Preserve SqlU(1)
                    SqlU(1) = "UPDATE co_register SET ACT_START_DATE='" & Sdate & "' WHERE ORDER_NO='" & MainForm.CustOrd.PROD_NO & "' AND LINE_NO='" & MainForm.CustOrd.LINE_NO & "';"
                End If

                Dim TmlEntityYGS As New MFG_ENTITY.Op(MainForm.Setting.Var_03_MySql_YGSP)
                TmlEntityYGS.ExecuteTransactionQuery(SqlU, ErrMsg:=ErrMsg)
                If ErrMsg.Length > 0 Then
                    WMsg.Message = "Cannot Update ACTUAL START DATE for the job!"
                    WMsg.ShowDialog()
                    Return Nothing
                End If
            End If

            MainForm.CustOrd = TmlEntityYGS.GetDatabaseTableAs_Object(Of POCO_YGSP.cust_ord)("INDEX_NO", MainForm.CustOrd.INDEX_NO, ErrMsg:=ErrMsg)
            CustOrd = MainForm.CustOrd
            If ErrMsg.Length > 0 Then
                WMsg.Message = "Cannot Read INDEX NO. Data"
                WMsg.ShowDialog()
                Return Nothing
            End If

            If Not IsNumeric(CustOrd.SHIP_LOT.ToString) Then
                WMsg.Message = "SHIP_LOT is Blank in Cust_Ord Table!"
                WMsg.ShowDialog()
                Return Nothing
            End If

#End Region

            Dim ProcessStepReturn As New CheckSheetStep

            ProcessStepReturn.ProcessNo = "20"
            ProcessStepReturn.ProcessStep = "Correct Unit Picked"
            ProcessStepReturn.Activity = "Picked Unit check"
            ProcessStepReturn.ToCheck = "Correct MSCODE of Capsule or Level 2.0 unit Picked"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.UserIput
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "20_01_00"
            ProcessStepReturn.ActivityToCheck = "Check unit Picked for Production"
            ProcessStepReturn.UserInputAction.UserActionMessage = "MS-CODE of Capsule/Level 2.0 unit Picked?"

            Dim Var1, Var2, Var3 As New String("")
            Dim Mscode As String = CustOrd.MS_CODE_BEFORE

            If Mscode Like "YTA6??-???????*" Then Var1 = Mscode.Remove(3, 1).Insert(3, "7")
            If Mscode Like "YTA7??-???????*" Then Var1 = Mscode.Remove(3, 1).Insert(3, "6")

            If Mscode Like "YTA???-J??????*" Then Var2 = Mscode.Remove(7, 1).Insert(7, "F")
            If Mscode Like "YTA???-F??????*" Then Var2 = Mscode.Remove(7, 1).Insert(7, "J")

            If Mscode Like "YTA???-???A???*" Then Var2 = Mscode.Remove(10, 1).Insert(7, "C")
            If Mscode Like "YTA???-???C???*" Then Var2 = Mscode.Remove(10, 1).Insert(7, "A")

            ProcessStepReturn.UserInputAction.UserInputList = {Mscode, Var1, Var2, Var3}
            ProcessStepReturn.UserInputAction.UserInputCorrect = CustOrd.MS_CODE_BEFORE

            Dim WriteFields As New List(Of OpenPdfOperation_x64.WriteField)
            Dim ResultValue As String = "[Correct] ✓"
            AddWriteField(ProcessStepReturn.StepNo, ResultValue, WriteFields:=WriteFields)
            AddWriteField("Sr_No_Before", CustOrd.SERIAL_NO_BEFORE, WriteFields:=WriteFields) 'Link to previous step
            AddWriteField("Prod_Date", CustOrd.ACTUAL_START_DATE, WriteFields:=WriteFields) 'Actual Start Date updated in this step
            AddWriteField("20_01_00_01", MainForm.Initial, WriteFields:=WriteFields)
            Dim ResultJson = AddResultTexts(WriteFields, ErrMsg:=ErrMsg)
            If ErrMsg.Length > 0 Then
                WMsg.Message = $"AddResultTexts() Error for Step:{ProcessStepReturn.StepNo}: {ErrMsg}"
                WMsg.ShowDialog()
                Return Nothing
            End If
            ProcessStepReturn.Result = Newtonsoft.Json.JsonConvert.SerializeObject(ResultJson, Newtonsoft.Json.Formatting.Indented)

        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    'Process Step-30 Plate Inspection
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

            Dim ResultValue_30_01_00 As String = "[Correct] ✓"
            Dim ResultValue_30_01_01 As String = ""
            If (CustOrd.SERIAL_NO Like "Y3*" Or CustOrd.SERIAL_NO Like "S5???????-M") And CustOrd.EU_COUNTRY = "SA" And Link.PlantID = "5Q00" Then
                ProcessStepReturn.UserInputAction.UserInputCorrect = "Made In KSA"
                ResultValue_30_01_01 = "Made in KSA"
            ElseIf (CustOrd.SERIAL_NO Like "Y3*" Or CustOrd.SERIAL_NO Like "S5*") And CustOrd.EU_COUNTRY <> "SA" And Link.PlantID = "5Q00" Then
                ProcessStepReturn.UserInputAction.UserInputCorrect = "Made In China"
                ResultValue_30_01_01 = "Made in China"
            ElseIf (CustOrd.SERIAL_NO Like "Y4*" Or CustOrd.SERIAL_NO Like "90*") And CustOrd.MS_CODE Like "*/JP*" And Link.PlantID = "6Z00" Then
                ProcessStepReturn.UserInputAction.UserInputCorrect = "Made In Japan"
                ResultValue_30_01_01 = "Made in Japan"
            Else
                WMsg.Message = $"Error for Step:{ProcessStepReturn.StepNo}: SerialNo:{CustOrd.SERIAL_NO}, EU-Country:{CustOrd.EU_COUNTRY} does not have COO rule in QCC App."
                WMsg.ShowDialog()
                Return Nothing
            End If

            Dim WriteFields As New List(Of OpenPdfOperation_x64.WriteField)
            'Dim ResultValue As String = "[Correct] ✓"
            AddWriteField(ProcessStepReturn.StepNo, ResultValue_30_01_00, WriteFields:=WriteFields)
            AddWriteField("30_01_01", ResultValue_30_01_01, WriteFields:=WriteFields)
            Dim ResultJson = AddResultTexts(WriteFields, ErrMsg:=ErrMsg)
            If ErrMsg.Length > 0 Then
                WMsg.Message = $"AddResultTexts() Error for Step:{ProcessStepReturn.StepNo}: {ErrMsg}"
                WMsg.ShowDialog()
                Return Nothing
            End If

            ProcessStepReturn.Result = Newtonsoft.Json.JsonConvert.SerializeObject(ResultJson, Newtonsoft.Json.Formatting.Indented)
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
            ProcessStepReturn.ToCheck = "Plate data correct?"
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

            Dim ResultValue_30_04_00 As String = ""
            Dim ResultValue_30_04_01 As String = ""
            If CustOrd.SERIAL_NO_BEFORE Like "S5WC*" Or CustOrd.SERIAL_NO_BEFORE Like "S5X1*" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "NamePlate_without_CE.jpg"
                ResultValue_30_04_00 = "[Yes] ✓"
                ResultValue_30_04_01 = "No CE Mark"
            Else
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "NamePlate_with_CE.jpg"
                ResultValue_30_04_00 = "[Yes] ✓"
                ResultValue_30_04_01 = "with CE Mark"
            End If

            Dim WriteFields As New List(Of OpenPdfOperation_x64.WriteField)
            'Dim ResultValue As String = "[Correct] ✓"
            AddWriteField(ProcessStepReturn.StepNo, ResultValue_30_04_00, WriteFields:=WriteFields)
            AddWriteField("30_04_01", ResultValue_30_04_01, WriteFields:=WriteFields)
            AddWriteField("30_04_00_01", MainForm.Initial, WriteFields:=WriteFields)
            Dim ResultJson = AddResultTexts(WriteFields, ErrMsg:=ErrMsg)
            If ErrMsg.Length > 0 Then
                WMsg.Message = $"AddResultTexts() Error for Step:{ProcessStepReturn.StepNo}: {ErrMsg}"
                WMsg.ShowDialog()
                Return Nothing
            End If

            ProcessStepReturn.Result = Newtonsoft.Json.JsonConvert.SerializeObject(ResultJson, Newtonsoft.Json.Formatting.Indented)

            Return ProcessStepReturn
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    'Process Step-40 Plate Mounting
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

            Dim WriteFields As New List(Of OpenPdfOperation_x64.WriteField)
            Dim ResultValue_40_01_00 As String = "[Correct] ✓"
            Dim ResultValue_40_01_01 As String = MainForm.DataPlateCorrect
            AddWriteField(ProcessStepReturn.StepNo, ResultValue_40_01_00, WriteFields:=WriteFields)
            AddWriteField("40_01_01", ResultValue_40_01_01, WriteFields:=WriteFields)
            Dim ResultJson = AddResultTexts(WriteFields, ErrMsg:=ErrMsg)
            If ErrMsg.Length > 0 Then
                WMsg.Message = $"AddResultTexts() Error for Step:{ProcessStepReturn.StepNo}: {ErrMsg}"
                WMsg.ShowDialog()
                Return Nothing
            End If

            ProcessStepReturn.Result = Newtonsoft.Json.JsonConvert.SerializeObject(ResultJson, Newtonsoft.Json.Formatting.Indented)

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
                ElseIf CustOrd.MS_CODE Like "YTA*/C3*" Then
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "BOUT_C3_NORMAL.jpg"
                    ProcessStepReturn.Result = "Circle-16,58.5,25.8"
                Else
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "BOUT_C3_NORMAL.jpg"
                    ProcessStepReturn.Result = "Circle-16,62.5,25.8"
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
                ElseIf CustOrd.MS_CODE_BEFORE Like "YTA[67]10-?????N?*" And CustOrd.MS_CODE Like "YTA[67]10-?????N?*" Then
                    ProcessStepReturn.SinglePointAction.SPI_Message = "[Indicator] Confirm Indicator Status?"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "No Change.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Added.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "No Change.jpg"
                    ProcessStepReturn.Result = "Circle-16,62,28.8"
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
                ElseIf CustOrd.MS_CODE_BEFORE Like "YTA[67]10-?????N?*" And CustOrd.MS_CODE Like "YTA[67]10-?????N?*" Then
                    ProcessStepReturn.SinglePointAction.SPI_Message = "[Nut/Stud] Confirm Nut/Stud usage for Indicator/Board fixing to case"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\70\" & "Indicator_W_Nut.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\70\" & "NoneInd_w_Stud.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "NoneInd_w_Stud.jpg"
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
            ElseIf CustOrd.MS_CODE Like "YTA[67]10-?????N?*" Then
                ProcessStepReturn.ProcessNo = "90"
                ProcessStepReturn.ProcessStep = "Modification / Assembly Checks"
                ProcessStepReturn.Activity = "Blind Cover Assembly"
                ProcessStepReturn.ToCheck = "Clean & No galling"
                ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
                ProcessStepReturn.Initial = Initial

                ProcessStepReturn.StepNo = "90_01_00"
                ProcessStepReturn.ActivityToCheck = "Blind Cover Assembling"
                ProcessStepReturn.SinglePointAction.SPI_Message = "[Indicator Side] Confirm Blind Cover Quality after Assembling?"
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
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\100\" & "WithA.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\100\" & "WithoutA.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "WithA.jpg"
                    ProcessStepReturn.Result = "Circle-16,47.4,33.3"
                ElseIf CustOrd.MS_CODE_BEFORE Like "YTA[67]10-???????*/A*" And CustOrd.MS_CODE Like "YTA[67]10-???????*/A*" Then
                    ProcessStepReturn.SinglePointAction.SPI_Message = "[/A Option] Lightning Arrestor Installed?"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\100\" & "WithA.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\100\" & "WithoutA.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "WithA.jpg"
                    ProcessStepReturn.Result = "Circle-16,47.4,33.3"
                ElseIf CustOrd.MS_CODE_BEFORE Like "YTA[67]10-???????*/A*" And Not CustOrd.MS_CODE Like "YTA[67]10-???????*/A*" Then
                    ProcessStepReturn.SinglePointAction.SPI_Message = "[/A Option] Confirm Lightning Arrestor Status?"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\100\" & "WithA.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\100\" & "WithoutA.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "WithoutA.jpg"
                    ProcessStepReturn.Result = "Circle-16,55,33.3"
                ElseIf Not CustOrd.MS_CODE_BEFORE Like "YTA[67]10-???????*/A*" And Not CustOrd.MS_CODE Like "YTA[67]10-???????*/A*" Then
                    ProcessStepReturn.SinglePointAction.SPI_Message = "[/A Option] Confirm Lightning Arrestor Status?"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\100\" & "WithoutA.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\100\" & "WithA.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "WithoutA.jpg"
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
            ProcessStepReturn.Activity = "Ground Screws attach to Unit"
            ProcessStepReturn.ToCheck = "INTERNAL & EXTERNAL GROUND SCREW ON CASE"
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
            ProcessStepReturn.Activity = "QR Nameplate Fixing"
            ProcessStepReturn.ToCheck = "Name plate with QR-Image is prepared"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "110_02_00"
            ProcessStepReturn.ActivityToCheck = "Name plate with QR Image"
            ProcessStepReturn.SinglePointAction.SPI_Message = "Confirm Nameplate with QR image is prepared"
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

    'HIPOT
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

    'CRC
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

    'FINAL ASSY & CHECK
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
            If MainForm.CustOrd.MS_CODE_BEFORE Like "YTA???-??????N*" Then
                ProcessStepReturn.ActivityToCheck = "Bracket Assy"
                ProcessStepReturn.SinglePointAction.SPI_Message = "Is Mounting Bracket [Removed] or [Not-changed] in Before modification unit?"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Removed.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "No Change.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "No Change.jpg"
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_170_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_170_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(1)
            ElseIf (MainForm.CustOrd.MS_CODE_BEFORE.Substring(13, 1) = MainForm.CustOrd.MS_CODE.Substring(13, 1)) Then
                ProcessStepReturn.ActivityToCheck = "Bracket Assy"
                ProcessStepReturn.SinglePointAction.SPI_Message = "Is Mounting Bracket [Removed] or [Not-changed] in Before modification unit?"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Removed.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "No Change.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "No Change.jpg"
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_170_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_170_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(1)
            Else
                ProcessStepReturn.ActivityToCheck = "Bracket Assy"
                ProcessStepReturn.SinglePointAction.SPI_Message = "Is Mounting Bracket [Removed] or [Not-changed] in Before modification unit?"
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
                If CustOrd.MS_CODE.Substring(13, 1) <> "N" Then
                    ProcessStepReturn.ActivityToCheck = "Bracket Assy"
                    ProcessStepReturn.SinglePointAction.SPI_Message = "Confirm Mounting Bracket added to Final Unit?"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "No Change.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Added.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "No Change.jpg" 'Answer No change but tickmark on added as per QCC
                    ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_170_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                    ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_170_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)
                    ProcessStepReturn.Result &= "$" & MainForm.Setting.Var_60_170_Position_Initial.Replace("Initial", MainForm.Initial)
                ElseIf CustOrd.MS_CODE.Substring(13, 1) = "N" Then
                    ProcessStepReturn.ActivityToCheck = "Bracket Assy"
                    ProcessStepReturn.SinglePointAction.SPI_Message = "Confirm Mounting Bracket added to Final Unit?"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "No Change.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Added.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "No Change.jpg"
                    ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_170_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                    ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_170_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(1)
                    ProcessStepReturn.Result &= "$" & MainForm.Setting.Var_60_170_Position_Initial.Replace("Initial", MainForm.Initial)
                End If
            Else
                If CustOrd.MS_CODE.Substring(13, 1) <> "N" Then
                    ProcessStepReturn.ActivityToCheck = "Bracket Assy"
                    ProcessStepReturn.SinglePointAction.SPI_Message = "Confirm Mounting Bracket added to Final Unit?"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "No Change.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Added.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Added.jpg"
                    ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_170_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                    ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_170_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)
                    ProcessStepReturn.Result &= "$" & MainForm.Setting.Var_60_170_Position_Initial.Replace("Initial", MainForm.Initial)
                ElseIf CustOrd.MS_CODE.Substring(13, 1) = "N" Then
                    ProcessStepReturn.ActivityToCheck = "Bracket Assy"
                    ProcessStepReturn.SinglePointAction.SPI_Message = "Confirm Mounting Bracket added to Final Unit?"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Removed.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Added.jpg"
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Removed.jpg"
                    ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_170_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                    ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_170_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(1)
                    ProcessStepReturn.Result &= "$" & MainForm.Setting.Var_60_170_Position_Initial.Replace("Initial", MainForm.Initial)
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

            Dim QicFolerPath As String = System.IO.Path.Combine(MainForm.Setting.Var_06_DocsStore, "Production Complete Documents\QICDOC\")
            If Not (System.IO.Directory.Exists(QicFolerPath & CustOrd.PROD_NO)) Then
                System.IO.Directory.CreateDirectory(QicFolerPath & CustOrd.PROD_NO)
            End If
            QicFolerPath = QicFolerPath & MainForm.CustOrd.PROD_NO & "\"
            Dim lotNo As String = MainForm.CustOrd.SHIP_LOT
            If Not IsNumeric(lotNo) Then
                ErrMsg = "Lot No. is not readable from database."
                Return Nothing
            End If
            If Not (System.IO.Directory.Exists(QicFolerPath & lotNo)) Then
                System.IO.Directory.CreateDirectory(QicFolerPath & lotNo)
            End If
            QicFolerPath = QicFolerPath & lotNo & "\"

            Dim QicFile As String = System.IO.Path.Combine(MainForm.Setting.Var_06_DocsStore, "Production Complete Documents\YTA_QIC\") & MainForm.CustOrd.SERIAL_NO
            Dim TargetFile As String = QicFolerPath & "LINE-" & MainForm.CustOrd.LINE_NO & "-" & MainForm.CustOrd.INDEX_NO & "(1)-CQIC"
            If System.IO.File.Exists(QicFile & ".pdf") And CustOrd.MTC = "GO" Then
                If Not System.IO.File.Exists(TargetFile & ".pdf") Then
                    System.IO.File.Copy(QicFile & ".pdf", TargetFile & ".pdf", True)
                End If
            Else
                Dim ExcelWorkBook As Spire.Xls.Workbook = New Spire.Xls.Workbook()
                ExcelWorkBook.LoadFromFile(QicFile & ".xls")
                ExcelWorkBook.Worksheets(0).PageSetup.IsFitToPage = True
                ExcelWorkBook.Worksheets(0).SaveToPdf(TargetFile & ".pdf")
            End If

            ProcessStepReturn.ViewDocAction.PdfPath_DocumentCheck = TargetFile & ".pdf"

            Return ProcessStepReturn

        Catch ex As Exception
            'Dim ProcessStepReturn As New CheckSheetStep
            'ProcessStepReturn.ProcessNo = "180"
            'ProcessStepReturn.ProcessStep = "Print QIC"
            'ProcessStepReturn.Activity = "Print Certificates"
            'ProcessStepReturn.ToCheck = "Printed YES || NO"
            'ProcessStepReturn.Method = CheckSheetStep.MethodOption.DocumentCheck
            'ProcessStepReturn.Initial = Initial

            'ProcessStepReturn.StepNo = "180_01_00"
            'ProcessStepReturn.ActivityToCheck = "Print QIC"
            'ProcessStepReturn.ViewDocAction.DocumentCheckMessage = "Runtime Error:" & ex.Message
            WMsg.Message = "ProcessStepNo180_01_00() Exception:" & ex.Message
            WMsg.ShowDialog()
            Return Nothing
            'Return ProcessStepReturn
        End Try
    End Function
    Public Function ProcessStepNo180_02_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "180"
            ProcessStepReturn.ProcessStep = "Print Device Information Sheet"
            ProcessStepReturn.Activity = "Print Certificates"
            ProcessStepReturn.ToCheck = "Printed YES || NO"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.DocumentCheck
            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.Result = ""
            ProcessStepReturn.StepNo = "180_02_00"
            ProcessStepReturn.ActivityToCheck = "Print Device Information"
            ProcessStepReturn.ViewDocAction.DocumentCheckMessage = "Printed correctly YES || NO ?"
            'ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_180_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            'ProcessStepReturn.Result &= "$" & MainForm.Setting.Var_60_180_Position_Initial.Replace("Initial", MainForm.Initial)
            'ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_180_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)

            Dim TargetFile As String = ""
            If CustOrd.MS_CODE Like "YTA[67]10-F*" Then
                Dim DeviceIDFolerPath As String = MainForm.Setting.Var_06_DocsStore & "\Production Complete Documents\Device ID\"
                If Not (System.IO.Directory.Exists(DeviceIDFolerPath & CustOrd.PROD_NO)) Then
                    System.IO.Directory.CreateDirectory(DeviceIDFolerPath & CustOrd.PROD_NO)
                End If
                DeviceIDFolerPath = DeviceIDFolerPath & CustOrd.PROD_NO & "\"
                Dim lotNo As String = CustOrd.SHIP_LOT
                If Not (System.IO.Directory.Exists(DeviceIDFolerPath & lotNo)) Then
                    System.IO.Directory.CreateDirectory(DeviceIDFolerPath & lotNo)
                End If
                DeviceIDFolerPath = DeviceIDFolerPath & lotNo & "\"
                TargetFile = DeviceIDFolerPath & CustOrd.INDEX_NO & "-DeviceID-Lot" & lotNo
                Dim QicPath As String = MainForm.Setting.Var_06_DocsStore & "\Production Complete Documents\YTA_QIC\"
                If System.IO.File.Exists(QicPath & CustOrd.SERIAL_NO & "_DEV.pdf") Then
                    Dim TargetFilePDF As String = TargetFile & ".pdf"
                    If Not System.IO.File.Exists(TargetFilePDF) Then
                        System.IO.File.Copy(QicPath & CustOrd.SERIAL_NO & "_DEV.pdf", TargetFilePDF, True)
                    End If
                Else
                    Dim QicFile As String = QicPath & CustOrd.SERIAL_NO & "_DEV.xls"
                    'Dim ExpExl As New ExportExcel
                    'ExpExl.ExcelSheet1_SaveAsPdf(QicFile, TargetFile, ErrMsg)
                    'If Len(ErrMsg) > 0 Then
                    '    Exit Function
                    'End If
                    Dim ExcelWorkBook As Spire.Xls.Workbook = New Spire.Xls.Workbook()
                    ExcelWorkBook.LoadFromFile(QicFile)
                    ExcelWorkBook.Worksheets(0).PageSetup.IsFitToPage = True
                    ExcelWorkBook.Worksheets(0).SaveToPdf(TargetFile & ".pdf")

                    System.IO.File.Copy(TargetFile & ".pdf", QicPath & CustOrd.SERIAL_NO & "_DEV.pdf", True)

                End If
            Else
                Dim DeviceIDFolerPath As String = MainForm.Setting.Var_06_DocsStore & "\Production Complete Documents\Device ID\"
                TargetFile = DeviceIDFolerPath & "Template\No_DeviceID"
            End If

            ProcessStepReturn.ViewDocAction.PdfPath_DocumentCheck = TargetFile & ".pdf"

            Return ProcessStepReturn

        Catch ex As Exception
            WMsg.Message = "ProcessStepNo180_02_00() Exception:" & ex.Message
            WMsg.ShowDialog()
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo180_03_00(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "180"
            ProcessStepReturn.ProcessStep = "Print Label"
            ProcessStepReturn.Activity = "Print Label"
            ProcessStepReturn.ToCheck = "Printed YES || NO"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.DocumentCheck
            ProcessStepReturn.Initial = Initial
            ProcessStepReturn.Result = ""
            ProcessStepReturn.StepNo = "180_03_00"
            ProcessStepReturn.ActivityToCheck = "Print LABEL"
            ProcessStepReturn.ViewDocAction.DocumentCheckMessage = "Box Label Printed correctly?"

            Dim BlankDoc As String = MainForm.Setting.Var_06_DocsStore & "\Production Complete Documents\YTA_Label\Template\YMA_YtaLabel.pdf"
            Dim FinalDoc As String = MainForm.Setting.Var_06_DocsStore & "\Production Complete Documents\YTA_Label\" & CustOrd.INDEX_NO & ".pdf"
            If System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(FinalDoc)) Then
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(FinalDoc))
            End If

            Dim COO As String = ""
            If (CustOrd.SERIAL_NO Like "Y3*" Or CustOrd.SERIAL_NO Like "S5???????-M") And CustOrd.EU_COUNTRY = "SA" Then
                COO = "Made In KSA"
            ElseIf CustOrd.MS_CODE Like "*/JP*" Then
                COO = "Made In Japan"
            Else
                COO = "Made In China"
            End If
            If System.IO.File.Exists(BlankDoc) Then
                Dim P_Doc = OpenPdfOperation_x64.FileOp.GetDocument(BlankDoc, ErrMsg)
                Dim FontToWrite As OpenPdfOperation_x64.FileOp.FontName = OpenPdfOperation_x64.FileOp.FontName.Arial

                Dim WPS As New List(Of OpenPdfOperation_x64.WriteTextParameters)

                Dim WriteS As Integer = 10
                Dim WriteX As Double = 54
                Dim WriteY As Double = 23.2
                Dim WP1 As New OpenPdfOperation_x64.WriteTextParameters With {
                                            .Name = FontToWrite,
                                            .Colour = OpenPdfOperation_x64.FileOp.FontColor.Black,
                                            .Style = OpenPdfOperation_x64.FileOp.FontStyle.Regular,
                                            .TextSize = Integer.Parse(WriteS),
                                            .TextValue = CustOrd.KCC_DATE,
                                            .X_Position = CDbl(WriteX / 100),
                                            .Y_Position = CDbl(WriteY / 100)}

                WriteS = 10
                WriteX = 51
                WriteY = 37.5
                Dim WP2 As New OpenPdfOperation_x64.WriteTextParameters With {
                                            .Name = FontToWrite,
                                            .Colour = OpenPdfOperation_x64.FileOp.FontColor.Black,
                                            .Style = OpenPdfOperation_x64.FileOp.FontStyle.Regular,
                                            .TextSize = Integer.Parse(WriteS),
                                            .TextValue = COO,
                                            .X_Position = CDbl(WriteX / 100),
                                            .Y_Position = CDbl(WriteY / 100)}

                WPS.Add(WP1)
                WPS.Add(WP2)
                OpenPdfOperation_x64.FileOp.PDF_WriteText(P_Doc, WPS, ErrMsg)

                P_Doc.Save(FinalDoc)
            End If
            ProcessStepReturn.ViewDocAction.PdfPath_DocumentCheck = FinalDoc

            Return ProcessStepReturn

        Catch ex As Exception
            WMsg.Message = "ProcessStepNo180_03_00() Exception:" & ex.Message
            WMsg.ShowDialog()
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
            WMsg.Message = "ProcessStepNo190_02_00() Exception:" & ex.Message
            WMsg.ShowDialog()
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
            If Mscode Like "YTA???-??2*" And CustOrd.ORD_INST_CONTECT1_W24 <> "RTD PT100 4WIRE" Then
                Range_1 = "S1: " & CustOrd.ORD_INST_MIN_T70 & " TO " & CustOrd.ORD_INST_MAX_T70 & " " & CustOrd.UNIT_T70
                Range_2 = "S2: " & CustOrd.ORD_INST_MIN_T71 & " TO " & CustOrd.ORD_INST_MAX_T71 & " " & CustOrd.UNIT_T71
                Dim I1 As Integer = 0
                Dim I2 As Integer = 0
                If Integer.TryParse(CustOrd.ORD_INST_MIN_T70, I1) And Integer.TryParse(CustOrd.ORD_INST_MAX_T70, I2) And
                     Integer.TryParse(CustOrd.ORD_INST_MIN_T71, I1) And Integer.TryParse(CustOrd.ORD_INST_MAX_T71, I2) Then
                    Var1 = "S1: " & CustOrd.ORD_INST_MIN_T70 & " TO " & (Integer.Parse(CustOrd.ORD_INST_MAX_T70) + 25).ToString & " " & CustOrd.UNIT_T70
                    Var1 &= "|| S2: " & CustOrd.ORD_INST_MIN_T71 & " TO " & (Integer.Parse(CustOrd.ORD_INST_MAX_T71) - 25).ToString & " " & CustOrd.UNIT_T71
                    Var2 = "S1: " & CustOrd.ORD_INST_MIN_T70 & " TO " & (Integer.Parse(CustOrd.ORD_INST_MAX_T70) - 25).ToString & " " & CustOrd.UNIT_T70
                    Var2 &= "|| S2: " & CustOrd.ORD_INST_MIN_T71 & " TO " & (Integer.Parse(CustOrd.ORD_INST_MAX_T71) + 25).ToString & " " & CustOrd.UNIT_T71
                    Var3 = "S1: " & CustOrd.ORD_INST_MIN_T70 & " TO " & (Integer.Parse(CustOrd.ORD_INST_MAX_T70) + 50).ToString & " " & CustOrd.UNIT_T70
                    Var3 &= "|| S2: " & CustOrd.ORD_INST_MIN_T71 & " TO " & (Integer.Parse(CustOrd.ORD_INST_MAX_T71) - 50).ToString & " " & CustOrd.UNIT_T71
                Else
                    Var1 = "S1: " & CustOrd.ORD_INST_MIN_T70 & " TO " & (Decimal.Parse(CustOrd.ORD_INST_MAX_T70) + 25).ToString & " " & CustOrd.UNIT_T70
                    Var1 &= "|| S2: " & CustOrd.ORD_INST_MIN_T71 & " TO " & (Decimal.Parse(CustOrd.ORD_INST_MAX_T71) - 25).ToString & " " & CustOrd.UNIT_T71
                    Var2 = "S1: " & CustOrd.ORD_INST_MIN_T70 & " TO " & (Decimal.Parse(CustOrd.ORD_INST_MAX_T70) - 25).ToString & " " & CustOrd.UNIT_T70
                    Var2 &= "|| S2: " & CustOrd.ORD_INST_MIN_T71 & " TO " & (Decimal.Parse(CustOrd.ORD_INST_MAX_T71) + 25).ToString & " " & CustOrd.UNIT_T71
                    Var3 = "S1: " & CustOrd.ORD_INST_MIN_T70 & " TO " & (Decimal.Parse(CustOrd.ORD_INST_MAX_T70) + 50).ToString & " " & CustOrd.UNIT_T70
                    Var3 &= "|| S2: " & CustOrd.ORD_INST_MIN_T71 & " TO " & (Decimal.Parse(CustOrd.ORD_INST_MAX_T71) - 50).ToString & " " & CustOrd.UNIT_T71
                End If
            Else
                Range_1 = "S1: " & CustOrd.ORD_INST_MIN_T70 & " TO " & CustOrd.ORD_INST_MAX_T70 & " " & CustOrd.UNIT_T70
                Range_2 = "S2: "
                Dim I1 As Integer = 0
                Dim I2 As Integer = 0
                If Integer.TryParse(CustOrd.ORD_INST_MIN_T70, I1) And Integer.TryParse(CustOrd.ORD_INST_MAX_T70, I2) Then
                    Var1 = "S1: " & CustOrd.ORD_INST_MIN_T70 & " TO " & (Integer.Parse(CustOrd.ORD_INST_MAX_T70) + 25).ToString & " " & CustOrd.UNIT_T70
                    Var1 &= "|| S2: "
                    Var2 = "S1: " & CustOrd.ORD_INST_MIN_T70 & " TO " & (Integer.Parse(CustOrd.ORD_INST_MAX_T70) - 25).ToString & " " & CustOrd.UNIT_T70
                    Var2 &= "|| S2: "
                    Var3 = "S1: " & CustOrd.ORD_INST_MIN_T70 & " TO " & (Integer.Parse(CustOrd.ORD_INST_MAX_T70) + 50).ToString & " " & CustOrd.UNIT_T70
                    Var3 &= "|| S2: "
                Else
                    Var1 = "S1: " & CustOrd.ORD_INST_MIN_T70 & " TO " & (Decimal.Parse(CustOrd.ORD_INST_MAX_T70) + 25).ToString & " " & CustOrd.UNIT_T70
                    Var1 &= "|| S2: "
                    Var2 = "S1: " & CustOrd.ORD_INST_MIN_T70 & " TO " & (Decimal.Parse(CustOrd.ORD_INST_MAX_T70) - 25).ToString & " " & CustOrd.UNIT_T70
                    Var2 &= "|| S2: "
                    Var3 = "S1: " & CustOrd.ORD_INST_MIN_T70 & " TO " & (Decimal.Parse(CustOrd.ORD_INST_MAX_T70) + 50).ToString & " " & CustOrd.UNIT_T70
                    Var3 &= "|| S2: "
                End If
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
    Public Function ProcessStepNo190_12_01(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep

            ProcessStepReturn.ProcessNo = "190"
            ProcessStepReturn.ProcessStep = "QR Label"
            ProcessStepReturn.Activity = "Scan QR to ensure correctmess of Data"
            ProcessStepReturn.ToCheck = "Correct Nameplate with QR is Prepared"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.MakeUsrInpt

            ProcessStepReturn.StepNo = "190_12_01"
            ProcessStepReturn.ActivityToCheck = "QR code correctness"
            ProcessStepReturn.MakeUserInputAction.UserActionMessage = "Scan QR on the Nameplate of the unit"
            ProcessStepReturn.MakeUserInputAction.UserInputOld = ""
            ProcessStepReturn.MakeUserInputAction.UserInputSaveConnectionString = MainForm.Setting.Var_03_MySql_YGSP
            ProcessStepReturn.MakeUserInputAction.UserInputSaveTableName = "QR_CHECK" 'only check for "MFR:YOKOGAWA;CAT:EXT S/N;S/N:"
            ProcessStepReturn.MakeUserInputAction.UserInputSaveTableField = ""

            ProcessStepReturn.Result = "" 'This is blank as 'MakeUserInputAction' is for saving user input to DB only
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
            ProcessStepReturn.SinglePointAction.SPI_Message = "Check /N4 Tag plate fixed correctly with RING (F9900GG)  YES  ||  NO"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\190\" & "N4_PlateYTA.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Not Applicable.jpg"
            If CustOrd.MS_CODE Like "YTA???-???????*/N4*" Then
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "N4_PlateYTA.jpg"
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
            ProcessStepReturn.ActivityToCheck = "/N4 Tag Plate has Tag number? If Yes, is it correct?"
            ProcessStepReturn.SinglePointAction.SPI_Message = "Check /N4 Tag number. Printed and Correct?  YES  ||  NO"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Yes.jpg"
            ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Not Applicable.jpg"
            If CustOrd.MS_CODE Like "YTA???-???????*/N4*" Then
                If CustOrd.TAG_NO_525 = "" Or CustOrd.TAG_NO_525.ToUpper = "BLANK" Then
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Not Applicable.jpg"
                    ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_1900_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                    ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_1900_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(1)
                Else
                    ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Yes.jpg"
                    ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_1900_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                    ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_1900_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)
                End If
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
            If Not (CustOrd.MS_CODE Like "YTA???-??????N*") Then
                Dim ProcessStepReturn As New CheckSheetStep
                ProcessStepReturn.ProcessNo = "190"
                ProcessStepReturn.ProcessStep = "Visually Inspect Unit"
                ProcessStepReturn.Activity = "Display||Clean||Lock Screw||Approval Plate||Tag Plate||N4 Plate||N4 Tagnumber||Bracket"
                ProcessStepReturn.ToCheck = "Correct || Not Correct"
                ProcessStepReturn.Method = CheckSheetStep.MethodOption.ProcedureStep
                ProcessStepReturn.Initial = Initial

                ProcessStepReturn.StepNo = "190_16_00"
                ProcessStepReturn.ActivityToCheck = "Mounting Bracket Check"
                ProcessStepReturn.ProcedureStepAction.ProcedureMessage = "Is all the items as below is Available?"
                ProcessStepReturn.ProcedureStepAction.SetSizeMode = ProcedureStep.SizeMode.CenterImage
                If CustOrd.MS_CODE Like "YTA???-??????B*" Then
                    ProcessStepReturn.ProcedureStepAction.ImagePath_ProcedureStep = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\210\" & "YTA_B_Bracket.jpg"
                ElseIf CustOrd.MS_CODE Like "YTA???-??????D*" Then
                    ProcessStepReturn.ProcedureStepAction.ImagePath_ProcedureStep = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\210\" & "YTA_D_Bracket.jpg"
                ElseIf CustOrd.MS_CODE Like "YTA???-??????J*" Then
                    ProcessStepReturn.ProcedureStepAction.ImagePath_ProcedureStep = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\210\" & "YTA_J_Bracket.jpg"
                ElseIf CustOrd.MS_CODE Like "YTA???-??????K*" Then
                    ProcessStepReturn.ProcedureStepAction.ImagePath_ProcedureStep = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\210\" & "YTA_K_Bracket.jpg"
                End If
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_1900_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_1900_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)
                ProcessStepReturn.Result &= "$" & MainForm.Setting.Var_60_1900_Position_Initial.Replace("Initial", Initial)
                Return ProcessStepReturn
            Else
                Dim ProcessStepReturn As New CheckSheetStep
                ProcessStepReturn.ProcessNo = "190"
                ProcessStepReturn.ProcessStep = "Visually Inspect Unit"
                ProcessStepReturn.Activity = "Display||Clean||Lock Screw||Approval Plate||Tag Plate||N4 Plate||N4 Tagnumber||Bracket"
                ProcessStepReturn.ToCheck = "Correct || Not Correct"
                ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
                ProcessStepReturn.Initial = Initial

                ProcessStepReturn.StepNo = "190_16_00"
                ProcessStepReturn.ActivityToCheck = "Mounting Bracket Correct?"
                ProcessStepReturn.SinglePointAction.SPI_Message = "Is the Mounting Bracket Parts correct?"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Yes.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Not Applicable.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Not Applicable.jpg"
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_1900_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_1900_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(1)
                ProcessStepReturn.Result &= "$" & MainForm.Setting.Var_60_1900_Position_Initial.Replace("Initial", Initial)
                Return ProcessStepReturn

            End If

        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    'PACKING
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
            If CustOrd.MS_CODE Like "YTA*/K[UFS]*" Then
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_2000_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_2000_Position_IM.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_2000_Position_EUDoC.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
            ElseIf Not (CustOrd.MS_CODE Like "YTA*/K[UFS]*") Then
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_2000_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_2000_Position_IM.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
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
                ProcessStepReturn.ActivityToCheck = "Mounting Brackets Selection"
                ProcessStepReturn.SinglePointAction.SPI_Message = "Is Mounting Bracket Required for this Job?  YES  ||  NO"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Not Applicable.jpg"
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_2000_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_2000_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(1)
            Else
                ProcessStepReturn.ActivityToCheck = "Mounting Brackets Selection"
                ProcessStepReturn.SinglePointAction.SPI_Message = "Is Mounting Bracket Required for this Job?  YES  ||  NO"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Yes.jpg"
                ProcessStepReturn.Result = Array.Find(MainForm.Setting.Var_60_2000_Position_Tick.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1)
                ProcessStepReturn.Result &= "$" & Array.Find(MainForm.Setting.Var_60_2000_Position_Circle.Split("|"), Function(x) x.StartsWith(ProcessStepReturn.StepNo)).Split("$")(1).Split(";")(0)
            End If
            Return ProcessStepReturn


        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ProcessStepNo200_05_01(ByVal Initial As String, ByVal CustOrd As POCO_YGSP.cust_ord, Optional ByRef ErrMsg As String = "") As CheckSheetStep
        Try

            Dim ProcessStepReturn As New CheckSheetStep
            ProcessStepReturn.ProcessNo = "200"
            ProcessStepReturn.ProcessStep = "Packing"
            ProcessStepReturn.Activity = "UnitKept||IM||EuDOC||QIC||Bracket||VisualCheck||Markings||OrderTag"
            ProcessStepReturn.ToCheck = "Correct||Not Correct||Details"
            ProcessStepReturn.Method = CheckSheetStep.MethodOption.SinglePntInst
            ProcessStepReturn.Initial = Initial

            ProcessStepReturn.StepNo = "200_05_01"

            If CustOrd.MS_CODE Like "YTA???-??????B*" Then
                ProcessStepReturn.ActivityToCheck = "Which type of Mounting Brackets is Prepared?"
                ProcessStepReturn.SinglePointAction.SPI_Message = "Mounting Bracket Type"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\210\" & "YTA_J_Bracket.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\210\" & "YTA_B_Bracket.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "YTA_B_Bracket.jpg"
            ElseIf CustOrd.MS_CODE Like "YTA???-??????D*" Then
                ProcessStepReturn.ActivityToCheck = "Which type of Mounting Brackets is Prepared?"
                ProcessStepReturn.SinglePointAction.SPI_Message = "Mounting Bracket Type"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\210\" & "YTA_D_Bracket.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\210\" & "YTA_K_Bracket.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "YTA_D_Bracket.jpg"
            ElseIf CustOrd.MS_CODE Like "YTA???-??????J*" Then
                ProcessStepReturn.ActivityToCheck = "Which type of Mounting Brackets is Prepared?"
                ProcessStepReturn.SinglePointAction.SPI_Message = "Mounting Bracket Type"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\210\" & "YTA_J_Bracket.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\210\" & "YTA_B_Bracket.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "YTA_J_Bracket.jpg"
            ElseIf CustOrd.MS_CODE Like "YTA???-??????K*" Then
                ProcessStepReturn.ActivityToCheck = "Which type of Mounting Brackets is Prepared?"
                ProcessStepReturn.SinglePointAction.SPI_Message = "Mounting Bracket Type"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\210\" & "YTA_K_Bracket.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\210\" & "YTA_D_Bracket.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "YTA_K_Bracket.jpg"
            ElseIf CustOrd.MS_CODE Like "YTA???-??????N*" Then
                ProcessStepReturn.ActivityToCheck = "Which type of Mounting Brackets is Prepared?"
                ProcessStepReturn.SinglePointAction.SPI_Message = "Mounting Bracket Type"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_1 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\210\" & "YTA_K_Bracket.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_2 = MainForm.Setting.Var_52_SinglePntInst_ImagePath & "YTA\Common\" & "Not Applicable.jpg"
                ProcessStepReturn.SinglePointAction.ImagePath_SPI_Correct = "Not Applicable.jpg"
            End If
            ProcessStepReturn.Result = ""
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
        ElseIf (NewModel Like "*/C3*") And (OldModel Like "*/C1*") Then
            Return False
        ElseIf (NewModel Like "*/C3*") And (OldModel Like "*/C3*") Then
            Return False
        ElseIf (NewModel Like "*/C1*") And (OldModel Like "*/C3*") Then
            Return False
        ElseIf (NewModel Like "*/C1*") And (OldModel Like "*/C1*") Then
            Return False
        ElseIf (NewModel Like "*/C[13]*") And (Not OldModel Like "*/C[123]*") Then
            Return False
        ElseIf (Not NewModel Like "*/C[123]*") And (OldModel Like "*/C[13]*") Then
            Return False
        ElseIf (NewModel Like "*/C[13]*") And (OldModel Like "*/C2*") Then
            Return True
        ElseIf (NewModel Like "*/C2*") And (OldModel Like "*/C[13]*") Then
            Return True
        ElseIf (NewModel Like "*/C2*") And (Not OldModel Like "*/C[123]*") Then
            Return True
        ElseIf ((NewModel Like "*/A*") And (Not OldModel Like "*/A*")) Or
               ((Not NewModel Like "*/A*") And (OldModel Like "*/A*")) Then
            Return True
        Else
            Return False
        End If
    End Function
#End Region

#Region "Report"
    ''' <summary>
    ''' Update Name and Text to WriteField object and add to List of WriteField
    ''' </summary>
    ''' <param name="Name">WriteField Name as String</param>
    ''' <param name="Text">WriteField Text as String</param>
    ''' <param name="WriteFields">WriteFields object for storing Name and Text</param>
    Private Sub AddWriteField(ByVal Name As String, ByVal Text As String, ByRef WriteFields As List(Of OpenPdfOperation_x64.WriteField))
        Try
            Dim WriteField As New OpenPdfOperation_x64.WriteField
            WriteField.Name = Name
            WriteField.Text = Text
            WriteFields.Add(WriteField)
        Catch ex As Exception

        End Try
    End Sub
    ''' <summary>
    ''' Add the WriteFields to ResultFields by filtering from MainForm.ReportTemplate.Fields
    ''' and update the Text from WriteFields to ResultFields.
    ''' </summary>
    ''' <param name="WriteFields">WriteFields object</param>
    ''' <param name="ErrMsg">Error string returned from Function, if any</param>
    ''' <returns></returns>
    Private Function AddResultTexts(ByVal WriteFields As List(Of OpenPdfOperation_x64.WriteField), Optional ByRef ErrMsg As String = "") As List(Of OpenPdfOperation_x64.Field)
        Dim ResultFields As New List(Of OpenPdfOperation_x64.Field)
        For Each WriteItem In WriteFields
            Dim FilteredField = GetFilteredFields(WriteItem.Name, ErrMsg:=ErrMsg)
            If FilteredField IsNot Nothing Then
                ResultFields.AddRange(FilteredField)
            Else
                Dim EmptyField As New OpenPdfOperation_x64.Field
                EmptyField.Name = WriteItem.Name
                EmptyField.Text = ""
                EmptyField.X = 0
                EmptyField.Y = 0
                EmptyField.FontSize = 10.0
                EmptyField.FontName = "Segoe UI Symbol"
                EmptyField.FontStyle = "Regular"
                EmptyField.FontColor = -11579393
                EmptyField.Align = "Center"
                ResultFields.Add(EmptyField)
            End If
        Next
        For Each ResultItem In ResultFields
            For Each WriteItem In WriteFields
                If ResultItem.Name = WriteItem.Name Then
                    If Not (ResultItem.X = 0 And ResultItem.Y = 0) Then
                        ResultItem.Text = WriteItem.Text
                    End If
                End If
            Next
        Next
        Return ResultFields
    End Function
    ''' <summary>
    ''' Filter the MainForm.ReportTemplate.Fields by Field.Name and return the List of OpenPdfOperation_x64.Field
    ''' </summary>
    ''' <param name="FilterFieldName">Filter Field Name as string</param>
    ''' <param name="ErrMsg">Error string returned from function, if any</param>
    ''' <returns></returns>
    Private Function GetFilteredFields(ByVal FilterFieldName As String, ByRef ErrMsg As String) As List(Of OpenPdfOperation_x64.Field)
        Dim SubReportTemplate As OpenPdfOperation_x64.Template = MainForm.ReportTemplate
        Dim ResultFields = SubReportTemplate.Fields.Where(Function(x) x.Name = FilterFieldName).ToList
        If ResultFields.Count = 0 Then
            Return Nothing
        Else
            Return ResultFields
        End If
    End Function
    ''' <summary>
    ''' Make a Template object .Fields from the ProcessReturnFields and other Template header from MainForm.ReportTemplate
    ''' </summary>
    ''' <param name="ProcessReturnFields">Fields as List(Of OpenPdfOperation_x64.Field)</param>
    ''' <returns>object of OpenPdfOperation_x64.Template to use in report printing</returns>
    Public Function MakeProcessReturnTemplate(ByVal ProcessReturnFields As List(Of OpenPdfOperation_x64.Field)) As OpenPdfOperation_x64.Template
        ' Create a template object
        Dim SubReportTemplate As OpenPdfOperation_x64.Template = MainForm.ReportTemplate
        Dim TemplateMaster As New OpenPdfOperation_x64.Template With {
            .FileName = SubReportTemplate.FileName,
            .Width = SubReportTemplate.Width,
            .Height = SubReportTemplate.Height,
            .Dpi = SubReportTemplate.Dpi,
            .ZoomFactor = SubReportTemplate.ZoomFactor,
            .Fields = New List(Of OpenPdfOperation_x64.Field)
        }
        'Create Clone
        Dim ChildTemplate = CloneTemplate(TemplateMaster)
        'Add Fields to ChildTemplate
        If ProcessReturnFields.Count > 0 Then
            For Each ResultItem In ProcessReturnFields
                ChildTemplate.Fields.Add(ResultItem)
            Next
        Else
            Dim EmptyField As New OpenPdfOperation_x64.Field
            EmptyField.Name = ""
            EmptyField.Text = ""
            EmptyField.X = 0
            EmptyField.Y = 0
            EmptyField.FontSize = 10.0
            EmptyField.FontName = "Segoe UI Symbol"
            EmptyField.FontStyle = "Regular"
            EmptyField.FontColor = -11579393
            EmptyField.Align = "Center"
            ChildTemplate.Fields.Add(EmptyField)
        End If
        Return ChildTemplate
    End Function
    ''' <summary>
    ''' Make a deep copy of Template object
    ''' </summary>
    ''' <param name="source">Source Template</param>
    ''' <returns>Deep copy of Source Template</returns>
    Public Function CloneTemplate(source As OpenPdfOperation_x64.Template) As OpenPdfOperation_x64.Template
        Dim clonedFields As New List(Of Field)
        For Each f In source.Fields
            clonedFields.Add(New Field With {
            .Name = f.Name,
            .Text = f.Text,
            .X = f.X,
            .Y = f.Y,
            .Align = f.Align,
            .FontColor = f.FontColor,
            .FontName = f.FontName,
            .FontSize = f.FontSize,
            .FontStyle = f.FontSize
        })
        Next

        Dim clonedTemplate As New Template With {
        .FileName = source.FileName,
        .Width = source.Width,
        .Height = source.Height,
        .Dpi = source.Dpi,
        .ZoomFactor = source.ZoomFactor,
        .Fields = clonedFields
    }

        Return clonedTemplate
    End Function
#End Region
End Class
