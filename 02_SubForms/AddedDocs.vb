Public Class AddedDocs

    Public AddedDocErrorMessage As String = ""
    Public INST_IM_OK As Boolean = False
    Public SFTY_IM_OK As Boolean = False
    Public EUDOC_OK As Boolean = False
    Private CustOrd As POCO_YGSP.cust_ord
    Private DocTbl As List(Of POCO_QA.t1_eudoc_table)
    Private CurrentIndexNo As String = ""
    Private Sub Button_Done_Click(sender As Object, e As EventArgs) Handles Button_Done.Click
        Try
            If CurrentIndexNo.Length = 0 Then
                INST_IM_OK = False
                SFTY_IM_OK = False
                EUDOC_OK = False
                AddedDocErrorMessage = "IndexNo. cannot be read."
            Else
                AddedDocErrorMessage = ""
            End If
            If INST_IM_OK = True And SFTY_IM_OK = True And EUDOC_OK = True Then
                'IM & SafetyIM always has "-", this needed to be changed to "."
                Dim IM = TextBox_IM.Text.ToUpper.Trim & " + " & TextBox_SIM.Text.ToUpper.Trim
                IM = IM.Replace(" ", "").Replace("-", ".")
                For Each Item In MainForm.AllCheckResult
                    If Not IsNothing(Item) Then
                        If Item.StepNo = MainForm.CurrentCheckPoint.StepNo Then
                            Item.Result = Item.Result.Replace("IM-", IM & "-")
                            If TextBox_EUDoC.Text.ToUpper.Length > 0 Then Item.Result = Item.Result.Replace("EUDOC-", TextBox_EUDoC.Text.ToUpper & "-")
                        End If
                    End If
                Next
                MainForm.CurrentCheckPoint.Result = MainForm.CurrentCheckPoint.Result.Replace("IM-", IM & "-")
                If TextBox_EUDoC.Text.ToUpper.Length > 0 Then
                    'EU document QR code alway begin with "EUDOC-", this is not needed in the final result
                    MainForm.CurrentCheckPoint.Result = MainForm.CurrentCheckPoint.Result.Replace("EUDOC-", TextBox_EUDoC.Text.ToUpper.Split("-")(1) & "-")
                End If
                MainForm.InspectionStatus(MainForm.CurrentCheckPoint, True)
                MainForm.wait(1)
                MainForm.Button2.PerformClick()
                Me.Hide()
            Else
                If INST_IM_OK = False Then AddedDocErrorMessage &= " [IM Not OK] " Else AddedDocErrorMessage &= " [IM OK] "
                If SFTY_IM_OK = False Then AddedDocErrorMessage &= " [Sfty-IM Not OK] " Else AddedDocErrorMessage &= " [Sfty-IM OK] "
                If EUDOC_OK = False And TextBox_EUDoC.Enabled = True Then
                    AddedDocErrorMessage &= " [EU-DoC Not OK] "
                ElseIf EUDOC_OK = True And TextBox_EUDoC.Enabled = True Then
                    AddedDocErrorMessage &= " [EU-DoC OK] "
                End If
                Label_Error.Text = AddedDocErrorMessage
                MainForm.InspectionStatus(MainForm.CurrentCheckPoint, False)
                MainForm.wait(1)
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub Button_Cancel_Click(sender As Object, e As EventArgs) Handles Button_Cancel.Click
        INST_IM_OK = False
        SFTY_IM_OK = False
        EUDOC_OK = False
        MainForm.InspectionStatus(MainForm.CurrentCheckPoint, False)
        MainForm.wait(1)
        Me.Hide()
    End Sub
    Private Sub TextBox_EUDoC_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox_EUDoC.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                EUDOC_OK = False
                If CurrentIndexNo.Length = 0 Then
                    EUDOC_OK = False
                    AddedDocErrorMessage = "IndexNo. cannot be read."
                Else
                    AddedDocErrorMessage = ""
                    If CustOrd.MS_CODE Like "EJ[AX]???[EA]-?????-?????*/K[UFS]*" Then
                        Dim EuDoc = DocTbl.Where(Function(x) x.DOC_NAME Like CustOrd.MODEL & "*").FirstOrDefault
                        If TextBox_EUDoC.Text.ToUpper Like EuDoc.CAT_DOC & "/" & Math.Round(Decimal.Parse(EuDoc.DOC_VERSION), 0) Then
                            EUDOC_OK = True
                        End If
                    ElseIf CustOrd.MS_CODE Like "YTA[67]??-???????*/K[UFS]*" Then
                        Dim EuDoc = DocTbl.Where(Function(x) x.DOC_NAME Like CustOrd.MODEL & "*").FirstOrDefault
                        If TextBox_EUDoC.Text.ToUpper Like EuDoc.CAT_DOC & "/" & Math.Round(Decimal.Parse(EuDoc.DOC_VERSION), 0) Then
                            EUDOC_OK = True
                        End If
                    Else
                        EUDOC_OK = True
                    End If
                End If
            End If

            If EUDOC_OK = True Then
                If TextBox_SIM.Enabled = True And SFTY_IM_OK = False Then
                    TextBox_SIM.Focus()
                ElseIf TextBox_IM.Enabled = True And INST_IM_OK = False Then
                    TextBox_IM.Focus()
                Else
                    Button_Done.PerformClick()
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub
    Private Sub TextBox_IM_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox_IM.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                INST_IM_OK = False
                If CurrentIndexNo.Length = 0 Then
                    INST_IM_OK = False
                    AddedDocErrorMessage = "IndexNo. cannot be read."
                Else
                    AddedDocErrorMessage = ""
                    If CustOrd.MS_CODE Like "EJ[AX]???[EA]-?????-?????*" Then
                        Dim Version = DocTbl.Where(Function(x) x.CAT_DOC = "IM01C25A01-01EN").Select(Function(y) y.DOC_VERSION).FirstOrDefault
                        If TextBox_IM.Text.ToUpper Like "IM 01C25A01-01E/" & Math.Round(Decimal.Parse(Version), 0) & "*" Then
                            INST_IM_OK = True
                        End If
                    ElseIf CustOrd.MS_CODE Like "YTA[67]??-???????*" Then
                        Dim Version = DocTbl.Where(Function(x) x.CAT_DOC = "IM01C50G01-01EN").Select(Function(y) y.DOC_VERSION).FirstOrDefault
                        If TextBox_IM.Text.ToUpper Like "IM 01C50G01-01EN/" & Math.Round(Decimal.Parse(Version), 0) & "*" Then
                            Dim TmlEntity As New MFG_ENTITY.Op(MainForm.Setting.Var_03_MySql_YGSP)
                            Dim ErrMsg As String = ""
                            Dim Non_RoHS = TmlEntity.GetDatabaseTableAs_List(Of POCO_YGSP.non_rohs)("MS_CODE", "%", "MS_CODE", "%", ErrMsg)
                            If ErrMsg.Length > 0 Then
                                INST_IM_OK = False
                                MsgBox("Could not read NonRoHS serial list. ERROR:" & ErrMsg)
                                Exit Sub
                            Else
                                Dim SL_Check = Non_RoHS.Where(Function(x) x.SERIAL_NO = CustOrd.SERIAL_NO_BEFORE).ToList
                                If SL_Check.Count > 0 Then
                                    If CustOrd.SERIAL_NO_BEFORE = SL_Check.Select(Function(y) y.SERIAL_NO).FirstOrDefault.ToString Then
                                        If TextBox_IM.Text.ToUpper Like "IM 01C50G01-01EN/" & Math.Round(Decimal.Parse(Version), 0) & "/21-YKSA" Then
                                            INST_IM_OK = True
                                        End If
                                    End If
                                Else
                                    If TextBox_IM.Text.ToUpper Like "IM 01C50G01-01EN/" & Math.Round(Decimal.Parse(Version), 0) & "/21-0006-E" Then
                                        INST_IM_OK = True
                                    End If
                                End If

                            End If
                        End If
                    End If
                End If
            End If

            If INST_IM_OK = True Then
                If TextBox_SIM.Enabled = True And SFTY_IM_OK = False Then
                    TextBox_SIM.Focus()
                ElseIf TextBox_EUDoC.Enabled = True And EUDOC_OK = False Then
                    TextBox_EUDoC.Focus()
                Else
                    Button_Done.PerformClick()
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub TextBox_SIM_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox_SIM.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                SFTY_IM_OK = False
                If CurrentIndexNo.Length = 0 Then
                    SFTY_IM_OK = False
                    AddedDocErrorMessage = "IndexNo. cannot be read."
                Else
                    AddedDocErrorMessage = ""
                    If CustOrd.MS_CODE Like "EJ[AX]???[EA]-?????-?????*" Then
                        Dim Version = DocTbl.Where(Function(x) x.CAT_DOC = "IM00C01C01-01Z1").Select(Function(y) y.DOC_VERSION).FirstOrDefault
                        If TextBox_SIM.Text.ToUpper Like "IM 00C01C01-01Z1/" & Math.Round(Decimal.Parse(Version), 0) & "*" Then
                            SFTY_IM_OK = True
                        End If
                    ElseIf CustOrd.MS_CODE Like "YTA[67]??-???????*" Then
                        Dim Version = DocTbl.Where(Function(x) x.CAT_DOC = "IM00C01C01-01Z1").Select(Function(y) y.DOC_VERSION).FirstOrDefault
                        If TextBox_SIM.Text.ToUpper Like "IM 00C01C01-01Z1/" & Math.Round(Decimal.Parse(Version), 0) & "*" Then
                            SFTY_IM_OK = True
                        End If
                    End If
                End If
            End If

            If SFTY_IM_OK = True Then
                If TextBox_IM.Enabled = True And INST_IM_OK = False Then
                    TextBox_IM.Focus()
                ElseIf TextBox_EUDoC.Enabled = True And EUDOC_OK = False Then
                    TextBox_EUDoC.Focus()
                Else
                    Button_Done.PerformClick()
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub
    Private Sub AddedDocs_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            CurrentIndexNo = MainForm.CustOrd.INDEX_NO
            If CurrentIndexNo.Length = 0 Then
                SFTY_IM_OK = False
                AddedDocErrorMessage = "IndexNo. cannot be read."
            Else
                Label_Error.Text = ""
                TextBox_EUDoC.Text = ""
                TextBox_IM.Text = ""
                TextBox_SIM.Text = ""
                AddedDocErrorMessage = ""
                'Dim TmlEntity As New MFG_ENTITY.Op(MainForm.Setting.Var_03_MySql_YGSP)
                'CustOrd = TmlEntity.GetDatabaseTableAs_Object(Of POCO_YGSP.cust_ord)("INDEX_NO", CurrentIndexNo, "INDEX_NO", CurrentIndexNo, AddedDocErrorMessage)
                CustOrd = MainForm.CustOrd
                If AddedDocErrorMessage.Length > 0 Then
                    TextBox_SIM.Enabled = False
                    TextBox_EUDoC.Enabled = False
                    TextBox_IM.Enabled = False
                    Label_Error.Text = AddedDocErrorMessage
                Else
                    TextBox_SIM.Enabled = True
                    TextBox_EUDoC.Enabled = True
                    TextBox_IM.Enabled = True
                    Label_Error.Text = ""
                End If
                If CustOrd.INDEX_NO.ToString.Length = 0 Then
                    TextBox_SIM.Enabled = False
                    TextBox_EUDoC.Enabled = False
                    TextBox_IM.Enabled = False
                    Label_Error.Text = "Could not read the YGS details of the IndexNo.:" & CurrentIndexNo
                End If
                Dim QAEntity As New MFG_ENTITY.Op(MainForm.Setting.Var_04_MySql_QA)
                DocTbl = QAEntity.GetDatabaseTableAs_List(Of POCO_QA.t1_eudoc_table)("DOC_STATUS", "RELS", "DOC_STATUS", "RELS", AddedDocErrorMessage)
                If AddedDocErrorMessage.Length > 0 Then
                    TextBox_SIM.Enabled = False
                    TextBox_EUDoC.Enabled = False
                    TextBox_IM.Enabled = False
                    Label_Error.Text = AddedDocErrorMessage
                Else
                    TextBox_SIM.Enabled = True
                    TextBox_EUDoC.Enabled = True
                    TextBox_IM.Enabled = True
                    Label_Error.Text = ""
                End If
                If DocTbl.Item(0).DOC_STATUS.ToString.Length = 0 Then
                    TextBox_SIM.Enabled = False
                    TextBox_EUDoC.Enabled = False
                    TextBox_IM.Enabled = False
                    Label_Error.Text = "Could not read the QA details of the IndexNo.:" & CurrentIndexNo
                End If

                If Label_Error.Text = "" Then
                    If (CustOrd.MS_CODE Like "YTA[67]??-???????*" Or CustOrd.MS_CODE Like "EJ[AX]???[EA]-?????-?????*") And CustOrd.MS_CODE Like "*/K[UFS]*" Then
                        TextBox_EUDoC.Enabled = True
                        EUDOC_OK = False
                    Else
                        EUDOC_OK = True
                        TextBox_EUDoC.Enabled = False
                    End If

                    If CustOrd.QTY_NO = "01" Then
                        INST_IM_OK = False
                        TextBox_IM.Enabled = True
                        SFTY_IM_OK = False
                        TextBox_SIM.Enabled = True
                    Else
                        INST_IM_OK = True
                        TextBox_IM.Enabled = False
                        SFTY_IM_OK = True
                        TextBox_SIM.Enabled = False
                    End If

                    If TextBox_IM.Enabled = True Then
                        TextBox_IM.Focus()
                    ElseIf TextBox_EUDoC.Enabled = True Then
                        TextBox_EUDoC.Focus()
                    ElseIf TextBox_IM.Enabled = False And TextBox_SIM.Enabled = False And TextBox_EUDoC.Enabled = False Then
                        EUDOC_OK = True
                        INST_IM_OK = True
                        SFTY_IM_OK = True
                        Button_Done.PerformClick()
                    End If
                End If



            End If

        Catch ex As Exception

        End Try
    End Sub
End Class

'Noted
'During load all the boolean results set to false and textboxed set to empty
'if the unit in non-atex, then du-doc boolean set to true
'user needs to scan the QR code in IM, SafetyIM and EU-Doc in the respective boxes
'The application returns the version from QA table and compire with  user input and set the respective boolean true
'for non-rohs YTA, the application return full list of serial number from ygsp table and then compire with 
'the before modificaiton serial of the specific indexno. if it matches, the IM needs 21-YKSA manual change. else 21-0006-E
'Button done will work only if all the boolean fields are true
'Button cancel will close the form, however the booleans will be set to false
'The IM and SafetyIM fields will be enabled only for Qty_No=01