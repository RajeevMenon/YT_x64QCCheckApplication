Imports Microsoft.SqlServer

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


#Region "IM and Safety-IM"
                'IM & SafetyIM always has "-", this needed to be changed to "|"
                Dim IM As String = "No-IM"
                Dim IM_ReplacementValue As String = ""

                If TextBox_IM.Text.ToUpper.Length > 0 Then
                    Dim IM_Name As String = TextBox_IM.Text.ToUpper.Trim.Split("/")(0) 'Same
                    Dim DocList = DocTbl.Where(Function(x) x.CAT_DOC = IM_Name).FirstOrDefault
                    If MainForm.CurrentQCC_Version = "1.3" Then IM_Name = IM_Name.Replace("-", "|") 'Same with - replaced with | for PrintQcc_Rev1()
                    Dim IM_Ver As String = DocList.DOC_VERSION '"(" & Decimal.Parse(TextBox_IM.Text.ToUpper.Trim.Split("/")(1) / 10).ToString("#") & ")" 'Devide by 10
                    Dim Final_IM_String As String = IM_Name & "(" & IM_Ver & ")" '& IM_Add
                    IM = Final_IM_String.ToUpper.Trim
                    Dim IM_ShortName As String = "IM"
                    If IM_ReplacementValue.Length > 0 Then
                        IM_ReplacementValue &= "," & IM_ShortName & "(" & IM_Ver & ")"
                    Else
                        IM_ReplacementValue &= IM_ShortName & "(" & IM_Ver & ")"
                    End If
                End If

                If TextBox_SIM.Text.ToUpper.Length > 0 Then
                    Dim SIM_Name As String = TextBox_SIM.Text.ToUpper.Trim.Split("/")(0) 'Same
                    Dim DocList = DocTbl.Where(Function(x) x.CAT_DOC = SIM_Name).FirstOrDefault
                    SIM_Name = SIM_Name.Replace("-", "|") 'Same with - replaced with | for PrintQcc_Rev1()
                    Dim SIM_Ver As String = DocList.DOC_VERSION
                    Dim Final_SIM_String As String = SIM_Name & "(" & SIM_Ver & ")"
                    IM &= "+" & Final_SIM_String.ToUpper.Trim
                    Dim SM_ShortName As String = "SF" 'DocList.CAT_DOC.Split("-")(1)
                    If IM_ReplacementValue.Length > 0 Then
                        IM_ReplacementValue &= "," & SM_ShortName & "(" & SIM_Ver & ")"
                    Else
                        IM_ReplacementValue &= SM_ShortName & "(" & SIM_Ver & ")"
                    End If
                End If

                If MainForm.CurrentQCC_Version = "1.3" Then
                    MainForm.CurrentCheckPoint.Result = MainForm.CurrentCheckPoint.Result.Replace("IM-", IM & "-")
                Else
                    MainForm.CurrentCheckPoint.Result = MainForm.CurrentCheckPoint.Result.Replace("IM", IM_ReplacementValue)
                End If

#End Region
#Region "EU Doc"
                If TextBox_EUDoC.Text.ToUpper.Length > 0 Then
                    'EU document QR code alway begin with "EUDOC-", this is not needed in the final result
                    Dim EU_CatDoc = TextBox_EUDoC.Text.ToUpper.Split("-")(1).Split("/")(0)
                    Dim DocList = DocTbl.Where(Function(x) x.CAT_DOC Like "*" & EU_CatDoc).FirstOrDefault
                    Dim EU_Doc = DocList.CAT_DOC.Replace("EUDOC-", "") 'TextBox_EUDoC.Text.Replace("EUDOC-", "")
                    Dim EU_DocEd = DocList.EDITION 'TextBox_EUDoC.Text.ToUpper.Split("/")(1)
                    Dim EU_DocVer = DocList.DOC_VERSION 'TextBox_EUDoC.Text.ToUpper.Split("/")(1)
                    Dim QCC_WriteText As String = EU_Doc & " (" & EU_DocVer & ")"
                    If MainForm.CurrentQCC_Version = "1.3" Then
                        MainForm.CurrentCheckPoint.Result = MainForm.CurrentCheckPoint.Result.Replace("EUDOC-", QCC_WriteText & "-")
                    Else
                        MainForm.CurrentCheckPoint.Result = MainForm.CurrentCheckPoint.Result.Replace("EUDOC", QCC_WriteText)
                    End If
                Else
                    MainForm.CurrentCheckPoint.Result = MainForm.CurrentCheckPoint.Result.Replace("EUDOC", "Not Applicable")
                End If
#End Region

                Dim TmlEntityYGS As New MFG_ENTITY.Op(Link.Mysql_YGSP_ConStr)
                Dim UpdateString As String = $"{TextBox_IM.Text},{TextBox_SIM.Text},{TextBox_EUDoC.Text}"
                Dim UpdateSql(0) As String
                UpdateSql(0) = $"UPDATE cust_ord SET DOC_ATTACH='{UpdateString}' WHERE INDEX_NO='" & CustOrd.INDEX_NO & "';"
                Dim ErrMsg As String = ""
                TmlEntityYGS.ExecuteTransactionQuery(UpdateSql, ErrMsg:=ErrMsg)
                If ErrMsg.Length > 0 Then
                    AddedDocErrorMessage = "Could not update details of due to " & ErrMsg
                    Label_Error.Text = AddedDocErrorMessage
                    MainForm.InspectionStatus(MainForm.CurrentCheckPoint, False)
                    MainForm.wait(1)
                    Exit Sub
                Else
                    Dim CustOrdUpdated = TmlEntityYGS.GetDatabaseTableAs_List(Of POCO_YGSP.cust_ord)("INDEX_NO", CustOrd.INDEX_NO, ErrMsg)
                    If ErrMsg.Length > 0 Then
                        AddedDocErrorMessage = "Could not read cust_ord details of due to " & ErrMsg
                        Label_Error.Text = AddedDocErrorMessage
                    End If
                    If CustOrdUpdated.Count = 1 Then
                        MainForm.CustOrd = CustOrdUpdated.Item(0)
                    End If
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
        'Following Rule Is applicable for generating IM QR
        '-------------------------------------------------
        'Format: CAT_DOC/DOC_VERSION  = EUDOC-F9220NL/17.0 or F9220NL/17.0
        Try
            If e.KeyCode = Keys.Enter Then
                EUDOC_OK = False
                If CurrentIndexNo.Length = 0 Then
                    EUDOC_OK = False
                    AddedDocErrorMessage = "IndexNo. cannot be read."
                Else
                    AddedDocErrorMessage = ""
                    If CustOrd.MS_CODE Like "YTA[67]??-???????*/K[UFS]*" Then
                        Dim EuDoc = DocTbl.Where(Function(x) x.DOC_NAME Like "*" & CustOrd.MODEL & "*EU-DOC").FirstOrDefault
                        If TextBox_EUDoC.Text.ToUpper = EuDoc.CAT_DOC & "/" & EuDoc.DOC_VERSION Then
                            EUDOC_OK = True
                        ElseIf TextBox_EUDoC.Text.ToUpper = EuDoc.CAT_DOC.Replace("EUDOC-", "") & "/" & EuDoc.DOC_VERSION Then
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
        'Following Rule Is applicable for generating IM QR
        '-------------------------------------------------
        'Format: CAT_DOC/DOC_VERSION  = IM01C50G01-01EN/7.2
        Try

            If e.KeyCode = Keys.Enter Then
                INST_IM_OK = False
                If CurrentIndexNo.Length = 0 Then
                    INST_IM_OK = False
                    AddedDocErrorMessage = "IndexNo. cannot be read."
                Else
                    AddedDocErrorMessage = ""
                    If CustOrd.MS_CODE Like "YTA[67]??-???????*" Then
                        'NonRoHS only for 5Q00
                        If Link.PlantID = "5Q00" Then
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
                                        Dim IM_Doc = DocTbl.Where(Function(x) x.CAT_DOC = "IM01C50G01-01ENZ").FirstOrDefault
                                        Dim QR As String = ""
                                        If IM_Doc.DOC_NAME.Split("+").Length > 1 Then
                                            QR = IM_Doc.CAT_DOC & "/" & Decimal.Parse(IM_Doc.DOC_VERSION * 10).ToString("#0")
                                            For i As Integer = 1 To IM_Doc.DOC_NAME.Split("+").Length - 1
                                                QR += "/" & IM_Doc.DOC_NAME.Split("+")(i).Trim.ToUpper
                                            Next
                                        Else
                                            QR = IM_Doc.CAT_DOC & "/" & Decimal.Parse(IM_Doc.DOC_VERSION * 10).ToString("#0")
                                        End If
                                        If TextBox_IM.Text.ToUpper Like QR.ToUpper Then
                                            INST_IM_OK = True
                                        End If
                                    End If
                                Else
                                    'Standard
                                    Dim IM_Doc = DocTbl.Where(Function(x) x.CAT_DOC = "IM01C50G01-01EN").FirstOrDefault
                                    Dim QR As String = IM_Doc.CAT_DOC & "/" & IM_Doc.DOC_VERSION
                                    Dim Latest_YHQ_QR As String = "IM 01C50G01-01E/" & CInt(CDbl(IM_Doc.DOC_VERSION)).ToString() 'Manage YHQ QR with space after IM 01C25A01-01EN
                                    Dim Latest_TML_QR As String = "IM01C50G01-01EN/" & IM_Doc.DOC_VERSION 'Manage TML QR without space after IM01C25A01-01EN
                                    If (TextBox_IM.Text.ToUpper = Latest_TML_QR.ToUpper) Or (TextBox_IM.Text.ToUpper = Latest_YHQ_QR.ToUpper) Then
                                        INST_IM_OK = True
                                    End If
                                End If
                            End If
                        Else
                            'Standard
                            Dim IM_Doc = DocTbl.Where(Function(x) x.CAT_DOC = "IM01C50G01-01EN").FirstOrDefault
                            Dim QR As String = IM_Doc.CAT_DOC & "/" & IM_Doc.DOC_VERSION
                            Dim Latest_YHQ_QR As String = "IM 01C50G01-01E/" & CInt(CDbl(IM_Doc.DOC_VERSION)).ToString() 'Manage YHQ QR with space after IM 01C25A01-01EN
                            Dim Latest_TML_QR As String = "IM01C50G01-01EN/" & IM_Doc.DOC_VERSION 'Manage TML QR without space after IM01C25A01-01EN
                            If (TextBox_IM.Text.ToUpper = Latest_TML_QR.ToUpper) Or (TextBox_IM.Text.ToUpper = Latest_YHQ_QR.ToUpper) Then
                                INST_IM_OK = True
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
        'Following Rule Is applicable for generating IM QR
        '-------------------------------------------------
        'Format: CAT_DOC/DOC_VERSION  = IM00C01C01-01Z1/1.0
        Try
            If e.KeyCode = Keys.Enter Then
                SFTY_IM_OK = False
                If CurrentIndexNo.Length = 0 Then
                    SFTY_IM_OK = False
                    AddedDocErrorMessage = "IndexNo. cannot be read."
                Else
                    AddedDocErrorMessage = ""
                    If CustOrd.MS_CODE Like "YTA[67]??-???????*" Then
                        Dim Document = DocTbl.Where(Function(x) x.CAT_DOC = "IM00C01C01-01Z1").ToList.FirstOrDefault
                        Dim QR As String = Document.CAT_DOC & "/" & Document.DOC_VERSION
                        Dim Latest_YHQ_QR As String = "IM 00C01C01-01Z1/" & CInt(CDbl(Document.DOC_VERSION)).ToString() 'Manage YHQ QR with space after IM 01C25A01-01EN
                        Dim Latest_TML_QR As String = "IM00C01C01-01Z1/" & Document.DOC_VERSION 'Manage TML QR without space after IM01C25A01-01EN
                        If (TextBox_SIM.Text.ToUpper = Latest_TML_QR.ToUpper) Or (TextBox_SIM.Text.ToUpper = Latest_YHQ_QR.ToUpper) Then
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

                CustOrd = MainForm.CustOrd

                If CustOrd.DOC_ATTACH.Split(",").Length = 3 Then
                    TextBox_IM.Text = CustOrd.DOC_ATTACH.Split(",")(0).Trim
                    TextBox_SIM.Text = CustOrd.DOC_ATTACH.Split(",")(1).Trim
                    TextBox_EUDoC.Text = CustOrd.DOC_ATTACH.Split(",")(2).Trim
                End If


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
                    If (CustOrd.MS_CODE Like "YTA[67]??-???????*") And CustOrd.MS_CODE Like "*/K[UFS]*" Then
                        TextBox_EUDoC.Enabled = True
                        EUDOC_OK = False
                    Else
                        EUDOC_OK = True
                        TextBox_EUDoC.Enabled = False
                    End If

                    'FZ2-YTA-01155
                    INST_IM_OK = False
                    TextBox_IM.Enabled = True
                    SFTY_IM_OK = False
                    TextBox_SIM.Enabled = True

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