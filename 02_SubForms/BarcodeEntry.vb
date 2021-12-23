Public Class BarcodeEntry
    Private Sub TextBox_Scan_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox_Scan.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                Dim ErrMsg As String = ""
                Dim TmlEntityYGS As New MFG_ENTITY.Op(MainForm.Setting.Var_03_MySql_YGSP)
                Dim TmlEntityQA As New MFG_ENTITY.Op(MainForm.Setting.Var_04_MySql_QA)
                Dim FieldName As String = ""
                Dim IndexNo = TextBox_Scan.Text
                If IndexNo.Length = 10 Then FieldName = "BAR" Else FieldName = "INDEX_NO"
                MainForm.CustOrd = TmlEntityYGS.GetDatabaseTableAs_Object(Of POCO_YGSP.cust_ord)(FieldName, IndexNo, FieldName, IndexNo, ErrMsg)
                If ErrMsg.Length > 0 Then
                    Label_Message.Text = ErrMsg
                Else
                    MainForm.QcData = TmlEntityQA.GetDatabaseTableAs_List(Of POCO_QA.yta_qcc_v1p2)("INDEX_NO", MainForm.CustOrd.INDEX_NO, "INDEX_NO", MainForm.CustOrd.INDEX_NO, ErrMsg)
                    If ErrMsg.Length > 0 Then
                        Label_Message.Text = ErrMsg
                    Else
                        Me.Close()
                        MainForm.Text = MainForm.Text.Split("-")(0) & "-" & " User: " & MainForm.Initial & ", Current Unit: " & MainForm.CustOrd.SERIAL_NO & " [ " & MainForm.CustOrd.MS_CODE & " ]"
                        MainForm.TextBox_Step.Text = MainForm.Setting.Var_08_StepsAllowed.Split(",")(0)
                        MainForm.FirstCheckPoint()
                    End If
                End If
            End If
        Catch ex As Exception
            Label_Message.Text = "Runtime Error:" & ex.Message.Substring(0, 50) & ".."
        End Try
    End Sub
End Class