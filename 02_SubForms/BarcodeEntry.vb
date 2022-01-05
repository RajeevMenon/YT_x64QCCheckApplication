Public Class BarcodeEntry
    Private Sub TextBox_Scan_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox_Scan.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                Dim ErrMsg As String = ""
                Dim TmlEntityYGS As New MFG_ENTITY.Op(MainForm.Setting.Var_03_MySql_YGSP)
                Dim TmlEntityQA As New MFG_ENTITY.Op(MainForm.Setting.Var_04_MySql_QA)
                Dim FieldName As String = ""
                Dim IndexNo = ""

                If TextBox_Scan.Text Like "*MFR:*S/N:*" Then
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