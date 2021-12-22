Public Class BarcodeEntry
    Private Sub TextBox_Scan_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox_Scan.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                Dim ErrMsg As String = ""
                Dim TmlEntity As New MFG_ENTITY.Op(MainForm.Setting.Var_03_MySql_YGSP)
                Dim FieldName As String = ""
                Dim IndexNo = TextBox_Scan.Text
                If IndexNo.Length = 10 Then FieldName = "BAR" Else FieldName = "INDEX_NO"
                MainForm.CustOrd = TmlEntity.GetDatabaseTableAs_Object(Of POCO_YGSP.cust_ord)(FieldName, IndexNo, FieldName, IndexNo, ErrMsg)
                If ErrMsg.Length > 0 Then
                    Label_Message.Text = ErrMsg
                Else
                    Me.Close()
                    MainForm.TextBox_Step.Text = 1
                    MainForm.Button2.PerformClick()
                End If
            End If
        Catch ex As Exception
            Label_Message.Text = "Runtime Error:" & ex.Message.Substring(0, 50) & ".."
        End Try
    End Sub
End Class