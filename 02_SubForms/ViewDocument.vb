Public Class ViewDocument
    Public curNaviageUrl As String
    Public webLoaded As Boolean
    Dim TmlEty As New MFG_ENTITY.Op(MainForm.Setting.Var_03_MySql_YGSP)
    Dim WMsg As New WarningForm
    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Me.Hide()
    End Sub
    Private Sub WebBrowser1_DocumentCompleted(sender As Object, e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles ViewBrowser.DocumentCompleted 'V8.3 ADDED
        webLoaded = True 'V8.3 ADDED
        curNaviageUrl = ViewBrowser.Url.AbsoluteUri
    End Sub
    Private Sub Button1_Click(sender As Object, e As System.EventArgs) Handles Button1.Click
        ViewBrowser.Navigate(curNaviageUrl)
        ViewBrowser.Refresh()
    End Sub
    Private Sub ViewDocument_Load(sender As Object, e As EventArgs) Handles Me.Load
        ViewBrowser.Navigate(curNaviageUrl)
        ViewBrowser.Refresh()
    End Sub
    Private Sub Button_GO_Click(sender As Object, e As EventArgs) Handles Button_GO.Click
        Try
            Dim ErrMsg As String = ""
            Dim Sql(0) As String
            Sql(0) = "UPDATE cust_ord SET QIC='GO' WHERE INDEX_NO='" & MainForm.CustOrd.INDEX_NO & "';"
            TmlEty.ExecuteTransactionQuery(Sql, ErrMsg:=ErrMsg)
            If ErrMsg.Length > 0 Then
                WMsg.Message = ErrMsg
                WMsg.ShowDialog()
                Me.Close()
            Else
                MainForm.CustOrd = TmlEty.GetDatabaseTableAs_Object(Of POCO_YGSP.cust_ord)("INDEX_NO", MainForm.CustOrd.INDEX_NO, ErrMsg:=ErrMsg)
                If ErrMsg.Length > 0 Then
                    WMsg.Message = ErrMsg
                    WMsg.ShowDialog()
                    Me.Close()
                End If
                MainForm.InspectionStatus(MainForm.CurrentCheckPoint, True)
                MainForm.Button2.PerformClick()
                Me.Close()
            End If

        Catch ex As Exception

        End Try
    End Sub
    Private Sub Button_NG_Click(sender As Object, e As EventArgs) Handles Button_NG.Click
        Try
            Dim ErrMsg As String = ""
            Dim Sql(0) As String
            Sql(0) = "UPDATE cust_ord SET QIC='NG' WHERE INDEX_NO='" & MainForm.CustOrd.INDEX_NO & "';"
            TmlEty.ExecuteTransactionQuery(Sql, ErrMsg:=ErrMsg)
            If ErrMsg.Length > 0 Then
                WMsg.Message = ErrMsg
                WMsg.ShowDialog()
                Me.Close()
            Else
                MainForm.CustOrd = TmlEty.GetDatabaseTableAs_Object(Of POCO_YGSP.cust_ord)("INDEX_NO", MainForm.CustOrd.INDEX_NO, ErrMsg:=ErrMsg)
                If ErrMsg.Length > 0 Then
                    WMsg.Message = ErrMsg
                    WMsg.ShowDialog()
                    Me.Close()
                End If
                Me.Close()
                MainForm.InspectionStatus(MainForm.CurrentCheckPoint, False)
                MainForm.ShowCompleteInspection()
            End If

        Catch ex As Exception

        End Try
    End Sub
End Class