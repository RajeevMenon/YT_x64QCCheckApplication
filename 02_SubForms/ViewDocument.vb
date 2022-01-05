Public Class ViewDocument
    Public curNaviageUrl As String
    Public webLoaded As Boolean
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
            MainForm.InspectionStatus(MainForm.CurrentCheckPoint, True)
            MainForm.Button2.PerformClick()
            Me.Close()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button_NG_Click(sender As Object, e As EventArgs) Handles Button_NG.Click
        Try
            Me.Close()
            MainForm.InspectionStatus(MainForm.CurrentCheckPoint, False)
            MainForm.ShowCompleteInspection()
        Catch ex As Exception

        End Try
    End Sub
End Class