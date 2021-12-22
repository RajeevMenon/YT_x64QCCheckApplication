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
End Class