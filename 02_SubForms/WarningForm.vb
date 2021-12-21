Public Class WarningForm
    Public Message As String
    Public Reply As Boolean = False
    Private Sub WarningForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            RichTextBox1.Text = Message
            PictureBox1.Load(System.Windows.Forms.Application.StartupPath & "\YTA_Images\Warning.png")
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Reply = True
        Me.Hide()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Reply = False
        Me.Hide()
    End Sub
End Class