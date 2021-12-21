Public Class SelectUserInput
    Public selectedItemValue As String = "NaN"
    Public inputValues() As String
    Public Message As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            selectedItemValue = ListView1.SelectedItems.Item(0).Text
            Me.Hide()
        Catch ex As Exception
            selectedItemValue = "NaN"
        End Try
    End Sub

    Private Sub SelectUserInput_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            RichTextBoxMessage.Text = Message
            ListView1.Items.Clear()
            If inputValues.Length > 0 Then
                For Each inputvalue In inputValues
                    ListView1.Items.Add(inputvalue)
                Next
            End If
        Catch ex As Exception
            RichTextBoxMessage.Text = ""
            ListView1.Items.Clear()
            selectedItemValue = "NaN"
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        selectedItemValue = "EXIT"
        Me.Hide()
    End Sub

    Private Sub ListView1_MouseDoubleClick(sender As Object, e As EventArgs) Handles ListView1.MouseDoubleClick
        Try
            selectedItemValue = ListView1.SelectedItems.Item(0).Text
            Me.Hide()
        Catch ex As Exception
            selectedItemValue = "NaN"
        End Try
    End Sub

    Private Sub ListView1_MouseClick(sender As Object, e As MouseEventArgs) Handles ListView1.MouseClick
        Try
            selectedItemValue = ListView1.SelectedItems.Item(0).Text
            Me.Hide()
        Catch ex As Exception
            selectedItemValue = "NaN"
        End Try
    End Sub
End Class