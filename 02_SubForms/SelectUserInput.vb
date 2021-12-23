Public Class SelectUserInput
    Public selectedItemValue As String = "NaN"
    Public inputValues() As String
    Public Message As String
    Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As IntPtr, ByVal nIndex As Integer) As Integer
    Public Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Integer) As Integer
    Public Const GWL_STYLE As Integer = (-16)
    Public Const WS_VSCROLL As Integer = &H200000
    Public Const WS_HSCROLL As Integer = &H100000

    Enum Status
        InActive
        Start
        Finish
    End Enum
    Public Result As Status = Status.InActive
    Dim rnd As New Random()
    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Try
            Result = Status.Finish
            selectedItemValue = ListView1.SelectedItems.Item(0).Text
            Me.Hide()
        Catch ex As Exception
            selectedItemValue = "NaN"
        End Try
    End Sub
    Private Sub SelectUserInput_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Result = Status.Start
            RichTextBoxMessage.Text = Message
            RichTextBox_AutoSize(RichTextBoxMessage)
            ListView1.Items.Clear()
            If inputValues.Length > 0 Then
                Dim ShuffledItems = inputValues.OrderBy(Function() rnd.Next).ToArray()
                For Each inputvalue In ShuffledItems
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
        Result = Status.InActive
        selectedItemValue = "EXIT"
        Me.Close()
        MainForm.InspectionStatus(MainForm.CurrentCheckPoint, False)
        MainForm.ShowCompleteInspection()
    End Sub
    Private Sub ListView1_MouseDoubleClick(sender As Object, e As EventArgs) Handles ListView1.MouseDoubleClick
        Try
            Result = Status.Finish
            selectedItemValue = ListView1.SelectedItems.Item(0).Text
            If selectedItemValue = MainForm.CurrentCheckPoint.UserInputAction.UserInputCorrect Then
                ListView1.BackColor = Color.Green
                MainForm.wait(1)
                MainForm.InspectionStatus(MainForm.CurrentCheckPoint, True)
                MainForm.Button2.PerformClick()
                Me.Close()
            Else
                ListView1.BackColor = Color.Red
                MainForm.wait(1)
                ListView1.BackColor = Color.White
                MainForm.InspectionStatus(MainForm.CurrentCheckPoint, False)
                ListView1.Items.Clear()
                If inputValues.Length > 0 Then
                    Dim ShuffledItems = inputValues.OrderBy(Function() Rnd.Next).ToArray()
                    For Each inputvalue In ShuffledItems
                        ListView1.Items.Add(inputvalue)
                    Next
                End If
            End If
        Catch ex As Exception
            selectedItemValue = "NaN"
        End Try
    End Sub
    Private Sub ListView1_MouseClick(sender As Object, e As MouseEventArgs) Handles ListView1.MouseClick
        Try
            Result = Status.Finish
            selectedItemValue = ListView1.SelectedItems.Item(0).Text
            If selectedItemValue = MainForm.CurrentCheckPoint.UserInputAction.UserInputCorrect Then
                ListView1.BackColor = Color.Green
                MainForm.InspectionStatus(MainForm.CurrentCheckPoint, True)
                MainForm.wait(1)
                MainForm.Button2.PerformClick()
                Me.Close()
            Else
                ListView1.BackColor = Color.Red
                MainForm.InspectionStatus(MainForm.CurrentCheckPoint, False)
                MainForm.wait(1)
                ListView1.BackColor = Color.White
                ListView1.Items.Clear()
                If inputValues.Length > 0 Then
                    Dim ShuffledItems = inputValues.OrderBy(Function() Rnd.Next).ToArray()
                    For Each inputvalue In ShuffledItems
                        ListView1.Items.Add(inputvalue)
                    Next
                End If
            End If
        Catch ex As Exception
            selectedItemValue = "NaN"
        End Try
    End Sub
    Private Sub RichTextBox_AutoSize(ByVal Rc As RichTextBox)
        Try

            'Center the text in the RichBox
            Rc.SelectAll()
            Rc.SelectionAlignment = HorizontalAlignment.Center
            'Is scrollbar visible?
            Dim bVScrollBar As Boolean
            bVScrollBar = ((GetWindowLong(Rc.Handle, GWL_STYLE) And WS_VSCROLL) = WS_VSCROLL)
            Select Case bVScrollBar
                Case True
                    'Scrollbar is visible - Make it smaller
                    Do
                        Rc.ZoomFactor = Rc.ZoomFactor - 0.01
                        bVScrollBar = ((GetWindowLong(Rc.Handle, GWL_STYLE) And WS_VSCROLL) = WS_VSCROLL)
                        'If the scrollbar is no longer visible we are done!
                        If bVScrollBar = False Then Exit Do
                    Loop
                Case False
                    'Scrollbar is not visible - Make it bigger
                    Do
                        Rc.ZoomFactor = Rc.ZoomFactor + 0.01
                        bVScrollBar = ((GetWindowLong(Rc.Handle, GWL_STYLE) And WS_VSCROLL) = WS_VSCROLL)
                        If bVScrollBar = True Then
                            Do
                                'Found the scrollbar, make smaller until bar is not visible
                                Rc.ZoomFactor = Rc.ZoomFactor - 0.01
                                bVScrollBar = ((GetWindowLong(Rc.Handle, GWL_STYLE) And WS_VSCROLL) = WS_VSCROLL)
                                'If the scrollbar is no longer visible we are done!
                                If bVScrollBar = False Then Exit Do
                            Loop
                            Exit Do
                        End If
                    Loop
            End Select
        Catch ex As Exception

        End Try
    End Sub

End Class