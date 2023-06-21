Public Class MakeUserInput
    Public InputValueOld As String
    Public InputValueNew As String
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
    Dim WMsg As New WarningForm
    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Try
            Result = Status.Finish
            InputValueNew = InputTextBox.Text.Trim
            Me.Hide()
        Catch ex As Exception
            InputValueNew = "NaN"
        End Try
    End Sub
    Private Sub SelectUserInput_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Result = Status.Start
            RichTextBoxMessage.Text = Message
            RichTextBox_AutoSize(RichTextBoxMessage)
            InputTextBox.Text = ""
            If InputValueOld.Length > 0 Then
                InputTextBox.Text = InputValueOld
            End If
            InputTextBox.Select()
        Catch ex As Exception
            RichTextBoxMessage.Text = ""
            InputValueNew = "NaN"
        End Try
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Result = Status.InActive
        InputValueNew = "EXIT"
        Me.Close()
        MainForm.InspectionStatus(MainForm.CurrentCheckPoint, False)
        MainForm.ShowCompleteInspection()
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
    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            MainForm.PrintQcc_Rev1()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub InputTextBox_KeyDown(sender As Object, e As KeyEventArgs) Handles InputTextBox.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                If MainForm.CurrentCheckPoint.MakeUserInputAction.UserInputSaveConnectionString.Length > 0 Then

                    If MainForm.CurrentCheckPoint.MakeUserInputAction.UserInputSaveTableName = "QR_CHECK" Then 'DLM QR Label. Here connectionstring is empty
                        Dim CheckString = "MFR:YOKOGAWA;CAT:EXT S/N;S/N:"
                        If InputTextBox.Text.Trim.ToUpper Like CheckString & MainForm.CustOrd.SERIAL_NO Then
                            InputTextBox.BackColor = Color.Green
                            MainForm.wait(1)
                            MainForm.InspectionStatus(MainForm.CurrentCheckPoint, True)
                            MainForm.Button2.PerformClick()
                            Me.Close()
                        End If
                    Else

                        Dim ErrMsg As String = ""
                        Dim TmlEty As New MFG_ENTITY.Op(MainForm.CurrentCheckPoint.MakeUserInputAction.UserInputSaveConnectionString)
                        Dim Sql(0) As String
                        Sql(0) = "UPDATE " & MainForm.CurrentCheckPoint.MakeUserInputAction.UserInputSaveTableName
                        Sql(0) &= " SET " & MainForm.CurrentCheckPoint.MakeUserInputAction.UserInputSaveTableField & " = "
                        Sql(0) &= "'" & InputTextBox.Text.Trim.ToUpper & "';"
                        If Sql(0).Length = 0 Then
                            WMsg.Message = "Error: No Data to Save to DB!"
                            WMsg.ShowDialog()
                            Exit Sub
                        End If
                        TmlEty.ExecuteTransactionQuery(Sql, ErrMsg:=ErrMsg)
                        If ErrMsg.Length = 0 Then
                            InputTextBox.BackColor = Color.Green
                            MainForm.wait(1)
                            MainForm.CurrentCheckPoint.MakeUserInputAction.UserInputSaveUserValue = InputTextBox.Text.Trim.ToUpper
                            MainForm.InspectionStatus(MainForm.CurrentCheckPoint, True)
                            MainForm.Button2.PerformClick()
                            Me.Close()
                        Else
                            InputTextBox.BackColor = Color.Red
                            MainForm.wait(1)
                            InputTextBox.BackColor = Color.White
                            MainForm.InspectionStatus(MainForm.CurrentCheckPoint, False)
                            WMsg.Message = "Error: " & ErrMsg
                            WMsg.ShowDialog()
                            Exit Sub
                        End If

                    End If




                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

End Class