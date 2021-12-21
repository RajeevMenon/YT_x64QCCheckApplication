
Public Class MainForm

    Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As IntPtr, ByVal nIndex As Integer) As Integer
    Public Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Integer) As Integer
    Public Const GWL_STYLE As Integer = (-16)
    Public Const WS_VSCROLL As Integer = &H200000
    Public Const WS_HSCROLL As Integer = &H100000
    Public Shared Setting As New Settings

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim Version = AppControl.GetVersion("C:\TML_INI\QualityControlCheckAppliation\")
        Me.Text = Me.Text & " [ Ver:" & Version & "]"
        Setting = AppControl.GetSettings("C:\TML_INI\QualityControlCheckAppliation\") 'System.Windows.Forms.Application.StartupPath)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            'Dim AD As AddedDocs = New AddedDocs
            'AD.TopLevel = False
            'PanelSubForm.Controls.Add(AD)
            'AD.Dock = DockStyle.Fill
            'AD.Show()

            Dim CheckPoint As CheckSheetStep = New CheckSheetStep

            CheckPoint = YTA_CheckSheet.ProcessStepNo(20, "46501497")
            If Not IsNothing(CheckPoint) Then
                RichTextBox_Step.Text = CheckPoint.ProcessNo & vbCrLf & CheckPoint.ProcessStep
                RichTextBox_ActivityToCheck.Text = CheckPoint.ActivityToCheck
                RichTextBox_AutoSize(RichTextBox_Step)
                RichTextBox_AutoSize(RichTextBox_ActivityToCheck)
                If CheckPoint.Method = CheckSheetStep.MethodOption.ProcedureStep Then
                    ProcedureStepView.TopLevel = False
                    ' ProcedureStepView.Parent = PanelSubForm
                    ProcedureStepView.InputFeatures(CheckPoint.ProcessStep, CheckPoint.ActivityToCheck, CheckPoint.ProcedureStepAction.ImagePath_ProcedureStep)
                    PanelSubForm.Controls.Add(ProcedureStepView)
                    'PanelSubForm.SetAutoScrollMargin(AD.Size.Width - PanelSubForm.Size.Width, AD.Size.Height - PanelSubForm.Size.Height)
                    ProcedureStepView.AutoScroll = True
                    ProcedureStepView.Dock = DockStyle.Fill
                    PanelSubForm.AutoScroll = True
                    ProcedureStepView.Show()
                End If
            End If

            CheckPoint = YTA_CheckSheet.ProcessStepNo(30, "46501497")
            If Not IsNothing(CheckPoint) Then
                RichTextBox_Step.Text = CheckPoint.ProcessNo & vbCrLf & CheckPoint.ProcessStep
                RichTextBox_ActivityToCheck.Text = CheckPoint.ActivityToCheck
                RichTextBox_AutoSize(RichTextBox_Step)
                RichTextBox_AutoSize(RichTextBox_ActivityToCheck)
                If CheckPoint.Method = CheckSheetStep.MethodOption.UserIput Then
                    SelectUserInput.TopLevel = False
                    SelectUserInput.Message = CheckPoint.UserInputAction.UserActionMessage
                    SelectUserInput.inputValues = CheckPoint.UserInputAction.UserInputList
                    PanelSubForm.Controls.Add(SelectUserInput)
                    SelectUserInput.AutoScroll = True
                    SelectUserInput.Dock = DockStyle.Fill
                    PanelSubForm.AutoScroll = True
                    SelectUserInput.Show()
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub MainForm_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        Try
            If (e.Modifiers = Keys.Control And e.KeyCode = Keys.S) Then
                MsgBox("hi")
            End If
        Catch ex As Exception

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
