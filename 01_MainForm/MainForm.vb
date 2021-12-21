
Public Class MainForm

    Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As IntPtr, ByVal nIndex As Integer) As Integer
    Public Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Integer) As Integer
    Public Const GWL_STYLE As Integer = (-16)
    Public Const WS_VSCROLL As Integer = &H200000
    Public Const WS_HSCROLL As Integer = &H100000
    Public Shared Setting As New Settings

    Public CurrentCheckPoint As CheckSheetStep

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim Version = AppControl.GetVersion("C:\TML_INI\QualityControlCheckAppliation\")
        Me.Text = Me.Text & " [ Ver:" & Version & "]"
        Setting = AppControl.GetSettings("C:\TML_INI\QualityControlCheckAppliation\") 'System.Windows.Forms.Application.StartupPath)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            Dim ErMsg As String = ""
            CurrentCheckPoint = New CheckSheetStep
            For StepNumber As Integer = Integer.Parse(RichTextBox_Step.Text) + 1 To 210
                RichTextBox_ActivityToCheck.Text = "Wait.."
                CurrentCheckPoint = YTA_CheckSheet.ProcessStepNo(StepNo:=StepNumber, Initial:="46501497", IndexNo:=100003859546, ErrMsg:=ErMsg)
                If Not IsDBNull(CurrentCheckPoint) Then
                    If Not IsNothing(CurrentCheckPoint) Then
                        If Not IsNothing(CurrentCheckPoint.ActivityToCheck) Then
                            RichTextBox_Step.Text = CurrentCheckPoint.ProcessNo '& vbCrLf & CheckPoint.ProcessStep
                            RichTextBox_ActivityToCheck.Text = CurrentCheckPoint.ActivityToCheck
                            RichTextBox_AutoSize(RichTextBox_Step)
                            RichTextBox_AutoSize(RichTextBox_ActivityToCheck)
                            If CurrentCheckPoint.Method = CheckSheetStep.MethodOption.ProcedureStep Then
                                ' ProcedureStepView.Parent = PanelSubForm
                                'PanelSubForm.SetAutoScrollMargin(AD.Size.Width - PanelSubForm.Size.Width, AD.Size.Height - PanelSubForm.Size.Height)
                                Dim PSV As New ProcedureStepView
                                PSV.TopLevel = False
                                PSV.InputFeatures(CurrentCheckPoint.ProcessStep, CurrentCheckPoint.ActivityToCheck, CurrentCheckPoint.ProcedureStepAction.ImagePath_ProcedureStep)
                                PanelSubForm.Controls.Add(PSV)
                                PSV.AutoScroll = True
                                PSV.Dock = DockStyle.Fill
                                PanelSubForm.AutoScroll = True
                                PSV.PictureBox1.Select()
                                PSV.Show()
                                Me.Refresh()
                                Exit For
                            ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.UserIput Then
                                Dim SUI As SelectUserInput = New SelectUserInput
                                SUI.TopLevel = False
                                SUI.Message = CurrentCheckPoint.UserInputAction.UserActionMessage
                                SUI.inputValues = CurrentCheckPoint.UserInputAction.UserInputList
                                PanelSubForm.Controls.Add(SUI)
                                SUI.AutoScroll = True
                                SUI.Dock = DockStyle.Fill
                                PanelSubForm.AutoScroll = True
                                SUI.ListView1.Select()
                                SUI.Show()
                                Exit For
                            End If
                        End If
                    End If
                End If
                If Not IsNothing(CurrentCheckPoint) Then
                    RichTextBox_Step.Text = CurrentCheckPoint.ProcessNo '& vbCrLf & CheckPoint.ProcessStep
                    RichTextBox_ActivityToCheck.Text = CurrentCheckPoint.ActivityToCheck
                    RichTextBox_AutoSize(RichTextBox_Step)
                    RichTextBox_AutoSize(RichTextBox_ActivityToCheck)
                    If CurrentCheckPoint.Method = CheckSheetStep.MethodOption.ProcedureStep Then
                        ' ProcedureStepView.Parent = PanelSubForm
                        'PanelSubForm.SetAutoScrollMargin(AD.Size.Width - PanelSubForm.Size.Width, AD.Size.Height - PanelSubForm.Size.Height)
                        Dim PSV As New ProcedureStepView
                        PSV.TopLevel = False
                        PSV.InputFeatures(CurrentCheckPoint.ProcessStep, CurrentCheckPoint.ActivityToCheck, CurrentCheckPoint.ProcedureStepAction.ImagePath_ProcedureStep)
                        PanelSubForm.Controls.Add(PSV)
                        PSV.AutoScroll = True
                        PSV.Dock = DockStyle.Fill
                        PanelSubForm.AutoScroll = True
                        PSV.PictureBox1.Select()
                        PSV.Show()
                        Me.Refresh()
                        Exit For
                    ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.UserIput Then
                        Dim SUI As SelectUserInput = New SelectUserInput
                        SUI.TopLevel = False
                        SUI.Message = CurrentCheckPoint.UserInputAction.UserActionMessage
                        SUI.inputValues = CurrentCheckPoint.UserInputAction.UserInputList
                        PanelSubForm.Controls.Add(SUI)
                        SUI.AutoScroll = True
                        SUI.Dock = DockStyle.Fill
                        PanelSubForm.AutoScroll = True
                        SUI.ListView1.Select()
                        SUI.Show()
                        Exit For
                    End If
                End If
            Next

        Catch ex As Exception

        End Try
    End Sub


    Public Sub wait(ByVal interval As Integer)
        Dim sw As New Stopwatch
        Dim wait_time As Double
        Dim Elapsed_time As Double
        sw.Start()
        Do While sw.ElapsedMilliseconds < (interval * 1000)
            ' Allows UI to remain responsive
            Application.DoEvents()
            Elapsed_time = (sw.ElapsedMilliseconds + 1) / 1000
            wait_time = interval - Elapsed_time
        Loop
        sw.Stop()
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
