
Public Class MainForm

    Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As IntPtr, ByVal nIndex As Integer) As Integer
    Public Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Integer) As Integer
    Public Const GWL_STYLE As Integer = (-16)
    Public Const WS_VSCROLL As Integer = &H200000
    Public Const WS_HSCROLL As Integer = &H100000
    Public Shared Setting As New Settings

    Public Initial As String
    Public Shared CurrentCheckPoint As CheckSheetStep
    Dim ErrMsg As String = ""
    Public CustOrd As POCO_YGSP.cust_ord
    Public Shared AllowedSteps As String()
    'Public Shared AllCheckResult As List(Of CheckSheetStep)
    Public Shared AllCheckResult As CheckSheetStep()
    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim Version = AppControl.GetVersion("C:\TML_INI\QualityControlCheckAppliation\")
        Me.Text = Me.Text & " [ Ver:" & Version & "]"
        Setting = AppControl.GetSettings("C:\TML_INI\QualityControlCheckAppliation\") 'System.Windows.Forms.Application.StartupPath)
        AllowedSteps = Setting.Var_08_StepsAllowed.Split(",")
        ReDim AllCheckResult(AllowedSteps.Length - 1)
        Dim BEY As New Login
        BEY.TopLevel = False
        PanelSubForm.Controls.Add(BEY)
        BEY.AutoScroll = True
        BEY.Dock = DockStyle.Fill
        PanelSubForm.AutoScroll = True
        BEY.UsernameTextBox.Select()
        BEY.Show()
        Me.Refresh()
    End Sub
    Public Sub LoadIndexScan()
        Try
            PanelSubForm.Controls.Clear()
            Dim BEY As New BarcodeEntry
            BEY.TopLevel = False
            PanelSubForm.Controls.Add(BEY)
            BEY.AutoScroll = True
            BEY.Dock = DockStyle.Fill
            PanelSubForm.AutoScroll = True
            BEY.Show()
            BEY.TextBox_Scan.Focus()
        Catch ex As Exception

        End Try
    End Sub
    Public Sub FirstCheckPoint()
        Try
            Dim ErMsg As String = ""
            CurrentCheckPoint = New CheckSheetStep
            Dim StepNumbers = Setting.Var_08_StepsAllowed.Split(",")

            RichTextBox_ActivityToCheck.Text = "Wait.."
            CurrentCheckPoint = YTA_CheckSheet.ProcessStepNo(StepNo:=Integer.Parse(StepNumbers(0)), Initial:=Initial, CustOrd, ErrMsg:=ErMsg)
            If Not IsDBNull(CurrentCheckPoint) Then
                If Not IsNothing(CurrentCheckPoint) Then
                    If Not IsNothing(CurrentCheckPoint.ActivityToCheck) Then
                        RichTextBox_Step.Text = CurrentCheckPoint.ProcessNo '& vbCrLf & CheckPoint.ProcessStep
                        TextBox_Step.Text = CurrentCheckPoint.StepNo
                        RichTextBox_ActivityToCheck.Text = CurrentCheckPoint.ActivityToCheck
                        RichTextBox_AutoSize(RichTextBox_Step)
                        RichTextBox_AutoSize(RichTextBox_ActivityToCheck)
                        RichTextBox_ActivityToCheck.Refresh()
                        PanelSubForm.Controls.Clear()

                        If CurrentCheckPoint.Method = CheckSheetStep.MethodOption.ProcedureStep Then
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

                        ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.UserIput Then
                            Dim SUI As SelectUserInput = New SelectUserInput
                            SUI.TopLevel = False
                            SUI.Message = CurrentCheckPoint.UserInputAction.UserActionMessage
                            SUI.inputValues = CurrentCheckPoint.UserInputAction.UserInputList
                            PanelSubForm.Controls.Add(SUI)
                            'SUI.BringToFront()
                            SUI.AutoScroll = True
                            SUI.Dock = DockStyle.Fill
                            PanelSubForm.AutoScroll = True
                            SUI.ListView1.Select()
                            SUI.Show()
                            Me.Refresh()

                        ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.SinglePntInst Then
                            Dim SPI As SinglePointInstruction = New SinglePointInstruction
                            SPI.TopLevel = False
                            SPI.InputFeatures(CurrentCheckPoint.SinglePointAction.ImagePath_SPI_1, CurrentCheckPoint.SinglePointAction.ImagePath_SPI_2, CurrentCheckPoint.SinglePointAction.SPI_Message)
                            PanelSubForm.Controls.Add(SPI)
                            SPI.AutoScroll = True
                            SPI.Dock = DockStyle.Fill
                            PanelSubForm.AutoScroll = True
                            SPI.Show()
                            Me.Refresh()
                        ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.DocumentCheck Then
                            Dim VDOC As ViewDocument = New ViewDocument
                            VDOC.TopLevel = False
                            VDOC.curNaviageUrl = CurrentCheckPoint.ViewDocAction.PdfPath_DocumentCheck
                            PanelSubForm.Controls.Add(VDOC)
                            VDOC.AutoScroll = True
                            VDOC.Dock = DockStyle.Fill
                            PanelSubForm.AutoScroll = True
                            VDOC.Show()
                            Me.Refresh()
                        End If
                    End If
                End If
            End If




        Catch ex As Exception

        End Try
    End Sub
    Public Sub LastCheckPoint()
        Try
            Dim ErMsg As String = ""
            CurrentCheckPoint = New CheckSheetStep
            Dim StepNumbers = Setting.Var_08_StepsAllowed.Split(",")

            RichTextBox_ActivityToCheck.Text = "Wait.."
            CurrentCheckPoint = YTA_CheckSheet.ProcessStepNo(StepNo:=Integer.Parse(StepNumbers(StepNumbers.Length - 1)), Initial:=Initial, CustOrd, ErrMsg:=ErMsg)
            If Not IsDBNull(CurrentCheckPoint) Then
                If Not IsNothing(CurrentCheckPoint) Then
                    If Not IsNothing(CurrentCheckPoint.ActivityToCheck) Then
                        RichTextBox_Step.Text = CurrentCheckPoint.ProcessNo '& vbCrLf & CheckPoint.ProcessStep
                        TextBox_Step.Text = CurrentCheckPoint.StepNo
                        RichTextBox_ActivityToCheck.Text = CurrentCheckPoint.ActivityToCheck
                        RichTextBox_AutoSize(RichTextBox_Step)
                        RichTextBox_AutoSize(RichTextBox_ActivityToCheck)
                        RichTextBox_ActivityToCheck.Refresh()
                        PanelSubForm.Controls.Clear()

                        If CurrentCheckPoint.Method = CheckSheetStep.MethodOption.ProcedureStep Then
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

                        ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.UserIput Then
                            Dim SUI As SelectUserInput = New SelectUserInput
                            SUI.TopLevel = False
                            SUI.Message = CurrentCheckPoint.UserInputAction.UserActionMessage
                            SUI.inputValues = CurrentCheckPoint.UserInputAction.UserInputList
                            PanelSubForm.Controls.Add(SUI)
                            'SUI.BringToFront()
                            SUI.AutoScroll = True
                            SUI.Dock = DockStyle.Fill
                            PanelSubForm.AutoScroll = True
                            SUI.ListView1.Select()
                            SUI.Show()
                            Me.Refresh()

                        ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.SinglePntInst Then
                            Dim SPI As SinglePointInstruction = New SinglePointInstruction
                            SPI.TopLevel = False
                            SPI.InputFeatures(CurrentCheckPoint.SinglePointAction.ImagePath_SPI_1, CurrentCheckPoint.SinglePointAction.ImagePath_SPI_2, CurrentCheckPoint.SinglePointAction.SPI_Message)
                            PanelSubForm.Controls.Add(SPI)
                            SPI.AutoScroll = True
                            SPI.Dock = DockStyle.Fill
                            PanelSubForm.AutoScroll = True
                            SPI.Show()
                            Me.Refresh()
                        ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.DocumentCheck Then
                            Dim VDOC As ViewDocument = New ViewDocument
                            VDOC.TopLevel = False
                            VDOC.curNaviageUrl = CurrentCheckPoint.ViewDocAction.PdfPath_DocumentCheck
                            PanelSubForm.Controls.Add(VDOC)
                            VDOC.AutoScroll = True
                            VDOC.Dock = DockStyle.Fill
                            PanelSubForm.AutoScroll = True
                            VDOC.Show()
                            Me.Refresh()
                        End If
                    End If
                End If
            End If




        Catch ex As Exception

        End Try
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            Dim LoopStarted As Boolean = False
            Dim ErMsg As String = ""
            CurrentCheckPoint = New CheckSheetStep
            For i As Integer = Array.IndexOf(AllowedSteps, TextBox_Step.Text) + 1 To AllowedSteps.Length - 1 Step 1
                LoopStarted = True
                RichTextBox_ActivityToCheck.Text = "Wait.."
                PanelSubForm.Controls.Clear()
                CurrentCheckPoint = YTA_CheckSheet.ProcessStepNo(StepNo:=Integer.Parse(AllowedSteps(i)), Initial:=Initial, CustOrd, ErrMsg:=ErMsg)
                If Not IsDBNull(CurrentCheckPoint) Then
                    If Not IsNothing(CurrentCheckPoint) Then
                        If Not IsNothing(CurrentCheckPoint.ActivityToCheck) Then
                            RichTextBox_Step.Text = CurrentCheckPoint.ProcessNo '& vbCrLf & CheckPoint.ProcessStep
                            TextBox_Step.Text = CurrentCheckPoint.StepNo
                            RichTextBox_ActivityToCheck.Text = CurrentCheckPoint.ActivityToCheck
                            RichTextBox_AutoSize(RichTextBox_Step)
                            RichTextBox_AutoSize(RichTextBox_ActivityToCheck)
                            RichTextBox_ActivityToCheck.Refresh()
                            PanelSubForm.Controls.Clear()

                            If CurrentCheckPoint.Method = CheckSheetStep.MethodOption.ProcedureStep Then
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
                                'SUI.BringToFront()
                                SUI.AutoScroll = True
                                SUI.Dock = DockStyle.Fill
                                PanelSubForm.AutoScroll = True
                                SUI.ListView1.Select()
                                SUI.Show()
                                Me.Refresh()
                                Exit For
                            ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.SinglePntInst Then
                                Dim SPI As SinglePointInstruction = New SinglePointInstruction
                                SPI.TopLevel = False
                                SPI.InputFeatures(CurrentCheckPoint.SinglePointAction.ImagePath_SPI_1, CurrentCheckPoint.SinglePointAction.ImagePath_SPI_2, CurrentCheckPoint.SinglePointAction.SPI_Message)
                                PanelSubForm.Controls.Add(SPI)
                                SPI.AutoScroll = True
                                SPI.Dock = DockStyle.Fill
                                PanelSubForm.AutoScroll = True
                                SPI.Show()
                                Me.Refresh()
                                Exit For
                            ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.DocumentCheck Then
                                Dim VDOC As ViewDocument = New ViewDocument
                                VDOC.TopLevel = False
                                VDOC.curNaviageUrl = CurrentCheckPoint.ViewDocAction.PdfPath_DocumentCheck
                                PanelSubForm.Controls.Add(VDOC)
                                VDOC.AutoScroll = True
                                VDOC.Dock = DockStyle.Fill
                                PanelSubForm.AutoScroll = True
                                VDOC.Show()
                                Me.Refresh()
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next

            If LoopStarted = False And RichTextBox_Step.BackColor <> Color.Yellow Then
                Dim AllPass As Boolean = True
                For Each item In AllCheckResult
                    If Not item.CheckResult Like "*True*" Then
                        AllPass = False
                        Exit For
                    End If
                Next
                If AllPass = True Then
                    RichTextBox_ActivityToCheck.Text = ""
                    RichTextBox_Step.Text = ""
                    TextBox_Step.Text = ""
                    RichTextBox_Step.BackColor = Color.White
                    PanelSubForm.Controls.Clear()
                    Dim BEY As New CompleteInspection
                    BEY.TopLevel = False
                    PanelSubForm.Controls.Add(BEY)
                    BEY.AutoScroll = True
                    BEY.Dock = DockStyle.Fill
                    PanelSubForm.AutoScroll = True
                    BEY.Show()
                    BEY.Button1.Select()
                Else
                    FirstCheckPoint()
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim LoopStarted As Boolean = False
            Dim ErMsg As String = ""
            CurrentCheckPoint = New CheckSheetStep
            Dim StartIndex As Integer = 0
            If TextBox_Step.Text = "" Then
                StartIndex = AllowedSteps.Length
            Else
                StartIndex = Array.IndexOf(AllowedSteps, TextBox_Step.Text)
            End If
            For i As Integer = StartIndex - 1 To 0 Step -1
                LoopStarted = True
                RichTextBox_ActivityToCheck.Text = "Wait.."
                PanelSubForm.Controls.Clear()
                CurrentCheckPoint = YTA_CheckSheet.ProcessStepNo(StepNo:=Integer.Parse(AllowedSteps(i)), Initial:="46501497", CustOrd, ErrMsg:=ErMsg)
                If Not IsDBNull(CurrentCheckPoint) Then
                    If Not IsNothing(CurrentCheckPoint) Then
                        If Not IsNothing(CurrentCheckPoint.ActivityToCheck) Then
                            RichTextBox_Step.Text = CurrentCheckPoint.ProcessNo '& vbCrLf & CheckPoint.ProcessStep
                            TextBox_Step.Text = CurrentCheckPoint.StepNo
                            RichTextBox_ActivityToCheck.Text = CurrentCheckPoint.ActivityToCheck
                            RichTextBox_AutoSize(RichTextBox_Step)
                            RichTextBox_AutoSize(RichTextBox_ActivityToCheck)
                            RichTextBox_ActivityToCheck.Refresh()
                            PanelSubForm.Controls.Clear()

                            If CurrentCheckPoint.Method = CheckSheetStep.MethodOption.ProcedureStep Then
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
                                'SUI.BringToFront()
                                SUI.AutoScroll = True
                                SUI.Dock = DockStyle.Fill
                                PanelSubForm.AutoScroll = True
                                SUI.ListView1.Select()
                                SUI.Show()
                                Me.Refresh()
                                Exit For
                            ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.SinglePntInst Then
                                Dim SPI As SinglePointInstruction = New SinglePointInstruction
                                SPI.TopLevel = False
                                SPI.InputFeatures(CurrentCheckPoint.SinglePointAction.ImagePath_SPI_1, CurrentCheckPoint.SinglePointAction.ImagePath_SPI_2, CurrentCheckPoint.SinglePointAction.SPI_Message)
                                PanelSubForm.Controls.Add(SPI)
                                SPI.AutoScroll = True
                                SPI.Dock = DockStyle.Fill
                                PanelSubForm.AutoScroll = True
                                SPI.Show()
                                Me.Refresh()
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next

            If LoopStarted = False And RichTextBox_Step.BackColor <> Color.Yellow Then
                Dim AllPass As Boolean = True
                For Each item In AllCheckResult
                    If Not item.CheckResult Like "*True*" Then
                        AllPass = False
                        Exit For
                    End If
                Next
                If AllPass = True Then
                    RichTextBox_ActivityToCheck.Text = ""
                    RichTextBox_Step.Text = ""
                    TextBox_Step.Text = ""
                    RichTextBox_Step.BackColor = Color.White
                    PanelSubForm.Controls.Clear()
                    Dim BEY As New CompleteInspection
                    BEY.TopLevel = False
                    PanelSubForm.Controls.Add(BEY)
                    BEY.AutoScroll = True
                    BEY.Dock = DockStyle.Fill
                    PanelSubForm.AutoScroll = True
                    BEY.Show()
                    BEY.Button1.Select()
                Else
                    LastCheckPoint()
                End If
            End If

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
    Public Function InspectionStatus(ByVal InspectionResult As CheckSheetStep, ByVal CheckResult As Boolean) As String
        Try
            InspectionResult.CheckResult = CheckResult
            AllCheckResult(Array.IndexOf(AllowedSteps, InspectionResult.StepNo)) = InspectionResult
            SetInspectionColor(InspectionResult.StepNo)
        Catch ex As Exception

        End Try
    End Function
    Public Sub SetInspectionColor(ByVal StepNo As String)
        Try
            If Not IsNothing(MainForm.AllCheckResult(Array.IndexOf(MainForm.AllowedSteps, StepNo.ToString))) Then
                If MainForm.AllCheckResult(Array.IndexOf(MainForm.AllowedSteps, StepNo.ToString)).StepNo = StepNo _
    And MainForm.AllCheckResult(Array.IndexOf(MainForm.AllowedSteps, StepNo.ToString)).CheckResult = True Then
                    RichTextBox_Step.BackColor = Color.LightGreen
                ElseIf MainForm.AllCheckResult(Array.IndexOf(MainForm.AllowedSteps, StepNo.ToString)).StepNo = StepNo _
                        And MainForm.AllCheckResult(Array.IndexOf(MainForm.AllowedSteps, StepNo.ToString)).CheckResult = False Then
                    RichTextBox_Step.BackColor = Color.OrangeRed
                Else
                    RichTextBox_Step.BackColor = Color.Yellow
                End If
            Else
                RichTextBox_Step.BackColor = Color.Yellow
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub RichTextBox_Step_Click(sender As Object, e As EventArgs) Handles RichTextBox_Step.Click
        Try
            Dim ErMsg As String = ""
            If Not IsNumeric(RichTextBox_Step.Text) Then
                ErMsg = "Cannot load Inspection point. Click Next/Previous buttons."
                MsgBox(ErrMsg, MsgBoxStyle.Critical, "Error")
                Exit Sub
            End If
            CurrentCheckPoint = New CheckSheetStep
            RichTextBox_ActivityToCheck.Text = "Wait.."
            CurrentCheckPoint = YTA_CheckSheet.ProcessStepNo(StepNo:=Integer.Parse(RichTextBox_Step.Text), Initial:=Initial, CustOrd, ErrMsg:=ErMsg)
            If Not IsDBNull(CurrentCheckPoint) Then
                If Not IsNothing(CurrentCheckPoint) Then
                    If Not IsNothing(CurrentCheckPoint.ActivityToCheck) Then
                        RichTextBox_Step.Text = CurrentCheckPoint.ProcessNo '& vbCrLf & CheckPoint.ProcessStep
                        TextBox_Step.Text = CurrentCheckPoint.StepNo
                        RichTextBox_ActivityToCheck.Text = CurrentCheckPoint.ActivityToCheck
                        RichTextBox_AutoSize(RichTextBox_Step)
                        RichTextBox_AutoSize(RichTextBox_ActivityToCheck)
                        RichTextBox_ActivityToCheck.Refresh()
                        PanelSubForm.Controls.Clear()

                        If CurrentCheckPoint.Method = CheckSheetStep.MethodOption.ProcedureStep Then
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

                        ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.UserIput Then
                            Dim SUI As SelectUserInput = New SelectUserInput
                            SUI.TopLevel = False
                            SUI.Message = CurrentCheckPoint.UserInputAction.UserActionMessage
                            SUI.inputValues = CurrentCheckPoint.UserInputAction.UserInputList
                            PanelSubForm.Controls.Add(SUI)
                            'SUI.BringToFront()
                            SUI.AutoScroll = True
                            SUI.Dock = DockStyle.Fill
                            PanelSubForm.AutoScroll = True
                            SUI.ListView1.Select()
                            SUI.Show()
                            Me.Refresh()

                        ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.SinglePntInst Then
                            Dim SPI As SinglePointInstruction = New SinglePointInstruction
                            SPI.TopLevel = False
                            SPI.InputFeatures(CurrentCheckPoint.SinglePointAction.ImagePath_SPI_1, CurrentCheckPoint.SinglePointAction.ImagePath_SPI_2, CurrentCheckPoint.SinglePointAction.SPI_Message)
                            PanelSubForm.Controls.Add(SPI)
                            SPI.AutoScroll = True
                            SPI.Dock = DockStyle.Fill
                            PanelSubForm.AutoScroll = True
                            SPI.Show()
                            Me.Refresh()
                        ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.DocumentCheck Then
                            Dim VDOC As ViewDocument = New ViewDocument
                            VDOC.TopLevel = False
                            VDOC.curNaviageUrl = CurrentCheckPoint.ViewDocAction.PdfPath_DocumentCheck
                            PanelSubForm.Controls.Add(VDOC)
                            VDOC.AutoScroll = True
                            VDOC.Dock = DockStyle.Fill
                            PanelSubForm.AutoScroll = True
                            VDOC.Show()
                            Me.Refresh()
                        End If
                    End If
                End If
            End If




        Catch ex As Exception

        End Try
    End Sub
End Class
