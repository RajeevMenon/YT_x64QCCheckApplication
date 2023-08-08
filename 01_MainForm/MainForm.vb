
Public Class MainForm

    Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As IntPtr, ByVal nIndex As Integer) As Integer
    Public Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Integer) As Integer
    Public Const GWL_STYLE As Integer = (-16)
    Public Const WS_VSCROLL As Integer = &H200000
    Public Const WS_HSCROLL As Integer = &H100000
    Public Shared Setting As New Settings

#Region "OLD VERSION"
    'Public Initial As String
    'Public Shared CurrentCheckPoint As CheckSheetStep
    'Dim ErrMsg As String = ""
    'Public CustOrd As POCO_YGSP.cust_ord
    'Public Hipot As POCO_YGSP.hipot_tb
    'Public YTA_Crc As POCO_YGSP.yta710_inspection_tb
    'Public QcData As List(Of POCO_QA.yta_qcc_v1p2)
    'Public DataPlateCorrect As String
    'Public DataPlateCheck(3) As String
    'Public Shared AllowedSteps As String()
    'Public Shared AllCheckResult As CheckSheetStep()
    'Dim TmlEntityQA As MFG_ENTITY.Op
    'Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles Me.Load
    '    Dim Version = AppControl.GetVersion("C:\TML_INI\QualityControlCheckAppliation\")
    '    Me.Text = Me.Text & " [ Ver:" & Version & "]"
    '    Setting = AppControl.GetSettings("C:\TML_INI\QualityControlCheckAppliation\") 'System.Windows.Forms.Application.StartupPath)
    '    TmlEntityQA = New MFG_ENTITY.Op(Setting.Var_04_MySql_QA)
    '    AllowedSteps = Setting.Var_08_StepsCurrent.Split(",")
    '    ReDim AllCheckResult(AllowedSteps.Length - 1)
    '    Dim BEY As New Login
    '    BEY.TopLevel = False
    '    PanelSubForm.Controls.Add(BEY)
    '    BEY.AutoScroll = True
    '    BEY.Dock = DockStyle.Fill
    '    PanelSubForm.AutoScroll = True
    '    BEY.UsernameTextBox.Select()
    '    BEY.Show()
    '    Me.Refresh()
    'End Sub
    'Public Sub LoadIndexScan()
    '    Try
    '        PanelSubForm.Controls.Clear()
    '        Dim BEY As New BarcodeEntry
    '        BEY.TopLevel = False
    '        PanelSubForm.Controls.Add(BEY)
    '        BEY.AutoScroll = True
    '        BEY.Dock = DockStyle.Fill
    '        PanelSubForm.AutoScroll = True
    '        BEY.Show()
    '        BEY.TextBox_Scan.Focus()
    '    Catch ex As Exception

    '    End Try
    'End Sub
    'Public Sub FirstCheckPoint()
    '    Try
    '        Dim ErMsg As String = ""
    '        CurrentCheckPoint = New CheckSheetStep
    '        Dim StepNumbers = Setting.Var_08_StepsCurrent.Split(",")

    '        RichTextBox_ActivityToCheck.Text = "Wait.."
    '        CurrentCheckPoint = YTA_CheckSheet.ProcessStepNo(StepNo:=Integer.Parse(StepNumbers(0)), Initial:=Initial, CustOrd, ErrMsg:=ErMsg)
    '        If Not IsDBNull(CurrentCheckPoint) Then
    '            If Not IsNothing(CurrentCheckPoint) Then
    '                If Not IsNothing(CurrentCheckPoint.ActivityToCheck) Then
    '                    RichTextBox_Step.Text = CurrentCheckPoint.ProcessNo '& vbCrLf & CheckPoint.ProcessStep
    '                    TextBox_Step.Text = CurrentCheckPoint.StepNo
    '                    RichTextBox_ActivityToCheck.Text = CurrentCheckPoint.ActivityToCheck
    '                    RichTextBox_AutoSize(RichTextBox_Step)
    '                    RichTextBox_AutoSize(RichTextBox_ActivityToCheck)
    '                    RichTextBox_ActivityToCheck.Refresh()
    '                    PanelSubForm.Controls.Clear()

    '                    If CurrentCheckPoint.Method = CheckSheetStep.MethodOption.ProcedureStep Then
    '                        Dim PSV As New ProcedureStepView
    '                        PSV.TopLevel = False
    '                        PSV.InputFeatures(CurrentCheckPoint.ProcessStep, CurrentCheckPoint.ActivityToCheck, CurrentCheckPoint.ProcedureStepAction.ImagePath_ProcedureStep)
    '                        PanelSubForm.Controls.Add(PSV)
    '                        PSV.AutoScroll = True
    '                        PSV.Dock = DockStyle.Fill
    '                        PanelSubForm.AutoScroll = True
    '                        PSV.PictureBox1.Select()
    '                        PSV.Show()
    '                        Me.Refresh()

    '                    ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.UserIput Then
    '                        Dim SUI As SelectUserInput = New SelectUserInput
    '                        SUI.TopLevel = False
    '                        SUI.Message = CurrentCheckPoint.UserInputAction.UserActionMessage
    '                        SUI.inputValues = CurrentCheckPoint.UserInputAction.UserInputList
    '                        PanelSubForm.Controls.Add(SUI)
    '                        'SUI.BringToFront()
    '                        SUI.AutoScroll = True
    '                        SUI.Dock = DockStyle.Fill
    '                        PanelSubForm.AutoScroll = True
    '                        SUI.ListView1.Select()
    '                        SUI.Show()
    '                        Me.Refresh()
    '                    ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.SinglePntInst Then
    '                        Dim SPI As SinglePointInstruction = New SinglePointInstruction
    '                        SPI.TopLevel = False
    '                        SPI.InputFeatures(CurrentCheckPoint.SinglePointAction.ImagePath_SPI_1, CurrentCheckPoint.SinglePointAction.ImagePath_SPI_2, CurrentCheckPoint.SinglePointAction.SPI_Message)
    '                        PanelSubForm.Controls.Add(SPI)
    '                        SPI.AutoScroll = True
    '                        SPI.Dock = DockStyle.Fill
    '                        PanelSubForm.AutoScroll = True
    '                        SPI.Show()
    '                        Me.Refresh()
    '                    ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.DocumentCheck Then
    '                        Dim VDOC As ViewDocument = New ViewDocument
    '                        VDOC.TopLevel = False
    '                        VDOC.curNaviageUrl = CurrentCheckPoint.ViewDocAction.PdfPath_DocumentCheck
    '                        PanelSubForm.Controls.Add(VDOC)
    '                        VDOC.AutoScroll = True
    '                        VDOC.Dock = DockStyle.Fill
    '                        PanelSubForm.AutoScroll = True
    '                        VDOC.Show()
    '                        Me.Refresh()
    '                    ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.AddedDocs Then
    '                        Dim VDOC As AddedDocs = New AddedDocs
    '                        VDOC.TopLevel = False
    '                        PanelSubForm.Controls.Add(VDOC)
    '                        VDOC.AutoScroll = True
    '                        VDOC.Dock = DockStyle.Fill
    '                        PanelSubForm.AutoScroll = True
    '                        VDOC.Show()
    '                        Me.Refresh()
    '                    End If
    '                End If
    '            End If
    '        End If

    '    Catch ex As Exception

    '    End Try
    'End Sub
    'Public Sub LastCheckPoint()
    '    Try
    '        Dim ErMsg As String = ""
    '        CurrentCheckPoint = New CheckSheetStep
    '        Dim StepNumbers = Setting.Var_08_StepsCurrent.Split(",")

    '        RichTextBox_ActivityToCheck.Text = "Wait.."
    '        CurrentCheckPoint = YTA_CheckSheet.ProcessStepNo(StepNo:=Integer.Parse(StepNumbers(StepNumbers.Length - 1)), Initial:=Initial, CustOrd, ErrMsg:=ErMsg)
    '        If Not IsDBNull(CurrentCheckPoint) Then
    '            If Not IsNothing(CurrentCheckPoint) Then
    '                If Not IsNothing(CurrentCheckPoint.ActivityToCheck) Then
    '                    RichTextBox_Step.Text = CurrentCheckPoint.ProcessNo '& vbCrLf & CheckPoint.ProcessStep
    '                    TextBox_Step.Text = CurrentCheckPoint.StepNo
    '                    RichTextBox_ActivityToCheck.Text = CurrentCheckPoint.ActivityToCheck
    '                    RichTextBox_AutoSize(RichTextBox_Step)
    '                    RichTextBox_AutoSize(RichTextBox_ActivityToCheck)
    '                    RichTextBox_ActivityToCheck.Refresh()
    '                    PanelSubForm.Controls.Clear()

    '                    If CurrentCheckPoint.Method = CheckSheetStep.MethodOption.ProcedureStep Then
    '                        Dim PSV As New ProcedureStepView
    '                        PSV.TopLevel = False
    '                        PSV.InputFeatures(CurrentCheckPoint.ProcessStep, CurrentCheckPoint.ActivityToCheck, CurrentCheckPoint.ProcedureStepAction.ImagePath_ProcedureStep)
    '                        PanelSubForm.Controls.Add(PSV)
    '                        PSV.AutoScroll = True
    '                        PSV.Dock = DockStyle.Fill
    '                        PanelSubForm.AutoScroll = True
    '                        PSV.PictureBox1.Select()
    '                        PSV.Show()
    '                        Me.Refresh()

    '                    ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.UserIput Then
    '                        Dim SUI As SelectUserInput = New SelectUserInput
    '                        SUI.TopLevel = False
    '                        SUI.Message = CurrentCheckPoint.UserInputAction.UserActionMessage
    '                        SUI.inputValues = CurrentCheckPoint.UserInputAction.UserInputList
    '                        PanelSubForm.Controls.Add(SUI)
    '                        'SUI.BringToFront()
    '                        SUI.AutoScroll = True
    '                        SUI.Dock = DockStyle.Fill
    '                        PanelSubForm.AutoScroll = True
    '                        SUI.ListView1.Select()
    '                        SUI.Show()
    '                        Me.Refresh()

    '                    ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.SinglePntInst Then
    '                        Dim SPI As SinglePointInstruction = New SinglePointInstruction
    '                        SPI.TopLevel = False
    '                        SPI.InputFeatures(CurrentCheckPoint.SinglePointAction.ImagePath_SPI_1, CurrentCheckPoint.SinglePointAction.ImagePath_SPI_2, CurrentCheckPoint.SinglePointAction.SPI_Message)
    '                        PanelSubForm.Controls.Add(SPI)
    '                        SPI.AutoScroll = True
    '                        SPI.Dock = DockStyle.Fill
    '                        PanelSubForm.AutoScroll = True
    '                        SPI.Show()
    '                        Me.Refresh()
    '                    ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.DocumentCheck Then
    '                        Dim VDOC As ViewDocument = New ViewDocument
    '                        VDOC.TopLevel = False
    '                        VDOC.curNaviageUrl = CurrentCheckPoint.ViewDocAction.PdfPath_DocumentCheck
    '                        PanelSubForm.Controls.Add(VDOC)
    '                        VDOC.AutoScroll = True
    '                        VDOC.Dock = DockStyle.Fill
    '                        PanelSubForm.AutoScroll = True
    '                        VDOC.Show()
    '                        Me.Refresh()
    '                    ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.AddedDocs Then
    '                        Dim VDOC As AddedDocs = New AddedDocs
    '                        VDOC.TopLevel = False
    '                        PanelSubForm.Controls.Add(VDOC)
    '                        VDOC.AutoScroll = True
    '                        VDOC.Dock = DockStyle.Fill
    '                        PanelSubForm.AutoScroll = True
    '                        VDOC.Show()
    '                        Me.Refresh()
    '                    End If
    '                End If
    '            End If
    '        End If

    '    Catch ex As Exception

    '    End Try
    'End Sub
    'Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
    '    Try
    '        Setting = AppControl.GetSettings("C:\TML_INI\QualityControlCheckAppliation\") 'System.Windows.Forms.Application.StartupPath)
    '        Dim LoopStarted As Boolean = False
    '        Dim ErMsg As String = ""
    '        CurrentCheckPoint = New CheckSheetStep
    '        For i As Integer = Array.IndexOf(AllowedSteps, TextBox_Step.Text) + 1 To AllowedSteps.Length - 1 Step 1
    '            LoopStarted = True
    '            RichTextBox_ActivityToCheck.Text = "Wait.."
    '            PanelSubForm.Controls.Clear()
    '            CurrentCheckPoint = YTA_CheckSheet.ProcessStepNo(StepNo:=Integer.Parse(AllowedSteps(i)), Initial:=Initial, CustOrd, ErrMsg:=ErMsg)
    '            If Not IsDBNull(CurrentCheckPoint) Then
    '                If Not IsNothing(CurrentCheckPoint) Then
    '                    If Not IsNothing(CurrentCheckPoint.ActivityToCheck) Then
    '                        RichTextBox_Step.Text = CurrentCheckPoint.ProcessNo '& vbCrLf & CheckPoint.ProcessStep
    '                        TextBox_Step.Text = CurrentCheckPoint.StepNo
    '                        RichTextBox_ActivityToCheck.Text = CurrentCheckPoint.ActivityToCheck
    '                        RichTextBox_AutoSize(RichTextBox_Step)
    '                        RichTextBox_AutoSize(RichTextBox_ActivityToCheck)
    '                        RichTextBox_ActivityToCheck.Refresh()
    '                        PanelSubForm.Controls.Clear()

    '                        If CurrentCheckPoint.Method = CheckSheetStep.MethodOption.ProcedureStep Then
    '                            Dim PSV As New ProcedureStepView
    '                            PSV.TopLevel = False
    '                            PSV.InputFeatures(CurrentCheckPoint.ProcessStep, CurrentCheckPoint.ActivityToCheck, CurrentCheckPoint.ProcedureStepAction.ImagePath_ProcedureStep)
    '                            PanelSubForm.Controls.Add(PSV)
    '                            PSV.AutoScroll = True
    '                            PSV.Dock = DockStyle.Fill
    '                            PanelSubForm.AutoScroll = True
    '                            PSV.PictureBox1.Select()
    '                            PSV.Show()
    '                            Me.Refresh()
    '                            Exit For
    '                        ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.UserIput Then
    '                            Dim SUI As SelectUserInput = New SelectUserInput
    '                            SUI.TopLevel = False
    '                            SUI.Message = CurrentCheckPoint.UserInputAction.UserActionMessage
    '                            SUI.inputValues = CurrentCheckPoint.UserInputAction.UserInputList
    '                            PanelSubForm.Controls.Add(SUI)
    '                            'SUI.BringToFront()
    '                            SUI.AutoScroll = True
    '                            SUI.Dock = DockStyle.Fill
    '                            PanelSubForm.AutoScroll = True
    '                            SUI.ListView1.Select()
    '                            SUI.Show()
    '                            Me.Refresh()
    '                            Exit For
    '                        ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.SinglePntInst Then
    '                            Dim SPI As SinglePointInstruction = New SinglePointInstruction
    '                            SPI.TopLevel = False
    '                            SPI.InputFeatures(CurrentCheckPoint.SinglePointAction.ImagePath_SPI_1, CurrentCheckPoint.SinglePointAction.ImagePath_SPI_2, CurrentCheckPoint.SinglePointAction.SPI_Message)
    '                            PanelSubForm.Controls.Add(SPI)
    '                            SPI.AutoScroll = True
    '                            SPI.Dock = DockStyle.Fill
    '                            PanelSubForm.AutoScroll = True
    '                            SPI.Show()
    '                            Me.Refresh()
    '                            Exit For
    '                        ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.DocumentCheck Then
    '                            Dim VDOC As ViewDocument = New ViewDocument
    '                            VDOC.TopLevel = False
    '                            VDOC.curNaviageUrl = CurrentCheckPoint.ViewDocAction.PdfPath_DocumentCheck
    '                            PanelSubForm.Controls.Add(VDOC)
    '                            VDOC.AutoScroll = True
    '                            VDOC.Dock = DockStyle.Fill
    '                            PanelSubForm.AutoScroll = True
    '                            VDOC.Show()
    '                            Me.Refresh()
    '                            Exit For
    '                        ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.AddedDocs Then
    '                            Dim VDOC As AddedDocs = New AddedDocs
    '                            VDOC.TopLevel = False
    '                            PanelSubForm.Controls.Add(VDOC)
    '                            VDOC.AutoScroll = True
    '                            VDOC.Dock = DockStyle.Fill
    '                            PanelSubForm.AutoScroll = True
    '                            VDOC.Show()
    '                            Me.Refresh()
    '                            Exit For
    '                        End If
    '                    End If
    '                End If
    '            End If
    '        Next

    '        If LoopStarted = False And Not (RichTextBox_Step.BackColor = Color.Yellow Or RichTextBox_Step.BackColor = Color.OrangeRed) Then
    '            Dim AllPass As Boolean = True
    '            For Each item In AllCheckResult
    '                If Not IsNothing(item) Then
    '                    If Not item.CheckResult Like "*True*" Then
    '                        AllPass = False
    '                        Exit For
    '                    End If
    '                Else
    '                    AllPass = False
    '                    Exit For
    '                End If
    '            Next
    '            If AllPass = True Then
    '                ShowCompleteInspection()
    '                'RichTextBox_ActivityToCheck.Text = ""
    '                'RichTextBox_Step.Text = ""
    '                'TextBox_Step.Text = ""
    '                'RichTextBox_Step.BackColor = Color.White
    '                'PanelSubForm.Controls.Clear()
    '                'Dim BEY As New CompleteInspection
    '                'BEY.TopLevel = False
    '                'PanelSubForm.Controls.Add(BEY)
    '                'BEY.AutoScroll = True
    '                'BEY.Dock = DockStyle.Fill
    '                'PanelSubForm.AutoScroll = True
    '                'BEY.Show()
    '                'BEY.Button1.Select()
    '            Else
    '                FirstCheckPoint()
    '            End If
    '        End If

    '    Catch ex As Exception

    '    End Try
    'End Sub
    'Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
    '    Try
    '        Setting = AppControl.GetSettings("C:\TML_INI\QualityControlCheckAppliation\") 'System.Windows.Forms.Application.StartupPath)
    '        Dim LoopStarted As Boolean = False
    '        Dim ErMsg As String = ""
    '        CurrentCheckPoint = New CheckSheetStep
    '        Dim StartIndex As Integer = 0
    '        If TextBox_Step.Text = "" Then
    '            StartIndex = AllowedSteps.Length
    '        Else
    '            StartIndex = Array.IndexOf(AllowedSteps, TextBox_Step.Text)
    '        End If
    '        For i As Integer = StartIndex - 1 To 0 Step -1
    '            LoopStarted = True
    '            RichTextBox_ActivityToCheck.Text = "Wait.."
    '            PanelSubForm.Controls.Clear()
    '            CurrentCheckPoint = YTA_CheckSheet.ProcessStepNo(StepNo:=Integer.Parse(AllowedSteps(i)), Initial:="46501497", CustOrd, ErrMsg:=ErMsg)
    '            If Not IsDBNull(CurrentCheckPoint) Then
    '                If Not IsNothing(CurrentCheckPoint) Then
    '                    If Not IsNothing(CurrentCheckPoint.ActivityToCheck) Then
    '                        RichTextBox_Step.Text = CurrentCheckPoint.ProcessNo '& vbCrLf & CheckPoint.ProcessStep
    '                        TextBox_Step.Text = CurrentCheckPoint.StepNo
    '                        RichTextBox_ActivityToCheck.Text = CurrentCheckPoint.ActivityToCheck
    '                        RichTextBox_AutoSize(RichTextBox_Step)
    '                        RichTextBox_AutoSize(RichTextBox_ActivityToCheck)
    '                        RichTextBox_ActivityToCheck.Refresh()
    '                        PanelSubForm.Controls.Clear()

    '                        If CurrentCheckPoint.Method = CheckSheetStep.MethodOption.ProcedureStep Then
    '                            Dim PSV As New ProcedureStepView
    '                            PSV.TopLevel = False
    '                            PSV.InputFeatures(CurrentCheckPoint.ProcessStep, CurrentCheckPoint.ActivityToCheck, CurrentCheckPoint.ProcedureStepAction.ImagePath_ProcedureStep)
    '                            PanelSubForm.Controls.Add(PSV)
    '                            PSV.AutoScroll = True
    '                            PSV.Dock = DockStyle.Fill
    '                            PanelSubForm.AutoScroll = True
    '                            PSV.PictureBox1.Select()
    '                            PSV.Show()
    '                            Me.Refresh()
    '                            Exit For
    '                        ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.UserIput Then
    '                            Dim SUI As SelectUserInput = New SelectUserInput
    '                            SUI.TopLevel = False
    '                            SUI.Message = CurrentCheckPoint.UserInputAction.UserActionMessage
    '                            SUI.inputValues = CurrentCheckPoint.UserInputAction.UserInputList
    '                            PanelSubForm.Controls.Add(SUI)
    '                            'SUI.BringToFront()
    '                            SUI.AutoScroll = True
    '                            SUI.Dock = DockStyle.Fill
    '                            PanelSubForm.AutoScroll = True
    '                            SUI.ListView1.Select()
    '                            SUI.Show()
    '                            Me.Refresh()
    '                            Exit For
    '                        ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.SinglePntInst Then
    '                            Dim SPI As SinglePointInstruction = New SinglePointInstruction
    '                            SPI.TopLevel = False
    '                            SPI.InputFeatures(CurrentCheckPoint.SinglePointAction.ImagePath_SPI_1, CurrentCheckPoint.SinglePointAction.ImagePath_SPI_2, CurrentCheckPoint.SinglePointAction.SPI_Message)
    '                            PanelSubForm.Controls.Add(SPI)
    '                            SPI.AutoScroll = True
    '                            SPI.Dock = DockStyle.Fill
    '                            PanelSubForm.AutoScroll = True
    '                            SPI.Show()
    '                            Me.Refresh()
    '                            Exit For
    '                        ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.AddedDocs Then
    '                            Dim VDOC As AddedDocs = New AddedDocs
    '                            VDOC.TopLevel = False
    '                            PanelSubForm.Controls.Add(VDOC)
    '                            VDOC.AutoScroll = True
    '                            VDOC.Dock = DockStyle.Fill
    '                            PanelSubForm.AutoScroll = True
    '                            VDOC.Show()
    '                            Me.Refresh()
    '                            Exit For
    '                        End If
    '                    End If
    '                End If
    '            End If
    '        Next

    '        If LoopStarted = False And Not (RichTextBox_Step.BackColor = Color.Yellow Or RichTextBox_Step.BackColor = Color.OrangeRed) Then
    '            Dim AllPass As Boolean = True
    '            For Each item In AllCheckResult
    '                If Not IsNothing(item) Then
    '                    If Not item.CheckResult Like "*True*" Then
    '                        AllPass = False
    '                        Exit For
    '                    End If
    '                Else
    '                    AllPass = False
    '                    Exit For
    '                End If
    '            Next
    '            If AllPass = True Then
    '                ShowCompleteInspection()
    '                'RichTextBox_ActivityToCheck.Text = ""
    '                'RichTextBox_Step.Text = ""
    '                'TextBox_Step.Text = ""
    '                'RichTextBox_Step.BackColor = Color.White
    '                'PanelSubForm.Controls.Clear()
    '                'Dim BEY As New CompleteInspection
    '                'BEY.TopLevel = False
    '                'PanelSubForm.Controls.Add(BEY)
    '                'BEY.AutoScroll = True
    '                'BEY.Dock = DockStyle.Fill
    '                'PanelSubForm.AutoScroll = True
    '                'BEY.Show()
    '                'BEY.Button1.Select()
    '            Else
    '                LastCheckPoint()
    '            End If
    '        End If

    '    Catch ex As Exception

    '    End Try
    'End Sub
    'Public Sub wait(ByVal interval As Integer)
    '    Dim sw As New Stopwatch
    '    Dim wait_time As Double
    '    Dim Elapsed_time As Double
    '    sw.Start()
    '    Do While sw.ElapsedMilliseconds < (interval * 1000)
    '        ' Allows UI to remain responsive
    '        Application.DoEvents()
    '        Elapsed_time = (sw.ElapsedMilliseconds + 1) / 1000
    '        wait_time = interval - Elapsed_time
    '    Loop
    '    sw.Stop()
    'End Sub
    'Private Sub RichTextBox_AutoSize(ByVal Rc As RichTextBox)
    '    Try

    '        'Center the text in the RichBox
    '        Rc.SelectAll()
    '        Rc.SelectionAlignment = HorizontalAlignment.Center
    '        'Is scrollbar visible?
    '        Dim bVScrollBar As Boolean
    '        bVScrollBar = ((GetWindowLong(Rc.Handle, GWL_STYLE) And WS_VSCROLL) = WS_VSCROLL)
    '        Select Case bVScrollBar
    '            Case True
    '                'Scrollbar is visible - Make it smaller
    '                Do
    '                    Rc.ZoomFactor = Rc.ZoomFactor - 0.01
    '                    bVScrollBar = ((GetWindowLong(Rc.Handle, GWL_STYLE) And WS_VSCROLL) = WS_VSCROLL)
    '                    'If the scrollbar is no longer visible we are done!
    '                    If bVScrollBar = False Then Exit Do
    '                Loop
    '            Case False
    '                'Scrollbar is not visible - Make it bigger
    '                Do
    '                    Rc.ZoomFactor = Rc.ZoomFactor + 0.01
    '                    bVScrollBar = ((GetWindowLong(Rc.Handle, GWL_STYLE) And WS_VSCROLL) = WS_VSCROLL)
    '                    If bVScrollBar = True Then
    '                        Do
    '                            'Found the scrollbar, make smaller until bar is not visible
    '                            Rc.ZoomFactor = Rc.ZoomFactor - 0.01
    '                            bVScrollBar = ((GetWindowLong(Rc.Handle, GWL_STYLE) And WS_VSCROLL) = WS_VSCROLL)
    '                            'If the scrollbar is no longer visible we are done!
    '                            If bVScrollBar = False Then Exit Do
    '                        Loop
    '                        Exit Do
    '                    End If
    '                Loop
    '        End Select

    '    Catch ex As Exception

    '    End Try
    'End Sub
    'Public Function InspectionStatus(ByVal InspectionResult As CheckSheetStep, ByVal CheckResult As Boolean) As String
    '    Try
    '        InspectionResult.CheckResult = CheckResult
    '        AllCheckResult(Array.IndexOf(AllowedSteps, InspectionResult.StepNo)) = InspectionResult
    '        SetInspectionColor(InspectionResult.StepNo)
    '    Catch ex As Exception

    '    End Try
    'End Function
    'Public Sub SetInspectionColor(ByVal StepNo As String, Optional ByVal ProcessNo As String = "")
    '    Try
    '        RichTextBox_Step.BackColor = Color.Yellow

    '        If Not IsNothing(AllCheckResult(Array.IndexOf(AllowedSteps, StepNo.ToString))) Then
    '            If AllCheckResult(Array.IndexOf(AllowedSteps, StepNo.ToString)).StepNo = StepNo _
    'And AllCheckResult(Array.IndexOf(AllowedSteps, StepNo.ToString)).CheckResult = True Then
    '                RichTextBox_Step.BackColor = Color.LightGreen
    '            ElseIf AllCheckResult(Array.IndexOf(AllowedSteps, StepNo.ToString)).StepNo = StepNo _
    '                    And AllCheckResult(Array.IndexOf(AllowedSteps, StepNo.ToString)).CheckResult = False Then
    '                RichTextBox_Step.BackColor = Color.OrangeRed
    '            Else
    '                RichTextBox_Step.BackColor = Color.Yellow
    '            End If
    '        Else
    '            If QcData.Count > 0 Then
    '                For Each Item In QcData
    '                    If Item.PROCESS_NO = Decimal.Parse(ProcessNo).ToString("N1") And Item.REMARK Like "*" & StepNo & "*" Then
    '                        If Item.REMARK Like "*GO*" Then
    '                            RichTextBox_Step.BackColor = Color.LightGreen
    '                        ElseIf Item.REMARK Like "*NG*" Then
    '                            RichTextBox_Step.BackColor = Color.OrangeRed
    '                        End If
    '                    End If
    '                Next
    '            End If
    '        End If
    '    Catch ex As Exception

    '    End Try
    'End Sub
    'Public Sub ShowCompleteInspection()
    '    Try
    '        RichTextBox_ActivityToCheck.Text = ""
    '        RichTextBox_Step.Text = ""
    '        TextBox_Step.Text = ""
    '        RichTextBox_Step.BackColor = Color.White
    '        PanelSubForm.Controls.Clear()
    '        Dim BEY As New CompleteInspection
    '        BEY.TopLevel = False
    '        PanelSubForm.Controls.Add(BEY)
    '        BEY.AutoScroll = True
    '        BEY.Dock = DockStyle.Fill
    '        PanelSubForm.AutoScroll = True
    '        BEY.Show()
    '        BEY.Button1.Select()
    '    Catch ex As Exception

    '    End Try
    'End Sub
    'Private Sub RichTextBox_Step_Click(sender As Object, e As EventArgs) Handles RichTextBox_Step.Click
    '    Try
    '        Dim ErMsg As String = ""
    '        If Not IsNumeric(RichTextBox_Step.Text) Then
    '            ErMsg = "Cannot load Inspection point. Click Next/Previous buttons."
    '            MsgBox(ErrMsg, MsgBoxStyle.Critical, "Error")
    '            Exit Sub
    '        End If
    '        CurrentCheckPoint = New CheckSheetStep
    '        RichTextBox_ActivityToCheck.Text = "Wait.."
    '        CurrentCheckPoint = YTA_CheckSheet.ProcessStepNo(StepNo:=Integer.Parse(RichTextBox_Step.Text), Initial:=Initial, CustOrd, ErrMsg:=ErMsg)
    '        If Not IsDBNull(CurrentCheckPoint) Then
    '            If Not IsNothing(CurrentCheckPoint) Then
    '                If Not IsNothing(CurrentCheckPoint.ActivityToCheck) Then
    '                    RichTextBox_Step.Text = CurrentCheckPoint.ProcessNo '& vbCrLf & CheckPoint.ProcessStep
    '                    TextBox_Step.Text = CurrentCheckPoint.StepNo
    '                    RichTextBox_ActivityToCheck.Text = CurrentCheckPoint.ActivityToCheck
    '                    RichTextBox_AutoSize(RichTextBox_Step)
    '                    RichTextBox_AutoSize(RichTextBox_ActivityToCheck)
    '                    RichTextBox_ActivityToCheck.Refresh()
    '                    PanelSubForm.Controls.Clear()

    '                    If CurrentCheckPoint.Method = CheckSheetStep.MethodOption.ProcedureStep Then
    '                        Dim PSV As New ProcedureStepView
    '                        PSV.TopLevel = False
    '                        PSV.InputFeatures(CurrentCheckPoint.ProcessStep, CurrentCheckPoint.ActivityToCheck, CurrentCheckPoint.ProcedureStepAction.ImagePath_ProcedureStep)
    '                        PanelSubForm.Controls.Add(PSV)
    '                        PSV.AutoScroll = True
    '                        PSV.Dock = DockStyle.Fill
    '                        PanelSubForm.AutoScroll = True
    '                        PSV.PictureBox1.Select()
    '                        PSV.Show()
    '                        Me.Refresh()

    '                    ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.UserIput Then
    '                        Dim SUI As SelectUserInput = New SelectUserInput
    '                        SUI.TopLevel = False
    '                        SUI.Message = CurrentCheckPoint.UserInputAction.UserActionMessage
    '                        SUI.inputValues = CurrentCheckPoint.UserInputAction.UserInputList
    '                        PanelSubForm.Controls.Add(SUI)
    '                        'SUI.BringToFront()
    '                        SUI.AutoScroll = True
    '                        SUI.Dock = DockStyle.Fill
    '                        PanelSubForm.AutoScroll = True
    '                        SUI.ListView1.Select()
    '                        SUI.Show()
    '                        Me.Refresh()

    '                    ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.SinglePntInst Then
    '                        Dim SPI As SinglePointInstruction = New SinglePointInstruction
    '                        SPI.TopLevel = False
    '                        SPI.InputFeatures(CurrentCheckPoint.SinglePointAction.ImagePath_SPI_1, CurrentCheckPoint.SinglePointAction.ImagePath_SPI_2, CurrentCheckPoint.SinglePointAction.SPI_Message)
    '                        PanelSubForm.Controls.Add(SPI)
    '                        SPI.AutoScroll = True
    '                        SPI.Dock = DockStyle.Fill
    '                        PanelSubForm.AutoScroll = True
    '                        SPI.Show()
    '                        Me.Refresh()
    '                    ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.DocumentCheck Then
    '                        Dim VDOC As ViewDocument = New ViewDocument
    '                        VDOC.TopLevel = False
    '                        VDOC.curNaviageUrl = CurrentCheckPoint.ViewDocAction.PdfPath_DocumentCheck
    '                        PanelSubForm.Controls.Add(VDOC)
    '                        VDOC.AutoScroll = True
    '                        VDOC.Dock = DockStyle.Fill
    '                        PanelSubForm.AutoScroll = True
    '                        VDOC.Show()
    '                        Me.Refresh()
    '                    End If
    '                End If
    '            End If
    '        End If




    '    Catch ex As Exception

    '    End Try
    'End Sub
    'Public Sub PrintQcc_Rev0()
    '    Try
    '        Dim ErrMsg As String = ""
    '        QcData = TmlEntityQA.GetDatabaseTableAs_List(Of POCO_QA.yta_qcc_v1p2)("INDEX_NO", CustOrd.INDEX_NO, "INDEX_NO", CustOrd.INDEX_NO, ErrMsg)
    '        If QcData.Count > 0 Then
    '            Dim BlankDoc As String = Setting.Var_06_DocsStore & "Production Release Documents\QC Check Sheets\" & CustOrd.PROD_NO & "\Line-" & CustOrd.LINE_NO & "-(Qty " & CustOrd.TOT_QTY & " Pcs)\" & CustOrd.INDEX_NO & "-QCSHEET.pdf"
    '            Dim FinalDoc As String = Setting.Var_06_DocsStore & "Production Complete Documents\Signed_QCC\" & CustOrd.PROD_NO & "\Line-" & CustOrd.LINE_NO & "\" & CustOrd.INDEX_NO & "-QCS-Signed.pdf"
    '            Dim UseDoc As New PdfSharp.Pdf.PdfDocument
    '            If System.IO.File.Exists(FinalDoc) Then
    '                Dim P_Doc = OpenPdfOperation.FileOp.GetDocument(BlankDoc, ErrMsg)
    '                For Each QcWrite In QcData
    '                    Dim WriteParams = QcWrite.CHECK_RESULT.Split("$")
    '                    For Each WriteParam In WriteParams
    '                        Dim WriteInput = WriteParam.Split("-")(0)
    '                        Dim WriteSXY = WriteParam.Split("-")(1)
    '                        Dim WriteS = WriteSXY.Split(",")(0)
    '                        Dim WriteX = WriteSXY.Split(",")(1)
    '                        Dim WriteY = WriteSXY.Split(",")(2)
    '                        Dim TextToWrite As String = WriteInput
    '                        Dim FontToWrite As OpenPdfOperation.FileOp.FontName = OpenPdfOperation.FileOp.FontName.Arial
    '                        If WriteInput = "Tick" Then
    '                            TextToWrite = "ü"
    '                            FontToWrite = OpenPdfOperation.FileOp.FontName.Wingdings
    '                        End If
    '                        Dim WP As New OpenPdfOperation.WriteTextParameters With {
    '                                    .Name = FontToWrite,
    '                                    .Colour = OpenPdfOperation.FileOp.FontColor.Black,
    '                                    .Style = OpenPdfOperation.FileOp.FontStyle.Regular,
    '                                    .TextSize = Integer.Parse(WriteS),
    '                                    .TextValue = TextToWrite,
    '                                    .X_Position = CDbl(WriteX / 100),
    '                                    .Y_Position = CDbl(WriteY / 100)}
    '                        Dim WPS As New List(Of OpenPdfOperation.WriteTextParameters)
    '                        WPS.Add(WP)
    '                        'OpenPdfOperation.FileOp.PDF_WriteText(FinalDoc, FinalDoc, WPS, ErrMsg)
    '                        OpenPdfOperation.FileOp.PDF_WriteText(P_Doc, WPS, ErrMsg)
    '                        If ErrMsg.Length > 0 Then
    '                            MsgBox(ErrMsg)
    '                            Exit Sub
    '                        End If
    '                    Next
    '                Next
    '                P_Doc.Save(FinalDoc)
    '                Process.Start(FinalDoc)
    '            ElseIf System.IO.File.Exists(BlankDoc) Then
    '                Dim P_Doc = OpenPdfOperation.FileOp.GetDocument(BlankDoc, ErrMsg)
    '                For Each QcWrite In QcData
    '                    Dim WriteParams = QcWrite.CHECK_RESULT.Split("$")
    '                    For Each WriteParam In WriteParams
    '                        Dim WriteInput = WriteParam.Split("-")(0)
    '                        Dim WriteSXY = WriteParam.Split("-")(1)
    '                        Dim WriteS = WriteSXY.Split(",")(0)
    '                        Dim WriteX = WriteSXY.Split(",")(1)
    '                        Dim WriteY = WriteSXY.Split(",")(2)
    '                        Dim TextToWrite As String = WriteInput
    '                        Dim FontToWrite As OpenPdfOperation.FileOp.FontName = OpenPdfOperation.FileOp.FontName.Arial
    '                        If WriteInput = "Tick" Then
    '                            TextToWrite = "ü"
    '                            FontToWrite = OpenPdfOperation.FileOp.FontName.Wingdings
    '                        ElseIf WriteInput = "Circle" Then
    '                            TextToWrite = "¡"
    '                            FontToWrite = OpenPdfOperation.FileOp.FontName.Wingdings
    '                        ElseIf WriteInput = "Cross" Then
    '                            TextToWrite = "û"
    '                            FontToWrite = OpenPdfOperation.FileOp.FontName.Wingdings
    '                        End If
    '                        Dim WP As New OpenPdfOperation.WriteTextParameters With {
    '                                    .Name = FontToWrite,
    '                                    .Colour = OpenPdfOperation.FileOp.FontColor.Blue,
    '                                    .Style = OpenPdfOperation.FileOp.FontStyle.Regular,
    '                                    .TextSize = Integer.Parse(WriteS),
    '                                    .TextValue = TextToWrite,
    '                                    .X_Position = CDbl(WriteX / 100),
    '                                    .Y_Position = CDbl(WriteY / 100)}
    '                        Dim WPS As New List(Of OpenPdfOperation.WriteTextParameters)
    '                        WPS.Add(WP)
    '                        System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(FinalDoc))
    '                        OpenPdfOperation.FileOp.PDF_WriteText(P_Doc, WPS, ErrMsg)
    '                        If ErrMsg.Length > 0 Then
    '                            MsgBox(ErrMsg)
    '                            Exit Sub
    '                        End If
    '                    Next
    '                Next
    '                P_Doc.Save(FinalDoc)
    '                Process.Start(FinalDoc)
    '            Else
    '                MsgBox("QC Template Not created in TML System. Please create it first.")
    '                Exit Sub
    '            End If



    '        End If

    '    Catch ex As Exception

    '    End Try
    'End Sub
    'Public Sub PrintQcc_Rev1()
    '    Try

    '        Dim ErrMsg As String = ""
    '        QcData = TmlEntityQA.GetDatabaseTableAs_List(Of POCO_QA.yta_qcc_v1p2)("INDEX_NO", CustOrd.INDEX_NO, "INDEX_NO", CustOrd.INDEX_NO, ErrMsg)
    '        If QcData.Count > 0 Then
    '            Dim BlankDoc As String = Setting.Var_06_DocsStore & "Production Release Documents\QC Check Sheets\" & CustOrd.PROD_NO & "\Line-" & CustOrd.LINE_NO & "-(Qty " & CustOrd.TOT_QTY & " Pcs)\" & CustOrd.INDEX_NO & "-QCSHEET.pdf"
    '            Dim FinalDoc As String = Setting.Var_06_DocsStore & "Production Complete Documents\Signed_QCC\" & CustOrd.PROD_NO & "\Line-" & CustOrd.LINE_NO & "\" & CustOrd.INDEX_NO & "-QCS-Signed.pdf"
    '            Dim UseDoc As New PdfSharp.Pdf.PdfDocument
    '            If System.IO.File.Exists(BlankDoc) Then
    '                Dim P_Doc = OpenPdfOperation.FileOp.GetDocument(BlankDoc, ErrMsg)
    '                For Each QcWrite In QcData
    '                    Dim WriteParams = QcWrite.CHECK_RESULT.Split("$")
    '                    For Each WriteParam In WriteParams
    '                        Dim WriteInput = WriteParam.Split("-")(0)
    '                        Dim WriteSXY = WriteParam.Split("-")(1)
    '                        Dim WriteS = WriteSXY.Split(",")(0)
    '                        Dim WriteX = WriteSXY.Split(",")(1)
    '                        Dim WriteY = WriteSXY.Split(",")(2)
    '                        Dim TextToWrite As String = WriteInput
    '                        Dim FontToWrite As OpenPdfOperation.FileOp.FontName = OpenPdfOperation.FileOp.FontName.Arial
    '                        If WriteInput = "Tick" Then
    '                            TextToWrite = "ü"
    '                            FontToWrite = OpenPdfOperation.FileOp.FontName.Wingdings
    '                        ElseIf WriteInput = "Circle" Then
    '                            TextToWrite = "¡"
    '                            FontToWrite = OpenPdfOperation.FileOp.FontName.Wingdings
    '                        ElseIf WriteInput = "Cross" Then
    '                            TextToWrite = "û"
    '                            FontToWrite = OpenPdfOperation.FileOp.FontName.Wingdings
    '                        End If
    '                        Dim WP As New OpenPdfOperation.WriteTextParameters With {
    '                                    .Name = FontToWrite,
    '                                    .Colour = OpenPdfOperation.FileOp.FontColor.Blue,
    '                                    .Style = OpenPdfOperation.FileOp.FontStyle.Regular,
    '                                    .TextSize = Integer.Parse(WriteS),
    '                                    .TextValue = TextToWrite,
    '                                    .X_Position = CDbl(WriteX / 100),
    '                                    .Y_Position = CDbl(WriteY / 100)}
    '                        Dim WPS As New List(Of OpenPdfOperation.WriteTextParameters)
    '                        WPS.Add(WP)
    '                        System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(FinalDoc))
    '                        OpenPdfOperation.FileOp.PDF_WriteText(P_Doc, WPS, ErrMsg)
    '                        If ErrMsg.Length > 0 Then
    '                            MsgBox(ErrMsg)
    '                            Exit Sub
    '                        End If
    '                    Next
    '                Next
    '                P_Doc.Save(FinalDoc)
    '                Process.Start(FinalDoc)
    '                'System.IO.File.Copy(FinalDoc, Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), True)
    '                'Process.Start(System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), System.IO.Path.GetFileName(FinalDoc)))
    '            Else
    '                MsgBox("QC Template Not created in TML System. Please create it first.")
    '                Exit Sub
    '            End If



    '        End If

    '    Catch ex As Exception

    '    End Try
    'End Sub
#End Region
#Region "NEW VERSION"
    Public Initial As String
    Public WorkerName As String
    Public WorkerID As String
    Public VersionText As String
    Public Shared CurrentCheckPoint As CheckSheetStep
    Dim ErrMsg As String = ""
    Public CustOrd As POCO_YGSP.cust_ord
    Public Shared Hipot As POCO_YGSP.hipot_tb
    Public YTA_Crc As POCO_YGSP.yta710_inspection_tb
    Public QcData As List(Of POCO_QA.yta_qcc_v1p2)
    Public QcSteps As List(Of POCO_QA.yta_qcc_steps)
    Public DataPlateCorrect As String
    Public DataPlateCheck(3) As String
    Public Shared AllowedSteps As String()
    Public Shared AllCheckResult As CheckSheetStep()
    Dim TmlEntityQA As MFG_ENTITY.Op
    Dim WMsg As New WarningForm

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        VersionText = AppControl.GetVersion("C:\TML_INI\QualityControlCheckAppliation\")
        Me.Text = "YTA QC Check" & " [ Ver:" & VersionText & "]" & " [ Station:" & My.Settings.Station & "]"

        RefreshSettings(Link.Network)
        TmlEntityQA = New MFG_ENTITY.Op(Setting.Var_04_MySql_QA)
        QcSteps = TmlEntityQA.GetDatabaseTableAs_List(Of POCO_QA.yta_qcc_steps)("PRODUCT", "YTA", "QCC_VER", "1.2")

        ContextMenuStation.Enabled = False

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
            If My.Settings.Station = "Not Set" Then
Repeat:
                Dim StationName = InputBox("Please select Station Name", "STATION", Setting.Var_08_StepsName)
                If Setting.Var_08_StepsName Like "*" & StationName & "*" Then
                    My.Settings.Station = StationName
                    My.Settings.Save()
                    Me.Text = "QC Check" & " [ Ver:" & VersionText & "]" & " [ Station:" & StationName & "]"
                Else
                    GoTo Repeat
                End If
            End If

            Setting.Var_08_StepsCurrent = String.Join(",", QcSteps.OrderBy(Function(x) x.SLNO).Where(Function(x) x.STATION = My.Settings.Station).Select(Of String)(Function(x) x.STEP_NO).ToArray)
            AllowedSteps = Setting.Var_08_StepsCurrent.Split(",")
            ReDim AllCheckResult(AllowedSteps.Length - 1)
            Dim IndexPos = Array.IndexOf(Setting.Var_08_StepsName.Split(","), My.Settings.Station)
            If IndexPos > 0 Then
                Setting.Var_08_StepsPrevous = String.Join(",", QcSteps.OrderBy(Function(x) x.SLNO).Where(Function(x) x.STATION = Setting.Var_08_StepsName.Split(",")(IndexPos - 1)).Select(Of String)(Function(x) x.STEP_NO).ToArray)
            Else
                Setting.Var_08_StepsPrevous = "" 'Setting.Var_08_StepsCurrent
            End If

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
            MsgBox("LoadIndexScan() Error:" & ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Public Sub FirstCheckPoint()
        Try
            Dim ErMsg As String = ""
            CurrentCheckPoint = New CheckSheetStep
            Dim StepNumbers = Setting.Var_08_StepsCurrent.Split(",")

            RichTextBox_ActivityToCheck.Text = "Wait.."
            CurrentCheckPoint = YTA_CheckSheet_v1m2s.ProcessStepNo(StepNo:=StepNumbers(0), Initial:=Initial, CustOrd, ErrMsg:=ErMsg)
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
                        ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.MakeUsrInpt Then
                            Dim SUI As MakeUserInput = New MakeUserInput
                            SUI.TopLevel = False
                            SUI.Message = CurrentCheckPoint.MakeUserInputAction.UserActionMessage
                            SUI.InputValueOld = CurrentCheckPoint.MakeUserInputAction.UserInputOld
                            SUI.InputValueNew = ""
                            PanelSubForm.Controls.Add(SUI)
                            'SUI.BringToFront()
                            SUI.AutoScroll = True
                            SUI.Dock = DockStyle.Fill
                            PanelSubForm.AutoScroll = True
                            SUI.InputTextBox.Select()
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
                        ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.AddedDocs Then
                            Dim VDOC As AddedDocs = New AddedDocs
                            VDOC.TopLevel = False
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
            Dim StepNumbers = Setting.Var_08_StepsCurrent.Split(",")

            RichTextBox_ActivityToCheck.Text = "Wait.."
            CurrentCheckPoint = YTA_CheckSheet_v1m2s.ProcessStepNo(StepNo:=StepNumbers(StepNumbers.Length - 1), Initial:=Initial, CustOrd, ErrMsg:=ErMsg)
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
                        ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.AddedDocs Then
                            Dim VDOC As AddedDocs = New AddedDocs
                            VDOC.TopLevel = False
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

    'Next (Rightside) Button
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            'Setting = AppControl.GetSettings("C:\TML_INI\QualityControlCheckAppliation\") 'System.Windows.Forms.Application.StartupPath)
            RefreshSettings(Link.Network)
            Dim LoopStarted As Boolean = False
            Dim ErMsg As String = ""
            CurrentCheckPoint = New CheckSheetStep
            If AllowedSteps(AllowedSteps.Length - 1) = TextBox_Step.Text Then
                RichTextBox_ActivityToCheck.Text = "No more Inspection Point forward.."
                GoTo LoopFinished
            End If
            For i As Integer = Array.IndexOf(AllowedSteps, TextBox_Step.Text) + 1 To AllowedSteps.Length - 1 Step 1
                LoopStarted = True
                RichTextBox_ActivityToCheck.Text = "Wait.."
                PanelSubForm.Controls.Clear()
                CurrentCheckPoint = YTA_CheckSheet_v1m2s.ProcessStepNo(StepNo:=AllowedSteps(i), Initial:=Initial, CustOrd, ErrMsg:=ErMsg)
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
                            ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.MakeUsrInpt Then
                                Dim SUI As MakeUserInput = New MakeUserInput
                                SUI.TopLevel = False
                                SUI.Message = CurrentCheckPoint.MakeUserInputAction.UserActionMessage
                                SUI.InputValueOld = CurrentCheckPoint.MakeUserInputAction.UserInputOld
                                SUI.InputValueNew = ""
                                PanelSubForm.Controls.Add(SUI)
                                'SUI.BringToFront()
                                SUI.AutoScroll = True
                                SUI.Dock = DockStyle.Fill
                                PanelSubForm.AutoScroll = True
                                SUI.InputTextBox.Select()
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
                            ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.AddedDocs Then
                                Dim VDOC As AddedDocs = New AddedDocs
                                VDOC.TopLevel = False
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

LoopFinished:
            If LoopStarted = False And Not (RichTextBox_Step.BackColor = Color.Yellow Or RichTextBox_Step.BackColor = Color.OrangeRed) Then
                Dim AllPass As Boolean = True
                For Each item In AllCheckResult
                    If Not IsNothing(item) Then
                        If Not item.CheckResult Like "*True*" Then
                            AllPass = False
                            Exit For
                        End If
                    Else
                        AllPass = False
                        Exit For
                    End If
                Next
                If AllPass = True Then
                    ShowCompleteInspection()
                    'RichTextBox_ActivityToCheck.Text = ""
                    'RichTextBox_Step.Text = ""
                    'TextBox_Step.Text = ""
                    'RichTextBox_Step.BackColor = Color.White
                    'PanelSubForm.Controls.Clear()
                    'Dim BEY As New CompleteInspection
                    'BEY.TopLevel = False
                    'PanelSubForm.Controls.Add(BEY)
                    'BEY.AutoScroll = True
                    'BEY.Dock = DockStyle.Fill
                    'PanelSubForm.AutoScroll = True
                    'BEY.Show()
                    'BEY.Button1.Select()
                Else
                    FirstCheckPoint()
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub

    'Previous (Leftside) Button
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            'Setting = AppControl.GetSettings("C:\TML_INI\QualityControlCheckAppliation\") 'System.Windows.Forms.Application.StartupPath)
            RefreshSettings(Link.Network)
            Dim LoopStarted As Boolean = False
            Dim ErMsg As String = ""
            CurrentCheckPoint = New CheckSheetStep
            Dim StartIndex As Integer = 0
            If TextBox_Step.Text = "" Then
                StartIndex = AllowedSteps.Length
            Else
                StartIndex = Array.IndexOf(AllowedSteps, TextBox_Step.Text)
            End If
            'If Array.IndexOf(AllowedSteps, TextBox_Step.Text) = 0 Then
            '    RichTextBox_ActivityToCheck.Text = "No more Inspection Point backward.."
            '    Exit Sub
            'End If
            For i As Integer = StartIndex - 1 To 0 Step -1
                LoopStarted = True
                RichTextBox_ActivityToCheck.Text = "Wait.."
                PanelSubForm.Controls.Clear()
                CurrentCheckPoint = YTA_CheckSheet_v1m2s.ProcessStepNo(StepNo:=AllowedSteps(i), Initial:="46501497", CustOrd, ErrMsg:=ErMsg)
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
                            ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.MakeUsrInpt Then
                                Dim SUI As MakeUserInput = New MakeUserInput
                                SUI.TopLevel = False
                                SUI.Message = CurrentCheckPoint.MakeUserInputAction.UserActionMessage
                                SUI.InputValueOld = CurrentCheckPoint.MakeUserInputAction.UserInputOld
                                SUI.InputValueNew = ""
                                PanelSubForm.Controls.Add(SUI)
                                'SUI.BringToFront()
                                SUI.AutoScroll = True
                                SUI.Dock = DockStyle.Fill
                                PanelSubForm.AutoScroll = True
                                SUI.InputTextBox.Select()
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
                            ElseIf CurrentCheckPoint.Method = CheckSheetStep.MethodOption.AddedDocs Then
                                Dim VDOC As AddedDocs = New AddedDocs
                                VDOC.TopLevel = False
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

            If LoopStarted = False And Not (RichTextBox_Step.BackColor = Color.Yellow Or RichTextBox_Step.BackColor = Color.OrangeRed) Then
                Dim AllPass As Boolean = True
                For Each item In AllCheckResult
                    If Not IsNothing(item) Then
                        If Not item.CheckResult Like "*True*" Then
                            AllPass = False
                            Exit For
                        End If
                    Else
                        AllPass = False
                        Exit For
                    End If
                Next
                If AllPass = True Then
                    ShowCompleteInspection()
                    'RichTextBox_ActivityToCheck.Text = ""
                    'RichTextBox_Step.Text = ""
                    'TextBox_Step.Text = ""
                    'RichTextBox_Step.BackColor = Color.White
                    'PanelSubForm.Controls.Clear()
                    'Dim BEY As New CompleteInspection
                    'BEY.TopLevel = False
                    'PanelSubForm.Controls.Add(BEY)
                    'BEY.AutoScroll = True
                    'BEY.Dock = DockStyle.Fill
                    'PanelSubForm.AutoScroll = True
                    'BEY.Show()
                    'BEY.Button1.Select()
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
    Public Sub SetInspectionColor(ByVal StepNo As String, Optional ByVal ProcessNo As String = "")
        Try
            RichTextBox_Step.BackColor = Color.Yellow

            If Not IsNothing(AllCheckResult(Array.IndexOf(AllowedSteps, StepNo.ToString))) Then
                If AllCheckResult(Array.IndexOf(AllowedSteps, StepNo.ToString)).StepNo = StepNo _
    And AllCheckResult(Array.IndexOf(AllowedSteps, StepNo.ToString)).CheckResult = True Then
                    RichTextBox_Step.BackColor = Color.LightGreen
                ElseIf AllCheckResult(Array.IndexOf(AllowedSteps, StepNo.ToString)).StepNo = StepNo _
                        And AllCheckResult(Array.IndexOf(AllowedSteps, StepNo.ToString)).CheckResult = False Then
                    RichTextBox_Step.BackColor = Color.OrangeRed
                Else
                    RichTextBox_Step.BackColor = Color.Yellow
                End If
            Else
                If QcData.Count > 0 Then
                    For Each Item In QcData
                        If Item.PROCESS_NO = ProcessNo And Item.REMARK Like "*" & StepNo & "*" Then
                            If Item.REMARK Like "*GO*" Then
                                RichTextBox_Step.BackColor = Color.LightGreen
                            ElseIf Item.REMARK Like "*NG*" Then
                                RichTextBox_Step.BackColor = Color.OrangeRed
                            End If
                        End If
                    Next
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub
    Public Sub ShowCompleteInspection()
        Try
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
    Public Sub PrintQcc_Rev0()
        Try
            Dim ErrMsg As String = ""
            QcData = TmlEntityQA.GetDatabaseTableAs_List(Of POCO_QA.yta_qcc_v1p2)("INDEX_NO", CustOrd.INDEX_NO, "INDEX_NO", CustOrd.INDEX_NO, ErrMsg)
            If QcData.Count > 0 Then
                Dim BlankDoc As String = Setting.Var_06_DocsStore & "Production Release Documents\QC Check Sheets\" & CustOrd.PROD_NO & "\Line-" & CustOrd.LINE_NO & "-(Qty " & CustOrd.TOT_QTY & " Pcs)\" & CustOrd.INDEX_NO & "-QCSHEET.pdf"
                Dim FinalDoc As String = Setting.Var_06_DocsStore & "Production Complete Documents\Signed_QCC\" & CustOrd.PROD_NO & "\Line-" & CustOrd.LINE_NO & "\" & CustOrd.INDEX_NO & "-QCS-Signed.pdf"
                Dim UseDoc As New PdfSharp.Pdf.PdfDocument
                If System.IO.File.Exists(FinalDoc) Then
                    Dim P_Doc = OpenPdfOperation.FileOp.GetDocument(BlankDoc, ErrMsg)
                    For Each QcWrite In QcData
                        Dim WriteParams = QcWrite.CHECK_RESULT.Split("$")
                        For Each WriteParam In WriteParams
                            Dim WriteInput = WriteParam.Split("-")(0)
                            Dim WriteSXY = WriteParam.Split("-")(1)
                            Dim WriteS = WriteSXY.Split(",")(0)
                            Dim WriteX = WriteSXY.Split(",")(1)
                            Dim WriteY = WriteSXY.Split(",")(2)
                            Dim TextToWrite As String = WriteInput
                            Dim FontToWrite As OpenPdfOperation.FileOp.FontName = OpenPdfOperation.FileOp.FontName.Arial
                            If WriteInput = "Tick" Then
                                TextToWrite = "ü"
                                FontToWrite = OpenPdfOperation.FileOp.FontName.Wingdings
                            End If
                            Dim WP As New OpenPdfOperation.WriteTextParameters With {
                                        .Name = FontToWrite,
                                        .Colour = OpenPdfOperation.FileOp.FontColor.Black,
                                        .Style = OpenPdfOperation.FileOp.FontStyle.Regular,
                                        .TextSize = Integer.Parse(WriteS),
                                        .TextValue = TextToWrite,
                                        .X_Position = CDbl(WriteX / 100),
                                        .Y_Position = CDbl(WriteY / 100)}
                            Dim WPS As New List(Of OpenPdfOperation.WriteTextParameters)
                            WPS.Add(WP)
                            'OpenPdfOperation.FileOp.PDF_WriteText(FinalDoc, FinalDoc, WPS, ErrMsg)
                            OpenPdfOperation.FileOp.PDF_WriteText(P_Doc, WPS, ErrMsg)
                            If ErrMsg.Length > 0 Then
                                MsgBox(ErrMsg)
                                Exit Sub
                            End If
                        Next
                    Next
                    P_Doc.Save(FinalDoc)
                    Process.Start(FinalDoc)
                ElseIf System.IO.File.Exists(BlankDoc) Then
                    Dim P_Doc = OpenPdfOperation.FileOp.GetDocument(BlankDoc, ErrMsg)
                    For Each QcWrite In QcData
                        Dim WriteParams = QcWrite.CHECK_RESULT.Split("$")
                        For Each WriteParam In WriteParams
                            Dim WriteInput = WriteParam.Split("-")(0)
                            Dim WriteSXY = WriteParam.Split("-")(1)
                            Dim WriteS = WriteSXY.Split(",")(0)
                            Dim WriteX = WriteSXY.Split(",")(1)
                            Dim WriteY = WriteSXY.Split(",")(2)
                            Dim TextToWrite As String = WriteInput
                            Dim FontToWrite As OpenPdfOperation.FileOp.FontName = OpenPdfOperation.FileOp.FontName.Arial
                            If WriteInput = "Tick" Then
                                TextToWrite = "ü"
                                FontToWrite = OpenPdfOperation.FileOp.FontName.Wingdings
                            ElseIf WriteInput = "Circle" Then
                                TextToWrite = "¡"
                                FontToWrite = OpenPdfOperation.FileOp.FontName.Wingdings
                            ElseIf WriteInput = "Cross" Then
                                TextToWrite = "û"
                                FontToWrite = OpenPdfOperation.FileOp.FontName.Wingdings
                            End If
                            Dim WP As New OpenPdfOperation.WriteTextParameters With {
                                        .Name = FontToWrite,
                                        .Colour = OpenPdfOperation.FileOp.FontColor.Blue,
                                        .Style = OpenPdfOperation.FileOp.FontStyle.Regular,
                                        .TextSize = Integer.Parse(WriteS),
                                        .TextValue = TextToWrite,
                                        .X_Position = CDbl(WriteX / 100),
                                        .Y_Position = CDbl(WriteY / 100)}
                            Dim WPS As New List(Of OpenPdfOperation.WriteTextParameters)
                            WPS.Add(WP)
                            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(FinalDoc))
                            OpenPdfOperation.FileOp.PDF_WriteText(P_Doc, WPS, ErrMsg)
                            If ErrMsg.Length > 0 Then
                                MsgBox(ErrMsg)
                                Exit Sub
                            End If
                        Next
                    Next
                    P_Doc.Save(FinalDoc)
                    Process.Start(FinalDoc)
                Else
                    MsgBox("QC Template Not created in TML System. Please create it first.")
                    Exit Sub
                End If



            End If

        Catch ex As Exception

        End Try
    End Sub
    Public Sub PrintQcc_Rev1()
        Try

            Dim ErrMsg As String = ""
            QcData = TmlEntityQA.GetDatabaseTableAs_List(Of POCO_QA.yta_qcc_v1p2)("INDEX_NO", CustOrd.INDEX_NO, "INDEX_NO", CustOrd.INDEX_NO, ErrMsg)
            If QcData.Count > 0 Then
                Dim BlankDoc As String = Setting.Var_06_DocsStore & "\Production Release Documents\QC Check Sheets\" & CustOrd.PROD_NO & "\Line-" & CustOrd.LINE_NO & "-(Qty " & CustOrd.TOT_QTY & " Pcs)\" & CustOrd.INDEX_NO & "-QCSHEET.pdf"
                Dim FinalDoc As String = Setting.Var_06_DocsStore & "\Production Complete Documents\Signed_QCC\" & CustOrd.PROD_NO & "\Line-" & CustOrd.LINE_NO & "\" & CustOrd.INDEX_NO & "-QCS-Signed.pdf"
                Dim UseDoc As New PdfSharp.Pdf.PdfDocument
                If System.IO.File.Exists(BlankDoc) Then
                    Dim P_Doc = OpenPdfOperation.FileOp.GetDocument(BlankDoc, ErrMsg)
                    For Each QcWrite In QcData
                        If QcWrite.INDEX_NO.ToString.Length > 0 Then
                            Dim WriteParams = QcWrite.CHECK_RESULT.Split("$")
                            For Each WriteParam In WriteParams
                                Dim WriteInput = WriteParam.Split("-")(0)
                                Dim WriteSXY = WriteParam.Split("-")(1)
                                Dim WriteS = WriteSXY.Split(",")(0)
                                Dim WriteX = WriteSXY.Split(",")(1)
                                Dim WriteY = WriteSXY.Split(",")(2)
                                Dim TextToWrite As String = WriteInput
                                Dim FontToWrite As OpenPdfOperation.FileOp.FontName = OpenPdfOperation.FileOp.FontName.Arial
                                If WriteInput = "Tick" Then
                                    TextToWrite = "ü"
                                    FontToWrite = OpenPdfOperation.FileOp.FontName.Wingdings
                                ElseIf WriteInput = "Circle" Then
                                    TextToWrite = "¡"
                                    FontToWrite = OpenPdfOperation.FileOp.FontName.Wingdings
                                ElseIf WriteInput = "Cross" Then
                                    TextToWrite = "û"
                                    FontToWrite = OpenPdfOperation.FileOp.FontName.Wingdings
                                End If
                                Dim WP As New OpenPdfOperation.WriteTextParameters With {
                                            .Name = FontToWrite,
                                            .Colour = OpenPdfOperation.FileOp.FontColor.Blue,
                                            .Style = OpenPdfOperation.FileOp.FontStyle.Regular,
                                            .TextSize = Integer.Parse(WriteS),
                                            .TextValue = TextToWrite,
                                            .X_Position = CDbl(WriteX / 100),
                                            .Y_Position = CDbl(WriteY / 100)}
                                Dim WPS As New List(Of OpenPdfOperation.WriteTextParameters)
                                WPS.Add(WP)
                                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(FinalDoc))
                                OpenPdfOperation.FileOp.PDF_WriteText(P_Doc, WPS, ErrMsg)
                                If ErrMsg.Length > 0 Then
                                    MsgBox(ErrMsg)
                                    Exit Sub
                                End If
                            Next
                        Else
                            MsgBox("Hi")
                        End If
                    Next
                    P_Doc.Save(FinalDoc)
                    'Process.Start(FinalDoc)
                    '"QCS-Signed"
                    For Each fname In System.IO.Directory.GetFiles(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments))
                        If fname Like "*QCS-Signed*" Then
                            Try
                                System.IO.File.Delete(fname)
                            Catch ex As Exception
                                Continue For
                            End Try
                        End If
                    Next
                    Dim FinalFileName As String = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), System.IO.Path.GetFileName(FinalDoc))
                    For i As Integer = 1 To 1
                        Try
                            System.IO.File.Copy(FinalDoc, FinalFileName, True)
                        Catch ex As Exception
                            Continue For
                        End Try
                    Next
                    Process.Start(FinalFileName)
                Else
                    MsgBox("QC Template Not created in TML System. Please create it first.")
                    Exit Sub
                End If



            End If

        Catch ex As Exception

        End Try
    End Sub

    'Private Sub RefreshSettings()
    '    Try
    '        Setting = AppControl.GetSettings("C:\TML_INI\QualityControlCheckAppliation\") 'System.Windows.Forms.Application.StartupPath)
    '    Catch ex As Exception
    '        'MsgBox("Settings files read error. Error: " & ex.Message)
    '        WMsg.Message = "Settings files read error. Error: " & ex.Message
    '        WMsg.ShowDialog()
    '    End Try
    'End Sub
    Private Sub RefreshSettings(ByVal Network As String)
        Try
            'Setting = AppControl.GetSettingsLive("C:\TML_INI\QualityControlCheckAppliation_EJA\")
            If Network = "FACTORY" Then
                Setting = AppControl.GetSettingsLive(Application.StartupPath & "\00_Settings")
            ElseIf Network = "DOMAIN" Then
                Setting = AppControl.GetSettingsDomain(Application.StartupPath & "\00_Settings")
            ElseIf Network = "LOCAL" Then
                Setting = AppControl.GetSettingsLocal(Application.StartupPath & "\00_Settings")
            Else
                WMsg.Message = "Settings files read error. Error: " & " Please select a Suitable NETWORK Type!"
                WMsg.ShowDialog()
            End If
            Setting.Var_01_ServerIP = Link.ServerAddress 'This is necessary for the application to work at YUAE and YKS
            Setting.Var_02_ServerPort = Link.PortNo 'This is necessary for the application to work at YUAE and YKS
            Setting.Var_03_MySql_YGSP = Link.Mysql_YGSP_ConStr 'This is necessary for the application to work at YUAE and YKS
            Setting.Var_04_MySql_QA = Link.Mysql_QA_ConStr 'This is necessary for the application to work at YUAE and YKS
            Setting.Var_05_Factory = Link.Factory 'This is necessary for the application to work at YUAE and YKS
            Setting.Var_06_DocsStore = Link.DocsStore 'This is necessary for the application to work at YUAE and YKS

        Catch ex As Exception
            'MsgBox("Settings files read error. Error: " & ex.Message)
            WMsg.Message = "Settings files read error. Error: " & ex.Message
            WMsg.ShowDialog()
        End Try
    End Sub


#End Region

#Region "Menu Strip"
    Private Sub SelectStationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectStationToolStripMenuItem.Click
        Try
            Dim Version = AppControl.GetVersion("C:\TML_INI\QualityControlCheckAppliation\")
            Me.Text = "QC Check" & " [ Ver:" & Version & "]"
Repeat:
            Dim StationName = InputBox("Please select Station Name", "STATION", Setting.Var_08_StepsName)
            If Setting.Var_08_StepsName Like "*" & StationName & "*" And Not StationName.Contains(",") And StationName.Length > 0 Then
                My.Settings.Station = StationName
                My.Settings.Save()
                Me.Text = Me.Text & " [ Station:" & StationName & "]"
            Else
                GoTo Repeat
            End If

            RefreshSettings(Link.Network)

            Setting.Var_08_StepsCurrent = String.Join(",", QcSteps.OrderBy(Function(x) x.SLNO).Where(Function(x) x.STATION = My.Settings.Station).Select(Of String)(Function(x) x.STEP_NO).ToArray)
            AllowedSteps = Setting.Var_08_StepsCurrent.Split(",")
            ReDim AllCheckResult(AllowedSteps.Length - 1)
            Dim IndexPos = Array.IndexOf(Setting.Var_08_StepsName.Split(","), My.Settings.Station)
            If IndexPos > 0 Then
                Setting.Var_08_StepsPrevous = String.Join(",", QcSteps.OrderBy(Function(x) x.SLNO).Where(Function(x) x.STATION = Setting.Var_08_StepsName.Split(",")(IndexPos - 1)).Select(Of String)(Function(x) x.STEP_NO).ToArray)
            Else
                Setting.Var_08_StepsPrevous = "" 'Setting.Var_08_StepsCurrent
            End If

            LoadIndexScan()

        Catch ex As Exception
            'MsgBox("Station Set Error:" & ex.Message, MsgBoxStyle.Critical, "Error")
            WMsg.Message = "Station Set Error:" & ex.Message
            WMsg.ShowDialog()
        End Try
    End Sub
    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        RefreshSettings(Link.Network)
    End Sub
    Private Sub RestartToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RestartToolStripMenuItem.Click
        Try
            VersionText = AppControl.GetVersion("C:\TML_INI\QualityControlCheckAppliation\")
            Me.Text = "YTA QC Check" & " [ Ver:" & VersionText & "]" & " [ Station:" & My.Settings.Station & "]"

            RefreshSettings(Link.Network)

            Setting.Var_08_StepsCurrent = String.Join(",", QcSteps.OrderBy(Function(x) x.SLNO).Where(Function(x) x.STATION = My.Settings.Station).Select(Of String)(Function(x) x.STEP_NO).ToArray)
            AllowedSteps = Setting.Var_08_StepsCurrent.Split(",")
            ReDim AllCheckResult(AllowedSteps.Length - 1)
            Dim IndexPos = Array.IndexOf(Setting.Var_08_StepsName.Split(","), My.Settings.Station)
            If IndexPos > 0 Then
                Setting.Var_08_StepsPrevous = String.Join(",", QcSteps.OrderBy(Function(x) x.SLNO).Where(Function(x) x.STATION = Setting.Var_08_StepsName.Split(",")(IndexPos - 1)).Select(Of String)(Function(x) x.STEP_NO).ToArray)
            Else
                Setting.Var_08_StepsPrevous = "" 'Setting.Var_08_StepsCurrent
            End If

            LoadIndexScan()

        Catch ex As Exception
            'MsgBox("Station Set Error:" & ex.Message, MsgBoxStyle.Critical, "Error")
            WMsg.Message = "Station Set Error:" & ex.Message
            WMsg.ShowDialog()
        End Try
    End Sub
#End Region
End Class
