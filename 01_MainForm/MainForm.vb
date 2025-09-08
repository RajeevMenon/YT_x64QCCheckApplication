
Imports System.IO
Imports Newtonsoft.Json
Imports Spire.Pdf.Exporting

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
    '    Dim Version = AppControl.GetVersion("C:\TML_INI\YT_x64QCCheckAppliation\")
    '    Me.Text = Me.Text & " [ Ver:" & Version & "]"
    '    Setting = AppControl.GetSettings("C:\TML_INI\YT_x64QCCheckAppliation\") 'System.Windows.Forms.Application.StartupPath)
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
    '        Setting = AppControl.GetSettings("C:\TML_INI\YT_x64QCCheckAppliation\") 'System.Windows.Forms.Application.StartupPath)
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
    '        Setting = AppControl.GetSettings("C:\TML_INI\YT_x64QCCheckAppliation\") 'System.Windows.Forms.Application.StartupPath)
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
    '                Dim P_Doc = OpenPdfOperation_x64.FileOp.GetDocument(BlankDoc, ErrMsg)
    '                For Each QcWrite In QcData
    '                    Dim WriteParams = QcWrite.CHECK_RESULT.Split("$")
    '                    For Each WriteParam In WriteParams
    '                        Dim WriteInput = WriteParam.Split("-")(0)
    '                        Dim WriteSXY = WriteParam.Split("-")(1)
    '                        Dim WriteS = WriteSXY.Split(",")(0)
    '                        Dim WriteX = WriteSXY.Split(",")(1)
    '                        Dim WriteY = WriteSXY.Split(",")(2)
    '                        Dim TextToWrite As String = WriteInput
    '                        Dim FontToWrite As OpenPdfOperation_x64.FileOp.FontName = OpenPdfOperation_x64.FileOp.FontName.Arial
    '                        If WriteInput = "Tick" Then
    '                            TextToWrite = "ü"
    '                            FontToWrite = OpenPdfOperation_x64.FileOp.FontName.Wingdings
    '                        End If
    '                        Dim WP As New OpenPdfOperation_x64.WriteTextParameters With {
    '                                    .Name = FontToWrite,
    '                                    .Colour = OpenPdfOperation_x64.FileOp.FontColor.Black,
    '                                    .Style = OpenPdfOperation_x64.FileOp.FontStyle.Regular,
    '                                    .TextSize = Integer.Parse(WriteS),
    '                                    .TextValue = TextToWrite,
    '                                    .X_Position = CDbl(WriteX / 100),
    '                                    .Y_Position = CDbl(WriteY / 100)}
    '                        Dim WPS As New List(Of OpenPdfOperation_x64.WriteTextParameters)
    '                        WPS.Add(WP)
    '                        'OpenPdfOperation_x64.FileOp.PDF_WriteText(FinalDoc, FinalDoc, WPS, ErrMsg)
    '                        OpenPdfOperation_x64.FileOp.PDF_WriteText(P_Doc, WPS, ErrMsg)
    '                        If ErrMsg.Length > 0 Then
    '                            MsgBox(ErrMsg)
    '                            Exit Sub
    '                        End If
    '                    Next
    '                Next
    '                P_Doc.Save(FinalDoc)
    '                Process.Start(FinalDoc)
    '            ElseIf System.IO.File.Exists(BlankDoc) Then
    '                Dim P_Doc = OpenPdfOperation_x64.FileOp.GetDocument(BlankDoc, ErrMsg)
    '                For Each QcWrite In QcData
    '                    Dim WriteParams = QcWrite.CHECK_RESULT.Split("$")
    '                    For Each WriteParam In WriteParams
    '                        Dim WriteInput = WriteParam.Split("-")(0)
    '                        Dim WriteSXY = WriteParam.Split("-")(1)
    '                        Dim WriteS = WriteSXY.Split(",")(0)
    '                        Dim WriteX = WriteSXY.Split(",")(1)
    '                        Dim WriteY = WriteSXY.Split(",")(2)
    '                        Dim TextToWrite As String = WriteInput
    '                        Dim FontToWrite As OpenPdfOperation_x64.FileOp.FontName = OpenPdfOperation_x64.FileOp.FontName.Arial
    '                        If WriteInput = "Tick" Then
    '                            TextToWrite = "ü"
    '                            FontToWrite = OpenPdfOperation_x64.FileOp.FontName.Wingdings
    '                        ElseIf WriteInput = "Circle" Then
    '                            TextToWrite = "¡"
    '                            FontToWrite = OpenPdfOperation_x64.FileOp.FontName.Wingdings
    '                        ElseIf WriteInput = "Cross" Then
    '                            TextToWrite = "û"
    '                            FontToWrite = OpenPdfOperation_x64.FileOp.FontName.Wingdings
    '                        End If
    '                        Dim WP As New OpenPdfOperation_x64.WriteTextParameters With {
    '                                    .Name = FontToWrite,
    '                                    .Colour = OpenPdfOperation_x64.FileOp.FontColor.Blue,
    '                                    .Style = OpenPdfOperation_x64.FileOp.FontStyle.Regular,
    '                                    .TextSize = Integer.Parse(WriteS),
    '                                    .TextValue = TextToWrite,
    '                                    .X_Position = CDbl(WriteX / 100),
    '                                    .Y_Position = CDbl(WriteY / 100)}
    '                        Dim WPS As New List(Of OpenPdfOperation_x64.WriteTextParameters)
    '                        WPS.Add(WP)
    '                        System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(FinalDoc))
    '                        OpenPdfOperation_x64.FileOp.PDF_WriteText(P_Doc, WPS, ErrMsg)
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
    '                Dim P_Doc = OpenPdfOperation_x64.FileOp.GetDocument(BlankDoc, ErrMsg)
    '                For Each QcWrite In QcData
    '                    Dim WriteParams = QcWrite.CHECK_RESULT.Split("$")
    '                    For Each WriteParam In WriteParams
    '                        Dim WriteInput = WriteParam.Split("-")(0)
    '                        Dim WriteSXY = WriteParam.Split("-")(1)
    '                        Dim WriteS = WriteSXY.Split(",")(0)
    '                        Dim WriteX = WriteSXY.Split(",")(1)
    '                        Dim WriteY = WriteSXY.Split(",")(2)
    '                        Dim TextToWrite As String = WriteInput
    '                        Dim FontToWrite As OpenPdfOperation_x64.FileOp.FontName = OpenPdfOperation_x64.FileOp.FontName.Arial
    '                        If WriteInput = "Tick" Then
    '                            TextToWrite = "ü"
    '                            FontToWrite = OpenPdfOperation_x64.FileOp.FontName.Wingdings
    '                        ElseIf WriteInput = "Circle" Then
    '                            TextToWrite = "¡"
    '                            FontToWrite = OpenPdfOperation_x64.FileOp.FontName.Wingdings
    '                        ElseIf WriteInput = "Cross" Then
    '                            TextToWrite = "û"
    '                            FontToWrite = OpenPdfOperation_x64.FileOp.FontName.Wingdings
    '                        End If
    '                        Dim WP As New OpenPdfOperation_x64.WriteTextParameters With {
    '                                    .Name = FontToWrite,
    '                                    .Colour = OpenPdfOperation_x64.FileOp.FontColor.Blue,
    '                                    .Style = OpenPdfOperation_x64.FileOp.FontStyle.Regular,
    '                                    .TextSize = Integer.Parse(WriteS),
    '                                    .TextValue = TextToWrite,
    '                                    .X_Position = CDbl(WriteX / 100),
    '                                    .Y_Position = CDbl(WriteY / 100)}
    '                        Dim WPS As New List(Of OpenPdfOperation_x64.WriteTextParameters)
    '                        WPS.Add(WP)
    '                        System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(FinalDoc))
    '                        OpenPdfOperation_x64.FileOp.PDF_WriteText(P_Doc, WPS, ErrMsg)
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
    Public StationName As String
    Public Initial As String
    Public WorkerName As String
    Public WorkerID As String
    Public LoginID As String
    Public VersionText As String
    Public Shared CurrentCheckPoint As CheckSheetStep
    Dim ErrMsg As String = ""
    Public CoHeader As POCO_YGSP.co_register
    Public CustOrd As POCO_YGSP.cust_ord
    Public Shared Hipot As POCO_YGSP.hipot_tb
    Public YTA_Crc As POCO_YGSP.yta710_inspection_tb


    Public DataPlateCorrect As String
    Public DataPlateCheck(3) As String
    Public Shared AllowedSteps As String()
    Public Shared AllCheckResult As CheckSheetStep()
    Dim TmlEntityQA As MFG_ENTITY.Op
    Dim TmlEntityYGS As New MFG_ENTITY.Op(Link.Mysql_YGSP_ConStr)
    Dim WMsg As New WarningForm

    'PdfReportTool
    Public ReportTemplate As New OpenPdfOperation_x64.Template
    Public JsonReportFile_STD As String = Application.StartupPath & "\05_Report_Templates\YMA-TML-SF-16-v1p4-STD.json"

    'QC Steps to Navigate
    Public QcSteps As List(Of POCO_QA.yta_qcc_steps)

    'Version control
    Public CurrentQCC_Version As String = "1.3"

    'Save FinalQcc to ProductionComplete Folder? Select True to Save.
    Dim SaveFinalDoc As Boolean = False

    'Version specific objects
    Dim CheckSheet_v1p3 As New YTA_CheckSheet_v1p3
    Dim CheckSheet_v1p4 As New YTA_CheckSheet_v1p4
    Public QcData_1p3 As List(Of POCO_QA.yta_qcc_v1p2)
    Public QcData_1p4 As List(Of POCO_QA.yta_qcc_v1p4)
    Public QcSteps_1p3_File As String = Application.StartupPath & "\04_QC_CheckSheet\QCC_Steps_v1p3.json"
    Public QcSteps_1p4_File As String = Application.StartupPath & "\04_QC_CheckSheet\QCC_Steps_v1p4.json"

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles Me.Load

        VersionText = AppControl.GetVersion(Application.StartupPath & "\00_Settings")
        Me.Text = $"YTA QC Check (x64)  [ App-Ver: { VersionText } ]  Station: [ {My.Settings.Station} ] QCC Ver: [ {CurrentQCC_Version} ] Plant: [ {Link.PlantID} ] Server: [ {Link.ServerAddress} ] QCC-Save: [ {SaveFinalDoc.ToString} ]"

        If My.Settings.Station.Length > 0 Then
            StationName = My.Settings.Station
        End If

        RefreshSettings(Link.Network)
        TmlEntityQA = New MFG_ENTITY.Op(Setting.Var_04_MySql_QA)
        If CurrentQCC_Version = "1.4" Then
            Dim QcStepsjson As String = System.IO.File.ReadAllText(QcSteps_1p4_File)
            QcSteps = JsonConvert.DeserializeObject(Of List(Of POCO_QA.yta_qcc_steps))(QcStepsjson)
        Else
            Dim QcStepsjson As String = System.IO.File.ReadAllText(QcSteps_1p3_File)
            QcSteps = JsonConvert.DeserializeObject(Of List(Of POCO_QA.yta_qcc_steps))(QcStepsjson)
        End If


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
                    Me.Text = $"YTA QC Check (x64)  [ App-Ver: { VersionText } ]  Station: [ {My.Settings.Station} ] QCC Ver: [ {CurrentQCC_Version} ] Plant: [ {Link.PlantID} ] Server: [ {Link.ServerAddress} ] QCC-Save: [ {SaveFinalDoc.ToString} ]"
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
            CurrentCheckPoint = YTA_CheckSheet_v1p3.ProcessStepNo(StepNo:=StepNumbers(0), Initial:=Initial, CustOrd, ErrMsg:=ErMsg)
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
                            PSV.InputFeatures(CurrentCheckPoint.ProcessStep, CurrentCheckPoint.ProcedureStepAction.ProcedureMessage, CurrentCheckPoint.ProcedureStepAction.ImagePath_ProcedureStep)
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
            CurrentCheckPoint = YTA_CheckSheet_v1p3.ProcessStepNo(StepNo:=StepNumbers(StepNumbers.Length - 1), Initial:=Initial, CustOrd, ErrMsg:=ErMsg)
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
                            PSV.InputFeatures(CurrentCheckPoint.ProcessStep, CurrentCheckPoint.ProcedureStepAction.ProcedureMessage, CurrentCheckPoint.ProcedureStepAction.ImagePath_ProcedureStep)
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
            'Setting = AppControl.GetSettings("C:\TML_INI\YT_x64QCCheckAppliation\") 'System.Windows.Forms.Application.StartupPath)
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
                CurrentCheckPoint = YTA_CheckSheet_v1p3.ProcessStepNo(StepNo:=AllowedSteps(i), Initial:=Initial, CustOrd, ErrMsg:=ErMsg)
                If Not IsDBNull(CurrentCheckPoint) Then
                    If Not IsNothing(CurrentCheckPoint) Then
                        If Not IsNothing(CurrentCheckPoint.ActivityToCheck) Then
                            RichTextBox_Step.Text = CurrentCheckPoint.ProcessNo '& vbCrLf & CheckPoint.ProcessStep
                            TextBox_Step.Text = CurrentCheckPoint.StepNo
                            ToolTip1.SetToolTip(RichTextBox_Step, TextBox_Step.Text)
                            RichTextBox_ActivityToCheck.Text = CurrentCheckPoint.ActivityToCheck
                            RichTextBox_AutoSize(RichTextBox_Step)
                            RichTextBox_AutoSize(RichTextBox_ActivityToCheck)
                            RichTextBox_ActivityToCheck.Refresh()
                            PanelSubForm.Controls.Clear()

                            If CurrentCheckPoint.Method = CheckSheetStep.MethodOption.ProcedureStep Then
                                Dim PSV As New ProcedureStepView
                                PSV.TopLevel = False
                                PSV.InputFeatures(CurrentCheckPoint.ProcessStep, CurrentCheckPoint.ProcedureStepAction.ProcedureMessage, CurrentCheckPoint.ProcedureStepAction.ImagePath_ProcedureStep)
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
            'Setting = AppControl.GetSettings("C:\TML_INI\YT_x64QCCheckAppliation\") 'System.Windows.Forms.Application.StartupPath)
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
                CurrentCheckPoint = YTA_CheckSheet_v1p3.ProcessStepNo(StepNo:=AllowedSteps(i), Initial:="46501497", CustOrd, ErrMsg:=ErMsg)
                If Not IsDBNull(CurrentCheckPoint) Then
                    If Not IsNothing(CurrentCheckPoint) Then
                        If Not IsNothing(CurrentCheckPoint.ActivityToCheck) Then
                            RichTextBox_Step.Text = CurrentCheckPoint.ProcessNo '& vbCrLf & CheckPoint.ProcessStep
                            TextBox_Step.Text = CurrentCheckPoint.StepNo
                            ToolTip1.SetToolTip(RichTextBox_Step, TextBox_Step.Text)
                            RichTextBox_ActivityToCheck.Text = CurrentCheckPoint.ActivityToCheck
                            RichTextBox_AutoSize(RichTextBox_Step)
                            RichTextBox_AutoSize(RichTextBox_ActivityToCheck)
                            RichTextBox_ActivityToCheck.Refresh()
                            PanelSubForm.Controls.Clear()

                            If CurrentCheckPoint.Method = CheckSheetStep.MethodOption.ProcedureStep Then
                                Dim PSV As New ProcedureStepView
                                PSV.TopLevel = False
                                PSV.InputFeatures(CurrentCheckPoint.ProcessStep, CurrentCheckPoint.ProcedureStepAction.ProcedureMessage, CurrentCheckPoint.ProcedureStepAction.ImagePath_ProcedureStep)
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
    Public Sub InspectionStatus(ByVal InspectionResult As CheckSheetStep, ByVal CheckResult As Boolean)
        Try
            InspectionResult.CheckResult = CheckResult
            AllCheckResult(Array.IndexOf(AllowedSteps, InspectionResult.StepNo)) = InspectionResult
            SetInspectionColor(InspectionResult.StepNo)
        Catch ex As Exception

        End Try
    End Sub
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
                If QcData_1p3.Count > 0 Then
                    For Each Item In QcData_1p3
                        If Item.PROCESS_NO = ProcessNo And Item.REMARK Like "*" & StepNo & "*" Then
                            If Item.REMARK Like "*GO*" Then
                                RichTextBox_Step.BackColor = Color.LightGreen
                            ElseIf Item.REMARK Like "*NG*" Then
                                RichTextBox_Step.BackColor = Color.OrangeRed
                            End If
                        End If
                    Next
                End If
                If QcData_1p4.Count > 0 Then
                    For Each Item In QcData_1p4
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
            CurrentCheckPoint = YTA_CheckSheet_v1p3.ProcessStepNo(StepNo:=Integer.Parse(RichTextBox_Step.Text), Initial:=Initial, CustOrd, ErrMsg:=ErMsg)
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
            QcData_1p3 = TmlEntityQA.GetDatabaseTableAs_List(Of POCO_QA.yta_qcc_v1p2)("INDEX_NO", CustOrd.INDEX_NO, "INDEX_NO", CustOrd.INDEX_NO, ErrMsg)
            If QcData_1p3.Count > 0 Then
                Dim BlankDoc As String = Setting.Var_06_DocsStore & "Production Release Documents\QC Check Sheets\" & CustOrd.PROD_NO & "\Line-" & CustOrd.LINE_NO & "-(Qty " & CustOrd.TOT_QTY & " Pcs)\" & CustOrd.INDEX_NO & "-QCSHEET.pdf"
                Dim FinalDoc As String = Setting.Var_06_DocsStore & "Production Complete Documents\Signed_QCC\" & CustOrd.PROD_NO & "\Line-" & CustOrd.LINE_NO & "\" & CustOrd.INDEX_NO & "-QCS-Signed.pdf"
                If SaveFinalDoc = False Then
                    Dim FinalFileName As String = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), System.IO.Path.GetFileName(FinalDoc))
                    FinalDoc = FinalFileName
                End If
                Dim UseDoc As New PdfSharp.Pdf.PdfDocument
                If System.IO.File.Exists(FinalDoc) Then
                    Dim P_Doc = OpenPdfOperation_x64.FileOp.GetDocument(BlankDoc, ErrMsg)
                    For Each QcWrite In QcData_1p3
                        Dim WriteParams = QcWrite.CHECK_RESULT.Split("$")
                        For Each WriteParam In WriteParams
                            Dim WriteInput = WriteParam.Split("-")(0)
                            WriteInput = WriteInput.Replace("|", "-")
                            Dim WriteSXY = WriteParam.Split("-")(1)
                            Dim WriteS = WriteSXY.Split(",")(0)
                            Dim WriteX = WriteSXY.Split(",")(1)
                            Dim WriteY = WriteSXY.Split(",")(2)
                            Dim TextToWrite As String = WriteInput
                            Dim FontToWrite As OpenPdfOperation_x64.FileOp.FontName = OpenPdfOperation_x64.FileOp.FontName.Arial
                            If WriteInput = "Tick" Then
                                TextToWrite = "ü"
                                FontToWrite = OpenPdfOperation_x64.FileOp.FontName.Wingdings
                            End If
                            Dim WP As New OpenPdfOperation_x64.WriteTextParameters With {
                                        .Name = FontToWrite,
                                        .Colour = OpenPdfOperation_x64.FileOp.FontColor.Black,
                                        .Style = OpenPdfOperation_x64.FileOp.FontStyle.Regular,
                                        .TextSize = Integer.Parse(WriteS),
                                        .TextValue = TextToWrite,
                                        .X_Position = CDbl(WriteX / 100),
                                        .Y_Position = CDbl(WriteY / 100)}
                            Dim WPS As New List(Of OpenPdfOperation_x64.WriteTextParameters)
                            WPS.Add(WP)
                            'OpenPdfOperation_x64.FileOp.PDF_WriteText(FinalDoc, FinalDoc, WPS, ErrMsg)
                            OpenPdfOperation_x64.FileOp.PDF_WriteText(P_Doc, WPS, ErrMsg)
                            If ErrMsg.Length > 0 Then
                                MsgBox(ErrMsg)
                                Exit Sub
                            End If
                        Next
                    Next
                    P_Doc.Save(FinalDoc)
                    Process.Start(FinalDoc)
                ElseIf System.IO.File.Exists(BlankDoc) Then
                    Dim P_Doc = OpenPdfOperation_x64.FileOp.GetDocument(BlankDoc, ErrMsg)
                    For Each QcWrite In QcData_1p3
                        Dim WriteParams = QcWrite.CHECK_RESULT.Split("$")
                        For Each WriteParam In WriteParams
                            Dim WriteInput = WriteParam.Split("-")(0)
                            Dim WriteSXY = WriteParam.Split("-")(1)
                            Dim WriteS = WriteSXY.Split(",")(0)
                            Dim WriteX = WriteSXY.Split(",")(1)
                            Dim WriteY = WriteSXY.Split(",")(2)
                            Dim TextToWrite As String = WriteInput
                            Dim FontToWrite As OpenPdfOperation_x64.FileOp.FontName = OpenPdfOperation_x64.FileOp.FontName.Arial
                            If WriteInput = "Tick" Then
                                TextToWrite = "ü"
                                FontToWrite = OpenPdfOperation_x64.FileOp.FontName.Wingdings
                            ElseIf WriteInput = "Circle" Then
                                TextToWrite = "¡"
                                FontToWrite = OpenPdfOperation_x64.FileOp.FontName.Wingdings
                            ElseIf WriteInput = "Cross" Then
                                TextToWrite = "û"
                                FontToWrite = OpenPdfOperation_x64.FileOp.FontName.Wingdings
                            End If
                            Dim WP As New OpenPdfOperation_x64.WriteTextParameters With {
                                        .Name = FontToWrite,
                                        .Colour = OpenPdfOperation_x64.FileOp.FontColor.Blue,
                                        .Style = OpenPdfOperation_x64.FileOp.FontStyle.Regular,
                                        .TextSize = Integer.Parse(WriteS),
                                        .TextValue = TextToWrite,
                                        .X_Position = CDbl(WriteX / 100),
                                        .Y_Position = CDbl(WriteY / 100)}
                            Dim WPS As New List(Of OpenPdfOperation_x64.WriteTextParameters)
                            WPS.Add(WP)
                            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(FinalDoc))
                            OpenPdfOperation_x64.FileOp.PDF_WriteText(P_Doc, WPS, ErrMsg)
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

            If IsDate(CustOrd.ACTUAL_FINISH_DATE) Then
                If My.Settings.Login <> "46501497" Then
                    WMsg.Message = "Current transmitter already finished on " & CustOrd.ACTUAL_FINISH_DATE & ". Please checK in Production Complete Documents Folder."
                    WMsg.ShowDialog()
                    Exit Sub
                Else
                    If MsgBox($"USER:{ My.Settings.Login} Do you want to run PrintQcc_Rev1() again with the DB CheckResult values and Print QCC?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                        WMsg.Message = "Current transmitter already finished on " & CustOrd.ACTUAL_FINISH_DATE & ". Please checK in Production Complete Documents Folder."
                        WMsg.ShowDialog()
                        Exit Sub
                    End If
                End If
            End If

            Dim ErrMsg As String = ""
            QcData_1p3 = TmlEntityQA.GetDatabaseTableAs_List(Of POCO_QA.yta_qcc_v1p2)("INDEX_NO", CustOrd.INDEX_NO, "INDEX_NO", CustOrd.INDEX_NO, ErrMsg)
            If QcData_1p3.Count > 0 Then
                RefreshSettings(Link.Network)
                Dim QcChecksheet As String = "YMA-TML-SF-16-v1p3-STD.pdf"
                Dim BlankDoc = Application.StartupPath & $"\05_Report_Templates\{QcChecksheet}"
                Dim FinalDoc As String = Setting.Var_06_DocsStore & "\Production Complete Documents\Signed_QCC\" & CustOrd.PROD_NO & "\Line-" & CustOrd.LINE_NO & "\" & CustOrd.INDEX_NO & "-QCS-Signed.pdf"
                If SaveFinalDoc = False Then
                    Dim FinalFileName As String = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), System.IO.Path.GetFileName(FinalDoc))
                    FinalDoc = FinalFileName
                End If


#Region "Fill Param Array  --YMA-TML-SF-16 Rev 1.2"


                Dim TextValue_SF16 As String = "This is a Test Text"
                Dim xPos_SF16 As Double = 0 'CDbl(PdfPage.Width) * CDbl(TextBox1.Text)
                Dim yPos_SF16 As Double = 0 'CDbl(PdfPage.Height) * CDbl(TextBox2.Text)
                Dim FontName_SF16 As String = "Arial" '"Verdana" '"Arial"
                Dim FontSize_SF16 As String = "10" '"12" '"10"
                Dim FontThick_SF16 As String = "Regular" '"Bold" '"Regular"
                Dim FontColor_SF16 As String = "Black"
                Dim FontSize_Small As String = "8" '"12" '"10"


                Dim Mark_YTA_Index As String = CustOrd.INDEX_NO
                Dim Mark_YTA_OrderNo As String = CustOrd.PROD_NO & "    Line: " & CustOrd.LINE_NO
                Dim Mark_MsCode_Before As String = CustOrd.MS_CODE_BEFORE
                Dim Mark_MsCode_After As String = CustOrd.MS_CODE
                Dim Mark_Serial_After As String = CustOrd.SERIAL_NO
                Dim Mark_PlateDataCorrect As String = ""
                Dim Mark_IM As String = ""
                Dim Mark_SafetyIM As String = ""
                Dim Mark_Header As String = ""
                Dim Mark_XJ_18 As String = "" '2 line Tag XJYTA710-18
                If CustOrd.XJ_NO = "XJYTA710-18" Then
                    Mark_XJ_18 = "(/Z → 2-LINE-TAG)"
                Else
                    Mark_XJ_18 = ""
                End If
                Dim Mark_XJ_24 As String = "" '5 Point IO XJYTA710-24
                If CustOrd.XJ_NO = "XJYTA710-24" Then
                    Mark_XJ_24 = "(/Z → 9P-I/O)"
                Else
                    Mark_XJ_24 = ""
                End If

                Dim BomLines = TmlEntityYGS.GetDatabaseTableAs_List(Of POCO_YGSP.common_bom_reserve_list)("RESNO", CoHeader.RESERVATION_NO, ErrMsg:=ErrMsg)
                If ErrMsg.Length > 0 Then
                    WMsg.Message = "common_bom_reserve_list Read Error: " & ErrMsg
                    WMsg.ShowDialog()
                    Exit Sub
                End If
                If BomLines.Count = 0 Then
                    Dim BomLinesIssued = TmlEntityYGS.GetDatabaseTableAs_List(Of POCO_YGSP.common_bom_reserve_list_issued)("RESNO", CoHeader.RESERVATION_NO, ErrMsg:=ErrMsg)
                    If ErrMsg.Length > 0 Then
                        WMsg.Message = "common_bom_reserve_list Read Error: " & ErrMsg
                        WMsg.ShowDialog()
                        Exit Sub
                    End If
                    If BomLinesIssued.Count = 0 Then
                        WMsg.Message = $"BOM {CoHeader.RESERVATION_NO} does not contains any YTA units"
                        WMsg.ShowDialog()
                        Exit Sub
                    Else
                        Dim PartCount = 0
                        For Each part In BomLinesIssued
                            If part.PARTNO Like "YTA*" Then
                                If part.PARTNO Like "*2007098838*" Then
                                    If CustOrd.QTY_NO = "01" Then
                                        Mark_PlateDataCorrect = "YES without CE"
                                        Mark_IM = "IM [Ed.7/21-YKSA]"
                                        Mark_SafetyIM = "SafetyIM [Ed.1]"
                                        Mark_Header = "non-RoHS10"
                                        PartCount += 1
                                    Else
                                        Mark_PlateDataCorrect = "YES without CE"
                                        Mark_IM = "IM and SafetyIM not required"
                                        Mark_SafetyIM = ""
                                        Mark_Header = "non-RoHS10"
                                        PartCount += 1
                                    End If
                                Else
                                    If CustOrd.QTY_NO = "01" Then
                                        Mark_PlateDataCorrect = "YES with CE"
                                        Mark_IM = "IM [Ed.7/21-0006]"
                                        Mark_SafetyIM = "SafetyIM [Ed.1]"
                                        Mark_Header = ""
                                    Else
                                        Mark_PlateDataCorrect = "YES with CE"
                                        Mark_IM = "IM and SafetyIM not required"
                                        Mark_SafetyIM = ""
                                        Mark_Header = ""
                                    End If
                                End If
                            End If
                        Next
                        If PartCount > 1 Then
                            ErrMsg = "Only 1 type of RoHS-6 YTA can be added"
                            WMsg.Message = "QCC File Print Error: " & ErrMsg
                            WMsg.ShowDialog()
                            Exit Sub
                        End If
                    End If
                Else
                    Dim PartCount = 0
                    For Each part In BomLines
                        If part.PARTNO Like "YTA*" Then
                            If part.PARTNO Like "*2007098838*" Then
                                If CustOrd.QTY_NO = "01" Then
                                    Mark_PlateDataCorrect = "YES without CE"
                                    Mark_IM = "IM [Ed.7/21-YKSA]"
                                    Mark_SafetyIM = "SafetyIM [Ed.1]"
                                    Mark_Header = "non-RoHS10"
                                    PartCount += 1
                                Else
                                    Mark_PlateDataCorrect = "YES without CE"
                                    Mark_IM = "IM and SafetyIM not required"
                                    Mark_SafetyIM = ""
                                    Mark_Header = "non-RoHS10"
                                    PartCount += 1
                                End If
                            Else
                                If CustOrd.QTY_NO = "01" Then
                                    Mark_PlateDataCorrect = "YES with CE"
                                    Mark_IM = "IM [Ed.7/21-0006]"
                                    Mark_SafetyIM = "SafetyIM [Ed.1]"
                                    Mark_Header = ""
                                Else
                                    Mark_PlateDataCorrect = "YES with CE"
                                    Mark_IM = "IM and SafetyIM not required"
                                    Mark_SafetyIM = ""
                                    Mark_Header = ""
                                End If
                            End If
                        End If
                    Next
                    If PartCount > 1 Then
                        ErrMsg = "Only 1 type of RoHS-6 YTA can be added"
                        WMsg.Message = "QCC File Print Error: " & ErrMsg
                        WMsg.ShowDialog()
                        Exit Sub
                    End If
                End If
                Mark_PlateDataCorrect = ""
                Mark_IM = ""
                Mark_SafetyIM = ""


                Dim CountryOfOrigin As String = "Original"
                If Len(Mark_Serial_After) > 5 Then
                    If CustOrd.EU_COUNTRY = "SA" And Mark_Serial_After Like "Y3???????" Then
                        CountryOfOrigin = "Made in KSA"
                    ElseIf CustOrd.EU_COUNTRY <> "SA" And Mark_Serial_After Like "91???????" Then
                        CountryOfOrigin = "Not allowed" '"Made in Japan"
                        ErrMsg = "Cannot make Made in Japan unit at YKSA TML"
                        WMsg.Message = "QCC File Print Error: " & ErrMsg
                        WMsg.ShowDialog()
                        Exit Sub
                    ElseIf CustOrd.EU_COUNTRY <> "SA" And (Mark_Serial_After Like "S5???????" Or Mark_Serial_After Like "Y3???????") Then
                        CountryOfOrigin = "Made in China"
                    ElseIf CustOrd.EU_COUNTRY <> "SA" And Mark_Serial_After Like "C2???????" Then
                        CountryOfOrigin = "Not allowed" '"Made in Singapore"
                        ErrMsg = "Cannot make Made in Singapore unit at YKSA TML"
                        WMsg.Message = "QCC File Print Error: " & ErrMsg
                        WMsg.ShowDialog()
                        Exit Sub
                    End If
                End If
                CountryOfOrigin = ""


                Dim MfgEntity As New MFG_ENTITY.Op(Link.Mysql_PMS_ConStr, MFG_ENTITY.Connections.DB_PMS)
                Dim SpecWrite_yta = MfgEntity.GetDatabaseAsModel_List(Of POCO_PMS.sf_templates)(New POCO_PMS.sf_templates, "DOC_NUMBER", "YMA-TML-SF-16", "REV", "1.2", ErrMsg)
                If ErrMsg.Length > 0 Then
                    WMsg.Message = "QCC File Print Error: " & ErrMsg
                    WMsg.ShowDialog()
                    Exit Sub
                End If
                Dim TextValue As String = "This is a Test Text"
                Dim xPos As Double = 0 'CDbl(PdfPage.Width) * CDbl(TextBox1.Text)
                Dim yPos As Double = 0 'CDbl(PdfPage.Height) * CDbl(TextBox2.Text)
                Dim FontName As String = "Arial" '"Verdana" '"Arial"
                Dim FontSize As String = "14" '"12" '"10"
                Dim FontThick As String = "Bold" '"Regular"
                Dim FontColor As String = "Black"

                Dim Pcount_YTA = -1
                Dim ParameterArray_SF16(100)
                For Each Parameter In SpecWrite_yta
                    Dim TempSpec_yta As New POCO_PMS.sf_templates
                    If Parameter.NAME = "INDEX_NO" Then
                        TextValue_SF16 = Mark_YTA_Index
                    ElseIf Parameter.NAME = "ORDER_NO" Then
                        TextValue_SF16 = Mark_YTA_OrderNo
                    ElseIf Parameter.NAME = "MD_BEFORE" Then
                        TextValue_SF16 = Mark_MsCode_Before
                    ElseIf Parameter.NAME = "MD_AFTER" Then
                        TextValue_SF16 = $"{Mark_MsCode_After}  {Mark_XJ_18}{Mark_XJ_24}"
                    ElseIf Parameter.NAME = "DATE_PROD" Then
                        TextValue_SF16 = "" 'Date.Parse(Unit.SCHEDULE).ToString("yyyy-MM-")
                    ElseIf Parameter.NAME = "SN_BEFORE" Then
                        TextValue_SF16 = "" 'SerialNoBefore
                    ElseIf Parameter.NAME = "SN_AFTER" Then
                        TextValue_SF16 = CustOrd.SERIAL_NO
                    ElseIf Parameter.NAME = "MAIN_IM_VER" Then
                        TextValue_SF16 = Mark_IM
                    ElseIf Parameter.NAME = "SAFE_IM_VER" Then
                        TextValue_SF16 = Mark_SafetyIM
                    ElseIf Parameter.NAME = "COO" Then
                        TextValue_SF16 = CountryOfOrigin
                    ElseIf Parameter.NAME = "PLATECHECK" Then
                        TextValue_SF16 = Mark_PlateDataCorrect
                    ElseIf Parameter.NAME = "HEADER" Then
                        TextValue_SF16 = Mark_Header
                    ElseIf Parameter.NAME = "XJYTA710-18" Then
                        TextValue_SF16 = Mark_XJ_18
                    ElseIf Parameter.NAME = "XJYTA710-24" Then
                        TextValue_SF16 = Mark_XJ_24
                    End If
                    TempSpec_yta = SpecWrite_yta.Where(Function(x) x.NAME = Parameter.NAME).First
                    xPos = TempSpec_yta.X_POS
                    yPos = TempSpec_yta.Y_POS
                    FontSize_SF16 = TempSpec_yta.FONT_SIZE
                    FontThick_SF16 = TempSpec_yta.FONT_THICK
                    FontName_SF16 = TempSpec_yta.FONT_NAME
                    Pcount_YTA += 1
                    ParameterArray_SF16.SetValue({TextValue_SF16, FontName_SF16, FontSize_SF16, FontThick_SF16, FontColor_SF16, xPos, yPos}, Pcount_YTA)
                Next
                ReDim Preserve ParameterArray_SF16(Pcount_YTA)

                Dim DocPDF As String = BlankDoc
                Dim P_Doc_Template As PdfSharp.Pdf.PdfDocument
                P_Doc_Template = PdfSharp.Pdf.IO.PdfReader.Open(DocPDF, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Modify)
                Dim PdfPage As PdfSharp.Pdf.PdfPage = P_Doc_Template.Pages(0)
                If Not System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(FinalDoc)) Then
                    System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(FinalDoc))
                End If
                PDF_AddText(DocPDF, FinalDoc, ParameterArray_SF16, ErrMsg)
                If Len(ErrMsg) > 0 Then
                    WMsg.Message = $"QC Template preperation PDF_AddText() Error: {ErrMsg}"
                    WMsg.ShowDialog()
                    Exit Sub
                Else
                    BlankDoc = FinalDoc 'The filled template will be the new blank document
                End If
#End Region

                'Write QCC Results
                Dim UseDoc As New PdfSharp.Pdf.PdfDocument
                If System.IO.File.Exists(BlankDoc) Then
                    Dim P_Doc = OpenPdfOperation_x64.FileOp.GetDocument(BlankDoc, ErrMsg)
                    Dim WPS As New List(Of OpenPdfOperation_x64.WriteTextParameters)

                    For Each QcWrite In QcData_1p3
                        If QcWrite.INDEX_NO.ToString.Length > 0 Then
                            'If QcWrite.PROCESS_NO = "200" Then
                            '    Dim Stopehere As String = ""
                            'End If
                            Dim WriteParams = QcWrite.CHECK_RESULT.Split("$")
                            For Each WriteParam In WriteParams
                                Dim WriteInput = WriteParam.Split("-")(0)
                                WriteInput = WriteInput.Replace("|", "-")
                                Dim WriteSXY = WriteParam.Split("-")(1)
                                Dim WriteS = WriteSXY.Split(",")(0)
                                Dim WriteX = WriteSXY.Split(",")(1)
                                Dim WriteY = WriteSXY.Split(",")(2)
                                Dim TextToWrite As String = WriteInput
                                Dim FontToWrite As OpenPdfOperation_x64.FileOp.FontName = OpenPdfOperation_x64.FileOp.FontName.Arial
                                If WriteInput = "Tick" Then
                                    TextToWrite = "ü"
                                    FontToWrite = OpenPdfOperation_x64.FileOp.FontName.Wingdings
                                ElseIf WriteInput = "Circle" Then
                                    TextToWrite = "¡"
                                    FontToWrite = OpenPdfOperation_x64.FileOp.FontName.Wingdings
                                ElseIf WriteInput = "Cross" Then
                                    TextToWrite = "û"
                                    FontToWrite = OpenPdfOperation_x64.FileOp.FontName.Wingdings
                                End If
                                Dim WP As New OpenPdfOperation_x64.WriteTextParameters With {
                                            .Name = FontToWrite,
                                            .Colour = OpenPdfOperation_x64.FileOp.FontColor.Blue,
                                            .Style = OpenPdfOperation_x64.FileOp.FontStyle.Regular,
                                            .TextSize = Integer.Parse(WriteS),
                                            .TextValue = TextToWrite,
                                            .X_Position = CDbl(WriteX / 100),
                                            .Y_Position = CDbl(WriteY / 100)}
                                WPS.Add(WP)
                            Next
                        Else
                            MsgBox("Hi")
                        End If
                    Next

                    'Save new inspection results
                    System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(FinalDoc))
                    OpenPdfOperation_x64.FileOp.PDF_WriteText(P_Doc, WPS, ErrMsg)
                    If ErrMsg.Length > 0 Then
                        MsgBox(ErrMsg)
                        Exit Sub
                    End If
                    P_Doc.Save(FinalDoc)

                    'Copy to MyDocuments to show locally
                    Dim FinalFileName = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), System.IO.Path.GetFileName(FinalDoc))
                    For i As Integer = 1 To 1
                        Try
                            System.IO.File.Copy(FinalDoc, FinalFileName, True)
                        Catch ex As Exception
                            Continue For
                        End Try
                    Next

                Else
                    MsgBox("QC Template Not created in TML System. Please create it first.")
                    Exit Sub
                End If


            End If

        Catch ex As Exception

        End Try
    End Sub
    Private Sub PDF_AddText(ByVal TemplatePDF As String, ByVal FinishedDoc As String, ByVal ParameterArray() As Object, ByRef ErrMsg As String)
        ErrMsg = ""
        Try

            Dim DocPDF As String = TemplatePDF
            Dim P_Doc As New PdfSharp.Pdf.PdfDocument
            P_Doc = PdfSharp.Pdf.IO.PdfReader.Open(DocPDF, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Modify)
            Dim PdfPage As PdfSharp.Pdf.PdfPage = P_Doc.Pages(0)
            Dim gfx As PdfSharp.Drawing.XGraphics = PdfSharp.Drawing.XGraphics.FromPdfPage(PdfPage)

            Dim RowCount As Integer = ParameterArray.GetLength(0) - 1

            Dim TextValue As String
            Dim X_pos As Double
            Dim Y_pos As Double
            Dim FontName As String
            Dim FontSize As String
            Dim FontThick As String
            Dim FontCol As String
            Dim FontStyle As PdfSharp.Drawing.XFontStyle
            Dim FontType As PdfSharp.Drawing.XFont
            Dim FontColor As New PdfSharp.Drawing.XSolidBrush

            For i As Integer = 0 To RowCount
                TextValue = ParameterArray(i)(0)
                FontName = ParameterArray(i)(1)
                FontSize = ParameterArray(i)(2)
                FontThick = ParameterArray(i)(3)
                If FontThick = "Bold" Then
                    FontStyle = PdfSharp.Drawing.XFontStyle.Bold
                ElseIf FontThick = "Regular" Then
                    FontStyle = PdfSharp.Drawing.XFontStyle.Regular
                ElseIf FontThick = "BoldItalic" Then
                    FontStyle = PdfSharp.Drawing.XFontStyle.BoldItalic
                End If
                FontType = New PdfSharp.Drawing.XFont(FontName, FontSize, FontStyle)
                FontCol = ParameterArray(i)(4)
                If FontCol = "Black" Then
                    FontColor = PdfSharp.Drawing.XBrushes.Black
                End If
                X_pos = CDbl(PdfPage.Width) * CDbl(ParameterArray(i)(5).ToString)
                Y_pos = CDbl(PdfPage.Height) * CDbl(ParameterArray(i)(6).ToString)
                gfx.DrawString(TextValue, FontType, FontColor, New PdfSharp.Drawing.XRect(X_pos, Y_pos, PdfPage.Width, PdfPage.Height), PdfSharp.Drawing.XStringFormats.TopLeft)
            Next

            P_Doc.Save(FinishedDoc)
        Catch ex As Exception

        End Try
    End Sub
    Private Sub RefreshSettings(ByVal Network As String)
        Try
            'Setting = AppControl.GetSettingsLive("C:\TML_INI\YT_x64QCCheckAppliation_EJA\")
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
    Public Sub PrintQcc_Rev2()
        Try


            If IsDate(CustOrd.ACTUAL_FINISH_DATE) Then
                If My.Settings.Login <> "46501497" Then
                    WMsg.Message = "Current transmitter already finished on " & CustOrd.ACTUAL_FINISH_DATE & ". Please checK in Production Complete Documents Folder."
                    WMsg.ShowDialog()
                    Exit Sub
                Else
                    If MsgBox($"USER:{ My.Settings.Login} Do you want to run PrintQcc_Rev2() again with the DB CheckResult values and Print QCC?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                        WMsg.Message = "Current transmitter already finished on " & CustOrd.ACTUAL_FINISH_DATE & ". Please checK in Production Complete Documents Folder."
                        WMsg.ShowDialog()
                        Exit Sub
                    End If
                End If
            End If

            Dim ErrMsg As String = ""
            QcData_1p4 = TmlEntityQA.GetDatabaseTableAs_List(Of POCO_QA.yta_qcc_v1p4)("INDEX_NO", CustOrd.INDEX_NO, "INDEX_NO", CustOrd.INDEX_NO, ErrMsg)
            If QcData_1p4.Count > 0 Then
                RefreshSettings(Link.Network)
                Dim BlankDoc As String = ""
                Dim FinalDoc As String = Setting.Var_06_DocsStore & "Production Complete Documents\Signed_QCC\" & CustOrd.PROD_NO & "\Line-" & CustOrd.LINE_NO & "\" & CustOrd.INDEX_NO & "-QCS-Signed.pdf"
                If SaveFinalDoc = False Then
                    Dim FinalFileName As String = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), System.IO.Path.GetFileName(FinalDoc))
                    FinalDoc = FinalFileName
                End If

Retry:
                Try
                    If System.IO.File.Exists(FinalDoc) Then
                        Using fs As New FileStream(FinalDoc, FileMode.Open, FileAccess.ReadWrite, FileShare.None)
                            ' File is not in use
                        End Using
                    End If
                Catch ex As IOException
                    ' File is in use
                    WMsg.Message = "QCC File Open Error: " & ex.Message
                    WMsg.ShowDialog()
                    If MsgBox("There is error in starting the QCC File. Do you want to re-open the file?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                        GoTo Retry
                    End If
                End Try

                Dim FinalTemplates As New List(Of OpenPdfOperation_x64.Template)
                For Each ProcessStep In QcData_1p4
                    Dim ReportResult = ProcessStep.CHECK_RESULT
                    LoggerHelper.LogInfo($"[CHECK_RESULT] for Index No: {CustOrd.INDEX_NO} is [ {ProcessStep.CHECK_RESULT} ]")
                    Dim Template As OpenPdfOperation_x64.Template = JsonConvert.DeserializeObject(Of OpenPdfOperation_x64.Template)(ReportResult)
                    FinalTemplates.Add(Template)
                Next
                Dim DistinctTemplateNames = FinalTemplates.Select(Function(X) X.FileName).Distinct.ToArray
                If DistinctTemplateNames.Count <> 1 Then
                    WMsg.Message = $"There are {DistinctTemplateNames.Count} Template Filenames to Write!"
                    WMsg.ShowDialog()
                    Exit Sub
                Else
                    BlankDoc = Application.StartupPath & $"\05_Report_Templates\{DistinctTemplateNames(0)}"
                    If System.IO.File.Exists(BlankDoc) Then
                        Dim WriteTemplate As New OpenPdfOperation_x64.Template
                        WriteTemplate.FileName = FinalTemplates.FirstOrDefault.FileName
                        WriteTemplate.Width = FinalTemplates.FirstOrDefault.Width
                        WriteTemplate.Height = FinalTemplates.FirstOrDefault.Height
                        WriteTemplate.ZoomFactor = FinalTemplates.FirstOrDefault.ZoomFactor
                        WriteTemplate.Dpi = FinalTemplates.FirstOrDefault.Dpi
                        WriteTemplate.Fields = New List(Of OpenPdfOperation_x64.Field)
                        For Each FinalTemplate In FinalTemplates
                            WriteTemplate.Fields.AddRange(FinalTemplate.Fields)
                        Next
                        If WriteTemplate.Fields.Count > 0 Then
                            OpenPdfOperation_x64.FileOp.PDF_XUnit_WriteJsonTextnBarcode(TemplatePDF:=BlankDoc, FinishedDoc:=FinalDoc, Param:=WriteTemplate, ErrMsg:=ErrMsg)
                            If ErrMsg.Length > 0 Then
                                MsgBox(ErrMsg)
                            End If
                        End If


                        Dim FinalFileName = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), System.IO.Path.GetFileName(FinalDoc))
                        For i As Integer = 1 To 1
                            Try
                                System.IO.File.Copy(FinalDoc, FinalFileName, True)
                            Catch ex As Exception
                                Continue For
                            End Try
                        Next
Retry_01:
                        Try
                            'open the final document
                            Process.Start(FinalFileName)
                        Catch ex As Exception
                            WMsg.Message = "QCC File Open Error: " & ex.Message
                            WMsg.ShowDialog()
                            If MsgBox("There is error in starting the QCC File. Do you want to re-open the file?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                                GoTo Retry_01
                            End If
                        End Try
                    Else
                        WMsg.Message = $"QC-Template: [ {DistinctTemplateNames(0)} ] not available in [ 13_Report_Templates ] Folder."
                        WMsg.ShowDialog()
                        Exit Sub
                    End If
                End If

            End If

        Catch ex As Exception
            'MsgBox(ex.Message)
            WMsg.Message = "PrintQcc_Rev1() Error:" & ex.Message
            WMsg.ShowDialog()
        End Try
    End Sub

#End Region

#Region "Menu Strip"
    Private Sub SelectStationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectStationToolStripMenuItem.Click
        Try
            'Dim Version = AppControl.GetVersion("C:\TML_INI\YT_x64QCCheckAppliation\")
            VersionText = AppControl.GetVersion(Application.StartupPath & "\00_Settings")
            Me.Text = $"YTA QC Check (x64)  [ App-Ver: { VersionText } ]  Station: [ {My.Settings.Station} ] QCC Ver: [ {CurrentQCC_Version} ] Plant: [ {Link.PlantID} ] Server: [ {Link.ServerAddress} ] QCC-Save: [ {SaveFinalDoc.ToString} ]"
Repeat:
            StationName = InputBox("Please select Station Name", "STATION", Setting.Var_08_StepsName)
            If Setting.Var_08_StepsName Like "*" & StationName & "*" And Not StationName.Contains(",") And StationName.Length > 0 Then
                My.Settings.Station = StationName
                My.Settings.Save()
                Me.Text = "YTA QC Check" & " [ Ver:" & VersionText & "]" & " [ Station:" & My.Settings.Station & "]"
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
            'VersionText = AppControl.GetVersion("C:\TML_INI\YT_x64QCCheckAppliation\")
            VersionText = AppControl.GetVersion(Application.StartupPath & "\00_Settings")
            Me.Text = $"YTA QC Check (x64)  [ App-Ver: { VersionText } ]  Station: [ {My.Settings.Station} ] QCC Ver: [ {CurrentQCC_Version} ] Plant: [ {Link.PlantID} ] Server: [ {Link.ServerAddress} ] QCC-Save: [ {SaveFinalDoc.ToString} ]"

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
    Private Sub GoToStepToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GoToStepToolStripMenuItem.Click
        Try
            If LoginID = "46501497" Then
                If Not IsNothing(CustOrd) Then
                    If CustOrd.MS_CODE.Length > 0 Then
                        Dim StepNo As String = ""
                        StepNo = InputBox("Input StepNo to go to", "STEP NO.", "")
                        If StepNo.Length > 0 Then
                            TextBox_Step.Text = StepNo
                            If AllowedSteps.Length > 0 Then
                                If Array.IndexOf(AllowedSteps, TextBox_Step.Text) - 1 >= 0 Then
                                    TextBox_Step.Text = AllowedSteps(Array.IndexOf(AllowedSteps, TextBox_Step.Text) - 1)
                                    Button2.PerformClick()
                                Else
                                    FirstCheckPoint()
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "PTR Reporting"
    Private Sub ReportPTRFailToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReportPTRFailToolStripMenuItem.Click
        Try
            Try
                Form_PTR.PtR_NC_Control1.CONNECTION_STRING_MAIN = Link.Mysql_YGSP_ConStr
                Form_PTR.PtR_NC_Control1.CONNECTION_STRING_PTR = Link.Mysql_QA_ConStr
                Form_PTR.PtR_NC_Control1.ASSEMBLY_LINE_ID = "YTA"
                Form_PTR.PtR_NC_Control1.STATION_ID = My.Settings.Station
                Form_PTR.PtR_NC_Control1.WAIT_WINDOW_PTR = 3
                With Form_PTR.PtR_NC_Control1
                    Form_PTR.Text = "PTR Reporting - Line:" & .ASSEMBLY_LINE_ID & ", Station:" & .STATION_ID
                End With
                Form_PTR.Show()
            Catch ex As Exception

            End Try
        Catch ex As Exception

        End Try
    End Sub

#End Region

#Region "QuickStep"
    Private Sub MainForm_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        Try
            If e.Alt And e.KeyCode = Keys.S Then
                TextBox_Step.Text = InputBox("Quick Step", "Quick Step", AllowedSteps(1))
                Dim i As Integer = Array.IndexOf(AllowedSteps, TextBox_Step.Text)
                If i > 0 Then
                    TextBox_Step.Text = AllowedSteps(i - 1)
                    Button2.PerformClick()
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub
#End Region


End Class
