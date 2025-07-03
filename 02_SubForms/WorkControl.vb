Public Class WorkControl

    Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As IntPtr, ByVal nIndex As Integer) As Integer
    Public Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Integer) As Integer
    Public Const GWL_STYLE As Integer = (-16)
    Public Const WS_VSCROLL As Integer = &H200000
    Public Const WS_HSCROLL As Integer = &H100000

    Public SqlConnection As String 'Mandatory input variable
    Public ResNo As String 'Mandatory input variable
    Public Action As String 'To use output variable: EXIT or CONTINUE
    Public ControlModel As New POCO_YGSP.controlcenter
    Private Sub WorkControl_Load(sender As Object, e As EventArgs) Handles Me.Load
        If SqlConnection.Length > 0 And ResNo.Length > 0 Then
            Action = ""
            Dim TEty As New MFG_ENTITY.Op(SqlConnection)
            Dim Cc = TEty.GetDatabaseTableAs_List(Of POCO_YGSP.controlcenter)("CONTROL_BOM", ResNo, "STATUS", "ACTIVE")
            If Cc.Count > 0 Then
                For Each SetControl In Cc
                    'Stations: CASE_ASSY,BUILD,HIPOT,PMS,CRC,FINALASSY,PACKING
                    'WorkControls:SCHEDULE, QCCS_PRINT, BOM_PRINT, BOM_PICK, MARKING, CASE_ASSY, BUILD,HIPOT, PMS,CRC, FINAL_INSP, PACKING, READY_MAIL, SHIP_MAIL
                    If (SetControl.CONTROL_PROCESS = MainForm.StationName) Or (SetControl.CONTROL_PROCESS = "FINAL_INSP" And MainForm.StationName = "FINALASSY") Then
                        RichTextBox_Header.Text = SetControl.CONTROL_TYPE
                        RichTextBox_ID.Text = SetControl.ULOCK
                        RichTextBox_Message.Text = SetControl.MESSAGE
                        ControlModel = SetControl
                        RichTextBox_AutoSize(RichTextBox_Header)
                        RichTextBox_AutoSize(RichTextBox_ID)
                        RichTextBox_AutoSize(RichTextBox_Message)
                        If ControlModel.CONTROL_TYPE = "STOP" Then
                            Action = "EXIT"
                            Exit Sub
                        Else
                            Action = "CONTINUE"
                            Exit Sub
                        End If
                    End If
                Next
            End If
        End If
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub WorkControl_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Button1.Focus()
    End Sub

End Class