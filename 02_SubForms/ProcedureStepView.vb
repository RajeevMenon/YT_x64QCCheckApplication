Public Class ProcedureStepView

    Public Message As String = ""
    Public WindowText As String = ""
    Public PicturePath As String = ""
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

    ''' <summary>
    ''' Assign the input parameters to load images and message in the SIP window when it opens to user.
    ''' </summary>
    ''' <param name="WindowText">Window header text to be displayed</param>
    ''' <param name="Message">Message to be displayed at the TOP</param>
    ''' <param name="PicturePath">Full path of the Picture file to be displayed</param>
    Public Sub InputFeatures(ByVal WindowText As String, ByVal Message As String, ByVal PicturePath As String)

        Result = Status.Start

        Try
            If WindowText <> "" Then
                Me.Text = WindowText
            Else
                Me.Text = "PROCEDURE VIEW"
            End If
            If Message <> "" Then
                RichTextBox1.Text = Message
            Else
                RichTextBox1.Text = ""
            End If
            RichTextBox_AutoSize(RichTextBox1)

            If PicturePath <> "" Then
                Dim Source As Image
                Source = Image.FromFile(PicturePath)

                Dim MaxHeight As Integer = PictureBox1.Height
                Dim AdjRatio As Decimal = MaxHeight / Source.Height
                Dim NewHeight As Integer = CInt(Source.Height * AdjRatio)
                Dim NewWidth As Integer = CInt(Source.Width * AdjRatio)
                Dim resizedImage As Image = ResizeImage(Source, NewWidth, NewHeight)
                PictureBox1.Image = resizedImage

                'If Source.Width > PictureBox1.ClientSize.Width OrElse Source.Height > PictureBox1.ClientSize.Height Then
                '    PictureBox1.SizeMode = PictureBoxSizeMode.AutoSize 'if image is larger than picturebox
                'Else
                '    PictureBox1.SizeMode = PictureBoxSizeMode.CenterImage 'if image is smaller than picturebox
                'End If
                If Not IsNothing(MainForm.CurrentCheckPoint.ProcedureStepAction.SetSizeMode) Then
                    If MainForm.CurrentCheckPoint.ProcedureStepAction.SetSizeMode = ProcedureStep.SizeMode.Normal Then PictureBox1.SizeMode = PictureBoxSizeMode.Normal
                    If MainForm.CurrentCheckPoint.ProcedureStepAction.SetSizeMode = ProcedureStep.SizeMode.AutoSize Then PictureBox1.SizeMode = PictureBoxSizeMode.AutoSize
                    If MainForm.CurrentCheckPoint.ProcedureStepAction.SetSizeMode = ProcedureStep.SizeMode.CenterImage Then PictureBox1.SizeMode = PictureBoxSizeMode.CenterImage
                    If MainForm.CurrentCheckPoint.ProcedureStepAction.SetSizeMode = ProcedureStep.SizeMode.StretchImage Then PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
                    If MainForm.CurrentCheckPoint.ProcedureStepAction.SetSizeMode = ProcedureStep.SizeMode.Zoom Then PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
                End If
                PictureBox1.Anchor = AnchorStyles.Top

                'Me.DoubleBuffered = True
                'Panel1.AutoScroll = True
                'Panel1.BackgroundImageLayout = ImageLayout.Center
                'Me.Refresh()
                'Panel1.BackgroundImage = Source
            End If

        Catch ex As Exception

        End Try
    End Sub

    ' Function to resize the image
    Private Function ResizeImage(ByVal img As Image, ByVal width As Integer, ByVal height As Integer) As Image
        ' Create a new bitmap with the desired size
        Dim bmp As New Bitmap(width, height)

        ' Use Graphics to draw the resized image
        Using g As Graphics = Graphics.FromImage(bmp)
            ' Set high-quality interpolation modes
            g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
            g.DrawImage(img, 0, 0, width, height)
        End Using

        Return bmp
    End Function

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Try
            Result = Status.Finish
            MainForm.InspectionStatus(MainForm.CurrentCheckPoint, True)
            MainForm.Button2.PerformClick()
            Me.Hide()
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Me.Close()
            MainForm.InspectionStatus(MainForm.CurrentCheckPoint, False)
            MainForm.ShowCompleteInspection()
        Catch ex As Exception

        End Try
    End Sub
End Class