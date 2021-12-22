''' <summary>
''' Library to handle Single Point instruction check point with Images.
''' </summary>
Public Class SinglePointInstruction

    Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As IntPtr, ByVal nIndex As Integer) As Integer
    Public Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Integer) As Integer
    Public Const GWL_STYLE As Integer = (-16)
    Public Const WS_VSCROLL As Integer = &H200000
    Public Const WS_HSCROLL As Integer = &H100000

    ''' <summary>
    ''' The path of Image file to be loaded on the LEFT view.
    ''' </summary>
    Dim LeftImage As String
    ''' <summary>
    ''' The Path of Imgae file to be loaded on the RIGHT view.
    ''' </summary>
    Dim RightImage As String
    ''' <summary>
    ''' Message to be displayed on the TOP message display area
    ''' </summary>
    Dim Message As String

    ''' <summary>
    ''' Return either 'LEFT' or 'RIGHT' as per user selection to use in caller subroutiens.
    ''' </summary>
    Public SelectedSideFile As String

    ''' <summary>
    ''' Assign the input parameters to load images and message in the SIP window when it opens to user.
    ''' </summary>
    ''' <param name="Image_1_Path">First Image file name with path to be loaded into any one picture box</param>
    ''' <param name="Image_2_Path">First Image file name with path to be loaded into any one picture box</param>
    ''' <param name="MessageText">Message to be displayed on the TOP message display area</param>
    Public Sub InputFeatures(ByVal Image_1_Path As String, ByVal Image_2_Path As String, ByVal MessageText As String)
        Try
            Dim rn As New Random
            If rn.Next(1, 3) = 1 Then 'Lowerbound is inclusive and upper bound is exclusive
                LeftImage = Image_1_Path
                RightImage = Image_2_Path
            Else
                RightImage = Image_1_Path
                LeftImage = Image_2_Path
            End If
            Message = MessageText
        Catch ex As Exception

        End Try
    End Sub

    Private Sub SinglePointInstruction_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            LoadImage()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub LoadImage()
        Try
            If LeftImage <> "" Then
                Dim Source As Image
                Source = Image.FromFile(LeftImage)
                If Source.Width > PictureBox1.ClientSize.Width OrElse Source.Height > PictureBox1.ClientSize.Height Then
                    PictureBox1.SizeMode = PictureBoxSizeMode.Zoom 'if image is larger than picturebox
                Else
                    PictureBox1.SizeMode = PictureBoxSizeMode.CenterImage 'if image is smaller than picturebox
                End If
                PictureBox1.Load(LeftImage) 'This is needed to load with file path as the file name is needed when user selects this input
            End If
            If RightImage <> "" Then
                Dim Source As Image
                Source = Image.FromFile(RightImage)
                If Source.Width > PictureBox2.ClientSize.Width OrElse Source.Height > PictureBox2.ClientSize.Height Then
                    PictureBox2.SizeMode = PictureBoxSizeMode.Zoom 'if image is larger than picturebox
                Else
                    PictureBox2.SizeMode = PictureBoxSizeMode.CenterImage 'if image is smaller than picturebox
                End If
                PictureBox2.Load(RightImage)
            End If
            RichTextBoxMessage.Text = Message
            RichTextBox_AutoSize(RichTextBoxMessage)
            PictureBox1.Select()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Try
            SelectedSideFile = System.IO.Path.GetFileName(PictureBox1.ImageLocation)
            If SelectedSideFile Like "*" & MainForm.CurrentCheckPoint.SinglePointAction.ImagePath_SPI_Correct & "*" Then
                MainForm.Button2.PerformClick()
                Me.Close()
            Else
                InputFeatures(MainForm.CurrentCheckPoint.SinglePointAction.ImagePath_SPI_1, MainForm.CurrentCheckPoint.SinglePointAction.ImagePath_SPI_2, MainForm.CurrentCheckPoint.SinglePointAction.SPI_Message)
                LoadImage()
                Me.Refresh()
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)
        Try
            SelectedSideFile = System.IO.Path.GetFileName(PictureBox2.ImageLocation)
            If SelectedSideFile Like "*" & MainForm.CurrentCheckPoint.SinglePointAction.ImagePath_SPI_Correct & "*" Then
                MainForm.Button2.PerformClick()
                Me.Close()
            Else
                InputFeatures(MainForm.CurrentCheckPoint.SinglePointAction.ImagePath_SPI_1, MainForm.CurrentCheckPoint.SinglePointAction.ImagePath_SPI_2, MainForm.CurrentCheckPoint.SinglePointAction.SPI_Message)
                LoadImage()
                Me.Refresh()
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Try
            SelectedSideFile = System.IO.Path.GetFileName(PictureBox1.ImageLocation)
            If SelectedSideFile = MainForm.CurrentCheckPoint.SinglePointAction.ImagePath_SPI_Correct Then
                MainForm.Button2.PerformClick()
                Me.Close()
            Else
                RichTextBoxMessage.BackColor = Color.Red
                MainForm.wait(1)
                RichTextBoxMessage.BackColor = Color.PaleGreen
                InputFeatures(MainForm.CurrentCheckPoint.SinglePointAction.ImagePath_SPI_1, MainForm.CurrentCheckPoint.SinglePointAction.ImagePath_SPI_2, MainForm.CurrentCheckPoint.SinglePointAction.SPI_Message)
                LoadImage()
                Me.Refresh()
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Try
            SelectedSideFile = System.IO.Path.GetFileName(PictureBox2.ImageLocation)
            If SelectedSideFile = MainForm.CurrentCheckPoint.SinglePointAction.ImagePath_SPI_Correct Then
                MainForm.Button2.PerformClick()
                Me.Close()
            Else
                RichTextBoxMessage.BackColor = Color.Red
                MainForm.wait(1)
                RichTextBoxMessage.BackColor = Color.PaleGreen
                InputFeatures(MainForm.CurrentCheckPoint.SinglePointAction.ImagePath_SPI_1, MainForm.CurrentCheckPoint.SinglePointAction.ImagePath_SPI_2, MainForm.CurrentCheckPoint.SinglePointAction.SPI_Message)
                LoadImage()
                Me.Refresh()
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