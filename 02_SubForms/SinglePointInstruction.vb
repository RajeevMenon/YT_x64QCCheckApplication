''' <summary>
''' Library to handle Single Point instruction check point with Images.
''' </summary>
Public Class SinglePointInst
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
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            SelectedSideFile = System.IO.Path.GetFileName(PictureBox1.ImageLocation)
            Me.Hide()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            SelectedSideFile = System.IO.Path.GetFileName(PictureBox2.ImageLocation)
            Me.Hide()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Try
            SelectedSideFile = System.IO.Path.GetFileName(PictureBox1.ImageLocation)
            Me.Hide()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Try
            SelectedSideFile = System.IO.Path.GetFileName(PictureBox2.ImageLocation)
            Me.Hide()
        Catch ex As Exception

        End Try
    End Sub
End Class