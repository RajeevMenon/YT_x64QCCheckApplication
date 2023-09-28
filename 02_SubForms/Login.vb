Imports System.Security.Cryptography 'MD5
Imports System.Text 'Encoding
Public Class Login

    ' TODO: Insert code to perform custom authentication using the provided username and password 
    ' (See https://go.microsoft.com/fwlink/?LinkId=35339).  
    ' The custom principal can then be attached to the current thread's principal as follows: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' where CustomPrincipal is the IPrincipal implementation used to perform authentication. 
    ' Subsequently, My.User will return identity information encapsulated in the CustomPrincipal object
    ' such as the username, display name, etc.

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        Dim loginID As String = UsernameTextBox.Text
        Dim psswd As String = PasswordTextBox.Text

        My.Settings.Login = UsernameTextBox.Text
        My.Settings.Save()

        Dim TmlEntity As New MFG_ENTITY.Op(MainForm.Setting.Var_03_MySql_YGSP)
        Dim ErrMsg As String = ""
        Dim LoginTbl = TmlEntity.GetDatabaseTableAs_Object(Of POCO_YGSP.login)("LOGIN", loginID, "LOGIN", loginID, ErrMsg)
        If ErrMsg.Length > 0 Then
            MsgBox(ErrMsg)
        Else
            If GetMD5(psswd) = LoginTbl.PASSWORD Then
                Me.Close()
                MainForm.Initial = LoginTbl.INITIAL & "(" & Date.Today.ToString("MMM/dd") & ")"
                MainForm.WorkerName = LoginTbl.PERSON
                MainForm.WorkerID = LoginTbl.INITIAL
                MainForm.ContextMenuStation.Enabled = True
                MainForm.LoadIndexScan()
            Else
                PasswordTextBox.BackColor = Color.Red
                MainForm.wait(1)
                PasswordTextBox.Text = ""
                PasswordTextBox.BackColor = Color.White
            End If
        End If
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        UsernameTextBox.Text = ""
        PasswordTextBox.Text = ""
    End Sub

    Private Sub Login_Load(sender As Object, e As EventArgs) Handles Me.Load
        If My.Settings.Login.Length > 0 Then
            UsernameTextBox.Text = My.Settings.Login
            PasswordTextBox.Select()
            PasswordTextBox.ScrollToCaret()
        End If
    End Sub
    Public Function GetMD5(ByVal Input As String) As String
        Using md5Hash As MD5 = MD5.Create()
            Dim data As Byte() = md5Hash.ComputeHash(Encoding.UTF8.GetBytes(Input)) ' Import System.Text for getting Encoding function
            Dim sBuilder As New StringBuilder()
            Dim i As Integer
            For i = 0 To data.Length - 1
                sBuilder.Append(data(i).ToString("x2"))
            Next i
            Return sBuilder.ToString()
        End Using
    End Function

End Class
