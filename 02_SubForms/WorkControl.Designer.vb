<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class WorkControl
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.RichTextBox_ID = New System.Windows.Forms.RichTextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.RichTextBox_Message = New System.Windows.Forms.RichTextBox()
        Me.RichTextBox_Header = New System.Windows.Forms.RichTextBox()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Yellow
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.RichTextBox_ID)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Controls.Add(Me.RichTextBox_Message)
        Me.Panel1.Controls.Add(Me.RichTextBox_Header)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(669, 251)
        Me.Panel1.TabIndex = 0
        '
        'RichTextBox_ID
        '
        Me.RichTextBox_ID.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.RichTextBox_ID.Location = New System.Drawing.Point(11, 6)
        Me.RichTextBox_ID.Name = "RichTextBox_ID"
        Me.RichTextBox_ID.Size = New System.Drawing.Size(122, 89)
        Me.RichTextBox_ID.TabIndex = 3
        Me.RichTextBox_ID.Text = ""
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(624, 6)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(32, 89)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "OK"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'RichTextBox_Message
        '
        Me.RichTextBox_Message.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.RichTextBox_Message.Location = New System.Drawing.Point(11, 101)
        Me.RichTextBox_Message.Name = "RichTextBox_Message"
        Me.RichTextBox_Message.Size = New System.Drawing.Size(645, 137)
        Me.RichTextBox_Message.TabIndex = 1
        Me.RichTextBox_Message.Text = ""
        '
        'RichTextBox_Header
        '
        Me.RichTextBox_Header.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.RichTextBox_Header.Location = New System.Drawing.Point(139, 6)
        Me.RichTextBox_Header.Name = "RichTextBox_Header"
        Me.RichTextBox_Header.Size = New System.Drawing.Size(479, 89)
        Me.RichTextBox_Header.TabIndex = 0
        Me.RichTextBox_Header.Text = ""
        '
        'WorkControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(669, 251)
        Me.ControlBox = False
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "WorkControl"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "WorkControl"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Panel1 As Panel
    Friend WithEvents Button1 As Button
    Friend WithEvents RichTextBox_Message As RichTextBox
    Friend WithEvents RichTextBox_Header As RichTextBox
    Friend WithEvents RichTextBox_ID As RichTextBox
End Class
