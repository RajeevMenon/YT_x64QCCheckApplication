<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ViewDocument
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
        Me.ViewBrowser = New System.Windows.Forms.WebBrowser()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Button_NG = New System.Windows.Forms.Button()
        Me.Button_GO = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ViewBrowser
        '
        Me.ViewBrowser.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ViewBrowser.Location = New System.Drawing.Point(0, 32)
        Me.ViewBrowser.MinimumSize = New System.Drawing.Size(20, 20)
        Me.ViewBrowser.Name = "ViewBrowser"
        Me.ViewBrowser.Size = New System.Drawing.Size(600, 436)
        Me.ViewBrowser.TabIndex = 2
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.Blue
        Me.Button1.Location = New System.Drawing.Point(478, 3)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(57, 23)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "Reload"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.ForeColor = System.Drawing.Color.Blue
        Me.Button3.Location = New System.Drawing.Point(538, 3)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(59, 23)
        Me.Button3.TabIndex = 1
        Me.Button3.Text = "Close"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.PaleGreen
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.Button_NG)
        Me.Panel1.Controls.Add(Me.Button_GO)
        Me.Panel1.Controls.Add(Me.Button3)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Controls.Add(Me.ViewBrowser)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(602, 470)
        Me.Panel1.TabIndex = 1
        '
        'Button_NG
        '
        Me.Button_NG.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_NG.ForeColor = System.Drawing.Color.Red
        Me.Button_NG.Location = New System.Drawing.Point(121, 3)
        Me.Button_NG.Name = "Button_NG"
        Me.Button_NG.Size = New System.Drawing.Size(57, 23)
        Me.Button_NG.TabIndex = 4
        Me.Button_NG.Text = "NG"
        Me.Button_NG.UseVisualStyleBackColor = True
        '
        'Button_GO
        '
        Me.Button_GO.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_GO.ForeColor = System.Drawing.Color.Blue
        Me.Button_GO.Location = New System.Drawing.Point(23, 3)
        Me.Button_GO.Name = "Button_GO"
        Me.Button_GO.Size = New System.Drawing.Size(57, 23)
        Me.Button_GO.TabIndex = 3
        Me.Button_GO.Text = "GO"
        Me.Button_GO.UseVisualStyleBackColor = True
        '
        'ViewDocument
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(602, 470)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "ViewDocument"
        Me.Text = "Document View"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ViewBrowser As System.Windows.Forms.WebBrowser
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Button_NG As Button
    Friend WithEvents Button_GO As Button
End Class
