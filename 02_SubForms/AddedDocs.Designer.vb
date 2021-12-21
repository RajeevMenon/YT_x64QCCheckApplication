<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AddedDocs
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Button_Done = New System.Windows.Forms.Button()
        Me.Button_Cancel = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox_IM = New System.Windows.Forms.TextBox()
        Me.TextBox_SIM = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBox_EUDoC = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label_Error = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Button_Done
        '
        Me.Button_Done.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Button_Done.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_Done.Location = New System.Drawing.Point(149, 395)
        Me.Button_Done.Name = "Button_Done"
        Me.Button_Done.Size = New System.Drawing.Size(124, 43)
        Me.Button_Done.TabIndex = 1
        Me.Button_Done.Text = "DONE"
        Me.Button_Done.UseVisualStyleBackColor = True
        '
        'Button_Cancel
        '
        Me.Button_Cancel.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Button_Cancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.Button_Cancel.Location = New System.Drawing.Point(306, 395)
        Me.Button_Cancel.Name = "Button_Cancel"
        Me.Button_Cancel.Size = New System.Drawing.Size(124, 43)
        Me.Button_Cancel.TabIndex = 2
        Me.Button_Cancel.Text = "CANCEL"
        Me.Button_Cancel.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 24.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(91, 39)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(417, 37)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "USER MANUAL(INST.-IM)"
        '
        'TextBox_IM
        '
        Me.TextBox_IM.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.TextBox_IM.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox_IM.Font = New System.Drawing.Font("Microsoft Sans Serif", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_IM.Location = New System.Drawing.Point(133, 79)
        Me.TextBox_IM.Name = "TextBox_IM"
        Me.TextBox_IM.Size = New System.Drawing.Size(331, 44)
        Me.TextBox_IM.TabIndex = 5
        '
        'TextBox_SIM
        '
        Me.TextBox_SIM.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.TextBox_SIM.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox_SIM.Font = New System.Drawing.Font("Microsoft Sans Serif", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_SIM.Location = New System.Drawing.Point(133, 193)
        Me.TextBox_SIM.Name = "TextBox_SIM"
        Me.TextBox_SIM.Size = New System.Drawing.Size(331, 44)
        Me.TextBox_SIM.TabIndex = 7
        '
        'Label2
        '
        Me.Label2.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 24.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(98, 154)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(393, 37)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "SAFETY MANUAL(S-IM)"
        '
        'TextBox_EUDoC
        '
        Me.TextBox_EUDoC.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.TextBox_EUDoC.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox_EUDoC.Font = New System.Drawing.Font("Microsoft Sans Serif", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_EUDoC.Location = New System.Drawing.Point(133, 316)
        Me.TextBox_EUDoC.Name = "TextBox_EUDoC"
        Me.TextBox_EUDoC.Size = New System.Drawing.Size(331, 44)
        Me.TextBox_EUDoC.TabIndex = 9
        '
        'Label3
        '
        Me.Label3.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 24.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(225, 280)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(141, 37)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "EU-DoC"
        '
        'Label_Error
        '
        Me.Label_Error.AutoSize = True
        Me.Label_Error.ForeColor = System.Drawing.Color.Red
        Me.Label_Error.Location = New System.Drawing.Point(17, 10)
        Me.Label_Error.Name = "Label_Error"
        Me.Label_Error.Size = New System.Drawing.Size(10, 13)
        Me.Label_Error.TabIndex = 10
        Me.Label_Error.Text = "."
        '
        'AddedDocs
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.PaleTurquoise
        Me.ClientSize = New System.Drawing.Size(615, 450)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label_Error)
        Me.Controls.Add(Me.TextBox_EUDoC)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TextBox_SIM)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBox_IM)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button_Cancel)
        Me.Controls.Add(Me.Button_Done)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "AddedDocs"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "SCAN DOCUMENT QR CODE"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button_Done As Button
    Friend WithEvents Button_Cancel As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents TextBox_IM As TextBox
    Friend WithEvents TextBox_SIM As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents TextBox_EUDoC As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Label_Error As Label
End Class
