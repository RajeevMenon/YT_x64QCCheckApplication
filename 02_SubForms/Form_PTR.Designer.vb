<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_PTR
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
        Me.PtR_NC_Control1 = New PTR_NC_Counter.PTR_NC_Control()
        Me.SuspendLayout()
        '
        'PtR_NC_Control1
        '
        Me.PtR_NC_Control1.ASSEMBLY_LINE_ID = ""
        Me.PtR_NC_Control1.BackColor = System.Drawing.Color.Black
        Me.PtR_NC_Control1.CONNECTION_STRING_MAIN = ""
        Me.PtR_NC_Control1.CONNECTION_STRING_PTR = ""
        Me.PtR_NC_Control1.Location = New System.Drawing.Point(-1, 4)
        Me.PtR_NC_Control1.Name = "PtR_NC_Control1"
        Me.PtR_NC_Control1.Size = New System.Drawing.Size(909, 524)
        Me.PtR_NC_Control1.STATION_ID = ""
        Me.PtR_NC_Control1.TabIndex = 0
        Me.PtR_NC_Control1.WAIT_WINDOW_PTR = 3
        '
        'Form_PTR
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(910, 528)
        Me.ControlBox = False
        Me.Controls.Add(Me.PtR_NC_Control1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "Form_PTR"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form_PTR"
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents PtR_NC_Control1 As PTR_NC_Counter.PTR_NC_Control
End Class
