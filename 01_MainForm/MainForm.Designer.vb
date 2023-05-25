<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainForm
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainForm))
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.RichTextBox_ActivityToCheck = New System.Windows.Forms.RichTextBox()
        Me.RichTextBox_Step = New System.Windows.Forms.RichTextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.PanelSubForm = New System.Windows.Forms.Panel()
        Me.TextBox_Step = New System.Windows.Forms.TextBox()
        Me.ContextMenuStation = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.SelectStationToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.PanelSubForm.SuspendLayout()
        Me.ContextMenuStation.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.[Single]
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.TableLayoutPanel2, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.PanelSubForm, 0, 1)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(12, 12)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 2
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 80.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(776, 426)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel2.ColumnCount = 4
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 5.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 15.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 75.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 5.0!))
        Me.TableLayoutPanel2.Controls.Add(Me.RichTextBox_ActivityToCheck, 2, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.RichTextBox_Step, 1, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.Button1, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.Button2, 3, 0)
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(4, 4)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 1
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(768, 78)
        Me.TableLayoutPanel2.TabIndex = 0
        '
        'RichTextBox_ActivityToCheck
        '
        Me.RichTextBox_ActivityToCheck.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RichTextBox_ActivityToCheck.Font = New System.Drawing.Font("Microsoft Sans Serif", 27.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RichTextBox_ActivityToCheck.Location = New System.Drawing.Point(156, 3)
        Me.RichTextBox_ActivityToCheck.Name = "RichTextBox_ActivityToCheck"
        Me.RichTextBox_ActivityToCheck.ReadOnly = True
        Me.RichTextBox_ActivityToCheck.Size = New System.Drawing.Size(570, 72)
        Me.RichTextBox_ActivityToCheck.TabIndex = 0
        Me.RichTextBox_ActivityToCheck.Text = ""
        '
        'RichTextBox_Step
        '
        Me.RichTextBox_Step.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RichTextBox_Step.Font = New System.Drawing.Font("Microsoft Sans Serif", 27.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RichTextBox_Step.Location = New System.Drawing.Point(41, 3)
        Me.RichTextBox_Step.Name = "RichTextBox_Step"
        Me.RichTextBox_Step.ReadOnly = True
        Me.RichTextBox_Step.Size = New System.Drawing.Size(109, 72)
        Me.RichTextBox_Step.TabIndex = 1
        Me.RichTextBox_Step.Text = ""
        '
        'Button1
        '
        Me.Button1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Button1.Location = New System.Drawing.Point(3, 3)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(32, 72)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "<"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Button2.Location = New System.Drawing.Point(732, 3)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(33, 72)
        Me.Button2.TabIndex = 3
        Me.Button2.Text = ">"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'PanelSubForm
        '
        Me.PanelSubForm.AutoScroll = True
        Me.PanelSubForm.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PanelSubForm.Controls.Add(Me.TextBox_Step)
        Me.PanelSubForm.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PanelSubForm.Location = New System.Drawing.Point(4, 89)
        Me.PanelSubForm.Name = "PanelSubForm"
        Me.PanelSubForm.Size = New System.Drawing.Size(768, 333)
        Me.PanelSubForm.TabIndex = 1
        '
        'TextBox_Step
        '
        Me.TextBox_Step.Location = New System.Drawing.Point(731, 3)
        Me.TextBox_Step.Name = "TextBox_Step"
        Me.TextBox_Step.Size = New System.Drawing.Size(32, 20)
        Me.TextBox_Step.TabIndex = 0
        '
        'ContextMenuStation
        '
        Me.ContextMenuStation.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SelectStationToolStripMenuItem, Me.ToolStripMenuItem1})
        Me.ContextMenuStation.Name = "ContextMenuStation"
        Me.ContextMenuStation.Size = New System.Drawing.Size(188, 70)
        '
        'SelectStationToolStripMenuItem
        '
        Me.SelectStationToolStripMenuItem.Name = "SelectStationToolStripMenuItem"
        Me.SelectStationToolStripMenuItem.Size = New System.Drawing.Size(187, 22)
        Me.SelectStationToolStripMenuItem.Text = "Select Station"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(187, 22)
        Me.ToolStripMenuItem1.Text = "Refresh Settings Read"
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.ContextMenuStrip = Me.ContextMenuStation
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.DoubleBuffered = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "MainForm"
        Me.Text = "QC Check"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.PanelSubForm.ResumeLayout(False)
        Me.PanelSubForm.PerformLayout()
        Me.ContextMenuStation.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents TableLayoutPanel2 As TableLayoutPanel
    Friend WithEvents PanelSubForm As Panel
    Friend WithEvents RichTextBox_ActivityToCheck As RichTextBox
    Friend WithEvents RichTextBox_Step As RichTextBox
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents TextBox_Step As TextBox
    Friend WithEvents ContextMenuStation As ContextMenuStrip
    Friend WithEvents SelectStationToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As ToolStripMenuItem
End Class
