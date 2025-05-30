<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.BTheader = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.BTexit = New System.Windows.Forms.Button()
        Me.BTSetting = New System.Windows.Forms.Button()
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.EndTaskMicrosoftExelToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.AddStartup = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'BTheader
        '
        Me.BTheader.BackColor = System.Drawing.Color.FromArgb(CType(CType(115, Byte), Integer), CType(CType(8, Byte), Integer), CType(CType(14, Byte), Integer))
        Me.BTheader.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.BTheader.Cursor = System.Windows.Forms.Cursors.SizeAll
        Me.BTheader.FlatAppearance.BorderSize = 0
        Me.BTheader.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BTheader.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BTheader.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.BTheader.Location = New System.Drawing.Point(12, 13)
        Me.BTheader.Name = "BTheader"
        Me.BTheader.Size = New System.Drawing.Size(223, 32)
        Me.BTheader.TabIndex = 1
        Me.BTheader.Text = "Office_Utilities_Tools"
        Me.BTheader.UseVisualStyleBackColor = False
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.Button1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button1.FlatAppearance.BorderSize = 0
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(12, 51)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(223, 28)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "End Task Microsoft Exel"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'BTexit
        '
        Me.BTexit.BackgroundImage = Global.Office_Utilities_Tools.My.Resources.Resources.btexit
        Me.BTexit.Cursor = System.Windows.Forms.Cursors.Hand
        Me.BTexit.FlatAppearance.BorderSize = 0
        Me.BTexit.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BTexit.Location = New System.Drawing.Point(204, 14)
        Me.BTexit.Name = "BTexit"
        Me.BTexit.Size = New System.Drawing.Size(30, 30)
        Me.BTexit.TabIndex = 3
        Me.BTexit.UseVisualStyleBackColor = True
        '
        'BTSetting
        '
        Me.BTSetting.BackgroundImage = Global.Office_Utilities_Tools.My.Resources.Resources.utilities_32
        Me.BTSetting.Cursor = System.Windows.Forms.Cursors.Hand
        Me.BTSetting.FlatAppearance.BorderSize = 0
        Me.BTSetting.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BTSetting.Location = New System.Drawing.Point(12, 15)
        Me.BTSetting.Name = "BTSetting"
        Me.BTSetting.Size = New System.Drawing.Size(30, 30)
        Me.BTSetting.TabIndex = 3
        Me.BTSetting.UseVisualStyleBackColor = True
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.BalloonTipText = "Office_Utilities_Tools"
        Me.NotifyIcon1.BalloonTipTitle = "Office_Utilities_Tools"
        Me.NotifyIcon1.ContextMenuStrip = Me.ContextMenuStrip1
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        Me.NotifyIcon1.Text = "Office_Utilities_Tools"
        Me.NotifyIcon1.Visible = True
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.EndTaskMicrosoftExelToolStripMenuItem, Me.ToolStripSeparator1, Me.AddStartup, Me.ExitToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(198, 98)
        '
        'EndTaskMicrosoftExelToolStripMenuItem
        '
        Me.EndTaskMicrosoftExelToolStripMenuItem.Image = Global.Office_Utilities_Tools.My.Resources.Resources.Microsoft_Excel_Logo1
        Me.EndTaskMicrosoftExelToolStripMenuItem.Name = "EndTaskMicrosoftExelToolStripMenuItem"
        Me.EndTaskMicrosoftExelToolStripMenuItem.Size = New System.Drawing.Size(197, 22)
        Me.EndTaskMicrosoftExelToolStripMenuItem.Text = "End Task Microsoft Exel"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(194, 6)
        '
        'AddStartup
        '
        Me.AddStartup.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.AddStartup.CheckOnClick = True
        Me.AddStartup.Name = "AddStartup"
        Me.AddStartup.Size = New System.Drawing.Size(197, 22)
        Me.AddStartup.Text = "Launch On Startup"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Image = Global.Office_Utilities_Tools.My.Resources.Resources._exit
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(197, 22)
        Me.ExitToolStripMenuItem.Text = "exit"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Black
        Me.ClientSize = New System.Drawing.Size(248, 92)
        Me.ContextMenuStrip = Me.ContextMenuStrip1
        Me.Controls.Add(Me.BTSetting)
        Me.Controls.Add(Me.BTexit)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.BTheader)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form1"
        Me.ShowInTaskbar = False
        Me.Text = "Office_Utilities_Tools"
        Me.TopMost = True
        Me.TransparencyKey = System.Drawing.Color.Black
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents BTheader As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents BTexit As Button
    Friend WithEvents BTSetting As Button
    Friend WithEvents NotifyIcon1 As NotifyIcon
    Friend WithEvents ContextMenuStrip1 As ContextMenuStrip
    Friend WithEvents ExitToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents EndTaskMicrosoftExelToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents AddStartup As ToolStripMenuItem
End Class
