<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Main
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Main))
        Me.btnGetLegacyScores = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtActivityLog = New System.Windows.Forms.TextBox()
        Me.MainWebBrowser = New System.Windows.Forms.WebBrowser()
        Me.cbOnlyRunRelativeStrength = New System.Windows.Forms.CheckBox()
        Me.cbForceRS = New System.Windows.Forms.CheckBox()
        Me.cbOnlyRunPerpetualIncome = New System.Windows.Forms.CheckBox()
        Me.cbRuleOfSix = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'btnGetLegacyScores
        '
        Me.btnGetLegacyScores.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnGetLegacyScores.Location = New System.Drawing.Point(989, 139)
        Me.btnGetLegacyScores.Margin = New System.Windows.Forms.Padding(4)
        Me.btnGetLegacyScores.Name = "btnGetLegacyScores"
        Me.btnGetLegacyScores.Size = New System.Drawing.Size(112, 31)
        Me.btnGetLegacyScores.TabIndex = 15
        Me.btnGetLegacyScores.Text = "Get Data"
        Me.btnGetLegacyScores.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(1013, 11)
        Me.Label2.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(51, 20)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "Tasks"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(15, 11)
        Me.Label1.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(58, 20)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "Activity"
        '
        'txtActivityLog
        '
        Me.txtActivityLog.Location = New System.Drawing.Point(15, 37)
        Me.txtActivityLog.Margin = New System.Windows.Forms.Padding(6)
        Me.txtActivityLog.Multiline = True
        Me.txtActivityLog.Name = "txtActivityLog"
        Me.txtActivityLog.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtActivityLog.Size = New System.Drawing.Size(901, 183)
        Me.txtActivityLog.TabIndex = 12
        '
        'MainWebBrowser
        '
        Me.MainWebBrowser.Location = New System.Drawing.Point(13, 229)
        Me.MainWebBrowser.MinimumSize = New System.Drawing.Size(20, 20)
        Me.MainWebBrowser.Name = "MainWebBrowser"
        Me.MainWebBrowser.ScriptErrorsSuppressed = True
        Me.MainWebBrowser.Size = New System.Drawing.Size(1192, 581)
        Me.MainWebBrowser.TabIndex = 16
        '
        'cbOnlyRunRelativeStrength
        '
        Me.cbOnlyRunRelativeStrength.AutoSize = True
        Me.cbOnlyRunRelativeStrength.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbOnlyRunRelativeStrength.Location = New System.Drawing.Point(925, 37)
        Me.cbOnlyRunRelativeStrength.Name = "cbOnlyRunRelativeStrength"
        Me.cbOnlyRunRelativeStrength.Size = New System.Drawing.Size(223, 24)
        Me.cbOnlyRunRelativeStrength.TabIndex = 17
        Me.cbOnlyRunRelativeStrength.Text = "Only Run Relaitive Strength"
        Me.cbOnlyRunRelativeStrength.UseVisualStyleBackColor = True
        '
        'cbForceRS
        '
        Me.cbForceRS.AutoSize = True
        Me.cbForceRS.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbForceRS.Location = New System.Drawing.Point(924, 104)
        Me.cbForceRS.Name = "cbForceRS"
        Me.cbForceRS.Size = New System.Drawing.Size(281, 24)
        Me.cbForceRS.TabIndex = 18
        Me.cbForceRS.Text = "Force Running of Relaitive Strength"
        Me.cbForceRS.UseVisualStyleBackColor = True
        '
        'cbOnlyRunPerpetualIncome
        '
        Me.cbOnlyRunPerpetualIncome.AutoSize = True
        Me.cbOnlyRunPerpetualIncome.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbOnlyRunPerpetualIncome.Location = New System.Drawing.Point(924, 58)
        Me.cbOnlyRunPerpetualIncome.Name = "cbOnlyRunPerpetualIncome"
        Me.cbOnlyRunPerpetualIncome.Size = New System.Drawing.Size(222, 24)
        Me.cbOnlyRunPerpetualIncome.TabIndex = 19
        Me.cbOnlyRunPerpetualIncome.Text = "Only Run Perpetual Income"
        Me.cbOnlyRunPerpetualIncome.UseVisualStyleBackColor = True
        '
        'cbRuleOfSix
        '
        Me.cbRuleOfSix.AutoSize = True
        Me.cbRuleOfSix.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbRuleOfSix.Location = New System.Drawing.Point(924, 81)
        Me.cbRuleOfSix.Name = "cbRuleOfSix"
        Me.cbRuleOfSix.Size = New System.Drawing.Size(138, 24)
        Me.cbRuleOfSix.TabIndex = 20
        Me.cbRuleOfSix.Text = "Run Rule of Six"
        Me.cbRuleOfSix.UseVisualStyleBackColor = True
        '
        'Main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1221, 822)
        Me.Controls.Add(Me.cbRuleOfSix)
        Me.Controls.Add(Me.cbOnlyRunPerpetualIncome)
        Me.Controls.Add(Me.cbForceRS)
        Me.Controls.Add(Me.cbOnlyRunRelativeStrength)
        Me.Controls.Add(Me.MainWebBrowser)
        Me.Controls.Add(Me.btnGetLegacyScores)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtActivityLog)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "Main"
        Me.Text = "Blue Chip Authority"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnGetLegacyScores As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents txtActivityLog As TextBox
    Friend WithEvents MainWebBrowser As WebBrowser
    Friend WithEvents cbOnlyRunRelativeStrength As CheckBox
    Friend WithEvents cbForceRS As CheckBox
    Friend WithEvents cbOnlyRunPerpetualIncome As CheckBox
    Friend WithEvents cbRuleOfSix As CheckBox
End Class
