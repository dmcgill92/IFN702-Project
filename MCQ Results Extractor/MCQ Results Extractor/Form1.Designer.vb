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
		Me.btnContinue = New System.Windows.Forms.Button()
		Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
		Me.LaunchPanel = New System.Windows.Forms.Panel()
		Me.ComparisonPanel = New System.Windows.Forms.Panel()
		Me.tcLeft = New System.Windows.Forms.TabControl()
		Me.tabMatched1 = New System.Windows.Forms.TabPage()
		Me.tabUnmatched1 = New System.Windows.Forms.TabPage()
		Me.btnSave = New System.Windows.Forms.Button()
		Me.fileLocation3 = New System.Windows.Forms.TextBox()
		Me.fileLocation2 = New System.Windows.Forms.TextBox()
		Me.fileLocation1 = New System.Windows.Forms.TextBox()
		Me.lblOutput = New System.Windows.Forms.Label()
		Me.lblStudents = New System.Windows.Forms.Label()
		Me.lblResults = New System.Windows.Forms.Label()
		Me.btnBrowse3 = New System.Windows.Forms.Button()
		Me.btnBrowse2 = New System.Windows.Forms.Button()
		Me.btnBrowse1 = New System.Windows.Forms.Button()
		Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
		Me.tcRight = New System.Windows.Forms.TabControl()
		Me.tabMatched2 = New System.Windows.Forms.TabPage()
		Me.tabUnmatched2 = New System.Windows.Forms.TabPage()
		Me.dgMatchedLeft = New System.Windows.Forms.DataGridView()
		Me.dgMatchedRight = New System.Windows.Forms.DataGridView()
		Me.dgUnmatchedLeft = New System.Windows.Forms.DataGridView()
		Me.dgUnmatchedRight = New System.Windows.Forms.DataGridView()
		Me.btnMatch = New System.Windows.Forms.Button()
		Me.LaunchPanel.SuspendLayout()
		Me.ComparisonPanel.SuspendLayout()
		Me.tcLeft.SuspendLayout()
		Me.tabMatched1.SuspendLayout()
		Me.tabUnmatched1.SuspendLayout()
		Me.tcRight.SuspendLayout()
		Me.tabMatched2.SuspendLayout()
		Me.tabUnmatched2.SuspendLayout()
		CType(Me.dgMatchedLeft, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.dgMatchedRight, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.dgUnmatchedLeft, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.dgUnmatchedRight, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.SuspendLayout()
		'
		'btnContinue
		'
		Me.btnContinue.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.btnContinue.Enabled = False
		Me.btnContinue.Location = New System.Drawing.Point(670, 394)
		Me.btnContinue.Name = "btnContinue"
		Me.btnContinue.Size = New System.Drawing.Size(75, 23)
		Me.btnContinue.TabIndex = 3
		Me.btnContinue.Text = "Continue"
		Me.btnContinue.UseVisualStyleBackColor = True
		'
		'OpenFileDialog1
		'
		Me.OpenFileDialog1.FileName = "OpenFileDialog1"
		'
		'LaunchPanel
		'
		Me.LaunchPanel.Controls.Add(Me.fileLocation3)
		Me.LaunchPanel.Controls.Add(Me.fileLocation2)
		Me.LaunchPanel.Controls.Add(Me.fileLocation1)
		Me.LaunchPanel.Controls.Add(Me.lblOutput)
		Me.LaunchPanel.Controls.Add(Me.lblStudents)
		Me.LaunchPanel.Controls.Add(Me.lblResults)
		Me.LaunchPanel.Controls.Add(Me.btnBrowse3)
		Me.LaunchPanel.Controls.Add(Me.btnBrowse2)
		Me.LaunchPanel.Controls.Add(Me.btnBrowse1)
		Me.LaunchPanel.Controls.Add(Me.btnContinue)
		Me.LaunchPanel.Dock = System.Windows.Forms.DockStyle.Fill
		Me.LaunchPanel.Location = New System.Drawing.Point(0, 0)
		Me.LaunchPanel.Name = "LaunchPanel"
		Me.LaunchPanel.Size = New System.Drawing.Size(800, 450)
		Me.LaunchPanel.TabIndex = 15
		'
		'ComparisonPanel
		'
		Me.ComparisonPanel.Controls.Add(Me.tcRight)
		Me.ComparisonPanel.Controls.Add(Me.tcLeft)
		Me.ComparisonPanel.Controls.Add(Me.btnMatch)
		Me.ComparisonPanel.Controls.Add(Me.btnSave)
		Me.ComparisonPanel.Dock = System.Windows.Forms.DockStyle.Fill
		Me.ComparisonPanel.Location = New System.Drawing.Point(0, 0)
		Me.ComparisonPanel.Name = "ComparisonPanel"
		Me.ComparisonPanel.Size = New System.Drawing.Size(800, 450)
		Me.ComparisonPanel.TabIndex = 14
		Me.ComparisonPanel.Visible = False
		'
		'tcLeft
		'
		Me.tcLeft.Controls.Add(Me.tabMatched1)
		Me.tcLeft.Controls.Add(Me.tabUnmatched1)
		Me.tcLeft.Location = New System.Drawing.Point(31, 12)
		Me.tcLeft.Name = "tcLeft"
		Me.tcLeft.SelectedIndex = 0
		Me.tcLeft.Size = New System.Drawing.Size(357, 295)
		Me.tcLeft.TabIndex = 6
		'
		'tabMatched1
		'
		Me.tabMatched1.AutoScroll = True
		Me.tabMatched1.Controls.Add(Me.dgMatchedLeft)
		Me.tabMatched1.Location = New System.Drawing.Point(4, 22)
		Me.tabMatched1.Name = "tabMatched1"
		Me.tabMatched1.Padding = New System.Windows.Forms.Padding(3)
		Me.tabMatched1.Size = New System.Drawing.Size(349, 269)
		Me.tabMatched1.TabIndex = 0
		Me.tabMatched1.Text = "Matched"
		Me.tabMatched1.UseVisualStyleBackColor = True
		'
		'tabUnmatched1
		'
		Me.tabUnmatched1.Controls.Add(Me.dgUnmatchedLeft)
		Me.tabUnmatched1.Location = New System.Drawing.Point(4, 22)
		Me.tabUnmatched1.Name = "tabUnmatched1"
		Me.tabUnmatched1.Padding = New System.Windows.Forms.Padding(3)
		Me.tabUnmatched1.Size = New System.Drawing.Size(349, 269)
		Me.tabUnmatched1.TabIndex = 1
		Me.tabUnmatched1.Text = "Unmatched"
		Me.tabUnmatched1.UseVisualStyleBackColor = True
		'
		'btnSave
		'
		Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.btnSave.Enabled = False
		Me.btnSave.Location = New System.Drawing.Point(670, 394)
		Me.btnSave.Name = "btnSave"
		Me.btnSave.Size = New System.Drawing.Size(75, 23)
		Me.btnSave.TabIndex = 4
		Me.btnSave.Text = "Save"
		Me.btnSave.UseVisualStyleBackColor = True
		'
		'fileLocation3
		'
		Me.fileLocation3.Anchor = System.Windows.Forms.AnchorStyles.Top
		Me.fileLocation3.Location = New System.Drawing.Point(220, 302)
		Me.fileLocation3.Name = "fileLocation3"
		Me.fileLocation3.ReadOnly = True
		Me.fileLocation3.Size = New System.Drawing.Size(358, 20)
		Me.fileLocation3.TabIndex = 13
		'
		'fileLocation2
		'
		Me.fileLocation2.Anchor = System.Windows.Forms.AnchorStyles.Top
		Me.fileLocation2.Location = New System.Drawing.Point(220, 192)
		Me.fileLocation2.Name = "fileLocation2"
		Me.fileLocation2.ReadOnly = True
		Me.fileLocation2.Size = New System.Drawing.Size(358, 20)
		Me.fileLocation2.TabIndex = 13
		'
		'fileLocation1
		'
		Me.fileLocation1.Anchor = System.Windows.Forms.AnchorStyles.Top
		Me.fileLocation1.Location = New System.Drawing.Point(220, 75)
		Me.fileLocation1.Name = "fileLocation1"
		Me.fileLocation1.ReadOnly = True
		Me.fileLocation1.Size = New System.Drawing.Size(358, 20)
		Me.fileLocation1.TabIndex = 13
		'
		'lblOutput
		'
		Me.lblOutput.Anchor = System.Windows.Forms.AnchorStyles.Top
		Me.lblOutput.AutoSize = True
		Me.lblOutput.Location = New System.Drawing.Point(351, 276)
		Me.lblOutput.Name = "lblOutput"
		Me.lblOutput.Size = New System.Drawing.Size(98, 13)
		Me.lblOutput.TabIndex = 12
		Me.lblOutput.Text = "Output file location:"
		Me.lblOutput.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		'
		'lblStudents
		'
		Me.lblStudents.Anchor = System.Windows.Forms.AnchorStyles.Top
		Me.lblStudents.Location = New System.Drawing.Point(339, 163)
		Me.lblStudents.Name = "lblStudents"
		Me.lblStudents.Size = New System.Drawing.Size(118, 13)
		Me.lblStudents.TabIndex = 12
		Me.lblStudents.Text = "Student name file location:"
		Me.lblStudents.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		'
		'lblResults
		'
		Me.lblResults.Anchor = System.Windows.Forms.AnchorStyles.Top
		Me.lblResults.AutoSize = True
		Me.lblResults.Location = New System.Drawing.Point(351, 49)
		Me.lblResults.Name = "lblResults"
		Me.lblResults.Size = New System.Drawing.Size(96, 13)
		Me.lblResults.TabIndex = 12
		Me.lblResults.Text = "Result file location:"
		Me.lblResults.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		'
		'btnBrowse3
		'
		Me.btnBrowse3.Anchor = System.Windows.Forms.AnchorStyles.Top
		Me.btnBrowse3.Location = New System.Drawing.Point(362, 328)
		Me.btnBrowse3.Name = "btnBrowse3"
		Me.btnBrowse3.Size = New System.Drawing.Size(75, 23)
		Me.btnBrowse3.TabIndex = 11
		Me.btnBrowse3.Text = "Browse..."
		Me.btnBrowse3.UseVisualStyleBackColor = True
		'
		'btnBrowse2
		'
		Me.btnBrowse2.Anchor = System.Windows.Forms.AnchorStyles.Top
		Me.btnBrowse2.Location = New System.Drawing.Point(362, 218)
		Me.btnBrowse2.Name = "btnBrowse2"
		Me.btnBrowse2.Size = New System.Drawing.Size(75, 23)
		Me.btnBrowse2.TabIndex = 11
		Me.btnBrowse2.Text = "Browse..."
		Me.btnBrowse2.UseVisualStyleBackColor = True
		'
		'btnBrowse1
		'
		Me.btnBrowse1.Anchor = System.Windows.Forms.AnchorStyles.Top
		Me.btnBrowse1.Location = New System.Drawing.Point(362, 101)
		Me.btnBrowse1.Name = "btnBrowse1"
		Me.btnBrowse1.Size = New System.Drawing.Size(75, 23)
		Me.btnBrowse1.TabIndex = 11
		Me.btnBrowse1.Text = "Browse..."
		Me.btnBrowse1.UseVisualStyleBackColor = True
		'
		'tcRight
		'
		Me.tcRight.Controls.Add(Me.tabMatched2)
		Me.tcRight.Controls.Add(Me.tabUnmatched2)
		Me.tcRight.Location = New System.Drawing.Point(411, 12)
		Me.tcRight.Name = "tcRight"
		Me.tcRight.SelectedIndex = 0
		Me.tcRight.Size = New System.Drawing.Size(357, 295)
		Me.tcRight.TabIndex = 6
		'
		'tabMatched2
		'
		Me.tabMatched2.AutoScroll = True
		Me.tabMatched2.Controls.Add(Me.dgMatchedRight)
		Me.tabMatched2.Location = New System.Drawing.Point(4, 22)
		Me.tabMatched2.Name = "tabMatched2"
		Me.tabMatched2.Padding = New System.Windows.Forms.Padding(3)
		Me.tabMatched2.Size = New System.Drawing.Size(349, 269)
		Me.tabMatched2.TabIndex = 0
		Me.tabMatched2.Text = "Matched"
		Me.tabMatched2.UseVisualStyleBackColor = True
		'
		'tabUnmatched2
		'
		Me.tabUnmatched2.Controls.Add(Me.dgUnmatchedRight)
		Me.tabUnmatched2.Location = New System.Drawing.Point(4, 22)
		Me.tabUnmatched2.Name = "tabUnmatched2"
		Me.tabUnmatched2.Padding = New System.Windows.Forms.Padding(3)
		Me.tabUnmatched2.Size = New System.Drawing.Size(349, 269)
		Me.tabUnmatched2.TabIndex = 1
		Me.tabUnmatched2.Text = "Unmatched"
		Me.tabUnmatched2.UseVisualStyleBackColor = True
		'
		'dgMatchedLeft
		'
		Me.dgMatchedLeft.AllowUserToAddRows = False
		Me.dgMatchedLeft.AllowUserToDeleteRows = False
		Me.dgMatchedLeft.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		Me.dgMatchedLeft.Dock = System.Windows.Forms.DockStyle.Fill
		Me.dgMatchedLeft.Location = New System.Drawing.Point(3, 3)
		Me.dgMatchedLeft.Name = "dgMatchedLeft"
		Me.dgMatchedLeft.Size = New System.Drawing.Size(343, 263)
		Me.dgMatchedLeft.TabIndex = 0
		'
		'dgMatchedRight
		'
		Me.dgMatchedRight.AllowUserToAddRows = False
		Me.dgMatchedRight.AllowUserToDeleteRows = False
		Me.dgMatchedRight.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		Me.dgMatchedRight.Dock = System.Windows.Forms.DockStyle.Fill
		Me.dgMatchedRight.Location = New System.Drawing.Point(3, 3)
		Me.dgMatchedRight.Name = "dgMatchedRight"
		Me.dgMatchedRight.Size = New System.Drawing.Size(343, 263)
		Me.dgMatchedRight.TabIndex = 0
		'
		'dgUnmatchedLeft
		'
		Me.dgUnmatchedLeft.AllowUserToAddRows = False
		Me.dgUnmatchedLeft.AllowUserToDeleteRows = False
		Me.dgUnmatchedLeft.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		Me.dgUnmatchedLeft.Dock = System.Windows.Forms.DockStyle.Fill
		Me.dgUnmatchedLeft.Location = New System.Drawing.Point(3, 3)
		Me.dgUnmatchedLeft.Name = "dgUnmatchedLeft"
		Me.dgUnmatchedLeft.Size = New System.Drawing.Size(343, 263)
		Me.dgUnmatchedLeft.TabIndex = 0
		'
		'dgUnmatchedRight
		'
		Me.dgUnmatchedRight.AllowUserToAddRows = False
		Me.dgUnmatchedRight.AllowUserToDeleteRows = False
		Me.dgUnmatchedRight.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		Me.dgUnmatchedRight.Dock = System.Windows.Forms.DockStyle.Fill
		Me.dgUnmatchedRight.Location = New System.Drawing.Point(3, 3)
		Me.dgUnmatchedRight.Name = "dgUnmatchedRight"
		Me.dgUnmatchedRight.ReadOnly = True
		Me.dgUnmatchedRight.Size = New System.Drawing.Size(343, 263)
		Me.dgUnmatchedRight.TabIndex = 0
		'
		'btnMatch
		'
		Me.btnMatch.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.btnMatch.Location = New System.Drawing.Point(362, 328)
		Me.btnMatch.Name = "btnMatch"
		Me.btnMatch.Size = New System.Drawing.Size(75, 23)
		Me.btnMatch.TabIndex = 4
		Me.btnMatch.Text = "Match"
		Me.btnMatch.UseVisualStyleBackColor = True
		'
		'Form1
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(800, 450)
		Me.Controls.Add(Me.ComparisonPanel)
		Me.Controls.Add(Me.LaunchPanel)
		Me.Name = "Form1"
		Me.ShowIcon = False
		Me.Text = "MCQ Results Extractor"
		Me.LaunchPanel.ResumeLayout(False)
		Me.LaunchPanel.PerformLayout()
		Me.ComparisonPanel.ResumeLayout(False)
		Me.tcLeft.ResumeLayout(False)
		Me.tabMatched1.ResumeLayout(False)
		Me.tabUnmatched1.ResumeLayout(False)
		Me.tcRight.ResumeLayout(False)
		Me.tabMatched2.ResumeLayout(False)
		Me.tabUnmatched2.ResumeLayout(False)
		CType(Me.dgMatchedLeft, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.dgMatchedRight, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.dgUnmatchedLeft, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.dgUnmatchedRight, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)

	End Sub
	Friend WithEvents btnContinue As Button
	Friend WithEvents OpenFileDialog1 As OpenFileDialog
	Friend WithEvents LaunchPanel As Panel
	Friend WithEvents fileLocation1 As TextBox
	Friend WithEvents btnBrowse1 As Button
	Friend WithEvents fileLocation3 As TextBox
	Friend WithEvents fileLocation2 As TextBox
	Friend WithEvents lblOutput As Label
	Friend WithEvents lblStudents As Label
	Friend WithEvents lblResults As Label
	Friend WithEvents btnBrowse3 As Button
	Friend WithEvents btnBrowse2 As Button
	Friend WithEvents ComparisonPanel As Panel
	Friend WithEvents tcLeft As TabControl
	Friend WithEvents tabMatched1 As TabPage
	Friend WithEvents tabUnmatched1 As TabPage
	Friend WithEvents btnSave As Button
	Friend WithEvents SaveFileDialog1 As SaveFileDialog
	Friend WithEvents tcRight As TabControl
	Friend WithEvents tabMatched2 As TabPage
	Friend WithEvents tabUnmatched2 As TabPage
	Friend WithEvents dgMatchedRight As DataGridView
	Friend WithEvents dgUnmatchedRight As DataGridView
	Friend WithEvents dgMatchedLeft As DataGridView
	Friend WithEvents dgUnmatchedLeft As DataGridView
	Friend WithEvents btnMatch As Button
End Class
