﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
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
		Me.pnlHeaderSelection = New System.Windows.Forms.Panel()
		Me.grpBox = New System.Windows.Forms.GroupBox()
		Me.radbtn1 = New System.Windows.Forms.RadioButton()
		Me.lblHeaderSelection = New System.Windows.Forms.Label()
		Me.fileLocation2 = New System.Windows.Forms.TextBox()
		Me.fileLocation1 = New System.Windows.Forms.TextBox()
		Me.lblStudents = New System.Windows.Forms.Label()
		Me.lblResults = New System.Windows.Forms.Label()
		Me.btnBrowse2 = New System.Windows.Forms.Button()
		Me.btnBrowse1 = New System.Windows.Forms.Button()
		Me.ComparisonPanel = New System.Windows.Forms.Panel()
		Me.lblProg = New System.Windows.Forms.Label()
		Me.progBar1 = New System.Windows.Forms.ProgressBar()
		Me.btnConfirm = New System.Windows.Forms.Button()
		Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
		Me.tcLeft = New System.Windows.Forms.TabControl()
		Me.tabMatched1 = New System.Windows.Forms.TabPage()
		Me.dgMatchedLeft = New System.Windows.Forms.DataGridView()
		Me.tabUnmatched1 = New System.Windows.Forms.TabPage()
		Me.dgUnmatchedLeft = New System.Windows.Forms.DataGridView()
		Me.tcRight = New System.Windows.Forms.TabControl()
		Me.tabMatched2 = New System.Windows.Forms.TabPage()
		Me.dgMatchedRight = New System.Windows.Forms.DataGridView()
		Me.tabUnmatched2 = New System.Windows.Forms.TabPage()
		Me.dgUnmatchedRight = New System.Windows.Forms.DataGridView()
		Me.lblMatched = New System.Windows.Forms.Label()
		Me.btnMatch = New System.Windows.Forms.Button()
		Me.btnBack = New System.Windows.Forms.Button()
		Me.btnSave = New System.Windows.Forms.Button()
		Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
		Me.bgwrkMatching = New System.ComponentModel.BackgroundWorker()
		Me.bgwrkDataHandler = New System.ComponentModel.BackgroundWorker()
		Me.bgwrkPartialMatch = New System.ComponentModel.BackgroundWorker()
		Me.LaunchPanel.SuspendLayout()
		Me.pnlHeaderSelection.SuspendLayout()
		Me.grpBox.SuspendLayout()
		Me.ComparisonPanel.SuspendLayout()
		Me.TableLayoutPanel1.SuspendLayout()
		Me.tcLeft.SuspendLayout()
		Me.tabMatched1.SuspendLayout()
		CType(Me.dgMatchedLeft, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.tabUnmatched1.SuspendLayout()
		CType(Me.dgUnmatchedLeft, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.tcRight.SuspendLayout()
		Me.tabMatched2.SuspendLayout()
		CType(Me.dgMatchedRight, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.tabUnmatched2.SuspendLayout()
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
		Me.LaunchPanel.Controls.Add(Me.pnlHeaderSelection)
		Me.LaunchPanel.Controls.Add(Me.fileLocation2)
		Me.LaunchPanel.Controls.Add(Me.fileLocation1)
		Me.LaunchPanel.Controls.Add(Me.lblStudents)
		Me.LaunchPanel.Controls.Add(Me.lblResults)
		Me.LaunchPanel.Controls.Add(Me.btnBrowse2)
		Me.LaunchPanel.Controls.Add(Me.btnBrowse1)
		Me.LaunchPanel.Controls.Add(Me.btnContinue)
		Me.LaunchPanel.Dock = System.Windows.Forms.DockStyle.Fill
		Me.LaunchPanel.Location = New System.Drawing.Point(0, 0)
		Me.LaunchPanel.Name = "LaunchPanel"
		Me.LaunchPanel.Size = New System.Drawing.Size(800, 450)
		Me.LaunchPanel.TabIndex = 15
		'
		'pnlHeaderSelection
		'
		Me.pnlHeaderSelection.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom), System.Windows.Forms.AnchorStyles)
		Me.pnlHeaderSelection.Controls.Add(Me.grpBox)
		Me.pnlHeaderSelection.Controls.Add(Me.lblHeaderSelection)
		Me.pnlHeaderSelection.Location = New System.Drawing.Point(228, 267)
		Me.pnlHeaderSelection.Name = "pnlHeaderSelection"
		Me.pnlHeaderSelection.Size = New System.Drawing.Size(345, 150)
		Me.pnlHeaderSelection.TabIndex = 16
		Me.pnlHeaderSelection.Visible = False
		'
		'grpBox
		'
		Me.grpBox.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom), System.Windows.Forms.AnchorStyles)
		Me.grpBox.AutoSize = True
		Me.grpBox.Controls.Add(Me.radbtn1)
		Me.grpBox.Location = New System.Drawing.Point(3, 29)
		Me.grpBox.Name = "grpBox"
		Me.grpBox.Padding = New System.Windows.Forms.Padding(1)
		Me.grpBox.Size = New System.Drawing.Size(339, 51)
		Me.grpBox.TabIndex = 16
		Me.grpBox.TabStop = False
		'
		'radbtn1
		'
		Me.radbtn1.Anchor = System.Windows.Forms.AnchorStyles.Top
		Me.radbtn1.AutoSize = True
		Me.radbtn1.Location = New System.Drawing.Point(57, 17)
		Me.radbtn1.Name = "radbtn1"
		Me.radbtn1.Size = New System.Drawing.Size(225, 17)
		Me.radbtn1.TabIndex = 14
		Me.radbtn1.TabStop = True
		Me.radbtn1.Text = "Exam [Total Pts: 100 Percentage] |823157"
		Me.radbtn1.UseVisualStyleBackColor = True
		'
		'lblHeaderSelection
		'
		Me.lblHeaderSelection.Anchor = System.Windows.Forms.AnchorStyles.Top
		Me.lblHeaderSelection.AutoSize = True
		Me.lblHeaderSelection.Location = New System.Drawing.Point(100, 11)
		Me.lblHeaderSelection.Name = "lblHeaderSelection"
		Me.lblHeaderSelection.Size = New System.Drawing.Size(145, 13)
		Me.lblHeaderSelection.TabIndex = 15
		Me.lblHeaderSelection.Text = "Select correct results header:"
		Me.lblHeaderSelection.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
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
		'lblStudents
		'
		Me.lblStudents.Anchor = System.Windows.Forms.AnchorStyles.Top
		Me.lblStudents.AutoSize = True
		Me.lblStudents.Location = New System.Drawing.Point(334, 163)
		Me.lblStudents.Name = "lblStudents"
		Me.lblStudents.Size = New System.Drawing.Size(132, 13)
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
		'ComparisonPanel
		'
		Me.ComparisonPanel.Controls.Add(Me.lblProg)
		Me.ComparisonPanel.Controls.Add(Me.progBar1)
		Me.ComparisonPanel.Controls.Add(Me.btnConfirm)
		Me.ComparisonPanel.Controls.Add(Me.TableLayoutPanel1)
		Me.ComparisonPanel.Controls.Add(Me.lblMatched)
		Me.ComparisonPanel.Controls.Add(Me.btnMatch)
		Me.ComparisonPanel.Controls.Add(Me.btnBack)
		Me.ComparisonPanel.Controls.Add(Me.btnSave)
		Me.ComparisonPanel.Dock = System.Windows.Forms.DockStyle.Fill
		Me.ComparisonPanel.Location = New System.Drawing.Point(0, 0)
		Me.ComparisonPanel.Name = "ComparisonPanel"
		Me.ComparisonPanel.Size = New System.Drawing.Size(800, 450)
		Me.ComparisonPanel.TabIndex = 14
		Me.ComparisonPanel.Visible = False
		'
		'lblProg
		'
		Me.lblProg.Anchor = System.Windows.Forms.AnchorStyles.Bottom
		Me.lblProg.Location = New System.Drawing.Point(325, 399)
		Me.lblProg.Name = "lblProg"
		Me.lblProg.Size = New System.Drawing.Size(150, 13)
		Me.lblProg.TabIndex = 11
		Me.lblProg.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.lblProg.Visible = False
		'
		'progBar1
		'
		Me.progBar1.Anchor = System.Windows.Forms.AnchorStyles.Top
		Me.progBar1.Location = New System.Drawing.Point(280, 370)
		Me.progBar1.Name = "progBar1"
		Me.progBar1.Size = New System.Drawing.Size(240, 20)
		Me.progBar1.TabIndex = 10
		Me.progBar1.Visible = False
		'
		'btnConfirm
		'
		Me.btnConfirm.Anchor = System.Windows.Forms.AnchorStyles.Bottom
		Me.btnConfirm.Location = New System.Drawing.Point(358, 328)
		Me.btnConfirm.Name = "btnConfirm"
		Me.btnConfirm.Size = New System.Drawing.Size(85, 23)
		Me.btnConfirm.TabIndex = 9
		Me.btnConfirm.Text = "Confirm Match"
		Me.btnConfirm.UseVisualStyleBackColor = True
		Me.btnConfirm.Visible = False
		'
		'TableLayoutPanel1
		'
		Me.TableLayoutPanel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
			Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.TableLayoutPanel1.ColumnCount = 2
		Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
		Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
		Me.TableLayoutPanel1.Controls.Add(Me.tcLeft, 0, 0)
		Me.TableLayoutPanel1.Controls.Add(Me.tcRight, 1, 0)
		Me.TableLayoutPanel1.Location = New System.Drawing.Point(25, 22)
		Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
		Me.TableLayoutPanel1.RowCount = 1
		Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
		Me.TableLayoutPanel1.Size = New System.Drawing.Size(750, 300)
		Me.TableLayoutPanel1.TabIndex = 8
		'
		'tcLeft
		'
		Me.tcLeft.Controls.Add(Me.tabMatched1)
		Me.tcLeft.Controls.Add(Me.tabUnmatched1)
		Me.tcLeft.Dock = System.Windows.Forms.DockStyle.Fill
		Me.tcLeft.Location = New System.Drawing.Point(3, 3)
		Me.tcLeft.Name = "tcLeft"
		Me.tcLeft.SelectedIndex = 0
		Me.tcLeft.Size = New System.Drawing.Size(369, 294)
		Me.tcLeft.TabIndex = 6
		'
		'tabMatched1
		'
		Me.tabMatched1.AutoScroll = True
		Me.tabMatched1.Controls.Add(Me.dgMatchedLeft)
		Me.tabMatched1.Location = New System.Drawing.Point(4, 22)
		Me.tabMatched1.Name = "tabMatched1"
		Me.tabMatched1.Padding = New System.Windows.Forms.Padding(3)
		Me.tabMatched1.Size = New System.Drawing.Size(361, 268)
		Me.tabMatched1.TabIndex = 0
		Me.tabMatched1.Text = "Matched"
		Me.tabMatched1.UseVisualStyleBackColor = True
		'
		'dgMatchedLeft
		'
		Me.dgMatchedLeft.AllowUserToAddRows = False
		Me.dgMatchedLeft.AllowUserToDeleteRows = False
		Me.dgMatchedLeft.AllowUserToResizeColumns = False
		Me.dgMatchedLeft.AllowUserToResizeRows = False
		Me.dgMatchedLeft.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
		Me.dgMatchedLeft.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		Me.dgMatchedLeft.Dock = System.Windows.Forms.DockStyle.Fill
		Me.dgMatchedLeft.Location = New System.Drawing.Point(3, 3)
		Me.dgMatchedLeft.MultiSelect = False
		Me.dgMatchedLeft.Name = "dgMatchedLeft"
		Me.dgMatchedLeft.RowHeadersVisible = False
		Me.dgMatchedLeft.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
		Me.dgMatchedLeft.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
		Me.dgMatchedLeft.Size = New System.Drawing.Size(355, 262)
		Me.dgMatchedLeft.TabIndex = 0
		'
		'tabUnmatched1
		'
		Me.tabUnmatched1.Controls.Add(Me.dgUnmatchedLeft)
		Me.tabUnmatched1.Location = New System.Drawing.Point(4, 22)
		Me.tabUnmatched1.Name = "tabUnmatched1"
		Me.tabUnmatched1.Padding = New System.Windows.Forms.Padding(3)
		Me.tabUnmatched1.Size = New System.Drawing.Size(361, 268)
		Me.tabUnmatched1.TabIndex = 1
		Me.tabUnmatched1.Text = "Unmatched"
		Me.tabUnmatched1.UseVisualStyleBackColor = True
		'
		'dgUnmatchedLeft
		'
		Me.dgUnmatchedLeft.AllowUserToAddRows = False
		Me.dgUnmatchedLeft.AllowUserToDeleteRows = False
		Me.dgUnmatchedLeft.AllowUserToResizeColumns = False
		Me.dgUnmatchedLeft.AllowUserToResizeRows = False
		Me.dgUnmatchedLeft.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
		Me.dgUnmatchedLeft.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		Me.dgUnmatchedLeft.Dock = System.Windows.Forms.DockStyle.Fill
		Me.dgUnmatchedLeft.Location = New System.Drawing.Point(3, 3)
		Me.dgUnmatchedLeft.MultiSelect = False
		Me.dgUnmatchedLeft.Name = "dgUnmatchedLeft"
		Me.dgUnmatchedLeft.RowHeadersVisible = False
		Me.dgUnmatchedLeft.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
		Me.dgUnmatchedLeft.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
		Me.dgUnmatchedLeft.Size = New System.Drawing.Size(355, 262)
		Me.dgUnmatchedLeft.TabIndex = 0
		'
		'tcRight
		'
		Me.tcRight.Controls.Add(Me.tabMatched2)
		Me.tcRight.Controls.Add(Me.tabUnmatched2)
		Me.tcRight.Dock = System.Windows.Forms.DockStyle.Fill
		Me.tcRight.Location = New System.Drawing.Point(378, 3)
		Me.tcRight.Name = "tcRight"
		Me.tcRight.SelectedIndex = 0
		Me.tcRight.Size = New System.Drawing.Size(369, 294)
		Me.tcRight.TabIndex = 6
		'
		'tabMatched2
		'
		Me.tabMatched2.AutoScroll = True
		Me.tabMatched2.Controls.Add(Me.dgMatchedRight)
		Me.tabMatched2.Location = New System.Drawing.Point(4, 22)
		Me.tabMatched2.Name = "tabMatched2"
		Me.tabMatched2.Padding = New System.Windows.Forms.Padding(3)
		Me.tabMatched2.Size = New System.Drawing.Size(361, 268)
		Me.tabMatched2.TabIndex = 0
		Me.tabMatched2.Text = "Matched"
		Me.tabMatched2.UseVisualStyleBackColor = True
		'
		'dgMatchedRight
		'
		Me.dgMatchedRight.AllowUserToAddRows = False
		Me.dgMatchedRight.AllowUserToDeleteRows = False
		Me.dgMatchedRight.AllowUserToResizeColumns = False
		Me.dgMatchedRight.AllowUserToResizeRows = False
		Me.dgMatchedRight.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
		Me.dgMatchedRight.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		Me.dgMatchedRight.Dock = System.Windows.Forms.DockStyle.Fill
		Me.dgMatchedRight.Location = New System.Drawing.Point(3, 3)
		Me.dgMatchedRight.MultiSelect = False
		Me.dgMatchedRight.Name = "dgMatchedRight"
		Me.dgMatchedRight.RowHeadersVisible = False
		Me.dgMatchedRight.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
		Me.dgMatchedRight.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
		Me.dgMatchedRight.Size = New System.Drawing.Size(355, 262)
		Me.dgMatchedRight.TabIndex = 0
		'
		'tabUnmatched2
		'
		Me.tabUnmatched2.Controls.Add(Me.dgUnmatchedRight)
		Me.tabUnmatched2.Location = New System.Drawing.Point(4, 22)
		Me.tabUnmatched2.Name = "tabUnmatched2"
		Me.tabUnmatched2.Padding = New System.Windows.Forms.Padding(3)
		Me.tabUnmatched2.Size = New System.Drawing.Size(361, 268)
		Me.tabUnmatched2.TabIndex = 1
		Me.tabUnmatched2.Text = "Unmatched"
		Me.tabUnmatched2.UseVisualStyleBackColor = True
		'
		'dgUnmatchedRight
		'
		Me.dgUnmatchedRight.AllowUserToAddRows = False
		Me.dgUnmatchedRight.AllowUserToDeleteRows = False
		Me.dgUnmatchedRight.AllowUserToResizeColumns = False
		Me.dgUnmatchedRight.AllowUserToResizeRows = False
		Me.dgUnmatchedRight.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
		Me.dgUnmatchedRight.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		Me.dgUnmatchedRight.Dock = System.Windows.Forms.DockStyle.Fill
		Me.dgUnmatchedRight.Location = New System.Drawing.Point(3, 3)
		Me.dgUnmatchedRight.MultiSelect = False
		Me.dgUnmatchedRight.Name = "dgUnmatchedRight"
		Me.dgUnmatchedRight.ReadOnly = True
		Me.dgUnmatchedRight.RowHeadersVisible = False
		Me.dgUnmatchedRight.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
		Me.dgUnmatchedRight.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
		Me.dgUnmatchedRight.Size = New System.Drawing.Size(355, 262)
		Me.dgUnmatchedRight.TabIndex = 0
		'
		'lblMatched
		'
		Me.lblMatched.Anchor = System.Windows.Forms.AnchorStyles.Bottom
		Me.lblMatched.Location = New System.Drawing.Point(325, 354)
		Me.lblMatched.Name = "lblMatched"
		Me.lblMatched.Size = New System.Drawing.Size(150, 13)
		Me.lblMatched.TabIndex = 7
		Me.lblMatched.Text = "Matched: 0/0"
		Me.lblMatched.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.lblMatched.Visible = False
		'
		'btnMatch
		'
		Me.btnMatch.Anchor = System.Windows.Forms.AnchorStyles.Bottom
		Me.btnMatch.Location = New System.Drawing.Point(362, 328)
		Me.btnMatch.Name = "btnMatch"
		Me.btnMatch.Size = New System.Drawing.Size(75, 23)
		Me.btnMatch.TabIndex = 4
		Me.btnMatch.Text = "Match"
		Me.btnMatch.UseVisualStyleBackColor = True
		'
		'btnBack
		'
		Me.btnBack.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
		Me.btnBack.Location = New System.Drawing.Point(53, 394)
		Me.btnBack.Name = "btnBack"
		Me.btnBack.Size = New System.Drawing.Size(92, 23)
		Me.btnBack.TabIndex = 4
		Me.btnBack.Text = "Back"
		Me.btnBack.UseVisualStyleBackColor = True
		'
		'btnSave
		'
		Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.btnSave.Location = New System.Drawing.Point(653, 394)
		Me.btnSave.Name = "btnSave"
		Me.btnSave.Size = New System.Drawing.Size(92, 23)
		Me.btnSave.TabIndex = 4
		Me.btnSave.Text = "Save and Close"
		Me.btnSave.UseVisualStyleBackColor = True
		Me.btnSave.Visible = False
		'
		'bgwrkMatching
		'
		Me.bgwrkMatching.WorkerReportsProgress = True
		'
		'bgwrkDataHandler
		'
		Me.bgwrkDataHandler.WorkerReportsProgress = True
		'
		'bgwrkPartialMatch
		'
		Me.bgwrkPartialMatch.WorkerReportsProgress = True
		'
		'Form1
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(800, 450)
		Me.Controls.Add(Me.ComparisonPanel)
		Me.Controls.Add(Me.LaunchPanel)
		Me.MinimumSize = New System.Drawing.Size(800, 400)
		Me.Name = "Form1"
		Me.ShowIcon = False
		Me.Text = "MCQ Results Extractor"
		Me.LaunchPanel.ResumeLayout(False)
		Me.LaunchPanel.PerformLayout()
		Me.pnlHeaderSelection.ResumeLayout(False)
		Me.pnlHeaderSelection.PerformLayout()
		Me.grpBox.ResumeLayout(False)
		Me.grpBox.PerformLayout()
		Me.ComparisonPanel.ResumeLayout(False)
		Me.TableLayoutPanel1.ResumeLayout(False)
		Me.tcLeft.ResumeLayout(False)
		Me.tabMatched1.ResumeLayout(False)
		CType(Me.dgMatchedLeft, System.ComponentModel.ISupportInitialize).EndInit()
		Me.tabUnmatched1.ResumeLayout(False)
		CType(Me.dgUnmatchedLeft, System.ComponentModel.ISupportInitialize).EndInit()
		Me.tcRight.ResumeLayout(False)
		Me.tabMatched2.ResumeLayout(False)
		CType(Me.dgMatchedRight, System.ComponentModel.ISupportInitialize).EndInit()
		Me.tabUnmatched2.ResumeLayout(False)
		CType(Me.dgUnmatchedRight, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)

	End Sub
	Friend WithEvents btnContinue As Button
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents LaunchPanel As Panel
    Friend WithEvents fileLocation1 As TextBox
    Friend WithEvents btnBrowse1 As Button
    Friend WithEvents fileLocation2 As TextBox
    Friend WithEvents lblStudents As Label
    Friend WithEvents lblResults As Label
    Friend WithEvents btnBrowse2 As Button
    Friend WithEvents ComparisonPanel As Panel
    Friend WithEvents btnSave As Button
    Friend WithEvents SaveFileDialog1 As SaveFileDialog
    Friend WithEvents btnMatch As Button
    Friend WithEvents lblMatched As Label
    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents tcLeft As TabControl
    Friend WithEvents tabMatched1 As TabPage
    Friend WithEvents dgMatchedLeft As DataGridView
    Friend WithEvents tabUnmatched1 As TabPage
    Friend WithEvents dgUnmatchedLeft As DataGridView
    Friend WithEvents tcRight As TabControl
    Friend WithEvents tabMatched2 As TabPage
    Friend WithEvents dgMatchedRight As DataGridView
    Friend WithEvents tabUnmatched2 As TabPage
    Friend WithEvents dgUnmatchedRight As DataGridView
    Friend WithEvents btnConfirm As Button
    Friend WithEvents bgwrkMatching As System.ComponentModel.BackgroundWorker
    Friend WithEvents bgwrkDataHandler As System.ComponentModel.BackgroundWorker
    Friend WithEvents progBar1 As ProgressBar
    Friend WithEvents lblProg As Label
	Friend WithEvents bgwrkPartialMatch As System.ComponentModel.BackgroundWorker
	Friend WithEvents pnlHeaderSelection As Panel
	Friend WithEvents radbtn1 As RadioButton
	Friend WithEvents lblHeaderSelection As Label
	Friend WithEvents grpBox As GroupBox
	Friend WithEvents btnBack As Button
End Class
