<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
		Me.btnBrowse1 = New System.Windows.Forms.Button()
		Me.btnContinue = New System.Windows.Forms.Button()
		Me.Label1 = New System.Windows.Forms.Label()
		Me.TextBox1 = New System.Windows.Forms.TextBox()
		Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
		Me.SuspendLayout()
		'
		'btnBrowse1
		'
		Me.btnBrowse1.Location = New System.Drawing.Point(371, 105)
		Me.btnBrowse1.Name = "btnBrowse1"
		Me.btnBrowse1.Size = New System.Drawing.Size(75, 23)
		Me.btnBrowse1.TabIndex = 0
		Me.btnBrowse1.Text = "Browse..."
		Me.btnBrowse1.UseVisualStyleBackColor = True
		'
		'btnContinue
		'
		Me.btnContinue.Location = New System.Drawing.Point(676, 396)
		Me.btnContinue.Name = "btnContinue"
		Me.btnContinue.Size = New System.Drawing.Size(75, 23)
		Me.btnContinue.TabIndex = 3
		Me.btnContinue.Text = "Continue"
		Me.btnContinue.UseVisualStyleBackColor = True
		'
		'Label1
		'
		Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.None
		Me.Label1.AutoSize = True
		Me.Label1.Location = New System.Drawing.Point(360, 48)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(96, 13)
		Me.Label1.TabIndex = 4
		Me.Label1.Text = "Result file location:"
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		'
		'TextBox1
		'
		Me.TextBox1.Location = New System.Drawing.Point(229, 73)
		Me.TextBox1.Name = "TextBox1"
		Me.TextBox1.Size = New System.Drawing.Size(358, 20)
		Me.TextBox1.TabIndex = 7
		'
		'OpenFileDialog1
		'
		Me.OpenFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
		'
		'Form1
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(800, 450)
		Me.Controls.Add(Me.TextBox1)
		Me.Controls.Add(Me.Label1)
		Me.Controls.Add(Me.btnContinue)
		Me.Controls.Add(Me.btnBrowse1)
		Me.Name = "Form1"
		Me.Text = "Form1"
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub

	Friend WithEvents btnBrowse1 As Button
	Friend WithEvents btnContinue As Button
	Friend WithEvents Label1 As Label
	Friend WithEvents TextBox1 As TextBox
	Friend WithEvents OpenFileDialog1 As OpenFileDialog
End Class
