Public Class Form1
	Dim er As New ExcelReader   'Creates the excel reader
	Private Sub btnBrowse1_Click(sender As Object, e As EventArgs) Handles btnBrowse1.Click
		With OpenFileDialog1                        'Opens a file browser dialog
			.Title = "Excel File with Results"      'Dialog box title
			.FileName = ""
			.Filter = "Excel File|*.xlsx;*.xls"     'Filter to only show Excel files

			Dim sFileName As String = ""

			If .ShowDialog() = DialogResult.OK Then
				sFileName = .FileName

				'Validates the file selected
				If Trim(sFileName) <> "" Then
					If OpenFileDialog1.CheckFileExists Then
						fileLocation1.Text = OpenFileDialog1.FileName
						er.ResultsFilePath = OpenFileDialog1.FileName
					Else
						Throw New System.Exception("Files does not exist.")
					End If
				End If
			End If
		End With
	End Sub

	Private Sub btnBrowse2_Click(sender As Object, e As EventArgs) Handles btnBrowse2.Click
		With OpenFileDialog1                        'Opens a file browser dialog
			.Title = "Excel File with Student Names"      'Dialog box title
			.FileName = ""
			.Filter = "Excel File|*.xlsx;*.xls"     'Filter to only show Excel files

			Dim sFileName As String = ""

			If .ShowDialog() = DialogResult.OK Then
				sFileName = .FileName

				'Validates the file selected and sets it's location
				If Trim(sFileName) <> "" Then
					If OpenFileDialog1.CheckFileExists Then
						fileLocation2.Text = OpenFileDialog1.FileName
						er.StudentFilePath = OpenFileDialog1.FileName
					Else
						Throw New System.Exception("Files does not exist.")
					End If
				End If
			End If
		End With
	End Sub

	Private Sub btnBrowse3_Click(sender As Object, e As EventArgs) Handles btnBrowse3.Click
		With SaveFileDialog1                        'Opens a file browser dialog
			.OverwritePrompt = True                 'Opens the overwrite prompt if file already exists
			.Title = "Save output file"             'Dialog box title
			.FileName = ""
			.Filter = "Excel File|*.xlsx;*.xls"     'Filter to only show Excel files

			Dim sFileName As String = ""

			If .ShowDialog() = DialogResult.OK Then
				sFileName = .FileName

				'Sets the file location
				If Trim(sFileName) <> "" Then
					fileLocation3.Text = SaveFileDialog1.FileName
					er.OutputFilePath = SaveFileDialog1.FileName
				End If
			End If
		End With
	End Sub

	Private Sub btnContinue_Click(sender As Object, e As EventArgs) Handles btnContinue.Click       'Continues to next panel
		LaunchPanel.Hide()
		ComparisonPanel.Show()
	End Sub

	Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

	End Sub
End Class
