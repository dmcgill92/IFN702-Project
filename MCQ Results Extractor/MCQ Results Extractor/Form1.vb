Public Class Form1
	Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnBrowse1.Click
		If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
			'ListFiles(FolderBrowserDialog1.SelectedPath)

		End If
	End Sub

	Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles btnContinue.Click

	End Sub

	Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

	End Sub

	Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

	End Sub

	Private Sub Label3_Click(sender As Object, e As EventArgs)

	End Sub

	Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

	End Sub
End Class
