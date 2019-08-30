Imports Excel = Microsoft.Office.Interop.Excel
Friend Class ExcelReader
	'Excel components
	Dim xlApp As Excel.Application
	Dim xlWorkBook As Excel.Workbook
	Dim xlWorkSheet As Excel.Worksheet
	Public Event VariableChanged()

	'File locations
	Dim _resultsFilePath As String
	Dim _studentFilePath As String
	Dim _outputFilePath As String       'Where the output is saved

	Public Property ResultsFilePath As String
		Get
			Return _resultsFilePath
		End Get
		Set(value As String)
			_resultsFilePath = value
			RaiseEvent VariableChanged()
		End Set
	End Property

	Public Property StudentFilePath As String
		Get
			Return _studentFilePath
		End Get
		Set(value As String)
			_studentFilePath = value
			RaiseEvent VariableChanged()
		End Set
	End Property

	Public Property OutputFilePath As String
		Get
			Return _outputFilePath
		End Get
		Set(value As String)
			_outputFilePath = value
			RaiseEvent VariableChanged()
		End Set
	End Property

	Private Sub Read_Excel(ByVal sFile As String)   'Not yet implemented
		xlApp = New Excel.Application
		xlWorkBook = xlApp.Workbooks.Open(sFile)
		xlWorkSheet = xlApp.Worksheets("Sheet1")
	End Sub

	Private Sub Validate_File_Locations() Handles Me.VariableChanged    'Activates the continue button when all file locations are input and valid
		Form1.btnContinue.Enabled = Not (_resultsFilePath = String.Empty OrElse _studentFilePath = String.Empty OrElse _outputFilePath = String.Empty)
	End Sub
End Class
