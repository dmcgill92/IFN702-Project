Imports Excel = Microsoft.Office.Interop.Excel
Friend Class ExcelReader
	'Excel components
	Dim xlApp As Excel.Application
	Dim xlWorkBook As Excel.Workbook
	Dim xlWorkSheet As Excel.Worksheet
	Dim range As Excel.Range
	Public Event VariableChanged()

	'File locations
	Dim _resultsFilePath As String
	Dim _studentFilePath As String
	Dim _outputFilePath As String         'Where the output is saved

	Dim _resultsStudentList As List(Of Student) = New List(Of Student)
	Dim _studentList As List(Of Student) = New List(Of Student)
	Dim _outputStudentList As List(Of Student) = New List(Of Student)

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

	Public Sub Read_Excel(ByVal sFile As String, ByVal listID As Integer)
		Dim startTime As DateTime
		Dim stopTime As DateTime
		Dim elapsedTime As TimeSpan

		startTime = Now

		xlApp = New Excel.Application
		xlWorkBook = xlApp.Workbooks.Open(sFile)
		xlWorkSheet = xlWorkBook.Sheets(1)
		Dim rowCount As Integer = xlWorkSheet.UsedRange.Rows.Count
		Dim colCount As Integer = xlWorkSheet.UsedRange.Columns.Count

		Dim fNames As List(Of String) = New List(Of String)
		Dim lNames As List(Of String) = New List(Of String)
		Dim sNum As List(Of String) = New List(Of String)
		Dim results As List(Of Integer) = New List(Of Integer)

		range = xlWorkSheet.Range("A2:A" & rowCount)
		For Each cell As Excel.Range In range.Cells
			fNames.Add(cell.Value)
		Next

		range = xlWorkSheet.Range("B2:B" & rowCount)
		For Each cell As Excel.Range In range.Cells
			lNames.Add(cell.Value)
		Next

		range = xlWorkSheet.Range("C2:C" & rowCount)
		For Each cell As Excel.Range In range.Cells
			sNum.Add(cell.Value)
		Next

		range = xlWorkSheet.Range("D2:D" & rowCount)
		For Each cell As Excel.Range In range.Cells
			results.Add(cell.Value)
		Next

		For index As Integer = 0 To rowCount - 2
			If listID = 1 Then
				_resultsStudentList.Add(New Student(fNames(index), lNames(index), sNum(index), results(index)))
			ElseIf listID = 2 Then
				_studentList.Add(New Student(fNames(index), lNames(index), sNum(index), results(index)))
			Else
				_outputStudentList.Add(New Student(fNames(index), lNames(index), sNum(index), results(index)))
			End If
		Next

		xlWorkBook.Close()
		xlApp.Quit()

		' CLEAN UP. (CLOSE INSTANCES OF EXCEL OBJECTS.)
		System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp) : xlApp = Nothing
		System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook) : xlWorkBook = Nothing
		System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet) : xlWorkSheet = Nothing

		stopTime = Now
		elapsedTime = stopTime.Subtract(startTime)
		Console.WriteLine(elapsedTime.TotalSeconds.ToString("0.00000") & " Finished")
	End Sub

	Private Sub Validate_File_Locations() Handles Me.VariableChanged    'Activates the continue button when all file locations are input and valid
		Form1.btnContinue.Enabled = Not (_resultsFilePath = String.Empty OrElse _studentFilePath = String.Empty OrElse _outputFilePath = String.Empty)
	End Sub
End Class
