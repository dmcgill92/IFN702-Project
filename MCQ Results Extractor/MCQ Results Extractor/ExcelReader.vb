Imports System.IO
Imports OfficeOpenXml
Imports Microsoft.VisualBasic.FileIO

Friend Class ExcelReader

    Public Event VariableChanged()

    'File locations
    Dim _resultsFilePath As String

    Dim _studentFilePath As String
	Dim _outputFilePath As String         'Where the output is saved

	Public _resultsStudentList As List(Of Student) = New List(Of Student)
    Public _studentList As List(Of Student) = New List(Of Student)
    Public _outputStudentList As List(Of Student) = New List(Of Student)

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

	Public Function FindColumns(workSheet As ExcelWorksheet, selectedHeader As String) As List(Of String)
		Dim tempList As List(Of String) = New List(Of String)
		tempList.Add("")
		tempList.Add("")
		tempList.Add("")
		tempList.Add("")
		Dim foundColumns As Integer = 0

		Dim colCount As Integer = workSheet.Dimension.End.Column
		Dim range As ExcelRange = workSheet.Cells(1, 1, 1, colCount)
		For index As Integer = 1 To colCount + 1

			If range(1, index).Address <> "" And range(1, index).Value <> Nothing Then
				If range(1, index).Value.ToString() = "Initial" OrElse range(1, index).Value.ToString() = "First Name" Then
					tempList(0) = range(1, index).Address()(0).ToString()
					foundColumns += 1
				ElseIf range(1, index).Value.ToString() = "Surname" OrElse range(1, index).Value.ToString() = "Last Name" Then
					tempList(1) = range(1, index).Address()(0).ToString()
					foundColumns += 1
				ElseIf range(1, index).Value.ToString() = "Student number" OrElse range(1, index).Value.ToString() = "Username" Then
					tempList(2) = range(1, index).Address()(0).ToString()
					foundColumns += 1
				ElseIf range(1, index).Value.ToString() = "Score" OrElse range(1, index).Value.ToString() = selectedHeader Then
					tempList(3) = range(1, index).Address()(0).ToString()
					foundColumns += 1
				End If
			End If
			If foundColumns = 4 Then
				Exit For
			End If
		Next
		Return tempList
	End Function

	Public Function ValidateFile(ByVal sFile As String, isResults As Boolean) As Boolean
		Dim isValid As Boolean = False
		Dim requiredHeaderCount As Integer = If(isResults, 4, 3)
		Dim headerCount As Integer = 0
		If Path.GetExtension(sFile) <> ".csv" Then
			Dim file As FileInfo = New FileInfo(sFile)
			Dim xlPackage As New ExcelPackage(file)
			Dim workSheet As ExcelWorksheet = xlPackage.Workbook.Worksheets(1)
			Dim colCount As Integer = workSheet.Dimension.End.Column
			Dim range As ExcelRange = workSheet.Cells(1, 1, 1, colCount)
			For i As Integer = 1 To colCount
				Dim header As String = range(1, i).Value
				If isResults Then
					Select Case header
						Case "Initial"
							headerCount += 1
						Case "Surname"
							headerCount += 1
						Case "Student number"
							headerCount += 1
						Case "Score"
							headerCount += 1
						Case Else
							Exit Select
					End Select
				Else
					Select Case header
						Case "First Name"
							headerCount += 1
						Case "Last Name"
							headerCount += 1
						Case "Username"
							headerCount += 1
						Case Else
							Exit Select
					End Select
				End If
				If headerCount = requiredHeaderCount Then
					isValid = True
					Exit For
				End If
			Next
		Else
			Dim tfp As New TextFieldParser(sFile)
			tfp.Delimiters = New String() {","}
			tfp.TextFieldType = FieldType.Delimited

			Dim headers As String() = tfp.ReadFields()
			For index As Integer = 0 To headers.Count
				If headers(index) <> "" And headers(index) <> Nothing Then
					If isResults Then
						Select Case headers(index)
							Case "Initial"
								headerCount += 1
							Case "Surname"
								headerCount += 1
							Case "Student number"
								headerCount += 1
							Case "Score"
								headerCount += 1
							Case Else
								Exit Select
						End Select
					Else
						Select Case headers(index)
							Case "First Name"
								headerCount += 1
							Case "Last Name"
								headerCount += 1
							Case "Username"
								headerCount += 1
							Case "Last Access"
								Exit Select
							Case Else
								headerCount += 1
						End Select
					End If
				End If
				If headerCount = requiredHeaderCount Then
					isValid = True
					Exit For
				End If
			Next
		End If
		Return isValid
	End Function

	Public Function Read_Excel(ByVal sFile As String, selectedHeader As String) As List(Of Student) 'Extracts the data from the excel file
		Dim tempList As List(Of Student) = New List(Of Student)
		If Path.GetExtension(sFile) <> ".csv" Then
			Dim file As FileInfo = New FileInfo(sFile)
			Dim xlPackage As New ExcelPackage(file)

			Dim workSheet As ExcelWorksheet = xlPackage.Workbook.Worksheets(1)
			Dim rowCount As Integer = workSheet.Dimension.End.Row
			Dim columnLetters As List(Of String) = FindColumns(workSheet, selectedHeader)

			For index As Integer = 2 To rowCount
				Dim fname As String = ""
				Dim lname As String = ""
				Dim sNum As String = ""
				Dim result As Integer = 0
				If workSheet.Cells(columnLetters(0) & index).Value <> Nothing Then
					fname = workSheet.Cells(columnLetters(0) & index).Value.ToString.TrimEnd
				End If
				If workSheet.Cells(columnLetters(1) & index).Value <> Nothing Then
					lname = workSheet.Cells(columnLetters(1) & index).Value.ToString.TrimEnd
				End If
				If workSheet.Cells(columnLetters(2) & index).Value <> Nothing Then
					sNum = workSheet.Cells(columnLetters(2) & index).Value.ToString.TrimEnd
				End If
				If workSheet.Cells(columnLetters(3) & index).Value <> Nothing Then
					result = workSheet.Cells(columnLetters(3) & index).Value
				End If
				Dim id As Integer = index - 1
				Dim tempStudent As Student = New Student(fname, lname, sNum, result, id, Nothing, Nothing, Nothing, Nothing)

				tempList.Add(tempStudent)
			Next

			xlPackage.Dispose()
		Else
			Dim dataIndices As Integer() = New Integer(3) {}
			Dim foundColumns As Integer = 0
			Dim tfp As New TextFieldParser(sFile)
			tfp.Delimiters = New String() {","}
			tfp.TextFieldType = FieldType.Delimited

			Dim headers As String() = tfp.ReadFields()
			For index As Integer = 0 To headers.Count
				If headers(index) <> "" And headers(index) <> Nothing Then
					If headers(index) = "Initial" OrElse headers(index) = "First Name" Then
						dataIndices(0) = index
						foundColumns += 1
					ElseIf headers(index) = "Surname" OrElse headers(index) = "Last Name" Then
						dataIndices(1) = index
						foundColumns += 1
					ElseIf headers(index) = "Student number" OrElse headers(index) = "Username" Then
						dataIndices(2) = index
						foundColumns += 1
					ElseIf headers(index) = "Score" OrElse headers(index) = selectedHeader Then
						dataIndices(3) = index
						foundColumns += 1
					End If
				End If
				If foundColumns = 4 Then
					Exit For
				End If
			Next

			While tfp.EndOfData = False
				Dim fields As String() = tfp.ReadFields()
				Dim tempStudent = New Student(fields(dataIndices(0)), fields(dataIndices(1)), fields(dataIndices(2)), fields(dataIndices(3)), tempList.Count + 1, Nothing, Nothing, Nothing, Nothing)
				tempList.Add(tempStudent)
			End While
		End If
		Return tempList
	End Function

	Public Function Match(ByRef rStudent As Student, ByRef studentList As List(Of Student)) As Integer()
        Dim isMatch As Boolean = True
        Dim removeIndices As Integer() = New Integer(2) {}
        For Each student As Student In studentList
			If rStudent.FirstName = "" OrElse rStudent.FirstName(0).ToString().ToLower() <> student.FirstName(0).ToString().ToLower() Then
				isMatch = False
				Continue For
			End If
			If rStudent.LastName.ToLower() <> student.LastName.ToLower() Then
				isMatch = False
				Continue For
			End If
			If "n" & rStudent.StudentNumber <> student.StudentNumber Then
				isMatch = False
				Continue For
			End If
			isMatch = True
            removeIndices(1) = student.Id
			student.Result = rStudent.Result
			rStudent.MatchFirstName = student.FirstName
			rStudent.MatchLastName = student.LastName
            rStudent.MatchStudentNumber = student.StudentNumber
            Exit For
        Next
        If isMatch Then
            removeIndices(0) = rStudent.Id
            Return removeIndices
        End If
        Return Nothing
    End Function

	Public Function Partial_Match_Similarity(studentA As Student, studentList As List(Of Student)) As List(Of Single)
		Dim simList As List(Of Single) = New List(Of Single)
		For Each studentB As Student In studentList
			Dim fnA = (If(studentA.FirstName.Length > 0, studentA.FirstName(0).ToString().ToLower(), ""))
			Dim fnB = (If(studentB.FirstName.Length > 0, studentB.FirstName(0).ToString().ToLower(), ""))
			Dim fnSim As Single = (If(fnA = fnB, 1, 0))
			Dim lnA = studentA.LastName.ToLower()
			Dim lnB = studentB.LastName.ToLower()
			Dim lnSim As Single = GetSimilarity(lnA, lnB)
			Dim snA = "n" & studentA.StudentNumber
			Dim snB = studentB.StudentNumber
			Dim snSim As Single = GetSimilarity(snA, snB)
			Dim totalSim As Single = (fnSim * 0.1F + (lnSim * 0.3F) + (snSim * 0.6F))
			studentB.Match = totalSim
			simList.Add(totalSim)
		Next
		Return simList
	End Function

	Public Function GetSimilarity(stringA As String, stringB As String) As Single
        Dim distance As Single = ComputeDistance(stringA, stringB)
        Dim maxLen As Single = stringA.Length
        If maxLen < stringB.Length Then
            maxLen = stringB.Length
        End If
        If maxLen = 0.0F Then
            Return 1.0F
        Else
            Return 1.0F - distance / maxLen
        End If
    End Function

    Private Function ComputeDistance(s As String, t As String) As Integer
        Dim n As Integer = s.Length
        Dim m As Integer = t.Length
        Dim distance As Integer(,) = New Integer(n, m) {}
        Dim cost As Integer = 0

        If n = 0 Then
            Return m
        End If
        If m = 0 Then
            Return n
        End If

        Dim i As Integer = 0
        While i <= n
            distance(i, 0) = Math.Min(Threading.Interlocked.Increment(i), i - 1)
        End While

        Dim j As Integer = 0
        While j <= m
            distance(0, j) = Math.Min(Threading.Interlocked.Increment(j), j - 1)
        End While

		For i = 1 To n
			For j = 1 To m
				cost = (If(t.Substring(j - 1, 1) = s.Substring(i - 1, 1), 0, 1))
				distance(i, j) = Math.Min(distance(i - 1, j) + 1, Math.Min(distance(i, j - 1) + 1, distance(i - 1, j - 1) + cost))
			Next
		Next

		Return distance(n, m)
    End Function

	Private Sub Validate_File_Locations() Handles Me.VariableChanged    'Activates the continue button when all file locations are input and valid
		Form1.btnContinue.Enabled = Not (_resultsFilePath = String.Empty OrElse _studentFilePath = String.Empty)
	End Sub

	Public Sub WriteToExcel(inPath As String, outPath As String, list As List(Of Student), selectedHeader As String)
		Dim file As FileInfo = New FileInfo(inPath)
		Dim fileOut As FileInfo = New FileInfo(outPath)
		Dim xlPackage As New ExcelPackage(file)

		Dim workSheet As ExcelWorksheet = xlPackage.Workbook.Worksheets(1)
		Dim rowCount As Integer = workSheet.Dimension.End.Row
		Dim columnLetters As List(Of String) = FindColumns(workSheet, selectedHeader)

		For index As Integer = 2 To rowCount
			For Each student In list
				If student.StudentNumber = workSheet.Cells(columnLetters(2) & index).Value.ToString.TrimEnd Then
					workSheet.Cells(columnLetters(3) & index).Value = student.Result
					Exit For
				End If
			Next
		Next

		xlPackage.SaveAs(fileOut)
	End Sub

	Public Function CheckHeaders(ByVal sFile As String) As List(Of String)
		Dim file As FileInfo = New FileInfo(sFile)
		Dim xlPackage As New ExcelPackage(file)
		Dim tempList As List(Of String) = New List(Of String)

		Dim workSheet As ExcelWorksheet = xlPackage.Workbook.Worksheets(1)
		Dim rowCount As Integer = workSheet.Dimension.End.Row
		Dim columnLetters As List(Of String) = New List(Of String)

		Dim colCount As Integer = workSheet.Dimension.End.Column
		Dim range As ExcelRange = workSheet.Cells(1, 1, 1, colCount)
		For index As Integer = 1 To colCount + 1

			If range(1, index).Value = Nothing Then
				Exit For
			End If
			Dim header As String = range(1, index).Value.ToString()
			Select Case header
				Case "Last Name"
					Exit Select
				Case "First Name"
					Exit Select
				Case "Username"
					Exit Select
				Case "Last Access"
					Exit Select
				Case Else
					tempList.Add(header)
			End Select
		Next
		Return tempList
	End Function

End Class