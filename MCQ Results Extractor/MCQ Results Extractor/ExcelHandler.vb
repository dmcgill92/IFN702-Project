Imports System.IO
Imports OfficeOpenXml
Imports Microsoft.VisualBasic.FileIO

Friend Class ExcelHandler
	Private Function FindColumns(workSheet As ExcelWorksheet, selectedHeader As String) As List(Of String)
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

	Public Function ReadExcel(ByVal sFile As String, selectedHeader As String) As List(Of Student) 'Extracts the data from the excel file
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