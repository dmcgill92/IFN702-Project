Public Class DataMatcher
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

	Public Function PartialMatchSimilarity(studentA As Student, studentList As List(Of Student)) As List(Of Single)
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

	Private Function GetSimilarity(stringA As String, stringB As String) As Single
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
End Class
