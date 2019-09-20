Imports System.IO
Imports OfficeOpenXml

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

    Public Function FindColumns(workSheet As ExcelWorksheet) As List(Of String)
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
                ElseIf range(1, index).Value.ToString() = "Score" OrElse range(1, index).Value.ToString().Contains("Total Pts") Then
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

    Public Function Read_Excel(ByVal sFile As String) As List(Of Student) 'Extracts the data from the excel file
        Dim file As FileInfo = New FileInfo(sFile)
        Dim xlPackage As New ExcelPackage(file)

        Dim workSheet As ExcelWorksheet = xlPackage.Workbook.Worksheets(1)
        Dim rowCount As Integer = workSheet.Dimension.End.Row
        Dim columnLetters As List(Of String) = FindColumns(workSheet)

        Dim tempList As List(Of Student) = New List(Of Student)

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
            Dim tempStudent As Student = New Student(fname, lname, sNum, result, id, Nothing)

            tempList.Add(tempStudent)
        Next

        xlPackage.Dispose()
        Return tempList
    End Function

    Public Function Full_Match(ByRef resultsFile As List(Of Student), ByRef studentFile As List(Of Student)) As List(Of List(Of Integer))
        Dim rRemoveList As List(Of Integer) = New List(Of Integer)
        Dim sRemoveList As List(Of Integer) = New List(Of Integer)
        Dim tempList As List(Of List(Of Integer)) = New List(Of List(Of Integer))
        Dim isMatch As Boolean
        For Each rStudent As Student In resultsFile
            isMatch = True
            For Each student As Student In studentFile
                If rStudent.FirstName = "" OrElse rStudent.FirstName(0).ToString().ToLower() <> student.FirstName(0).ToString().ToLower() Then
                    'Console.WriteLine("{0} : {1}", rStudent.FirstName, student.FirstName)
                    isMatch = False
                    Continue For
                End If
                If rStudent.LastName.ToLower() <> student.LastName.ToLower() Then
                    'Console.WriteLine("{0} : {1}", rStudent.LastName, student.LastName)
                    isMatch = False
                    Continue For
                End If
                If "n" & rStudent.StudentNumber <> student.StudentNumber Then
                    'Console.WriteLine("{0} : {1}", rStudent.StudentNumber, student.StudentNumber)
                    isMatch = False
                    Continue For
                End If
                isMatch = True
                sRemoveList.Add(student.Id)
                student.Result = rStudent.Result
                Exit For
            Next
            If isMatch Then
                rRemoveList.Add(rStudent.Id)
                Continue For
            End If

        Next
        tempList.Add(rRemoveList)
        tempList.Add(sRemoveList)
        Return tempList
    End Function

    Public Sub Partial_Match_Similarity(studentA As Student, studentList As List(Of Student))
        Dim simList As List(Of Single) = New List(Of Single)
        For Each studentB As Student In studentList
            Dim fnA = (If(studentA.FirstName.Length > 0, studentA.FirstName(0).ToString().ToLower(), ""))
            Dim fnB = (If(studentB.FirstName.Length > 0, studentB.FirstName(0).ToString().ToLower(), ""))
            Dim fnSim As Single = (If(fnA = fnB, 1, 0))
            'Console.WriteLine("{0} : {1} - Match = {2}", fnA, fnB, fnSim)
            Dim lnA = studentA.LastName.ToLower()
            Dim lnB = studentB.LastName.ToLower()
            Dim lnSim As Single = GetSimilarity(lnA, lnB)
            'Console.WriteLine("{0} : {1} - Match = {2}", lnA, lnB, lnSim)
            Dim snA = "n" & studentA.StudentNumber
            Dim snB = studentB.StudentNumber
            Dim snSim As Single = GetSimilarity(snA, snB)
            'Console.WriteLine("{0} : {1} - Match = {2}", snA, snB, snSim)
            Dim totalSim As Single = (fnSim + (lnSim * 2) + (snSim * 3)) / 6
            studentB.Match = totalSim
            'Console.WriteLine("{0} : {1} - Match = {2}%", studentA.LastName, studentB.LastName, (totalSim * 100).ToString("N2"))
        Next
    End Sub

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

        'For x As Integer = 0 To m
        '    For y As Integer = 0 To n
        '        Console.Write("{0} ", distance(y, x))
        '    Next
        '    Console.WriteLine()
        'Next

        Return distance(n, m)
    End Function

    Private Sub Validate_File_Locations() Handles Me.VariableChanged    'Activates the continue button when all file locations are input and valid
        Form1.btnContinue.Enabled = Not (_resultsFilePath = String.Empty OrElse _studentFilePath = String.Empty)
    End Sub

End Class