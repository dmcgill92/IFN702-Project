Imports OfficeOpenXml
Imports System.IO

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
                ElseIf range(1, index).Value.ToString() = "Score" OrElse range(1, index).Value.ToString().Contains("Exam") Then
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
            Dim fName As String = workSheet.Cells(columnLetters(0) & index).Value.ToString
            Dim lName As String = workSheet.Cells(columnLetters(1) & index).Value.ToString
            Dim sNum As String = workSheet.Cells(columnLetters(2) & index).Value.ToString
            Dim result As Integer = workSheet.Cells(columnLetters(3) & index).Value
            Dim id As Integer = index - 1
            Dim tempStudent As Student = New Student(fName, lName, sNum, result, id, Nothing)


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
                If rStudent.FirstName <> student.FirstName Then
                    isMatch = False
                    Continue For
                End If
                If rStudent.LastName <> student.LastName Then
                    isMatch = False
                    Continue For
                End If
                If rStudent.StudentNumber <> student.StudentNumber Then
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
    Private Sub Validate_File_Locations() Handles Me.VariableChanged    'Activates the continue button when all file locations are input and valid
        Form1.btnContinue.Enabled = Not (_resultsFilePath = String.Empty OrElse _studentFilePath = String.Empty)
    End Sub
End Class
