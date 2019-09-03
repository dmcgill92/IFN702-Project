Imports Excel = Microsoft.Office.Interop.Excel

Imports OfficeOpenXml.Style
Imports OfficeOpenXml
Imports System
Imports System.IO
Imports MCQ_Results_Extractor

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

    Public Property ResultsStudentList As List(Of Student)
        Get
            Return _resultsStudentList
        End Get
        Set(value As List(Of Student))
            _resultsStudentList = value
        End Set
    End Property

    Public Property StudentList As List(Of Student)
        Get
            Return _studentList
        End Get
        Set(value As List(Of Student))
            _studentList = value
        End Set
    End Property

    Public Property OutputStudentList As List(Of Student)
        Get
            Return _outputStudentList
        End Get
        Set(value As List(Of Student))
            _outputStudentList = value
        End Set
    End Property

    Public Sub Read_Excel(ByVal sFile As String, ByVal listID As Integer) 'Extracts the data from the excel file
        Dim file As FileInfo = New FileInfo(sFile)
        Dim xlPackage As New ExcelPackage(file)

        Dim workSheet = xlPackage.Workbook.Worksheets(1)
        Dim rowCount As Integer = workSheet.Dimension.End.Row

        For index As Integer = 2 To rowCount
            Dim fName As String = workSheet.Cells("A" & index).Value.ToString
            Dim lName As String = workSheet.Cells("B" & index).Value.ToString
            Dim sNum As String = workSheet.Cells("C" & index).Value.ToString
            Dim result As Integer = workSheet.Cells("D" & index).Value
            Dim tempStudent As Student = New Student(fName, lName, sNum, result)

            Select Case listID
                Case 1
                    ResultsStudentList.Add(tempStudent)
                Case 2
                    StudentList.Add(tempStudent)
                Case 3
                    OutputStudentList.Add(tempStudent)
                Case Else
                    Throw New System.Exception("Invalid List ID")
            End Select
        Next

        xlPackage.Dispose()
    End Sub

    Private Sub Validate_File_Locations() Handles Me.VariableChanged    'Activates the continue button when all file locations are input and valid
		Form1.btnContinue.Enabled = Not (_resultsFilePath = String.Empty OrElse _studentFilePath = String.Empty OrElse _outputFilePath = String.Empty)
	End Sub
End Class
