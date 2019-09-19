Imports System.ComponentModel

Public Class Student
    Dim _firstName As String
    Dim _lastName As String
    Dim _studentNumber As String
    Dim _result As Integer
    Dim _id As Integer

    Dim _match As Single

    Public Sub New(firstName As String, lastName As String, studentNumber As String, result As Integer, id As Integer, match As Single)
        _firstName = firstName
        _lastName = lastName
        _studentNumber = studentNumber
        _result = result
        _id = id
        _match = match
    End Sub

    Public Sub New(student As Student)
        _firstName = student.FirstName
        _lastName = student.LastName
        _studentNumber = student.StudentNumber
        _result = student.Result
        _id = student.Id
        _match = student.Match
    End Sub

    <DisplayName("First Name")>
    Public Property FirstName As String
        Get
            Return _firstName
        End Get
        Set(value As String)
            _firstName = value
        End Set
    End Property

    <DisplayName("Last Name")>
    Public Property LastName As String
        Get
            Return _lastName
        End Get
        Set(value As String)
            _lastName = value
        End Set
    End Property

    <DisplayName("Student Number")>
    Public Property StudentNumber As String
        Get
            Return _studentNumber
        End Get
        Set(value As String)
            _studentNumber = value
        End Set
    End Property

    Public Property Result As Integer
        Get
            Return _result
        End Get
        Set(value As Integer)
            _result = value
        End Set
    End Property

    <DisplayName("Match %")>
    Public Property Match As Single
        Get
            Return _match
        End Get
        Set(value As Single)
            _match = value
        End Set
    End Property

    <Browsable(False)>
    Public Property Id As Integer
        Get
            Return _id
        End Get
        Set(value As Integer)
            _id = value
        End Set
    End Property
End Class
