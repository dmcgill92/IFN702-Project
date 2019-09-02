Public Class Student
	Dim _firstName As String
	Dim _lastName As String
	Dim _studentNumber As String
	Dim _result As Integer

	Public Sub New(firstName As String, lastName As String, studentNumber As String, result As Integer)
		_firstName = firstName
		_lastName = lastName
		_studentNumber = studentNumber
		_result = result
	End Sub

	Public Property FirstName As String
		Get
			Return _firstName
		End Get
		Set(value As String)
			_firstName = value
		End Set
	End Property

	Public Property LastName As String
		Get
			Return _lastName
		End Get
		Set(value As String)
			_lastName = value
		End Set
	End Property

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
End Class
