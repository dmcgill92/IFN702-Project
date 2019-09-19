Imports Equin.ApplicationFramework

Public Class DataHandler
    Public Sub RemoveDataRange(dataGrid As DataGridView, ByVal ids As List(Of Integer))
        For index As Integer = dataGrid.Rows.Count - 1 To 0 Step -1
            Dim student As Student = dataGrid.Rows(index).DataBoundItem.Object
            If ids.Contains(student.Id) Then
                dataGrid.Rows.Remove(dataGrid.Rows(index))
            End If
        Next
    End Sub

    Public Sub AddDataRange(dataGrid As DataGridView, fromGrid As DataGridView, ByVal ids As List(Of Integer))
        For Each row As DataGridViewRow In fromGrid.Rows
            Dim student As Student = row.DataBoundItem
            If ids.Contains(student.Id) Then
                dataGrid.Rows.Add(row)
            End If
        Next
    End Sub

    Public Sub AddIds(ByRef editList As List(Of Student), ByVal fromList As List(Of Student), ByVal ids As List(Of Integer))
        For Each student As Student In fromList
            If ids.Contains(student.Id) Then
                editList.Add(student)
            End If
        Next
    End Sub

    Public Sub BindData(dataGrid As DataGridView, studentList As List(Of Student))
        Dim blView As BindingListView(Of Student) = New BindingListView(Of Student)(studentList)
        dataGrid.AutoGenerateColumns = True
        dataGrid.DataSource = blView
        dataGrid.BindingContext = New BindingContext()
        dataGrid.Columns(0).FillWeight = 2
        dataGrid.Columns(1).FillWeight = 2
        dataGrid.Columns(2).FillWeight = 2
        dataGrid.Columns(3).FillWeight = 1
        dataGrid.Columns(4).Visible = False
    End Sub

    Public Sub ConfirmMatch(student As Student, removeGrid As DataGridView, addGrid As DataGridView)

    End Sub
End Class
