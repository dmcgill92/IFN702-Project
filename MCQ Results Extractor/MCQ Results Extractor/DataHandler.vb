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

    Public Sub AddToGrid(student As Student, grid As DataGridView, list As List(Of Student))
        list.Add(student)
        UpdateGrid(grid, list)
    End Sub

    Public Sub RemoveFromGrid(student As Student, grid As DataGridView, list As List(Of Student))
        list.Remove(student)
        UpdateGrid(grid, list)
    End Sub

    Public Sub UpdateGrid(grid As DataGridView, list As List(Of Student))
        Dim sortedOrder As SortOrder = grid.SortOrder
        Dim sortColumnIndex As Integer = -1
        If sortedOrder <> SortOrder.None Then
            sortColumnIndex = grid.SortedColumn.Index
        End If
        Dim bindListView As BindingListView(Of Student) = New BindingListView(Of Student)(list)
        grid.DataSource = bindListView
        grid.BindingContext = New BindingContext()
        If sortedOrder = SortOrder.None Then
            Return
        End If
        grid.Sort(grid.Columns(sortColumnIndex), sortedOrder - 1)
    End Sub

End Class