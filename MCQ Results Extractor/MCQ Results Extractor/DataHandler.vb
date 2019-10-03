Imports Equin.ApplicationFramework

Public Class DataHandler
    Public Sub RemoveFromGridByID(id As Integer, grid As DataGridView, list As List(Of Student))
        For Each student As Student In list
            If student.Id = id Then
                list.Remove(student)
                UpdateGrid(grid, list)
                Exit For
            End If
        Next
    End Sub

    Public Sub AddToGridByID(id As Integer, grid As DataGridView, listA As List(Of Student), listB As List(Of Student))
        For Each student As Student In listB
            If student.Id = id Then
                listA.Add(student)
                UpdateGrid(grid, listA)
                Exit For
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
        dataGrid.Columns(5).Visible = False
        dataGrid.Columns(6).Visible = False
    End Sub

    Public Sub AddToGrid(student As Student, grid As DataGridView, list As List(Of Student))
        list.Add(student)
        UpdateGrid(grid, list)
    End Sub

    Public Sub RemoveFromGrid(student As Student, grid As DataGridView, list As List(Of Student))
        list.Remove(student)
        UpdateGrid(grid, list)
    End Sub

    Public Sub UpdateGrid(ByRef grid As DataGridView, ByRef list As List(Of Student))
        If grid.InvokeRequired Then
            grid.Invoke(New UpdateGridInvoker(AddressOf UpdateGrid), grid, list)
        Else
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
        End If
    End Sub

    Private Delegate Sub UpdateGridInvoker(ByRef grid As DataGridView, ByRef list As List(Of Student))

End Class