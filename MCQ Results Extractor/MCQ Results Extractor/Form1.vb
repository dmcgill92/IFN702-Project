Imports System.ComponentModel
Imports Equin.ApplicationFramework

Public Class Form1
    Dim er As New ExcelReader   'Creates the excel reader

    Dim numMatched As Integer
    Dim totalStudents As Integer

    Dim originalResultsStudentList As List(Of Student)
    Dim originalStudentList As List(Of Student)

    Dim leftMatchedList As List(Of Student)
    Dim leftUnmatchedList As List(Of Student)
    Dim rightMatchedList As List(Of Student)
    Dim rightUnmatchedList As List(Of Student)

    Dim bindings As List(Of BindingListView(Of Student)) = New List(Of BindingListView(Of Student))()

    Private Sub btnBrowse1_Click(sender As Object, e As EventArgs) Handles btnBrowse1.Click
        With OpenFileDialog1                        'Opens a file browser dialog
            .Title = "Excel File with Results"      'Dialog box title
            .FileName = ""
            .Filter = "Excel File|*.xlsx;*.xls"     'Filter to only show Excel files

            Dim sFileName As String = ""

            If .ShowDialog() = DialogResult.OK Then
                sFileName = .FileName

                'Validates the file selected
                If Trim(sFileName) <> "" Then
                    If OpenFileDialog1.CheckFileExists Then
                        fileLocation1.Text = OpenFileDialog1.FileName
                        er.ResultsFilePath = OpenFileDialog1.FileName
                    Else
                        Throw New System.Exception("Files does not exist.")
                    End If
                End If
            End If
        End With
    End Sub

    Private Sub btnBrowse2_Click(sender As Object, e As EventArgs) Handles btnBrowse2.Click
        With OpenFileDialog1                        'Opens a file browser dialog
            .Title = "Excel File with Student Names"      'Dialog box title
            .FileName = ""
            .Filter = "Excel File|*.xlsx;*.xls"     'Filter to only show Excel files

            Dim sFileName As String = ""

            If .ShowDialog() = DialogResult.OK Then
                sFileName = .FileName

                'Validates the file selected and sets it's location
                If Trim(sFileName) <> "" Then
                    If OpenFileDialog1.CheckFileExists Then
                        fileLocation2.Text = OpenFileDialog1.FileName
                        er.StudentFilePath = OpenFileDialog1.FileName
                    Else
                        Throw New System.Exception("Files does not exist.")
                    End If
                End If
            End If
        End With
    End Sub


    Private Sub btnContinue_Click(sender As Object, e As EventArgs) Handles btnContinue.Click       'Continues to next panel
        'Sets the unmatched tab to active
        tcLeft.SelectTab(tabUnmatched1)
        tcRight.SelectTab(tabUnmatched2)
        tcLeft.TabPages.Remove(tabMatched1)
        tcRight.TabPages.Remove(tabMatched2)


        'Extract data from Excel
        originalResultsStudentList = er.Read_Excel(er.ResultsFilePath)
        leftUnmatchedList = New List(Of Student)(originalResultsStudentList)
        Dim blView As BindingListView(Of Student) = New BindingListView(Of Student)(leftUnmatchedList)
        dgUnmatchedLeft.DataSource = blView
        dgUnmatchedLeft.BindingContext = New BindingContext()
        dgUnmatchedLeft.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
        dgUnmatchedLeft.Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
        dgUnmatchedLeft.Columns(0).FillWeight = 2
        dgUnmatchedLeft.Columns(1).FillWeight = 2
        dgUnmatchedLeft.Columns(2).FillWeight = 2
        dgUnmatchedLeft.Columns(3).FillWeight = 1

        originalStudentList = er.Read_Excel(er.StudentFilePath)
        totalStudents = originalStudentList.Count
        rightUnmatchedList = New List(Of Student)(originalStudentList)
        Dim blView2 As BindingListView(Of Student) = New BindingListView(Of Student)(rightUnmatchedList)
        dgUnmatchedRight.DataSource = blView2
        dgUnmatchedRight.BindingContext = New BindingContext()
        dgUnmatchedRight.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
        dgUnmatchedRight.Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
        dgUnmatchedRight.Columns(0).FillWeight = 2
        dgUnmatchedRight.Columns(1).FillWeight = 2
        dgUnmatchedRight.Columns(2).FillWeight = 2
        dgUnmatchedRight.Columns(3).FillWeight = 1
        dgUnmatchedRight.Sort(dgUnmatchedRight.Columns(1), ListSortDirection.Ascending)

        'Switch the current panel
        LaunchPanel.Hide()
        ComparisonPanel.Show()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub BtnMatch_Click(sender As Object, e As EventArgs) Handles btnMatch.Click

        Dim tempList As List(Of List(Of Integer)) = er.Full_Match(leftUnmatchedList, rightUnmatchedList)
        Dim leftIds As List(Of Integer) = tempList(0)
        Dim rightIds As List(Of Integer) = tempList(1)
        numMatched = rightIds.Count

        leftMatchedList = New List(Of Student)
        AddIds(leftMatchedList, leftUnmatchedList, leftIds)
        Dim blView As BindingListView(Of Student) = New BindingListView(Of Student)(leftMatchedList)
        dgMatchedLeft.AutoGenerateColumns = True
        dgMatchedLeft.DataSource = blView
        dgMatchedLeft.BindingContext = New BindingContext()
        dgMatchedLeft.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
        dgMatchedLeft.Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
        dgMatchedLeft.Columns(0).FillWeight = 2
        dgMatchedLeft.Columns(1).FillWeight = 2
        dgMatchedLeft.Columns(2).FillWeight = 2
        dgMatchedLeft.Columns(3).FillWeight = 1


        rightMatchedList = New List(Of Student)
        AddIds(rightMatchedList, rightUnmatchedList, rightIds)
        Dim blView2 As BindingListView(Of Student) = New BindingListView(Of Student)(rightMatchedList)
        dgMatchedRight.AutoGenerateColumns = True
        dgMatchedRight.DataSource = blView2
        dgMatchedRight.BindingContext = New BindingContext()
        dgMatchedRight.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
        dgMatchedRight.Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
        dgMatchedRight.Columns(0).FillWeight = 2
        dgMatchedRight.Columns(1).FillWeight = 2
        dgMatchedRight.Columns(2).FillWeight = 2
        dgMatchedRight.Columns(3).FillWeight = 1
        dgMatchedLeft.Sort(dgMatchedLeft.Columns(1), ListSortDirection.Ascending)

        RemoveDataRange(dgUnmatchedLeft, leftIds)
        RemoveDataRange(dgUnmatchedRight, rightIds)

        tcLeft.TabPages.Insert(0, tabMatched1)
        tcRight.TabPages.Insert(0, tabMatched2)
        tcLeft.SelectTab(tabMatched1)
        tcRight.SelectTab(tabMatched2)
        lblMatched.Visible = True
        lblMatched.Text = "Matched: " & numMatched & "/" & totalStudents
        btnMatch.Enabled = False
    End Sub

    Private Sub tabControl_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tcLeft.SelectedIndexChanged, tcRight.SelectedIndexChanged
        Dim tabControl As TabControl = CType(sender, TabControl)
        Dim index As Integer = tabControl.SelectedIndex

        tcLeft.SelectTab(index)
        tcRight.SelectTab(index)
    End Sub

    Private Sub RemoveDataRange(dataGrid As DataGridView, ByVal ids As List(Of Integer))

        For index As Integer = dataGrid.Rows.Count - 1 To 0 Step -1
            Dim student As Student = dataGrid.Rows(index).DataBoundItem.Object
            If ids.Contains(student.Id) Then
                dataGrid.Rows.Remove(dataGrid.Rows(index))
            End If
        Next
    End Sub

    Private Sub AddDataRange(dataGrid As DataGridView, fromGrid As DataGridView, ByVal ids As List(Of Integer))

        For Each row As DataGridViewRow In fromGrid.Rows
            Dim student As Student = row.DataBoundItem
            If ids.Contains(student.Id) Then
                dataGrid.Rows.Add(row)
            End If
        Next
    End Sub

    Private Sub AddIds(ByRef editList As List(Of Student), ByVal fromList As List(Of Student), ByVal ids As List(Of Integer))
        For Each student As Student In fromList
            If ids.Contains(student.Id) Then
                editList.Add(student)
            End If
        Next
    End Sub

    Private Sub Form1_ResizeBegin(sender As Object, e As EventArgs) Handles MyBase.ResizeBegin
        SuspendLayout()
    End Sub

    Private Sub Form1_ResizeEnd(sender As Object, e As EventArgs) Handles MyBase.ResizeEnd
        ResumeLayout(True)
    End Sub

    Private Sub DgMatchedLeft_Scroll(sender As Object, e As ScrollEventArgs) Handles dgMatchedLeft.Scroll
        dgMatchedRight.FirstDisplayedScrollingRowIndex = dgMatchedLeft.FirstDisplayedScrollingRowIndex
    End Sub

    Private Sub DgMatchedRight_Scroll(sender As Object, e As ScrollEventArgs) Handles dgMatchedRight.Scroll
        dgMatchedLeft.FirstDisplayedScrollingRowIndex = dgMatchedRight.FirstDisplayedScrollingRowIndex
    End Sub

    Private Sub DgMatchedLeft_Sorted(sender As Object, e As EventArgs) Handles dgMatchedLeft.Sorted
        If dgMatchedRight.SortOrder <> Nothing Then
            If dgMatchedRight.SortedColumn.Index <> dgMatchedLeft.SortedColumn.Index Or dgMatchedRight.SortOrder <> dgMatchedLeft.SortOrder Then
                dgMatchedRight.Sort(dgMatchedRight.Columns(dgMatchedLeft.SortedColumn.Index), dgMatchedLeft.SortOrder - 1)
            End If
        Else
            dgMatchedRight.Sort(dgMatchedRight.Columns(dgMatchedLeft.SortedColumn.Index), dgMatchedLeft.SortOrder - 1)
        End If
    End Sub

    Private Sub DgMatchedRight_Sorted(sender As Object, e As EventArgs) Handles dgMatchedRight.Sorted
        If dgMatchedLeft.SortOrder <> Nothing Then
            If dgMatchedRight.SortedColumn.Index <> dgMatchedLeft.SortedColumn.Index Or dgMatchedRight.SortOrder <> dgMatchedLeft.SortOrder Then
                dgMatchedLeft.Sort(dgMatchedLeft.Columns(dgMatchedRight.SortedColumn.Index), dgMatchedRight.SortOrder - 1)

            End If
        Else
            dgMatchedLeft.Sort(dgMatchedLeft.Columns(dgMatchedRight.SortedColumn.Index), dgMatchedRight.SortOrder - 1)
        End If
    End Sub
End Class
