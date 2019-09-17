Imports System.ComponentModel
Imports Equin.ApplicationFramework

Public Class Form1
    Dim er As New ExcelReader   'Creates the excel reader
    Dim dh As New DataHandler

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
        dh.BindData(dgUnmatchedLeft, leftUnmatchedList)
        dgUnmatchedLeft.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
        dgUnmatchedLeft.Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable

        originalStudentList = er.Read_Excel(er.StudentFilePath)
        totalStudents = originalStudentList.Count
        rightUnmatchedList = New List(Of Student)(originalStudentList)
        dh.BindData(dgUnmatchedRight, rightUnmatchedList)
        dgUnmatchedRight.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
        dgUnmatchedRight.Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
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
        dh.AddIds(leftMatchedList, leftUnmatchedList, leftIds)
        dh.BindData(dgMatchedLeft, leftMatchedList)
        dgMatchedLeft.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
        dgMatchedLeft.Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable

        rightMatchedList = New List(Of Student)
        dh.AddIds(rightMatchedList, rightUnmatchedList, rightIds)
        dh.BindData(dgMatchedRight, rightMatchedList)
        dgMatchedRight.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
        dgMatchedRight.Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
        dgMatchedLeft.Sort(dgMatchedLeft.Columns(1), ListSortDirection.Ascending)

        dh.RemoveDataRange(dgUnmatchedLeft, leftIds)
        dh.RemoveDataRange(dgUnmatchedRight, rightIds)

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

    Private Sub DgUnmatchedLeft_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgUnmatchedLeft.CellClick
        Dim row As DataGridViewRow = dgUnmatchedLeft.SelectedRows(0)
        Dim selectedStudent As Student = row.DataBoundItem.Object
        Dim simList As List(Of Single) = er.Partial_Match_Similarity(selectedStudent, rightUnmatchedList)
        'dh.SetSimilarity(simList)
    End Sub
End Class
