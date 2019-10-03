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

    Dim leftIDs As List(Of Integer)
    Dim rightIDs As List(Of Integer)

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
        dgUnmatchedLeft.Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
        dgUnmatchedLeft.Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable
        dgUnmatchedLeft.Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable

        originalStudentList = er.Read_Excel(er.StudentFilePath)
        totalStudents = originalStudentList.Count
        rightUnmatchedList = New List(Of Student)(originalStudentList)
        dh.BindData(dgUnmatchedRight, rightUnmatchedList)
        dgUnmatchedRight.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
        dgUnmatchedRight.Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
        dgUnmatchedRight.Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable
        dgUnmatchedRight.Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
        dgUnmatchedRight.Sort(dgUnmatchedRight.Columns(1), ListSortDirection.Ascending)

        'Switch the current panel
        LaunchPanel.Hide()
        ComparisonPanel.Show()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub BtnMatch_Click(sender As Object, e As EventArgs) Handles btnMatch.Click
        progBar1.Visible = True
        lblProg.Visible = True
        dgUnmatchedLeft.Enabled = False
        dgUnmatchedRight.Enabled = False
        bgwrkMatching.RunWorkerAsync()
    End Sub

    Private Sub tabControl_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tcLeft.SelectedIndexChanged, tcRight.SelectedIndexChanged
        Dim tabControl As TabControl = CType(sender, TabControl)
        Dim index As Integer = tabControl.SelectedIndex
        If index = 1 Then
            btnConfirm.Visible = True
        Else
            btnConfirm.Visible = False
        End If
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

    Private Sub DgMatchedRight_Sorted(sender As Object, e As EventArgs) Handles dgMatchedRight.Sorted
        Dim sortCol As Integer = 0
        If dgMatchedLeft.SortOrder <> Nothing Then
            sortCol = dgMatchedRight.SortedColumn.Index + 4
            If dgMatchedLeft.SortedColumn.Index <> sortCol Or dgMatchedLeft.SortOrder <> dgMatchedRight.SortOrder Then
                dgMatchedLeft.Sort(dgMatchedLeft.Columns(sortCol), dgMatchedRight.SortOrder - 1)
            End If
        Else
            sortCol = dgMatchedRight.SortedColumn.Index + 4
            dgMatchedLeft.Sort(dgMatchedLeft.Columns(sortCol), dgMatchedRight.SortOrder - 1)
        End If
    End Sub

    Private Sub DgUnmatchedLeft_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgUnmatchedLeft.CellClick
        Dim row As DataGridViewRow = dgUnmatchedLeft.SelectedRows(0)
        Dim selectedStudent As Student = row.DataBoundItem.Object
        er.Partial_Match_Similarity(selectedStudent, rightUnmatchedList)
        dgUnmatchedRight.Sort(dgUnmatchedRight.Columns(4), ListSortDirection.Descending)
    End Sub

    Private Sub BtnConfirm_Click(sender As Object, e As EventArgs) Handles btnConfirm.Click
        Dim studentA As Student = dgUnmatchedLeft.SelectedRows(0).DataBoundItem.Object
        dh.RemoveFromGrid(studentA, dgUnmatchedLeft, leftUnmatchedList)
        dh.AddToGrid(studentA, dgMatchedLeft, leftMatchedList)

        Dim studentB As Student = dgUnmatchedRight.SelectedRows(0).DataBoundItem.Object
        dh.RemoveFromGrid(studentB, dgUnmatchedRight, rightUnmatchedList)
        studentB.Result = studentA.Result
        studentB.MatchLastName = studentA.LastName
        studentA.MatchStudentNumber = studentB.StudentNumber
        dh.AddToGrid(studentB, dgMatchedRight, rightMatchedList)

        If dgUnmatchedLeft.Rows.Count > 0 Then
            Dim row As DataGridViewRow = dgUnmatchedLeft.SelectedRows(0)
            Dim selectedStudent As Student = row.DataBoundItem.Object
            er.Partial_Match_Similarity(selectedStudent, rightUnmatchedList)
            dgUnmatchedRight.Sort(dgUnmatchedRight.Columns(4), ListSortDirection.Descending)
        Else
            btnConfirm.Enabled = False
        End If
    End Sub

    Private Sub DgMatchedLeft_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles dgMatchedLeft.RowsAdded
        UpdateMatchCount()
    End Sub

    Private Sub DgMatchedLeft_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs) Handles dgMatchedLeft.RowsRemoved
        UpdateMatchCount()
    End Sub

    Private Sub UpdateMatchCount()
        numMatched = dgMatchedLeft.Rows.Count
        lblMatched.Text = "Matched: " & numMatched & "/" & totalStudents
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgwrkMatching.DoWork

        leftIDs = New List(Of Integer)
        rightIDs = New List(Of Integer)
        Dim iteration As Integer = 0
        For Each rStudent As Student In leftUnmatchedList
            Dim ids = er.Match(rStudent, rightUnmatchedList)
            If Not ids Is Nothing Then
                leftIDs.Add(ids(0))
                rightIDs.Add(ids(1))
            End If
            iteration += 1
            bgwrkMatching.ReportProgress(iteration / totalStudents * 100)
            Console.WriteLine("Matching: {0}/{1}", iteration, totalStudents)
        Next
    End Sub

    Private Sub BgwrkMatching_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles bgwrkMatching.ProgressChanged
        progBar1.Value = e.ProgressPercentage
        lblProg.Text = String.Format("Matching {0}/{1}", e.ProgressPercentage * totalStudents / 100, totalStudents)
    End Sub

    Private Sub BgwrkMatching_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles bgwrkMatching.RunWorkerCompleted
        leftMatchedList = New List(Of Student)
        rightMatchedList = New List(Of Student)
        lblMatched.Visible = True
        Console.WriteLine("Finished Matching")

        dh.AddIds(leftMatchedList, leftUnmatchedList, leftIDs)
        dh.BindData(dgMatchedLeft, leftMatchedList)
        dgMatchedLeft.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
        dgMatchedLeft.Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
        dgMatchedLeft.Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable
        dgMatchedLeft.Columns(2).FillWeight = 1.1F
        dgMatchedLeft.Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
        Console.WriteLine("Added to left matched list")

        dh.AddIds(rightMatchedList, rightUnmatchedList, rightIDs)
        dh.BindData(dgMatchedRight, rightMatchedList)
        dgMatchedRight.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
        dgMatchedRight.Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
        dgMatchedRight.Sort(dgMatchedRight.Columns(1), ListSortDirection.Ascending)
        Console.WriteLine("Added to right matched list")

        bgwrkDataHandler.RunWorkerAsync()
    End Sub

    Private Sub BgwrkDataHandler_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgwrkDataHandler.DoWork
        Dim count As Integer = 0
        Dim total As Integer = leftIDs.Count + rightIDs.Count
        For Each id As Integer In leftIDs
            dh.RemoveFromGridByID(id, dgUnmatchedLeft, leftUnmatchedList)
            count += 1
            bgwrkDataHandler.ReportProgress(count / total * 100)
            Console.WriteLine("Removing from left: {0}/{1}", count, total)
        Next

        For Each id As Integer In rightIDs
            dh.RemoveFromGridByID(id, dgUnmatchedRight, rightUnmatchedList)
            count += 1
            bgwrkDataHandler.ReportProgress(count / total * 100)
            Console.WriteLine("Removing from right: {0}/{1}", count, total)
        Next
    End Sub

    Private Sub BgwrkDataHandler_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles bgwrkDataHandler.ProgressChanged
        progBar1.Value = e.ProgressPercentage
        lblProg.Text = String.Format("Updating Data...")
    End Sub

    Private Sub BgwrkDataHandler_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles bgwrkDataHandler.RunWorkerCompleted
        Console.WriteLine("Removing finished")
        tcLeft.TabPages.Insert(0, tabMatched1)
        tcRight.TabPages.Insert(0, tabMatched2)
        tcLeft.SelectTab(tabMatched1)
        tcRight.SelectTab(tabMatched2)

        dgUnmatchedRight.Columns(4).Visible = True
        dgUnmatchedRight.Columns(4).AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        dgUnmatchedRight.Columns(4).Width = 50
        dgUnmatchedRight.Columns(4).DefaultCellStyle.Format = "0.##%"

        Dim row As DataGridViewRow = dgUnmatchedLeft.SelectedRows(0)
        Dim selectedStudent As Student = row.DataBoundItem.Object
        er.Partial_Match_Similarity(selectedStudent, rightUnmatchedList)
        dgUnmatchedRight.Sort(dgUnmatchedRight.Columns(4), ListSortDirection.Descending)

        dgUnmatchedLeft.Enabled = True
        dgUnmatchedRight.Enabled = True

        dgUnmatchedLeft.Columns(1).SortMode = DataGridViewColumnSortMode.Automatic
        dgUnmatchedLeft.Columns(2).SortMode = DataGridViewColumnSortMode.Automatic

        progBar1.Visible = False
        lblProg.Visible = False
        lblMatched.Visible = True
        UpdateMatchCount()
        btnMatch.Visible = False
    End Sub

    Private Sub DgUnmatchedRight_CellPainting(sender As Object, e As DataGridViewCellPaintingEventArgs) Handles dgUnmatchedRight.CellPainting
        If e.RowIndex = -1 Then
            e.Paint(e.CellBounds, DataGridViewPaintParts.All And Not DataGridViewPaintParts.ContentBackground)
            e.Handled = True
        End If
    End Sub
End Class
