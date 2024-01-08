Public Class sortTest
    Private lvwColumnSorter As ListViewColumnSorter

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Create an instance of a ListView column sorter and assign it 
        ' to the ListView control.
        lvwColumnSorter = New ListViewColumnSorter()
        Me.ListView1.ListViewItemSorter = lvwColumnSorter

    End Sub


    Private Sub sortTest_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim columnheader As ColumnHeader    ' Used for creating column headers.
        Dim listviewitem As ListViewItem    ' Used for creating ListView items.

        ' Make sure that the view is set to show details.
        ListView1.View = View.Details

        ' Create some ListView items consisting of first and last names.
        listviewitem = New ListViewItem("Mike")
        listviewitem.SubItems.Add("Nash")
        listviewitem.SubItems.Add(40)
        Me.ListView1.Items.Add(listviewitem)

        listviewitem = New ListViewItem("Kim")
        listviewitem.SubItems.Add("Abercrombie")
        listviewitem.SubItems.Add(22)
        Me.ListView1.Items.Add(listviewitem)

        listviewitem = New ListViewItem("Sunil")
        listviewitem.SubItems.Add("Koduri")
        listviewitem.SubItems.Add(9)
        Me.ListView1.Items.Add(listviewitem)

        listviewitem = New ListViewItem("Birgit")
        listviewitem.SubItems.Add("Seidl")
        listviewitem.SubItems.Add(33)
        Me.ListView1.Items.Add(listviewitem)

        ' Create some column headers for the data.
        columnheader = New ColumnHeader()
        columnheader.Text = "First Name"
        Me.ListView1.Columns.Add(columnheader)

        columnheader = New ColumnHeader()
        columnheader.Text = "Last Name"
        Me.ListView1.Columns.Add(columnheader)

        columnheader = New ColumnHeader()
        columnheader.Text = "Age"
        Me.ListView1.Columns.Add(columnheader)

        ' Loop through and size each column header to fit the column header text.
        For Each columnheader In Me.ListView1.Columns
            columnheader.Width = -2
        Next

    End Sub

    Private Sub ListView1_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles ListView1.ColumnClick
        ' Determine if the clicked column is already the column that is 
        ' being sorted.
        If (e.Column = lvwColumnSorter.SortColumn) Then
            ' Reverse the current sort direction for this column.
            If (lvwColumnSorter.Order = SortOrder.Ascending) Then
                lvwColumnSorter.Order = SortOrder.Descending
            Else
                lvwColumnSorter.Order = SortOrder.Ascending
            End If
        Else
            ' Set the column number that is to be sorted; default to ascending.
            lvwColumnSorter.SortColumn = e.Column
            lvwColumnSorter.Order = SortOrder.Ascending
        End If

        ' Perform the sort with these new sort options.
        Me.ListView1.Sort()

    End Sub

    Private Sub ListView1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged

    End Sub
End Class