Imports System.Text.RegularExpressions

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ListView1.Items.Add("Some thing blah")
        ListView1.Items.Add("Beep beep")
        ListView1.Items.Add("Boop boop")
        ListView1.Items.Add("Beep boop")
    End Sub

    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView1.SelectedIndexChanged
        MsgBox($"Selected: {ListView1.SelectedItems(0).Text}")
    End Sub

    Private Sub Search_Click(sender As Object, e As EventArgs) Handles Search.Click
        Dim searchPattern As String = InputBox("Enter search pattern")
        Dim matchingItems = ListView1.Items.Cast(Of ListViewItem).Where(Function(item As ListViewItem) As Boolean
                                                                            Return Regex.IsMatch(item.Text, searchPattern)
                                                                        End Function)

        matchingItems.ToList().ForEach(Sub(item As ListViewItem)
                                           item.Selected = True
                                       End Sub)
    End Sub

End Class
