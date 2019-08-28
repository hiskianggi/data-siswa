Public Class Form2
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        System.Diagnostics.Process.Start("READ ME.txt")
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Me.LinkLabel1.LinkVisited = True
        System.Diagnostics.Process.Start("http://internet.zone.id")
    End Sub
End Class