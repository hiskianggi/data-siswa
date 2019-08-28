Imports System.Data.OleDb
Public Class Form3
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call koneksi()
        If conn.State <> ConnectionState.Closed Then conn.Close()
        conn.Open()
        cmd = New OleDbCommand("SELECT * From @U_P  WHERE us = '" & TextBox1.Text & "' and pw = '" & TextBox2.Text & "'", conn)
        dr = cmd.ExecuteReader()
        If (dr.HasRows()) Then
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox1.Focus()
            Me.Hide()
            form1.Show()
        Else
            MsgBox("Username & Password Anda Salah!", MessageBoxButtons.OK, "Login gagal")
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox1.Focus()
        End If
    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Focus()
        koneksi()
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call koneksi()
            If conn.State <> ConnectionState.Closed Then conn.Close()
            conn.Open()
            cmd = New OleDbCommand("SELECT * From @U_P  WHERE us = '" & TextBox1.Text & "' and pw = '" & TextBox2.Text & "'", conn)
            dr = cmd.ExecuteReader()
            If (dr.HasRows()) Then
                Me.Hide()
                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox1.Focus()
                form1.Show()
            Else
                MsgBox("Username & Password Anda Salah!", MessageBoxButtons.OK, "Login gagal")
                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox1.Focus()
            End If
        End If
    End Sub
End Class