Imports System.Data.OleDb
Public Class form1
    Sub TampilGrid()
        Call koneksi()
        da = New OleDbDataAdapter("select * from DataSiswa", conn)
        ds = New DataSet
        da.Fill(ds)
        DataGridView1.DataSource = ds.Tables(0)
        DataGridView1.ReadOnly = True
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        koneksi()
        TampilGrid()
        BersihkanTeks()
    End Sub

    Sub BersihkanTeks()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBoxCari.Clear()
        ComboBox6.Text = "Banyumanis"
        ComboBox3.Text = "RPL"
        ComboBoxCari.Text = "Nama"
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Then
            MessageBox.Show("Salah Satu atau Beberapa Kolom Belum Terisi!!", "Aplikasi Data Siswa", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
        Else
            Call koneksi()
            Dim simpan As String
            simpan = "insert into DataSiswa values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & ComboBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & ComboBox6.Text & "')"
            cmd = New OleDbCommand(simpan, conn)
            cmd.ExecuteNonQuery()
            MessageBox.Show("Tersimpan!", "Aplikasi Data Siswa", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            BersihkanTeks()
            TampilGrid()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox2.Text = "" Then
            MessageBox.Show("Anda Belum Memilih Data Yang Akan Dihapus!", "Aplikasi Data Siswa", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
        Else
            If MessageBox.Show("Yakin Ingin Menghapus" & vbNewLine & "" & vbNewLine & "         NIS = " & TextBox2.Text & "?", "Peringatan!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                Call koneksi()
                cmd = New OleDbCommand("delete from DataSiswa where NIS ='" & TextBox2.Text & "'", conn)
                cmd.ExecuteNonQuery()
                MessageBox.Show("Terhapus!", "Aplikasi Data Siswa", MessageBoxButtons.OK, MessageBoxIcon.Information)
                BersihkanTeks()
                TampilGrid()
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        On Error Resume Next
        TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        TextBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
        ComboBox3.Text = DataGridView1.Rows(e.RowIndex).Cells(2).Value
        TextBox4.Text = DataGridView1.Rows(e.RowIndex).Cells(3).Value
        TextBox5.Text = DataGridView1.Rows(e.RowIndex).Cells(4).Value
        ComboBox6.Text = DataGridView1.Rows(e.RowIndex).Cells(5).Value
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or ComboBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or ComboBox6.Text = "" Then
            MessageBox.Show("Salah Satu atau Beberapa Kolom Belum Terisi", "Aplikasi Data Siswa", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
        Else
            Call koneksi()
            cmd = New OleDbCommand("update DataSiswa SET Nama='" & TextBox1.Text & "',KompetensiKeahlian='" & ComboBox3.Text & "',Alamat='" & TextBox4.Text & "',NoHP='" & TextBox5.Text & "',Rayon='" & ComboBox6.Text & "'WHERE NIS='" & TextBox2.Text & "'", conn)
            cmd.ExecuteNonQuery()
            MessageBox.Show("Updated Sucess!", "Aplikasi Data Siswa", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            BersihkanTeks()
            TampilGrid()
        End If
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If TextBox1.Text = "" And TextBox2.Text = "" And TextBox4.Text = "" And TextBox5.Text = "" And TextBoxCari.Text = "" And ComboBox6.Text = "Banyumanis" And ComboBox3.Text = "RPL" And ComboBoxCari.Text = "Nama" Then
            MessageBox.Show("Data Sudah Direset!", "Aplikasi Data Siswa", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            BersihkanTeks()
            TampilGrid()
        End If
    End Sub
    Sub Search()
        cmd = New OleDbCommand("select * from DataSiswa where " & ComboBoxCari.Text & " like '%" & TextBoxCari.Text & "%'", conn)
        dr = cmd.ExecuteReader
        dr.Read()
        koneksi()
        da = New OleDbDataAdapter("select * from DataSiswa where " & ComboBoxCari.Text & " like '%" & TextBoxCari.Text & "%'", conn)
        ds = New DataSet
        da.Fill(ds)
        If dr.HasRows Then
            ds = New DataSet
            da.Fill(ds)
            DataGridView1.DataSource = ds.Tables(0)
        Else
            MessageBox.Show("Item Tidak Ditemukan!", "Aplikasi Data Siswa", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox1.Select()
        End If
    End Sub
    Private Sub ButtonCari_Click(sender As Object, e As EventArgs) Handles ButtonCari.Click
        If TextBoxCari.Text = "" Or ComboBoxCari.Text = "" Then
            MessageBox.Show("Masukkan Kategori Dan Isi Pencarian!!!", "Aplikasi Data Siswa", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
        Else
            koneksi()
            Search()
        End If
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Not ((e.KeyChar >= "0" And e.KeyChar <= "9") Or e.KeyChar = vbBack) Then e.Handled() = True
    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        If Not ((e.KeyChar >= "0" And e.KeyChar <= "9") Or ((e.KeyChar = "+") Or e.KeyChar = vbBack)) Then e.Handled() = True
    End Sub

    Private Sub TextBoxCari_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBoxCari.KeyDown
        If e.KeyCode = Keys.Enter Then
            Search()
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Form2.Show()
    End Sub

    Private Sub DataGridView1_KeyDown(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyDown
        If e.KeyCode = Keys.Delete Then
            If MessageBox.Show("Yakin Ingin Menghapus" & vbNewLine & "" & vbNewLine & "         NIS = " & TextBox2.Text & "?", "Peringatan!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                Call koneksi()
                cmd = New OleDbCommand("delete from DataSiswa where NIS ='" & TextBox2.Text & "'", conn)
                cmd.ExecuteNonQuery()
                MessageBox.Show("Terhapus!", "Aplikasi Data Siswa", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
            TampilGrid()
            BersihkanTeks()
        End If
    End Sub

    Private Sub ComboBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox6.KeyDown
        If e.KeyCode = Keys.Enter Then
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Then

            Else
                Call koneksi()
                Dim simpan As String
                simpan = "insert into DataSiswa values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & ComboBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & ComboBox6.Text & "')"
                cmd = New OleDbCommand(simpan, conn)
                cmd.ExecuteNonQuery()
                MessageBox.Show("Tersimpan!", "Aplikasi Data Siswa", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                BersihkanTeks()
                TampilGrid()
            End If
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If MessageBox.Show("Yakin Ingin Log Out?", "Peringatan!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
            Me.Close()
            MessageBox.Show("Berhasil Log Out!", "Aplikasi Data Siswa", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Form3.Show()
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class
