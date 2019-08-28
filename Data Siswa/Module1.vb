Imports System.Data.OleDb
Module Module1
    Public conn As New OleDbConnection
    Public cmd As OleDbCommand
    Public da As OleDbDataAdapter
    Public ds As DataSet
    Public dr As OleDbDataReader
    Sub koneksi()
        conn = New OleDbConnection("provider=microsoft.jet.oledb.4.0;Data Source=db_datasiswa.mdb;")
        conn.Open()
    End Sub
End Module