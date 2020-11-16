Imports System.Data.OleDb
Module Модуль_Output_Data
    Public db_path As String = Application.StartupPath
    'Public db_path As String = "\\mgts-dfs-05.mgts.corp.net\Folders$\117925"
    Dim bd_name As String = "SuperJet 2.0"
    Dim db_pass As String = "" 'пароль базы данных
    Dim db_con As OleDbConnection = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & db_path & "\" & bd_name & ".mdb" & ";Jet OLEDB:Database Password=" & db_pass)
    Dim db_com As New OleDbCommand
    Public Sub Output_DGV(dgv As DataGridView, i As String)
        Dim db_table As New DataTable
        Try
            db_con.Open()
            Dim db_adapter As New OleDbDataAdapter(i, db_con)
            db_adapter.Fill(db_table)
            dgv.DataSource = db_table
            db_con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            db_con.Close()
        End Try
    End Sub
End Module
