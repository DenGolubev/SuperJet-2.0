Imports System.Data.OleDb
Module Модуль_создания_таблиц
    Public db_path As String = Application.StartupPath
    'Public db_path As String = "\\mgts-dfs-05.mgts.corp.net\Folders$\117925"
    Dim bd_name As String = "SuperJet 2.0"
    Dim db_pass As String = "" 'пароль базы данных
    Dim db_con As OleDbConnection = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & db_path & "\" & bd_name & ".mdb" & ";Jet OLEDB:Database Password=" & db_pass)
    Dim db_com As New OleDbCommand
    Public Sub Временная_таблица(i As String)
        Try
            db_com.Connection = db_con
            db_con.Open()
            db_com.CommandText = i
            db_com.ExecuteNonQuery()
            db_con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            db_con.Close()
        End Try
    End Sub

    Public Sub Поля_временной_таблицы(i As String)
        Try
            db_com.Connection = db_con
            db_con.Open()
            db_com.CommandText = i
            db_com.ExecuteNonQuery()
            db_con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            db_con.Close()
        End Try
    End Sub
End Module
