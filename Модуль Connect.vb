Imports System.Data.OleDb
Module Модуль_Connect
    Public db_path As String = Application.StartupPath
    'Public db_path As String = "\\mgts-dfs-05.mgts.corp.net\Folders$\117925"
    'Public db_path As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
    Public bd_name As String = "SuperJet 2.0"
    Public db_pass As String = "" 'пароль базы данных
    Public db_con As OleDbConnection = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & db_path & "\" & bd_name & ".mdb" & ";Jet OLEDB:Database Password=" & db_pass)
    Public db_com As OleDbCommand
    Public db_adapter As New OleDbDataAdapter

    Sub con_open()
        'db_com.Connection = db_con
        db_con.Open()
    End Sub

    Sub con_close()
        'db_com.Connection = db_con
        db_con.Close()
    End Sub

    Sub con_status(индикатор As Button)
        Try
            If db_con.State = ConnectionState.Open Then
                индикатор.BackColor = Color.LightGreen
            Else
                индикатор.BackColor = Color.Red
            End If
        Catch ex As Exception
            индикатор.BackColor = Color.Red
            MsgBox(ex.Message & " " & "Скорее всего не инициализированно соединение с базой данных", MsgBoxStyle.Critical)
        End Try

    End Sub

End Module
