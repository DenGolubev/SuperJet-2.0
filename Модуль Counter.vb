Imports System.Data.OleDb
Module Модуль_Counter
    'Public db_path As String = "\\mgts-dfs-05.mgts.corp.net\Folders$\117925"
    Public db_path As String = Application.StartupPath
    Dim bd_name As String = "SuperJet 2.0"
    Dim db_pass As String = "" 'пароль базы данных
    Dim db_con As OleDbConnection = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & db_path & "\" & bd_name & ".mdb" & ";Jet OLEDB:Database Password=" & db_pass)
    Dim db_com As New OleDbCommand
    Public Function Count_data(ByRef count As Integer, i As String) As Integer
        Dim db_table As New DataTable
        Try
            db_con.Open()
            Dim db_adapter As New OleDbDataAdapter(i, db_con)
            db_adapter.Fill(db_table)
            count = db_table.Rows.Count
            If db_table.Rows.Count = 0 Then
                count = 0
            End If
            db_con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            db_con.Close()
        End Try
        Return count
    End Function

    Public Function Output_sum(ByVal i As String, ByRef a As Integer) As Integer
        Try
            db_com.Connection = db_con
            db_con.Open()
            db_com.CommandText = i
            a = db_com.ExecuteScalar
            db_con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            db_con.Close()
        End Try
        Return a
    End Function

    Public Function Count_data_sum(ByRef count As Integer, ByVal Col As Integer, i As String) As Integer
        Dim db_table As New DataTable
        Dim a As Integer
        Try
            db_con.Open()
            Dim db_adapter As New OleDbDataAdapter(i, db_con)
            db_adapter.Fill(db_table)
            For y As Integer = 0 To db_table.Rows.Count - 1
                If db_table.Rows(y)(Col) IsNot Nothing Then
                    a += db_table.Rows(y)(Col)
                Else
                    a = 0
                End If
            Next
            count = a
            db_con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            db_con.Close()
        End Try
        Return count
    End Function

    Public Function TT_time_percent(ByRef value As Integer, ByVal all_tt As Integer, ByVal all_tt_good As Integer) As Integer
        If all_tt = 0 Then
            value = 0
        Else
            value = all_tt_good / all_tt
        End If
        Return value
    End Function

    Public Function Text_Box_Count(ByRef a As Double, ByVal text_box1 As TextBox, ByVal text_box2 As TextBox) As Integer
        Dim b, c As Integer
        b = CInt(text_box1.Text)
        c = CInt(text_box2.Text)
        If b = 0 Or c = 0 Then
            a = 0
        Else
            a = b / c
        End If
        Return a
    End Function

    Public Function konvertacia(ByRef a As Double, ByVal b As Integer, ByVal c As Integer)
        If b = 0 Or c = 0 Then
            a = 0
        Else
            a = c / b
        End If
        Return a
    End Function

    Public Function konvertacia_uslug(ByRef a As Double, ByVal b As Integer, ByVal c As Integer, ByVal d As Integer)
        If b = 0 Or c + d = 0 Then
            a = 0
        Else
            a = (d + c) / b
        End If
        Return a
    End Function

    Public Function kabel(ByRef a As Integer, ByVal b As Integer, ByVal c As Integer, ByVal d As Integer, ByVal e As Integer, ByVal f As Integer)
        a = (b * 60) + (c * 15) + (d * 2) + (e * 15) + (f * 2)
        Return a
    End Function

    
End Module
