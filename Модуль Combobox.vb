Imports System.Data.OleDb
Module Модуль_Combobox

    Public Sub Combo_box(c_box As ComboBox, command As String, d_member As String, v_member As String)
        Try
            Dim dt As New DataTable()
            Dim rs As New OleDbDataAdapter(command, db_con)
            rs.Fill(dt)
            c_box.DataSource = dt
            c_box.DisplayMember = (d_member)
            c_box.ValueMember = v_member
            c_box.SelectedIndex = -1
            c_box.Update()
            c_box.Refresh()
            rs.Dispose()
        Catch ex As Exception
            db_con.Close()
        End Try
    End Sub
    ' "Select Last_name + ' ' + First_name + ' ' + Midl_name as Familia, Tab_n from Stuff"

End Module
