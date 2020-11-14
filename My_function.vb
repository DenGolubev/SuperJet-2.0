Imports System.Data.OleDb
Class My_function
    Private db_com As New OleDbCommand
    Public Function Rec_datateble_new(ByVal db_table As DataTable, ByRef pb_bar As ProgressBar)
        Dim CurentRow As DataRow
        Dim y As Integer
        For i As Integer = 0 To db_table.Rows.Count - 1 Step 1
            CurentRow = db_table.Rows.Item(i)
            Dim CurentCellArray() = CurentRow.ItemArray
            Dim mass_storage() = {CurentCellArray(0), CDate(CurentCellArray(1)), CurentCellArray(2), CurentCellArray(3), CurentCellArray(4), CurentCellArray(5), CurentCellArray(6),
            CurentCellArray(7), CurentCellArray(8), CurentCellArray(9)}
            Dim a As String = getMd5Hash(mass_storage(0))
            Debug.WriteLine(a)
            db_com.Connection = db_con
            db_con.Open()
            db_com.CommandText = "insert into [time_table_UTP] (Doc_sklad, Date_time, Sklad, Kod_sklad, Area, Tab_N, Ispolnitel, Kod_OEBS, Oborudovanie, Kolichestvo) 
            values ('" & a & "','" & mass_storage(1) & "', '" & mass_storage(2) & "', '" & mass_storage(3) & "', '" & mass_storage(4) & "', '" & mass_storage(5) & "', 
            '" & mass_storage(6) & "', '" & mass_storage(7) & "', '" & mass_storage(8) & "', '" & mass_storage(9) & "')"
            db_com.ExecuteNonQuery()
            db_con.Close()
            pb_bar.Minimum = 0
            pb_bar.Maximum = db_table.Rows.Count
            y += 1
            pb_bar.Value = y
            pb_bar.PerformStep()
        Next
    End Function

End Class
