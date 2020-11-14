Imports System.Data.OleDb
Imports System.Text
Imports System.Security.Cryptography
Module Модуль_Input
    Dim db_com As New OleDbCommand
    Public Sub rec_datateble(ByVal db_table As DataTable, ByRef pb_bar As ProgressBar)
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
            y = y + 1
            pb_bar.Value = y
            pb_bar.PerformStep()
        Next
    End Sub

    Public Sub rec_datateble_TaskList(ByVal db_table As DataTable, ByRef pb_bar As ProgressBar)
        Dim CurentRow As DataRow
        Dim y As Integer
        For i As Integer = 0 To db_table.Rows.Count - 1 Step 1
            CurentRow = db_table.Rows.Item(i)
            Dim CurentCellArray() = CurentRow.ItemArray
            Dim mass_storage() = {CurentCellArray(0), CurentCellArray(1), CurentCellArray(2), CurentCellArray(3), CurentCellArray(4), CurentCellArray(5), CurentCellArray(6),
            CDate(CurentCellArray(7)), CurentCellArray(8), CurentCellArray(9), CurentCellArray(10), CurentCellArray(11), CurentCellArray(12), CurentCellArray(13), CurentCellArray(14),
            CurentCellArray(15), CurentCellArray(16), CurentCellArray(17), CurentCellArray(18), CurentCellArray(19), CurentCellArray(20), CurentCellArray(21), CurentCellArray(22),
            CurentCellArray(23), CurentCellArray(24)}
            db_com.Connection = db_con
            db_con.Open()
            db_com.CommandText = "insert into time_table ( UIN, WFMN, TELN, Ispolnitel, Status, Type, Time_on, DTCS, Nevipolnenye, Comment, Area, Time_start, PhoneN, Address, 
Source, Internet_tarif, Internet_status, TV_tarif, TV_status, TV1_tarif, TV1_status, TV2_tarif, TV2_status, Povtorki, Priznak) 
values (" & mass_storage(0) & ", " & mass_storage(1) & ", " & mass_storage(2) & ", '" & mass_storage(3) & "', '" & mass_storage(4) & "', '" & mass_storage(5) & "', 
'" & mass_storage(6) & "', '" & mass_storage(7) & "', '" & mass_storage(8) & "', '" & mass_storage(9) & "', '" & mass_storage(10) & "', '" & mass_storage(11) & "', 
'" & mass_storage(12) & "', '" & mass_storage(13) & "', '" & mass_storage(14) & "', '" & mass_storage(15) & "', '" & mass_storage(16) & "', '" & mass_storage(17) & "', 
'" & mass_storage(18) & "', '" & mass_storage(19) & "', '" & mass_storage(20) & "', '" & mass_storage(21) & "', '" & mass_storage(22) & "', " & mass_storage(23) & ", 
'" & mass_storage(24) & "')"
            db_com.ExecuteNonQuery()
            db_con.Close()
            'Debug.WriteLine(TypeName(mass_storage(7)))
            Форма_загрузки_TaskList.ProgressBar1.Minimum = 0
            Форма_загрузки_TaskList.ProgressBar1.Maximum = db_table.Rows.Count
            y += 1
            Форма_загрузки_TaskList.ProgressBar1.Value = y
            Форма_загрузки_TaskList.ProgressBar1.PerformStep()
            Форма_загрузки_TaskList.TextBox1.Text = y
        Next
    End Sub

End Module
