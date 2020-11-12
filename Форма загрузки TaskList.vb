Imports System.Data.OleDb
Imports System.ComponentModel
Imports System.Threading
Public Class Форма_загрузки_TaskList
    Dim a As Integer
    Dim update_count, insert_count As Integer
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim db_table As New DataTable
        Dim OPFADI As New OpenFileDialog
        Dim status As Boolean
        OPFADI.Title = "Открываем TaskList"
        OPFADI.FileName = "Tasklist"
        OPFADI.Filter = "Файлы Excel | *.xls"
        If OPFADI.ShowDialog = DialogResult.OK Then
            TaskList_con(OPFADI.FileName, db_table, status)
            If status = True Then
                MsgBox("Загрузка данных прошла успешно. Загружено - " & db_table.Rows.Count & " строк")
                Временная_таблица("DELETE FROM [time_table]")
                rec_datateble_TaskList(db_table, ProgressBar1)
            End If
        End If
        '        update_count = Count_data(a, "UPDATE [TaskList Total] INNER JOIN [time_table] On [TaskList Total].[UIN] = [time_table].[UIN] Set [TaskList Total].[Status] = [time_table].[Status], 
        '[TaskList Total].[Ispolnitel] = [time_table].[Ispolnitel]")
        '        insert_count = Count_data(a, "Insert into [TaskList Total] Select * from time_table where [UIN] not in (Select [UIN] from [TaskList Total])")
        '        TextBox1.Text = update_count
        '        TextBox13.Text = insert_count
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Rec(DataGridView3)
        Input_string("Insert into [Equipment] Select * from [time_table_UTP] where [time_table_UTP].[Doc_sklad] not in (Select [Doc_sklad] from [Equipment])")
        Count_data(TextBox4.Text, "Select [UIN] from [TaskList Total]")
        Output_DGV(DataGridView3, "Select * from [TaskList Total]")
        Временная_таблица("DELETE FROM [time_table]")
        ProgressBar1.Value = 0
        Count_data(a, "Select [UIN] from [time_table]")
        TextBox1.Text = a
        Count_data(a, "Select [UIN] from [TaskList Total]")
        TextBox4.Text = a
        Count_status()
        Output_DGV(DataGridView3, "Select * from [TaskList Total]")
    End Sub

    Private Sub Форма_Загрузка_TskList_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Count_data(a, "Select [UIN] from [TaskList Total]")
        TextBox4.Text = a
        Count_status()
    End Sub

    Sub Count_status()

        Count_data(a, "Select [UIN] from [TaskList Total] where [Status] = 'Выполнена'")
        TextBox5.Text = a

        Count_data(a, "Select [UIN] from [TaskList Total] where [Status] = 'Не выполнена'")
        TextBox2.Text = a

        Count_data(a, "Select [UIN] from [TaskList Total] where [Status] = 'Запланирована'")
        TextBox6.Text = a

        Count_data(a, "Select [UIN] from [TaskList Total] where [Status] = 'Назначена'")
        TextBox7.Text = a

        Count_data(a, "Select [UIN] from [TaskList Total] where [Status] = 'Взята в работу'")
        TextBox8.Text = a

        Count_data(a, "Select [UIN] from [TaskList Total] where [Status] = 'В работе'")
        TextBox9.Text = a

        Count_data(a, "Select [UIN] from [TaskList Total] where [Status] = 'Завершено исполнение'")
        TextBox10.Text = a

        Count_data(a, "Select [UIN] from [TaskList Total] where [Status] = 'Отменена'")
        TextBox11.Text = a

        Count_data(a, "Select [UIN] from [TaskList Total] where [Status] = 'Открыта'")
        TextBox12.Text = a

    End Sub

    Private Sub Форма_загрузки_TaskList_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Временная_таблица("DELETE FROM [time_table]")
        Временная_таблица("DELETE FROM [time_table_UTP]")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim db_table As New DataTable
        Dim OPFADI As New OpenFileDialog
        Dim status As Boolean
        OPFADI.Title = "Открываем UTP"
        OPFADI.FileName = "UTP"
        OPFADI.Filter = "Файлы Excel | *.xls"
        If OPFADI.ShowDialog = DialogResult.OK Then
            excel_con(OPFADI.FileName, db_table, status)
            If status = True Then
                MsgBox("Загрузка данных прошла успешно. Загружено - " & db_table.Rows.Count & " строк")
                Временная_таблица("DELETE FROM [time_table_UTP]")
                rec_datateble(db_table, ProgressBar1)
            End If
        End If
    End Sub

End Class
