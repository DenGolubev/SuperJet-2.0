Public Class Форма_36_часов
    Dim d1, d2 As New Date
    Private Sub Форма_36_часов_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Button1.Hide()
        Label4.Text = Now().ToShortDateString
        CheckBox1.Text = "Текущий месяц"
        DateTimePicker1.Hide()
        DateTimePicker2.Hide()
        Combo_box(ComboBox1, "Select Last_name + ' ' + First_name + ' ' + Midl_name as Familia, Tab_n from Stuff", "Familia", "Tab_n")
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        checked_check(ON_Month(d1), OFF_Month(d2), CheckBox1, DateTimePicker1, DateTimePicker2)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        TextBox5.Text = Convert.ToString(ComboBox1.SelectedValue)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        RadioButton1.Checked = False
        RadioButton2.Checked = False
        RadioButton3.Checked = False
        Dim a As Integer
        TextBox1.Text = Count_data(a, "Select UIN from [TaskList Total] where Type in ('Обслуживание FTTB', 'Обслуживание PON (weld)', 'Обслуживание абонентов ADSL', 'Обслуживание абонентов OTA',
        'Обслуживание абонентов OTA (ЛАО - Загород)', 'Обслуживание абонентов PON', 'Обслуживание ЦТВ') and Status in ('Выполнена') and [Ispolnitel] LIKE '%" & TextBox5.Text & "%' 
        and [DTCS] BETWEEN CDATE('" & DateTimePicker1.Value.ToShortDateString & "') and CDATE('" & DateTimePicker2.Value.ToShortDateString & "')")
        TextBox2.Text = Count_data(a, "Select UIN from [TaskList Total] where Type in ('Обслуживание FTTB', 'Обслуживание PON (weld)', 'Обслуживание абонентов ADSL', 'Обслуживание абонентов OTA',
        'Обслуживание абонентов OTA (ЛАО - Загород)', 'Обслуживание абонентов PON', 'Обслуживание ЦТВ') and Status in ('Запланирована') and [Ispolnitel] LIKE '%" & TextBox5.Text & "%' 
        and [DTCS] BETWEEN CDATE('" & DateTimePicker1.Value.ToShortDateString & "') and CDATE('" & DateTimePicker2.Value.ToShortDateString & "')")
        TextBox3.Text = Count_data(a, "Select UIN from [TaskList Total] where Type in ('Обслуживание FTTB', 'Обслуживание PON (weld)', 'Обслуживание абонентов ADSL', 'Обслуживание абонентов OTA',
        'Обслуживание абонентов OTA (ЛАО - Загород)', 'Обслуживание абонентов PON', 'Обслуживание ЦТВ') and Status in ('Выполнена') and [Ispolnitel] LIKE '%" & TextBox5.Text & "%' 
        and DateDiff('h',Time_on ,Time_start) <= 36 and [DTCS] BETWEEN CDATE('" & DateTimePicker1.Value.ToShortDateString & "') and CDATE('" & DateTimePicker2.Value.ToShortDateString & "')")
        TextBox7.Text = Count_data(a, "Select UIN from [TaskList Total] where Type in ('Обслуживание FTTB', 'Обслуживание PON (weld)', 'Обслуживание абонентов ADSL', 'Обслуживание абонентов OTA',
        'Обслуживание абонентов OTA (ЛАО - Загород)', 'Обслуживание абонентов PON', 'Обслуживание ЦТВ') and Status in ('Выполнена') and [Ispolnitel] LIKE '%" & TextBox5.Text & "%' 
        and DateDiff('h',Time_on ,Time_start) <= 24 and [DTCS] BETWEEN CDATE('" & DateTimePicker1.Value.ToShortDateString & "') and CDATE('" & DateTimePicker2.Value.ToShortDateString & "')")
        TextBox8.Text = Count_data(a, "Select UIN from [TaskList Total] where Type in ('Обслуживание FTTB', 'Обслуживание PON (weld)', 'Обслуживание абонентов ADSL', 'Обслуживание абонентов OTA',
        'Обслуживание абонентов OTA (ЛАО - Загород)', 'Обслуживание абонентов PON', 'Обслуживание ЦТВ') and Status in ('Выполнена') and [Ispolnitel] LIKE '%" & TextBox5.Text & "%' 
        and DateDiff('h',Time_on ,Time_start) > 36 and [DTCS] BETWEEN CDATE('" & DateTimePicker1.Value.ToShortDateString & "') and CDATE('" & DateTimePicker2.Value.ToShortDateString & "')")
        TextBox9.Text = Count_data(a, "Select UIN from [TaskList Total] where Type in ('Обслуживание FTTB', 'Обслуживание PON (weld)', 'Обслуживание абонентов ADSL', 'Обслуживание абонентов OTA',
        'Обслуживание абонентов OTA (ЛАО - Загород)', 'Обслуживание абонентов PON', 'Обслуживание ЦТВ') and Status in ('Выполнена') and [Ispolnitel] LIKE '%" & TextBox5.Text & "%' 
        and DateDiff('h',Time_on ,Time_start) > 24 and [DTCS] BETWEEN CDATE('" & DateTimePicker1.Value.ToShortDateString & "') and CDATE('" & DateTimePicker2.Value.ToShortDateString & "')")

        Output_DGV(DataGridView1, "Select  DateDiff('h',Time_on ,Time_start) as 'Время жизни в часах', TELN, Ispolnitel, Status, Type, Time_on, Time_start, Address from [TaskList Total] 
        where Type in ('Обслуживание FTTB', 'Обслуживание PON (weld)', 'Обслуживание абонентов ADSL', 'Обслуживание абонентов OTA', 'Обслуживание абонентов OTA (ЛАО - Загород)', 'Обслуживание абонентов PON', 
        'Обслуживание ЦТВ') and Status in ('Выполнена') and DateDiff('h',Time_on , Time_start) > 36 and 
       [DTCS] BETWEEN CDATE('" & DateTimePicker1.Value.ToShortDateString & "') and CDATE('" & DateTimePicker2.Value.ToShortDateString & "') Order By DTCS")

        TextBox4.Text = Format(konvertacia(a, CInt(TextBox1.Text), CInt(TextBox3.Text)), "0.00 %")
        TextBox6.Text = Format(konvertacia(a, CInt(TextBox1.Text), CInt(TextBox7.Text)), "0.00 %")

        Button1.Show()

    End Sub



    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            Output_DGV(DataGridView1, "Select  DateDiff('h',Time_on ,Time_start) as 'Время жизни в часах', TELN, Ispolnitel, Status, Type, Time_on, Time_start, Address from [TaskList Total] 
    where Type in ('Обслуживание FTTB', 'Обслуживание PON (weld)', 'Обслуживание абонентов ADSL', 'Обслуживание абонентов OTA', 'Обслуживание абонентов OTA (ЛАО - Загород)', 'Обслуживание абонентов PON', 
    'Обслуживание ЦТВ') and Status in ('Запланирована') and DateDiff('h',Time_on ,'" & Now & "') > 12 Order By DTCS")
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            Output_DGV(DataGridView1, "Select  DateDiff('h',Time_on ,Time_start) as 'Время жизни в часах', TELN, Ispolnitel, Status, Type, Time_on, Time_start, Address from [TaskList Total] 
    where Type in ('Обслуживание FTTB', 'Обслуживание PON (weld)', 'Обслуживание абонентов ADSL', 'Обслуживание абонентов OTA', 'Обслуживание абонентов OTA (ЛАО - Загород)', 'Обслуживание абонентов PON', 
    'Обслуживание ЦТВ') and Status in ('Запланирована') and DateDiff('h',Time_on ,'" & Now & "') > 24 Order By DTCS")
        End If
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        Output_DGV(DataGridView1, "Select  DateDiff('h',Time_on ,Time_start) as 'Время жизни в часах', TELN, Ispolnitel, Status, Type, Time_on, Time_start, Address from [TaskList Total] 
    where Type in ('Обслуживание FTTB', 'Обслуживание PON (weld)', 'Обслуживание абонентов ADSL', 'Обслуживание абонентов OTA', 'Обслуживание абонентов OTA (ЛАО - Загород)', 'Обслуживание абонентов PON', 
    'Обслуживание ЦТВ') and Status in ('Выполнена') Order By DTCS")
    End Sub

End Class