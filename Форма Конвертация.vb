Imports System.Data.OleDb

Public Class Форма_Конвертация
    Dim d1, d2 As New Date
    Private Sub Форма_Конвертация_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ON_Month(d1)
        OFF_Month(d2)
        TextBox5.TextAlign = HorizontalAlignment.Center
        TextBox7.TextAlign = HorizontalAlignment.Center
        TextBox14.TextAlign = HorizontalAlignment.Center
        Combo_box(ComboBox1, "Select Last_name + ' ' + First_name + ' ' + Midl_name as Familia, Tab_n from Stuff", "Familia", "Tab_n")

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        New_PDF(DataGridView1)
        DGV_PDF(DataGridView1, ComboBox1)
    End Sub

    Private Sub Prosmotr()
        Dim a As Integer
        'Все выполненные наряды
        TextBox2.Text = Count_data(a, "Select [UIN] from [TaskList Total] where [Status] = ('Выполнена') and [Ispolnitel] LIKE '%" & TextBox1.Text & "%' 
        and [DTCS] BETWEEN CDATE('" & DateTimePicker1.Value.ToShortDateString & "') and CDATE('" & DateTimePicker2.Value.ToShortDateString & "')")

        'Все первички
        TextBox16.Text = Count_data(a, "Select [UIN] from [TaskList Total] where [Status] in ('Выполнена', 'Не выполнена') and [Type] in ('Подключение FTTB', 'Первичное переключение абонентов на PON', 
        'Подключение/переключение на PON с ОРК', 'Переключение с другого оператора', 'Инсталляция PON с ОРШ') and [Ispolnitel]  LIKE '%" & TextBox1.Text & "%' 
        and [DTCS] BETWEEN CDATE('" & DateTimePicker1.Value.ToShortDateString & "') and CDATE('" & DateTimePicker2.Value.ToShortDateString & "')")

        'Все выполненные первички
        TextBox3.Text = Count_data(a, "Select [UIN] from [TaskList Total] where [Status] = ('Выполнена') and [Type] in ('Подключение FTTB', 'Первичное переключение абонентов на PON', 
        'Подключение/переключение на PON с ОРК', 'Переключение с другого оператора', 'Инсталляция PON с ОРШ') and [Ispolnitel]  LIKE '%" & TextBox1.Text & "%' 
        and [DTCS] BETWEEN CDATE('" & DateTimePicker1.Value.ToShortDateString & "') and CDATE('" & DateTimePicker2.Value.ToShortDateString & "')")

        'Все не выполненные первички
        TextBox4.Text = Count_data(a, "Select [UIN] from [TaskList Total] where [Status] = ('Не выполнена') and [Type] in ('Подключение FTTB', 'Первичное переключение абонентов на PON', 
        'Подключение/переключение на PON с ОРК', 'Переключение с другого оператора', 'Инсталляция PON с ОРШ') and [Ispolnitel]  LIKE '%" & TextBox1.Text & "%' 
        and [DTCS] BETWEEN CDATE('" & DateTimePicker1.Value.ToShortDateString & "') and CDATE('" & DateTimePicker2.Value.ToShortDateString & "')")

        'Сданные наряды
        TextBox6.Text = Count_data(a, "Select [UIN] from [TaskList Total] where [Status_sdano] = 'Сдано' and [Ispolnitel]  LIKE '%" & TextBox1.Text & "%' 
        and [DTCS] BETWEEN CDATE('" & DateTimePicker1.Value.ToShortDateString & "') and CDATE('" & DateTimePicker2.Value.ToShortDateString & "')")

        'Наряды переданные курьеру
        TextBox20.Text = Count_data(a, "Select [UIN] from [TaskList Total] where [Status_sdano] = 'Передано курьеру' and [Ispolnitel]  LIKE '%" & TextBox1.Text & "%' 
        and [DTCS] BETWEEN CDATE('" & DateTimePicker1.Value.ToShortDateString & "') and CDATE('" & DateTimePicker2.Value.ToShortDateString & "')")

        'Возвраты наряды
        TextBox30.Text = Count_data(a, "Select [UIN] from [TaskList Total] where [Status_sdano] = 'Возврат' and [Ispolnitel]  LIKE '%" & TextBox1.Text & "%'")

        'Всего услуги интернет
        TextBox18.Text = Count_data(a, "Select [UIN] from [TaskList Total] where [Status] in ('Выполнена', 'Не выполнена') and [Type] in ('Подключение FTTB', 'Первичное переключение абонентов на PON', 
        'Подключение/переключение на PON с ОРК', 'Переключение с другого оператора', 'Инсталляция PON с ОРШ', 'Подключение услуг на PON') and [Internet_tarif] not in('') 
        and [Internet_status]  in ( 'Выполнена', 'Не выполнена', '') and [Ispolnitel]  LIKE '%" & TextBox1.Text & "%' and [DTCS] BETWEEN CDATE('" & DateTimePicker1.Value.ToShortDateString & "') 
        and CDATE('" & DateTimePicker2.Value.ToShortDateString & "')")

        'Всего услуги ТВ
        TextBox19.Text = Count_data(a, "Select [UIN] from [TaskList Total] where [Status] in ('Выполнена', 'Не выполнена') and [Type] in ('Подключение FTTB', 'Первичное переключение абонентов на PON', 
        'Подключение/переключение на PON с ОРК', 'Переключение с другого оператора', 'Инсталляция PON с ОРШ', 'Подключение услуг на PON') and [TV_tarif] not in ('') 
        and [TV_status] in ('Выполнена', 'Не выполнена', '') and [Ispolnitel]  LIKE '%" & TextBox1.Text & "%' and [DTCS] BETWEEN CDATE('" & DateTimePicker1.Value.ToShortDateString & "') 
        and CDATE('" & DateTimePicker2.Value.ToShortDateString & "')")

        'Выполнено интернет ФТТБ
        TextBox8.Text = Count_data(a, "Select [UIN] from [TaskList Total] where [Status] = ('Выполнена') and [Type] = ('Подключение FTTB') and [Internet_status] = ('Выполнена') 
        and [Ispolnitel]  LIKE '%" & TextBox1.Text & "%' and [DTCS] BETWEEN CDATE('" & DateTimePicker1.Value.ToShortDateString & "') and CDATE('" & DateTimePicker2.Value.ToShortDateString & "')")

        'Выполнено ТВ ФТТБ
        TextBox9.Text = Count_data(a, "Select [UIN] from [TaskList Total] where [Status] = ('Выполнена') and [Type] = ('Подключение FTTB') and [TV_status] = ('Выполнена') 
        and [Ispolnitel]  LIKE '%" & TextBox1.Text & "%' and [DTCS] BETWEEN CDATE('" & DateTimePicker1.Value.ToShortDateString & "') and CDATE('" & DateTimePicker2.Value.ToShortDateString & "')")

        'Выполнено интернет ПОН
        TextBox13.Text = Count_data(a, "Select [UIN] from [TaskList Total] where [Status] = ('Выполнена') and [Type] in ('Первичное переключение абонентов на PON', 'Подключение/переключение на PON с ОРК', 
        'Переключение с другого оператора', 'Инсталляция PON с ОРШ', 'Подключение услуг на PON') and [Internet_status] = ('Выполнена') and [Ispolnitel]  LIKE '%" & TextBox1.Text & "%' 
        and [DTCS] BETWEEN CDATE('" & DateTimePicker1.Value.ToShortDateString & "') and CDATE('" & DateTimePicker2.Value.ToShortDateString & "')")

        'Выполнено ТВ ПОН
        TextBox12.Text = Count_data(a, "Select [UIN] from [TaskList Total] where [Status] = ('Выполнена') and [Type] in ('Первичное переключение абонентов на PON', 'Подключение/переключение на PON с ОРК', 
        'Переключение с другого оператора', 'Инсталляция PON с ОРШ', 'Подключение услуг на PON') and [TV_status] = ('Выполнена') 
        and [Ispolnitel]  LIKE '%" & TextBox1.Text & "%' and [DTCS] BETWEEN CDATE('" & DateTimePicker1.Value.ToShortDateString & "') and CDATE('" & DateTimePicker2.Value.ToShortDateString & "')")

        'Выполнено TT ФТТБ
        TextBox10.Text = Count_data(a, "Select [UIN] from [TaskList Total] where [Status] = ('Выполнена') and [Type] = ('Обслуживание FTTB') and [Ispolnitel]  LIKE '%" & TextBox1.Text & "%'
        and [DTCS] BETWEEN CDATE('" & DateTimePicker1.Value.ToShortDateString & "') and CDATE('" & DateTimePicker2.Value.ToShortDateString & "')")

        'Выполнено TT ПОН
        TextBox11.Text = Count_data(a, "Select UIN from [TaskList Total] where [Type] in ('Обслуживание PON (weld)', 'Обслуживание абонентов PON') and Status = ('Выполнена') 
        and [DTCS] BETWEEN CDATE('" & DateTimePicker1.Value.ToShortDateString & "') and CDATE('" & DateTimePicker2.Value.ToShortDateString & "') and [Ispolnitel]  LIKE '%" & TextBox1.Text & "%'")

        Output_DGV(DataGridView1, "Select TELN, Ispolnitel, Status, Type, Time_on, DTCS, Status_sdano, Internet_tarif, Internet_status, TV_tarif, TV_status from [TaskList Total] where 
        [Status] = 'Выполнена' and [Ispolnitel] LIKE '%" & TextBox1.Text & "%' and [DTCS] BETWEEN CDATE('" & DateTimePicker1.Value.ToShortDateString & "') and 
        CDATE('" & DateTimePicker2.Value.ToShortDateString & "')")

        'Получение UTP на складе
        TextBox29.Text = Count_data_sum(a, 9, "Select * from [Equipment] where [Oborudovanie] = 'Кабель U/UTP cat 5е PVC нг-LS  4х2х0,51мм' and [Date_time] 
        BETWEEN CDATE('" & DateTimePicker1.Value.ToShortDateString & "') and CDATE('" & DateTimePicker2.Value.ToShortDateString & "') and [Tab_N] LIKE '%" & TextBox1.Text & "%'")

        Output_DGV(DataGridView2, "Select * from [Equipment] where [Oborudovanie] = 'Кабель U/UTP cat 5е PVC нг-LS  4х2х0,51мм' and [Date_time] 
        BETWEEN CDATE('" & DateTimePicker1.Value.ToShortDateString & "') and CDATE('" & DateTimePicker2.Value.ToShortDateString & "') and [Tab_N] LIKE '%" & TextBox1.Text & "%' order by [Date_time]")

        Output_DGV(DataGridView3, "Select * from [TaskList Total] where [Status_sdano] not in ('Сдано', 'Передано курьеру') and [Ispolnitel]  LIKE '%" & TextBox1.Text & "%'")

        'Всего услуги интернет без статуса
        TextBox34.Text = Count_data(a, "Select [UIN] from [TaskList Total] where Status = ('Выполнена') and [Type] in ('Подключение FTTB', 'Первичное переключение абонентов на PON', 
        'Подключение/переключение на PON с ОРК', 'Переключение с другого оператора', 'Инсталляция PON с ОРШ', 'Подключение услуг на PON') and [Internet_tarif] not in('') and [Internet_status] in ('') 
        and [Ispolnitel]  LIKE '%" & TextBox1.Text & "%' and [DTCS] BETWEEN CDATE('" & DateTimePicker1.Value.ToShortDateString & "') and CDATE('" & DateTimePicker2.Value.ToShortDateString & "')")

        'Всего услуги ТВ без статуса
        TextBox33.Text = Count_data(a, "Select [UIN] from [TaskList Total] where Status = ('Выполнена') and [Type] in ('Подключение FTTB', 'Первичное переключение абонентов на PON', 
        'Подключение/переключение на PON с ОРК', 'Переключение с другого оператора', 'Инсталляция PON с ОРШ', 'Подключение услуг на PON') and [TV_tarif] not in ('') and [TV_status] in ('') 
        and [Ispolnitel]  LIKE '%" & TextBox1.Text & "%' and [DTCS] BETWEEN CDATE('" & DateTimePicker1.Value.ToShortDateString & "') and CDATE('" & DateTimePicker2.Value.ToShortDateString & "')")

    End Sub

    Sub konv()
        Dim a As Double
        'Конвертация сданных нарядов
        TextBox7.Text = Format(konvertacia_uslug(a, CInt(TextBox2.Text), CInt(TextBox6.Text), CInt(TextBox20.Text)), "0.00 %")
        'Конвертация по нарядам
        TextBox5.Text = Format(konvertacia(a, CInt(TextBox16.Text), CInt(TextBox3.Text)), "0.00 %")
        'Конвертация интернет
        TextBox15.Text = Format(konvertacia_uslug(a, CInt(TextBox18.Text), CInt(TextBox8.Text), CInt(TextBox13.Text)), "0.00 %")
        'Конвертация ТВ
        TextBox17.Text = Format(konvertacia_uslug(a, CInt(TextBox19.Text), CInt(TextBox9.Text), CInt(TextBox12.Text)), "0.00 %")
        'Конвертация ТТ в 24 часах
        TextBox28.Text = Format(konvertacia(a, CInt(TextBox21.Text), CInt(TextBox24.Text)), "0.00 %")
        'Конвертация ТТ в 36 часах
        TextBox26.Text = Format(konvertacia(a, CInt(TextBox21.Text), CInt(TextBox23.Text)), "0.00 %")
        ' Расход кабеля
        TextBox14.Text = Format((CInt(TextBox8.Text) * 60) + (CInt(TextBox9.Text) * 15) + (CInt(TextBox10.Text) * 2) + (CInt(TextBox11.Text) * 2) + (CInt(TextBox12.Text) * 15), "0,00 м")
    End Sub

    Private Sub Prosmotr_tt()
        Dim a As Integer
        'Выполненые ТТ
        TextBox21.Text = Count_data(a, "Select UIN from [TaskList Total] where Type in ('Обслуживание FTTB', 'Обслуживание PON (weld)', 'Обслуживание абонентов ADSL', 'Обслуживание абонентов OTA',
        'Обслуживание абонентов OTA (ЛАО - Загород)', 'Обслуживание абонентов PON', 'Обслуживание ЦТВ') and Status = ('Выполнена') 
        and [DTCS] BETWEEN CDATE('" & DateTimePicker1.Value.ToShortDateString & "') and CDATE('" & DateTimePicker2.Value.ToShortDateString & "') and [Ispolnitel]  LIKE '%" & TextBox1.Text & "%'")
        'Запланированные ТТ
        TextBox22.Text = Count_data(a, "Select UIN from [TaskList Total] where Type in ('Обслуживание FTTB', 'Обслуживание PON (weld)', 'Обслуживание абонентов ADSL', 'Обслуживание абонентов OTA',
        'Обслуживание абонентов OTA (ЛАО - Загород)', 'Обслуживание абонентов PON', 'Обслуживание ЦТВ') and Status = ('Запланирована') 
        and [DTCS] BETWEEN CDATE('" & DateTimePicker1.Value.ToShortDateString & "') and CDATE('" & DateTimePicker2.Value.ToShortDateString & "') and [Ispolnitel]  LIKE '%" & TextBox1.Text & "%'")
        'Выполненные в 36 часов
        TextBox23.Text = Count_data(a, "Select UIN from [TaskList Total] where Type in ('Обслуживание FTTB', 'Обслуживание PON (weld)', 'Обслуживание абонентов ADSL', 'Обслуживание абонентов OTA',
        'Обслуживание абонентов OTA (ЛАО - Загород)', 'Обслуживание абонентов PON', 'Обслуживание ЦТВ') and Status = ('Выполнена') 
        and DateDiff('h',Time_on ,Time_start) <= 36 and [DTCS] BETWEEN CDATE('" & DateTimePicker1.Value.ToShortDateString & "') and CDATE('" & DateTimePicker2.Value.ToShortDateString & "')
        and [Ispolnitel]  LIKE '%" & TextBox1.Text & "%'")
        'Выполненные в 24 часа
        TextBox24.Text = Count_data(a, "Select UIN from [TaskList Total] where Type in ('Обслуживание FTTB', 'Обслуживание PON (weld)', 'Обслуживание абонентов ADSL', 'Обслуживание абонентов OTA',
        'Обслуживание абонентов OTA (ЛАО - Загород)', 'Обслуживание абонентов PON', 'Обслуживание ЦТВ') and Status = ('Выполнена') 
        and DateDiff('h',Time_on ,Time_start) <= 24 and [DTCS] BETWEEN CDATE('" & DateTimePicker1.Value.ToShortDateString & "') and CDATE('" & DateTimePicker2.Value.ToShortDateString & "')
        and [Ispolnitel]  LIKE '%" & TextBox1.Text & "%'")
        'Выполненные более 36 часов 
        TextBox27.Text = Count_data(a, "Select UIN from [TaskList Total] where Type in ('Обслуживание FTTB', 'Обслуживание PON (weld)', 'Обслуживание абонентов ADSL', 'Обслуживание абонентов OTA',
        'Обслуживание абонентов OTA (ЛАО - Загород)', 'Обслуживание абонентов PON', 'Обслуживание ЦТВ') and Status = ('Выполнена') 
        and DateDiff('h',Time_on ,Time_start) > 36 and [DTCS] BETWEEN CDATE('" & DateTimePicker1.Value.ToShortDateString & "') and CDATE('" & DateTimePicker2.Value.ToShortDateString & "')
        and [Ispolnitel]  LIKE '%" & TextBox1.Text & "%'")
        'Выполненные более 24 часов
        TextBox25.Text = Count_data(a, "Select UIN from [TaskList Total] where Type in ('Обслуживание FTTB', 'Обслуживание PON (weld)', 'Обслуживание абонентов ADSL', 'Обслуживание абонентов OTA',
        'Обслуживание абонентов OTA (ЛАО - Загород)', 'Обслуживание абонентов PON', 'Обслуживание ЦТВ') and Status = ('Выполнена') 
        and DateDiff('h',Time_on ,Time_start) > 24 and [DTCS] BETWEEN CDATE('" & DateTimePicker1.Value.ToShortDateString & "') and CDATE('" & DateTimePicker2.Value.ToShortDateString & "')
        and [Ispolnitel]  LIKE '%" & TextBox1.Text & "%'")

    End Sub



    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        TextBox1.Text = Convert.ToString(ComboBox1.SelectedValue)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Output_DGV(DataGridView3, "Select TELN, Ispolnitel, Status, Type, DTCS from [TaskList Total] where [Status] in ('Выполнена') and [Type] in ('Обслуживание FTTB')
        and [TELN] in (select [TELN] from [TaskList Total] where [Type] in ('Подключение FTTB') and [Status] in ('Выполнена')) order by [TELN]
        union
        Select TELN, Ispolnitel, Status, Type, DTCS from [TaskList Total] where [Status] in ('Выполнена') and [Type] in ('Подключение FTTB')
        and [TELN] in (select [TELN] from [TaskList Total] where [Type] in ('Обслуживание FTTB') and [Status] in ('Выполнена'))")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Output_DGV(DataGridView3, "Select TELN, Ispolnitel, Status, Type, DTCS, Povtorki from [TaskList Total] where [Status] in ('Выполнена') and [Type] in ('Обслуживание FTTB')
        and exists (select [TELN], count(TELN) from [TaskList Total] group by TELN HAVING count(TELN)>1) order by [TELN]")
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        checked_check(ON_Month(d1), OFF_Month(d2), CheckBox1, DateTimePicker1, DateTimePicker2)
    End Sub

    Private Sub TextBox34_TextChanged(sender As Object, e As EventArgs) Handles TextBox34.TextChanged

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Prosmotr()
        Prosmotr_tt()
        konv()
    End Sub


End Class