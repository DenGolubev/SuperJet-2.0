
Public Class Форма_табло_информации
    Dim d1, d2 As Date
    Private Sub Форма_табло_информации_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ON_Month(d1)
        OFF_Month(d2)
        TextBox15.Text = d1
        TextBox14.Text = d2
        Prosmotr_tt()
        Prosmotr()

        Debug.WriteLine(TypeName(ON_Month(d1)))
        Debug.WriteLine(TypeName(OFF_Month(d2)))
    End Sub

    Private Sub Prosmotr()
        Dim a As Integer

        'Все выполненные наряды
        TextBox18.Text = Count_data(a, "Select [UIN] from [TaskList Total] where [Status] = 'Выполнена' and [DTCS] BETWEEN CDATE('" & d1 & "') and CDATE('" & d2 & "') 
        and [Ispolnitel] not like '% НПП_%'")
        'Все первички
        TextBox17.Text = Count_data(a, "Select [UIN] from [TaskList Total] where [Status] in ('Выполнена', 'Не выполнена') and [Type] in ('Подключение FTTB', 'Первичное переключение абонентов на PON', 
        'Подключение/переключение на PON с ОРК', 'Переключение с другого оператора', 'Инсталляция PON с ОРШ') and [DTCS] BETWEEN CDATE('" & d1 & "') and CDATE('" & d2 & "') 
        and [Ispolnitel] not like '% НПП_%'")
        'Все выполненные первички
        TextBox19.Text = Count_data(a, "Select [UIN] from [TaskList Total] where [Status] in ('Выполнена') and [Type] in ('Подключение FTTB', 'Первичное переключение абонентов на PON', 
        'Подключение/переключение на PON с ОРК', 'Переключение с другого оператора', 'Инсталляция PON с ОРШ') and [DTCS] BETWEEN CDATE('" & d1 & "') and CDATE('" & d2 & "')
        and [Ispolnitel] not like '% НПП_%'")
        'Все не выполненные первички
        TextBox20.Text = Count_data(a, "Select [UIN] from [TaskList Total] where [Status] in ('Не выполнена') and [Type] in ('Подключение FTTB', 'Первичное переключение абонентов на PON', 
        'Подключение/переключение на PON с ОРК', 'Переключение с другого оператора', 'Инсталляция PON с ОРШ') and [DTCS] BETWEEN CDATE('" & d1 & "') and CDATE('" & d2 & "')
        and [Ispolnitel] not like '% НПП_%'")
        'Всего услуги интернет
        TextBox22.Text = Count_data(a, "Select [UIN] from [TaskList Total] where [Status] in ('Выполнена', 'Не выполнена') and [Type] in ('Подключение FTTB', 'Первичное переключение абонентов на PON', 
        'Подключение/переключение на PON с ОРК', 'Переключение с другого оператора', 'Инсталляция PON с ОРШ', 'Подключение услуг на PON') and [Internet_status] in ('Выполнена', 'Не выполнена') 
        and [DTCS] BETWEEN CDATE('" & d1 & "') and CDATE('" & d2 & "') and [Ispolnitel] not like '% НПП_%'")
        'Всего услуги ТВ
        TextBox21.Text = Count_data(a, "Select [UIN] from [TaskList Total] where [Status] in ('Выполнена', 'Не выполнена') and [Type] in ('Подключение FTTB', 'Первичное переключение абонентов на PON', 
        'Подключение/переключение на PON с ОРК', 'Переключение с другого оператора', 'Инсталляция PON с ОРШ', 'Подключение услуг на PON') and [TV_status] in ('Выполнена', 'Не выполнена') 
        and [DTCS] BETWEEN CDATE('" & d1 & "') and CDATE('" & d2 & "') and [Ispolnitel] not like '% НПП_%'")
        'Выполнено интернет ФТТБ
        TextBox28.Text = Count_data(a, "Select [UIN] from [TaskList Total] where [Status] in ('Выполнена') and [Type] in ('Подключение FTTB') and [Internet_status] in ('Выполнена') 
        and [DTCS] BETWEEN CDATE('" & d1 & "') and CDATE('" & d2 & "') and [Ispolnitel] not like '% НПП_%'")
        'Выполнено ТВ ФТТБ
        TextBox29.Text = Count_data(a, "Select [UIN] from [TaskList Total] where [Status] in ('Выполнена') and [Type] in ('Подключение FTTB') and [TV_status] in ('Выполнена') 
        and [DTCS] BETWEEN CDATE('" & d1 & "') and CDATE('" & d2 & "') and [Ispolnitel] not like '% НПП_%'")
        'Выполнено интернет ПОН
        TextBox30.Text = Count_data(a, "Select [UIN] from [TaskList Total] where [Status] in ('Выполнена') and [Type] in ('Первичное переключение абонентов на PON', 'Подключение/переключение на PON с ОРК', 
        'Переключение с другого оператора', 'Инсталляция PON с ОРШ', 'Подключение услуг на PON') and [Internet_status] in ('Выполнена') 
        and [DTCS] BETWEEN CDATE('" & d1 & "') and CDATE('" & d2 & "') and [Ispolnitel] not like '% НПП_%'")
        'Выполнено ТВ ПОН
        TextBox31.Text = Count_data(a, "Select [UIN] from [TaskList Total] where [Status] in ('Выполнена') and [Type] in ('Первичное переключение абонентов на PON', 'Подключение/переключение на PON с ОРК', 
        'Переключение с другого оператора', 'Инсталляция PON с ОРШ', 'Подключение услуг на PON') and [TV_status] in ('Выполнена') 
        and [DTCS] BETWEEN CDATE('" & d1 & "') and CDATE('" & d2 & "') and [Ispolnitel] not like '% НПП_%'")
        'Сданные наряды
        TextBox26.Text = Count_data(a, "Select [UIN] from [TaskList Total] where [Status_sdano] in ('Сдано', 'Передано курьеру') and [DTCS] BETWEEN CDATE('" & d1 & "') and CDATE('" & d2 & "')
        and [Ispolnitel] not like '% НПП_%'")
        'Конвертациия по нарядам
        TextBox16.Text = Format(konvertacia(a, CInt(TextBox17.Text), CInt(TextBox19.Text)), "0.00 %")
        'Конвертациия по сданным нарядам
        TextBox25.Text = Format(konvertacia(a, CInt(TextBox18.Text), CInt(TextBox26.Text)), "0.00 %")
        'Конвертация интернет
        TextBox27.Text = Format(konvertacia_uslug(a, CInt(TextBox22.Text), CInt(TextBox28.Text), CInt(TextBox30.Text)), "0.00 %")
        'Конвертация ТВ
        TextBox23.Text = Format(konvertacia_uslug(a, CInt(TextBox21.Text), CInt(TextBox29.Text), CInt(TextBox31.Text)), "0.00 %")
        'Расход кабеля
        TextBox24.Text = kabel(a, CInt(TextBox28.Text), CInt(TextBox29.Text), CInt(TextBox30.Text), CInt(TextBox31.Text), CInt(TextBox6.Text))
        'Взято кабеля на складе
        TextBox1.Text = Count_data_sum(a, 9, "Select * from [Equipment] where [Oborudovanie] = 'Кабель U/UTP cat 5е PVC нг-LS  4х2х0,51мм' and [Date_time] BETWEEN CDATE('" & d1 & "') and CDATE('" & d2 & "')
        and [Ispolnitel] not like '% НПП_%'")
        TextBox3.Text = Count_data(a, "Select [UIN] from [TaskList Total] where [Status] in ('Выполнена', 'Не выполнена') and [Type] in ('Подключение FTTB', 'Первичное переключение абонентов на PON', 
        'Подключение/переключение на PON с ОРК', 'Переключение с другого оператора', 'Инсталляция PON с ОРШ', 'Подключение услуг на PON') and [Internet_status] in ('') 
        and [DTCS] BETWEEN CDATE('" & d1 & "') and CDATE('" & d2 & "') and [Ispolnitel] not like '% НПП_%'")
        TextBox2.Text = Count_data(a, "Select [UIN] from [TaskList Total] where [Status] in ('Выполнена', 'Не выполнена') and [Type] in ('Подключение FTTB', 'Первичное переключение абонентов на PON', 
        'Подключение/переключение на PON с ОРК', 'Переключение с другого оператора', 'Инсталляция PON с ОРШ', 'Подключение услуг на PON') and [TV_status] in ('') 
        and [DTCS] BETWEEN CDATE('" & d1 & "') and CDATE('" & d2 & "') and [Ispolnitel] not like '% НПП_%'")

    End Sub

    Private Sub Prosmotr_tt()
        Dim a As Double
        'Выполнено всего ТТ
        TextBox6.Text = Count_data(a, "Select UIN from [TaskList Total] where Type in ('Обслуживание FTTB', 'Обслуживание PON (weld)', 'Обслуживание абонентов ADSL', 'Обслуживание абонентов OTA',
        'Обслуживание абонентов OTA (ЛАО - Загород)', 'Обслуживание абонентов PON', 'Обслуживание ЦТВ') and Status in ('Выполнена') 
        and [DTCS] BETWEEN CDATE('" & d1 & "') and CDATE('" & d2 & "') and [Ispolnitel] not like '% НПП_%'")
        'Запланировано всего ТТ
        TextBox7.Text = Count_data(a, "Select UIN from [TaskList Total] where Type in ('Обслуживание FTTB', 'Обслуживание PON (weld)', 'Обслуживание абонентов ADSL', 'Обслуживание абонентов OTA',
        'Обслуживание абонентов OTA (ЛАО - Загород)', 'Обслуживание абонентов PON', 'Обслуживание ЦТВ') and Status in ('Запланирована') 
        and [DTCS] BETWEEN CDATE('" & d1 & "') and CDATE('" & d2 & "') and [Ispolnitel] not like '% НПП_%'")
        'Выполнено всего ТТ в 36 часов
        TextBox8.Text = Count_data(a, "Select UIN from [TaskList Total] where Type in ('Обслуживание FTTB', 'Обслуживание PON (weld)', 'Обслуживание абонентов ADSL', 'Обслуживание абонентов OTA',
        'Обслуживание абонентов OTA (ЛАО - Загород)', 'Обслуживание абонентов PON', 'Обслуживание ЦТВ') and Status in ('Выполнена') 
        and DateDiff('h',Time_on ,Time_start) <= 36 and [DTCS] BETWEEN CDATE('" & d1 & "') and CDATE('" & d2 & "') and [Ispolnitel] not like '% НПП_%'")
        'Выполнено всего ТТ в 24 часа
        TextBox13.Text = Count_data(a, "Select UIN from [TaskList Total] where Type in ('Обслуживание FTTB', 'Обслуживание PON (weld)', 'Обслуживание абонентов ADSL', 'Обслуживание абонентов OTA',
        'Обслуживание абонентов OTA (ЛАО - Загород)', 'Обслуживание абонентов PON', 'Обслуживание ЦТВ') and Status in ('Выполнена') 
        and DateDiff('h',Time_on ,Time_start) <= 24 and [DTCS] BETWEEN CDATE('" & d1 & "') and CDATE('" & d2 & "') and [Ispolnitel] not like '% НПП_%'")
        'Выполнено всего ТТ более 36 часов
        TextBox11.Text = Count_data(a, "Select UIN from [TaskList Total] where Type in ('Обслуживание FTTB', 'Обслуживание PON (weld)', 'Обслуживание абонентов ADSL', 'Обслуживание абонентов OTA',
        'Обслуживание абонентов OTA (ЛАО - Загород)', 'Обслуживание абонентов PON', 'Обслуживание ЦТВ') and Status in ('Выполнена') 
        and DateDiff('h',Time_on ,Time_start) > 36 and [DTCS] BETWEEN CDATE('" & d1 & "') and CDATE('" & d2 & "') and [Ispolnitel] not like '% НПП_%'")
        'Выполнено всего ТТ более 24 часов
        TextBox9.Text = Count_data(a, "Select UIN from [TaskList Total] where Type in ('Обслуживание FTTB', 'Обслуживание PON (weld)', 'Обслуживание абонентов ADSL', 'Обслуживание абонентов OTA',
        'Обслуживание абонентов OTA (ЛАО - Загород)', 'Обслуживание абонентов PON', 'Обслуживание ЦТВ') and Status in ('Выполнена') 
        and DateDiff('h',Time_on ,Time_start) > 24 and [DTCS] BETWEEN CDATE('" & d1 & "') and CDATE('" & d2 & "') and [Ispolnitel] not like '% НПП_%'")
        'Конвертация 24 часа
        TextBox12.Text = Format(konvertacia(a, TextBox6.Text, TextBox13.Text), "0.00 %")
        'Конвертация 36 часа
        TextBox10.Text = Format(konvertacia(a, TextBox6.Text, TextBox8.Text), "0.00 %")

    End Sub


End Class