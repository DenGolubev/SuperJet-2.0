Module Модуль_date_time

    Public Function ON_Month(ByRef date_on As Date) As Date
        date_on = DateAdd("m", 0, DateSerial(Year(Today), Month(Today), 1))
        date_on.ToShortDateString()
        Return date_on
    End Function

    Public Function OFF_Month(ByRef date_off As Date) As Date
        date_off = DateAdd("m", 1, DateSerial(Year(Today), Month(Today), 0))
        date_off.ToShortDateString()
        Return date_off
    End Function

    Structure new_data
        Dim d1, d2 As DateTime
    End Structure

    Public Function checked_check(ByVal date_on As Date, ByVal date_off As Date, ch_box As CheckBox, dt_pick1 As DateTimePicker, dt_pick2 As DateTimePicker)
        Dim data As New new_data
        If ch_box.Checked = True Then
            dt_pick1.Enabled = True
            dt_pick2.Enabled = True
            dt_pick1.Show()
            dt_pick2.Show()
            data.d1 = dt_pick1.Value.ToShortDateString
            data.d2 = dt_pick2.Value.ToShortDateString
            ch_box.Text = "Выбор диапазон дат"
        Else
            dt_pick1.Enabled = False
            dt_pick2.Enabled = False
            dt_pick1.Hide()
            dt_pick2.Hide()
            data.d1 = ON_Month(date_on)
            data.d2 = OFF_Month(date_off)
            ch_box.Text = "Текущий месяц"
        End If
        Return data
    End Function


End Module
