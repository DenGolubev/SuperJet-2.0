﻿Imports System.Data.OleDb
Module Модуль_check
    Dim db_com As New OleDbCommand
    Dim db_reader As OleDbDataReader
    Dim db_table As New DataTable
    Public Function check_internet_status(ByVal a As String)
        Dim UIN As Long
        UIN = Convert.ToInt64(a)
        Dim internet_status As Boolean = True
        Dim count_tarif As Integer = Count_data(count_tarif, "Select [UIN] from [TaskList Total] where [UIN] = " & UIN & " and [Status] in ('Выполнена', 'Не выполнена') and [Type] in ('Подключение FTTB', 'Первичное переключение абонентов на PON', 
        'Подключение/переключение на PON с ОРК', 'Переключение с другого оператора', 'Инсталляция PON с ОРШ', 'Подключение услуг на PON') and [Internet_tarif] not in ('')")
        Dim count_status As Integer = Count_data(count_status, "Select [UIN] from [TaskList Total] where [UIN] = " & UIN & " and [Status] in ('Выполнена', 'Не выполнена') and [Type] in ('Подключение FTTB', 'Первичное переключение абонентов на PON', 
        'Подключение/переключение на PON с ОРК', 'Переключение с другого оператора', 'Инсталляция PON с ОРШ', 'Подключение услуг на PON') and [Internet_status] in ('')")
        If count_tarif > 0 Then
            If count_status > 0 Then
                internet_status = False
            End If
        End If


        Return internet_status
    End Function

    Public Function check_TV_status(ByVal a As String)
        Dim UIN As Long
        UIN = Convert.ToInt64(a)
        Dim TV_status As Boolean = True
        Dim count_status As Integer = Count_data(count_status, "Select [UIN] from [TaskList Total] where [UIN] = " & UIN & " and [Status] in ('Выполнена', 'Не выполнена') and [Type] in ('Подключение FTTB', 'Первичное переключение абонентов на PON', 
        'Подключение/переключение на PON с ОРК', 'Переключение с другого оператора', 'Инсталляция PON с ОРШ', 'Подключение услуг на PON') and [TV_status] in ('')")
        Dim count_tarif As Integer = Count_data(count_tarif, "Select [UIN] from [TaskList Total] where [UIN] = " & UIN & " and [Status] in ('Выполнена', 'Не выполнена') and [Type] in ('Подключение FTTB', 'Первичное переключение абонентов на PON', 
        'Подключение/переключение на PON с ОРК', 'Переключение с другого оператора', 'Инсталляция PON с ОРШ', 'Подключение услуг на PON') and [TV_tarif] not in ('')")

        If count_tarif > 0 Then
            If count_status > 0 Then
                TV_status = False
            End If
        End If

        Return TV_status
    End Function

End Module
