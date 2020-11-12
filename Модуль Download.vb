Imports System.Data.OleDb
Module Модуль_Download
    Public con_excel As New OleDbConnection
    Public Function excel_con(ByVal file_name As String, ByRef db_table As DataTable, ByRef status As Boolean)
        Dim db_con_excel As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source='" & file_name & "';Extended Properties =""Excel 8.0;HDR=Yes;IMEX=1"";")
        Dim db_adapter As New OleDbDataAdapter("SELECT [Документ], Format([Дата], 'dd.MM.yyyy'), [Склад], [Код складской организации], [СП], [Табельный номер], [Технический специалист], 
[Код позиции OEBS], [Номенклатура], [Количество] FROM [TDSheet$] where [СП] = 'Жуковский' and [Номенклатура] = 'Кабель U/UTP cat 5е PVC нг-LS  4х2х0,51мм'", db_con_excel)
        db_adapter.Fill(db_table)
        Try
            For i As Integer = db_table.Rows.Count - 1 To 0 Step -1
                Dim row As DataRow = db_table.Rows(i)
                If row.Item(0) Is Nothing Then
                    db_table.Rows.Remove(row)
                ElseIf row.Item(0).ToString Is Nothing Then
                    db_table.Rows.Remove(row)
                End If
            Next
            For i = 0 To db_table.Rows.Count - 1
                For j = 0 To db_table.Columns.Count - 1
                    If db_table.Rows(i)(j).ToString Is Nothing Then
                        db_table.Rows(i)(j) = "Нет данных"
                    End If
                Next
            Next
            status = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        Return status
    End Function

    Public Function TaskList_con(ByVal file_name As String, ByRef db_table As DataTable, ByRef status As Boolean)
        Try
            Dim db_con_excel As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source='" & file_name & "';Extended Properties =""Excel 8.0;HDR=Yes;IMEX=1"";")
            Dim db_adapter As New OleDbDataAdapter("SELECT [Уникальный номер наряда], [Номер наряда в WFM], [Телефон], [Исполнитель], [Статус], [Тип], [Время создания наряда], 
      Format([Дата перехода в текущий статус], 'dd.MM.yyyy'), [Причина не выполнения заказа], [Комментарий монтажника], [Район], [Время начала], [Контактный телефон], [Адрес ADSL, OTA и ТТ], 
[Источник заявки], [Интернет тариф], [Интернет статус], [Телевидение тариф], [Информация аметист статус], [ТВ мультирум1 №асрз], 
[ТВ мультирум1 статус], [ТВ мультирум2 №асрз], [ТВ мультирум2 статус], [Кол-во повторных выходов], [Признак статуса] FROM [TaskList$] where [Статус] not in ('Отменена')", db_con_excel)
            db_adapter.Fill(db_table)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        Try
            For i As Integer = db_table.Rows.Count - 1 To 0 Step -1
                Dim row As DataRow = db_table.Rows(i)
                If row.Item(0) Is Nothing Then
                    db_table.Rows.Remove(row)
                ElseIf row.Item(0).ToString = "" Then
                    db_table.Rows.Remove(row)
                End If
            Next
            For i = 0 To db_table.Rows.Count - 1
                For j = 0 To db_table.Columns.Count - 1
                    If db_table.Rows(i)(j).ToString Is DBNull.Value Then
                        db_table.Rows(i)(j) = "Нет данных"
                    End If
                Next
            Next
            status = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        Return status
    End Function

End Module
