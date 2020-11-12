Imports System.Data.OleDb
Module Модуль_записи_реестра_нарядов
    Public Sub Запись_в_реестр()
        Dim поле_дата As DateTime
        поле_дата = Форма_приема_нарядов.DateTimePicker1.Text
        Dim поле_UIN As String
        поле_UIN = Форма_приема_нарядов.TextBox2.Text
        Dim поле_TABN As String
        поле_TABN = Форма_приема_нарядов.TextBox5.Text
        Try
            db_com.Connection = db_con
            con_open()
            Input_string("Insert into Реестр_сданных ([Дата], [UIN], [WFMN], [TELN], [TABN]) values ('" & поле_дата & "', '" & поле_UIN & "', '" & поле_TABN & "')")
            db_com.ExecuteNonQuery()
            Input_string("Insert into Реестр_общий ([Дата], [UIN], [WFMN], [TELN], [TABN]) values ('" & поле_дата & "', '" & поле_UIN & "', '" & поле_TABN & "')")
            db_com.ExecuteNonQuery()
            con_close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            con_close()
        End Try
        Форма_приема_нарядов.TextBox2.Clear()
        Форма_приема_нарядов.TextBox5.Clear()
        Форма_приема_нарядов.TextBox2.Focus()
    End Sub
End Module