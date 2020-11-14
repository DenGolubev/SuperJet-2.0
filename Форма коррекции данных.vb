Imports System.Data.OleDb
Public Class Форма_коррекции_данных
    Dim db_com As New OleDbCommand

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ComboBox1.Text = "" Then
            MsgBox("Необходимо выбрать статус услуги Интернет")
        Else
            Try
                db_com.Connection = db_con
                db_con.Open()
                db_com.CommandText = "Update [TaskList Total] set [Internet_status] = '" & ComboBox1.Text & "' where [UIN] LIKE '%" & Label2.Text & "%'"
                db_com.ExecuteNonQuery()
                db_con.Close()
                Me.Hide()
                Форма_приема_нарядов.Show()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If



    End Sub

    Private Sub Форма_коррекции_данных_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label2.Text = Форма_приема_нарядов.TextBox2.Text
        Label5.Text = check_internet_status()
        Label6.Text = check_TV_status()
    End Sub
End Class