Imports System.Data.OleDb
Public Class Форма_коррекции_данных
    Dim db_com As New OleDbCommand

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim a As String = Label2.Text

        If ComboBox1.Text <> "" Then
            While check_internet_status(a) = False

                Try
                    db_com.Connection = db_con
                    db_con.Open()
                    db_com.CommandText = "Update [TaskList Total] set [Internet_status] = '" & ComboBox1.Text & "' where [UIN] LIKE '%" & Label2.Text & "%'"
                    db_com.ExecuteNonQuery()
                    db_con.Close()
                    Me.Hide()
                    Форма_приема_нарядов.Show()
                    Форма_приема_нарядов.TextBox2.Focus()
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

            End While

            While check_TV_status(a) = False

                Try
                    db_com.Connection = db_con
                    db_con.Open()
                    db_com.CommandText = "Update [TaskList Total] set [TV_status] = '" & ComboBox1.Text & "' where [UIN] LIKE '%" & Label2.Text & "%'"
                    db_com.ExecuteNonQuery()
                    db_con.Close()
                    Me.Hide()
                    Форма_приема_нарядов.Show()
                    Форма_приема_нарядов.TextBox2.Focus()
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

            End While
        ElseIf ComboBox1.Text = "" Then
            MsgBox("Необходимо выбрать статус услуги")
        End If

    End Sub

    Private Sub Форма_коррекции_данных_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim a As String = Форма_приема_нарядов.TextBox2.Text
        Label2.Text = a
        Label5.Text = check_internet_status(a)
        Label6.Text = check_TV_status(a)
    End Sub
End Class