Imports System.ComponentModel
Imports System.IO
Public Class Форма_Главная
    Dim a, b, c As Single
    Dim d, f As String
    Dim день_недели As New List(Of String) From {"Воскресенье", "Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота"}
    Dim таймер, таймер1 As New Timer
    Public db_path As String = Application.StartupPath
    Public bd_name As String = "SuperJet 2.0"
    Public bak As String = "SuperJet bak"
    Public db_pass As String = "" 'пароль базы данных
    Dim c_load As String
    Dim c_save, b_save As String


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Форма_MDI.Show()
        Форма_табло_информации.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        File.Delete(db_path & "\" & bd_name & ".mdb")
        File.Copy(db_path & "\" & bak & ".mdb", db_path & "\" & bd_name & ".mdb")
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        Dim x As Integer
        For x = 0 To 100000
            con_close()
            con_status(Button2)
            Timer2.Stop()
        Next
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        a = Date.Now.Hour
        b = Date.Now.Minute
        c = Date.Now.Second
        d = Date.Now.DayOfWeek
        f = Date.Now.ToLongDateString

        TextBox1.Text = a
        TextBox1.Refresh()

        TextBox2.Text = b
        TextBox2.Refresh()

        TextBox3.Text = c
        TextBox3.Refresh()

        TextBox4.Text = день_недели(d)
        TextBox4.Refresh()

        TextBox5.Text = f
        TextBox5.Refresh()

    End Sub

    Private Sub Форма_Главная_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        c_load = file_1(a)
        Timer1.Enabled = True
        Timer2.Enabled = True
        con_open()
        con_status(Button2)
        Label6.TextAlign = ContentAlignment.MiddleCenter
        Label6.Text = "Backup database data" & " " & vbCrLf & File.GetLastAccessTime(db_path & "\" & bak & ".mdb").ToString
        Debug.WriteLine(c_load)

    End Sub

    Private Sub Форма_Главная_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        c_save = file_1(a)
        If c_load = c_save Then
            MsgBox("Изменений в БД не было")
            Application.Exit()
        Else
            Dim y As DialogResult = MsgBox("Изменения в БД были - необходимо записать резервную копию БД", MsgBoxStyle.OkCancel)
            Select Case y
                Case DialogResult.OK
                    File.Delete(db_path & "\" & bak & ".mdb")
                    File.Copy(db_path & "\" & bd_name & ".mdb", db_path & "\" & bak & ".mdb")
                Case DialogResult.Cancel
                    Application.Exit()
            End Select
        End If
    End Sub

    Public Sub Удалить_Файл(FilePath As String)
        If FilePath <> "" And IO.File.Exists(FilePath) Then
            Application.DoEvents()
            Do Until IO.File.Exists(FilePath) = False
                Application.DoEvents()
                On Error Resume Next
                Kill(FilePath)
            Loop
            If IO.File.Exists(FilePath) = False Then
                MsgBox("Файл удален.")
            Else
                MsgBox("Не могу удалить файл")
            End If
        Else
            MsgBox("Файл не найден.")
        End If
    End Sub



End Class