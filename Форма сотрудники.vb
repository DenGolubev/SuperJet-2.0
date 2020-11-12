Imports System.Data.OleDb
Public Class Форма_сотрудники
    Dim a As Integer
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim db_com As New OleDbCommand
        Dim l_name, f_name, m_name As String
        Dim tab_n As Integer
        l_name = TextBox1.Text
        f_name = TextBox2.Text
        m_name = TextBox3.Text
        tab_n = TextBox4.Text
        Input_string("Insert into Stuff (Last_name, First_name, Midl_name, Tab_n) values ('" & l_name & "', '" & f_name & "', '" & m_name & "', '" & tab_n & "')")
        Count_data(a, "select Tab_n from Stuff")
        TextBox5.Text = a
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox1.Focus()
    End Sub

    Private Sub Форма_сотрудники_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Count_data(a, "select Tab_n from Stuff")
        TextBox5.Text = a
        Input_Russian()
    End Sub
End Class