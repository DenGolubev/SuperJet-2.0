Imports System.ComponentModel
Public Class Форма_просмотра_БД
    Private Sub Форма_просмотра_БД_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Output_DGV(DataGridView1, "Select [WFMN], [TELN], [Ispolnitel], [Status], [Status_sdano], [Data_sdachi] from [TaskList Total] where [Status_sdano] in ('Сдано') ORDER BY [Data_sdachi]")
        Output_DGV(DataGridView2, "Select [WFMN], [TELN], [Ispolnitel], [Status], [Status_sdano], [Data_sdachi] from [TaskList Total] where [Status_sdano] in ('Возврат') ORDER BY [Data_sdachi]")
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        Dim y As String = TextBox1.Text
        Output_DGV(DataGridView3, "Select * from [TaskList Total] where [WFMN] = '" & y & "' or [TELN] = '" & y & "'")
    End Sub
End Class