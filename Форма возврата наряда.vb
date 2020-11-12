Imports System.Data.OleDb
Imports System.Globalization

Public Class Форма_возврата_наряда
    Dim a As Integer
    Dim tmr As New Timer With {.Interval = 500}
    Dim db_com As New OleDbCommand
    Dim db_reader As OleDbDataReader
    Dim db_table As New DataTable

    Private Sub IndicatorFlash(ByVal fLight As Boolean)
        Dim gr As Graphics = Me.CreateGraphics
        Dim br As SolidBrush
        If fLight Then 'зажигаем индикатор
            br = New SolidBrush(Color.Red)
        Else 'гасим индикатор
            br = New SolidBrush(Me.BackColor)
        End If
        Dim rc As New Rectangle(0, 0, Me.Height, Me.Width)
        gr.FillRectangle(br, rc)
    End Sub

    Private Sub Tick(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'гасим индикатор
        IndicatorFlash(False)
        tmr.Stop()
    End Sub

    Sub Индикатор()
        If Поиск_статус_нарядов() > 0 Then
            IndicatorFlash(True)
            tmr.Start()
            MsgBox("Возможно дубль")
        End If
    End Sub

    Private Function Поиск_статус_нарядов()
        Dim i As Integer
        Dim UIN As Long
        UIN = TextBox2.Text
        db_com.Connection = db_con
        db_con.Open()
        db_com.CommandText = "Select [UIN], [Status_sdano] from [TaskList Total] where [UIN] = " & UIN & " and [Status_sdano] = 'Возврат'"
        db_reader = db_com.ExecuteReader()
        db_table.Load(db_reader)
        i = db_table.Rows.Count
        db_con.Close()
        Return i
    End Function

    Private Function поиск_нарядов()
        Dim i As Integer
        Dim UIN As Long
        UIN = TextBox2.Text
        Debug.WriteLine(UIN)
        db_com.Connection = db_con
        db_con.Open()
        db_com.CommandText = "select [UIN], [WFMN], [TELN], [Ispolnitel], [Status], [Status_sdano], [Data_sdachi] from [TaskList Total] where [UIN] = " & UIN & ""
        db_reader = db_com.ExecuteReader()
        db_table.Load(db_reader)
        db_con.Close()
        i = db_table.Rows.Count
        DataGridView1.DataSource = db_table
        Return i
    End Function

    Private Sub Форма_приема_нарядов_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Output_DGV(DataGridView1, "Select count([UIN]) as Quantity, [Data_sdachi] from [TaskList Total] where [Status_sdano] = 'Возврат' group by [Data_sdachi]")
        TextBox3.Text = Count_data(a, "Select UIN from [TaskList Total] where Status_sdano = 'Возврат'")
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        db_table.Clear()
        If e.KeyCode = Keys.Enter Then
            Dim a As String = TextBox2.Text
            TextBox2.Text = Mid(a, 2, 19)
            e.SuppressKeyPress = True
            If Поиск_статус_нарядов() > 0 Then
                Индикатор()
                TextBox2.Focus()
                TextBox2.Clear()
                TextBox5.Clear()
            Else
                поиск_нарядов()
                TextBox5.Focus()
            End If
        End If
    End Sub

    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim a As String = TextBox5.Text
            TextBox5.Text = Mid(a, 7, 6)
            e.SuppressKeyPress = True
            Button1.Focus()
        End If
    End Sub

    Sub Rec_data()
        db_com.Connection = db_con
        db_con.Open()
        db_com.CommandText = "Update [TaskList Total] set Status_sdano = 'Возврат', Data_sdachi = '" & Format(Now, "dd/MM/yyyy") & "', Fakt_Ispolnitel = '" & TextBox5.Text & "' 
where UIN Like '%" & TextBox2.Text & "%'"
        db_com.ExecuteNonQuery()
        db_con.Close()
        TextBox2.Clear()
        TextBox5.Clear()
        TextBox2.Focus()
        Output_DGV(DataGridView1, "Select [WFMN], [TELN], [Ispolnitel], [Status], [Status_sdano], [Data_sdachi] from [TaskList Total] where [Status_sdano] in ('Возврат')")
        Count_data(TextBox3.Text, "Select [UIN] from [TaskList Total] where [Status_sdano] = 'Возврат'")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Rec_data()
    End Sub
End Class