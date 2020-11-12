Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.Globalization
Public Class Форма_приема_нарядов
    Dim a As Integer
    Dim tmr As New Timer With {.Interval = 500}
    Dim db_com As New OleDbCommand
    Dim db_reader As OleDbDataReader
    Dim db_table As New DataTable

    Private Sub Форма_приема_нарядов_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        AddHandler tmr.Tick, AddressOf Tick
        Output_DGV(DataGridView1, "Select count([UIN]) as Quantity, [Data_sdachi] from [TaskList Total] where [Status_sdano] = 'Сдано' group by [Data_sdachi]")
        Output_DGV(DataGridView2, "Select count([UIN]) as Quantity, [Data_sdachi] from [TaskList Total] where [Status_sdano] = 'Передано курьеру' group by [Data_sdachi]")
        TextBox1.Text = Count_data(a, "Select [UIN] from [TaskList Total] where [Status] = 'Выполнена'")
        TextBox3.Text = Count_data(a, "Select UIN from [TaskList Total] where [Status_sdano] in ('Сдано', 'Передано курьеру')")
        TextBox4.Text = Format(Text_Box_Count(a, TextBox3, TextBox1), "0.00 %")
        TextBox4.TextAlign = HorizontalAlignment.Center
    End Sub
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

    Sub Индикатор()
        If Поиск_статус_нарядов() > 0 Then
            IndicatorFlash(True)
            tmr.Start()
            MsgBox("Возможно дубль")
        End If
    End Sub

    Private Function поиск_нарядов()
        Dim i As Integer
        Dim UIN As Long
        UIN = TextBox2.Text
        Debug.WriteLine(UIN)
        db_com.Connection = db_con
        db_con.Open()
        db_com.CommandText = "select [UIN] from [TaskList Total] where [UIN] = " & UIN & ""
        db_reader = db_com.ExecuteReader()
        db_table.Load(db_reader)
        db_con.Close()
        i = db_table.Rows.Count
        'DataGridView3.DataSource = db_table
        Return i
    End Function

    Private Function Поиск_статус_нарядов()
        Dim i As Integer
        Dim UIN As Long
        UIN = TextBox2.Text
        db_com.Connection = db_con
        db_con.Open()
        db_com.CommandText = "Select [UIN], [Status_sdano] from [TaskList Total] where [UIN] = " & UIN & " and [Status_sdano] = 'Сдано'"
        db_reader = db_com.ExecuteReader()
        db_table.Load(db_reader)
        i = db_table.Rows.Count
        db_con.Close()
        Return i
    End Function

    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim a As String = TextBox5.Text
            TextBox5.Text = Mid(a, 7, 6)
            e.SuppressKeyPress = True
            Button1.Focus()
        End If
    End Sub


    Sub Rec_data()
        Try
            db_com.Connection = db_con
            db_con.Open()
            db_com.CommandText = "Update [TaskList Total] set [Status_sdano] = 'Сдано', [Data_sdachi] = '" & Format(Now, "dd/MM/yyyy") & "', [Fakt_Ispolnitel] = '" & TextBox5.Text & "' 
where [UIN] LIKE '%" & TextBox2.Text & "%'"
            db_com.ExecuteNonQuery()
            db_con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Sub Rec_data_courier()
        Try
            db_com.Connection = db_con
            db_con.Open()
            db_com.CommandText = "Update [TaskList Total] set [Status_sdano] = 'Передано курьеру', [Data_sdachi] = '" & Format(Now, "dd/MM/yyyy") & "' where [Status_sdano] = 'Сдано'"
            db_com.ExecuteNonQuery()
            db_con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Rec_data()
        TextBox1.Text = Count_data(a, "Select [UIN] from [TaskList Total] where [Status] = 'Выполнена'")
        TextBox3.Text = Count_data(a, "Select [UIN] from [TaskList Total] where [Status_sdano] = 'Сдано'")
        TextBox4.Text = Format(Text_Box_Count(a, TextBox3, TextBox1), "0.00 %")
        Output_DGV(DataGridView1, "Select count([UIN]) as Quantity, [Data_sdachi] from [TaskList Total] where [Status_sdano] in ('Сдано') group by [Data_sdachi]")
        TextBox2.Clear()
        TextBox5.Clear()
        TextBox2.Focus()

    End Sub



    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Rec_data_courier()
        Output_DGV(DataGridView2, "Select count([UIN]) as Quantity, [Data_sdachi] from [TaskList Total] where [Status_sdano] = 'Передано курьеру' group by [Data_sdachi]")
        Output_DGV(DataGridView1, "Select count([UIN]) as Quantity, [Data_sdachi] from [TaskList Total] where [Status_sdano] = 'Сдано' group by [Data_sdachi]")
    End Sub

    Private Sub Форма_приема_нарядов_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Me.Dispose()
    End Sub
End Class