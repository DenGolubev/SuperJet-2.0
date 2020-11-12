Public Class Форма_MDI

    Private Sub Форма_MDI_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized
        Me.CenterToScreen()

    End Sub

    Private Sub ПриемНарядовToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПриемНарядовToolStripMenuItem.Click
        Создание_дочерних_форм(Форма_приема_нарядов, Me)
    End Sub

    Private Sub ПросмотрНарядовToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПросмотрНарядовToolStripMenuItem.Click
        Создание_дочерних_форм(Форма_просмотра_БД, Me)
    End Sub

    Private Sub ЗагрузкаTaskListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЗагрузкаTaskListToolStripMenuItem.Click
        Создание_дочерних_форм(Форма_загрузки_TaskList, Me)
    End Sub

    Private Sub ПриемСотрудниковToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПриемСотрудниковToolStripMenuItem.Click
        Создание_дочерних_форм(Форма_сотрудники, Me)
    End Sub

    Private Sub КонвертацияToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles КонвертацияToolStripMenuItem.Click
        Создание_дочерних_форм(Форма_Конвертация, Me)
    End Sub

    Private Sub ВозвратьНарядовToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВозвратьНарядовToolStripMenuItem.Click
        Создание_дочерних_форм(Форма_возврата_наряда, Me)
    End Sub

    Private Sub ЧасовToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЧасовToolStripMenuItem.Click
        Создание_дочерних_форм(Форма_36_часов, Me)
    End Sub

    Private Sub Форма_MDI_GotFocus(sender As Object, e As EventArgs) Handles Me.GotFocus
        Form_close(Форма_табло_информации)
    End Sub
End Class