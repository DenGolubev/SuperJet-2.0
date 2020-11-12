Module Модуль_создания_дочерних_форм
    Public Sub Создание_дочерних_форм(NameForm As Form, MDI As Форма_MDI)
        Dim ChildForm As New Form
        ChildForm = NameForm
        ChildForm.MdiParent = MDI
        ChildForm.Show()
    End Sub

    Public Sub Form_close(NameForm As Form)
        NameForm.Dispose()
        NameForm.Close()
    End Sub
End Module
