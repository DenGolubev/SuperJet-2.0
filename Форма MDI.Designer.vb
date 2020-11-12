<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Форма_MDI
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Является обязательной для конструктора форм Windows Forms
    Private components As System.ComponentModel.IContainer

    'Примечание: следующая процедура является обязательной для конструктора форм Windows Forms
    'Для ее изменения используйте конструктор форм Windows Form.  
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.РеестрНарядовToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПриемНарядовToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ВозвратьНарядовToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПросмотрНарядовToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ЗагрузкаДанныхToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ЗагрузкаTaskListToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.СотрудникиToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПриемСотрудниковToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.УвольнениеСотрудниковToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ОтчетыToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.КонвертацияToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ЧасовToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.РеестрНарядовToolStripMenuItem, Me.ЗагрузкаДанныхToolStripMenuItem, Me.СотрудникиToolStripMenuItem, Me.ОтчетыToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(1507, 24)
        Me.MenuStrip1.TabIndex = 1
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'РеестрНарядовToolStripMenuItem
        '
        Me.РеестрНарядовToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ПриемНарядовToolStripMenuItem, Me.ВозвратьНарядовToolStripMenuItem, Me.ПросмотрНарядовToolStripMenuItem})
        Me.РеестрНарядовToolStripMenuItem.Name = "РеестрНарядовToolStripMenuItem"
        Me.РеестрНарядовToolStripMenuItem.Size = New System.Drawing.Size(104, 20)
        Me.РеестрНарядовToolStripMenuItem.Text = "Реестр нарядов"
        '
        'ПриемНарядовToolStripMenuItem
        '
        Me.ПриемНарядовToolStripMenuItem.Name = "ПриемНарядовToolStripMenuItem"
        Me.ПриемНарядовToolStripMenuItem.Size = New System.Drawing.Size(179, 22)
        Me.ПриемНарядовToolStripMenuItem.Text = "Прием нарядов"
        '
        'ВозвратьНарядовToolStripMenuItem
        '
        Me.ВозвратьНарядовToolStripMenuItem.Name = "ВозвратьНарядовToolStripMenuItem"
        Me.ВозвратьНарядовToolStripMenuItem.Size = New System.Drawing.Size(179, 22)
        Me.ВозвратьНарядовToolStripMenuItem.Text = "Возврать нарядов"
        '
        'ПросмотрНарядовToolStripMenuItem
        '
        Me.ПросмотрНарядовToolStripMenuItem.Name = "ПросмотрНарядовToolStripMenuItem"
        Me.ПросмотрНарядовToolStripMenuItem.Size = New System.Drawing.Size(179, 22)
        Me.ПросмотрНарядовToolStripMenuItem.Text = "Просмотр нарядов"
        '
        'ЗагрузкаДанныхToolStripMenuItem
        '
        Me.ЗагрузкаДанныхToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ЗагрузкаTaskListToolStripMenuItem})
        Me.ЗагрузкаДанныхToolStripMenuItem.Name = "ЗагрузкаДанныхToolStripMenuItem"
        Me.ЗагрузкаДанныхToolStripMenuItem.Size = New System.Drawing.Size(111, 20)
        Me.ЗагрузкаДанныхToolStripMenuItem.Text = "Загрузка данных"
        '
        'ЗагрузкаTaskListToolStripMenuItem
        '
        Me.ЗагрузкаTaskListToolStripMenuItem.Name = "ЗагрузкаTaskListToolStripMenuItem"
        Me.ЗагрузкаTaskListToolStripMenuItem.Size = New System.Drawing.Size(165, 22)
        Me.ЗагрузкаTaskListToolStripMenuItem.Text = "Загрузка TaskList"
        '
        'СотрудникиToolStripMenuItem
        '
        Me.СотрудникиToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ПриемСотрудниковToolStripMenuItem, Me.УвольнениеСотрудниковToolStripMenuItem})
        Me.СотрудникиToolStripMenuItem.Name = "СотрудникиToolStripMenuItem"
        Me.СотрудникиToolStripMenuItem.Size = New System.Drawing.Size(85, 20)
        Me.СотрудникиToolStripMenuItem.Text = "Сотрудники"
        '
        'ПриемСотрудниковToolStripMenuItem
        '
        Me.ПриемСотрудниковToolStripMenuItem.Name = "ПриемСотрудниковToolStripMenuItem"
        Me.ПриемСотрудниковToolStripMenuItem.Size = New System.Drawing.Size(213, 22)
        Me.ПриемСотрудниковToolStripMenuItem.Text = "Прием сотрудников"
        '
        'УвольнениеСотрудниковToolStripMenuItem
        '
        Me.УвольнениеСотрудниковToolStripMenuItem.Name = "УвольнениеСотрудниковToolStripMenuItem"
        Me.УвольнениеСотрудниковToolStripMenuItem.Size = New System.Drawing.Size(213, 22)
        Me.УвольнениеСотрудниковToolStripMenuItem.Text = "Увольнение сотрудников"
        '
        'ОтчетыToolStripMenuItem
        '
        Me.ОтчетыToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.КонвертацияToolStripMenuItem, Me.ЧасовToolStripMenuItem})
        Me.ОтчетыToolStripMenuItem.Name = "ОтчетыToolStripMenuItem"
        Me.ОтчетыToolStripMenuItem.Size = New System.Drawing.Size(60, 20)
        Me.ОтчетыToolStripMenuItem.Text = "Отчеты"
        '
        'КонвертацияToolStripMenuItem
        '
        Me.КонвертацияToolStripMenuItem.Name = "КонвертацияToolStripMenuItem"
        Me.КонвертацияToolStripMenuItem.Size = New System.Drawing.Size(197, 22)
        Me.КонвертацияToolStripMenuItem.Text = "Карточка монтажника"
        '
        'ЧасовToolStripMenuItem
        '
        Me.ЧасовToolStripMenuItem.Name = "ЧасовToolStripMenuItem"
        Me.ЧасовToolStripMenuItem.Size = New System.Drawing.Size(197, 22)
        Me.ЧасовToolStripMenuItem.Text = "36 часов"
        '
        'Форма_MDI
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1507, 590)
        Me.Controls.Add(Me.MenuStrip1)
        Me.IsMdiContainer = True
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "Форма_MDI"
        Me.Text = "Форма_MDI"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents РеестрНарядовToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ПриемНарядовToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ВозвратьНарядовToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ПросмотрНарядовToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ЗагрузкаДанныхToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ЗагрузкаTaskListToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents СотрудникиToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ПриемСотрудниковToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents УвольнениеСотрудниковToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ОтчетыToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents КонвертацияToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ЧасовToolStripMenuItem As ToolStripMenuItem
End Class
