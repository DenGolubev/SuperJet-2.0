Imports System.Globalization
Module Модуль_Input_Language
    Public Sub Input_English()
        InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(New CultureInfo("en-US"))
    End Sub

    Public Sub Input_Russian()
        InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(New CultureInfo("ru-RU"))
    End Sub
End Module
