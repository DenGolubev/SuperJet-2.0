Imports System.Security.Cryptography
Imports System.Text
Imports System.IO
Module Модуль_сравнения_БД
    Public db_path As String = Application.StartupPath
    'Public db_path As String = "\\mgts-dfs-05.mgts.corp.net\Folders$\117925"
    Public bd_name As String = "SuperJet 2.0"
    Public bak As String = "SuperJet bak"
    Public db_pass As String = "" 'пароль базы данных

    Function file_1(ByVal md5code As String) As String
        Dim md5 As MD5CryptoServiceProvider = New MD5CryptoServiceProvider
        Dim f As FileStream = New FileStream(db_path & "\" & bd_name & ".mdb", FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
        f = New FileStream(db_path & "\" & bd_name & ".mdb", FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
        md5.ComputeHash(f)
        Dim ObjFSO As Object = CreateObject("Scripting.FileSystemObject")
        Dim objFile = ObjFSO.GetFile(db_path & "\" & bd_name & ".mdb")

        Dim hash As Byte() = md5.Hash
        Dim buff As StringBuilder = New StringBuilder
        Dim hashByte As Byte
        For Each hashByte In hash
            buff.Append(String.Format("{0:X1}", hashByte))
        Next

        md5code = buff.ToString()
        Return md5code

        f.Close()
        f.Dispose()
    End Function

    Function file_2(ByVal md51code As String) As String
        Dim md51 As MD5CryptoServiceProvider = New MD5CryptoServiceProvider
        Dim f1 As FileStream = New FileStream(db_path & "\" & bak & ".mdb", FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
        f1 = New FileStream(db_path & "\" & bak & ".mdb", FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
        md51.ComputeHash(f1)
        Dim ObjFSO1 As Object = CreateObject("Scripting.FileSystemObject")
        Dim objFile1 = ObjFSO1.GetFile(db_path & "\" & bak & ".mdb")

        Dim hash1 As Byte() = md51.Hash
        Dim buff1 As StringBuilder = New StringBuilder
        Dim hashByte1 As Byte
        For Each hashByte1 In hash1
            buff1.Append(String.Format("{0:X1}", hashByte1))
        Next

        md51code = buff1.ToString()
        Return md51code
        f1.Close()
        f1.Dispose()
    End Function
End Module
