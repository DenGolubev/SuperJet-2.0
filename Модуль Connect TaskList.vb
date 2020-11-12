Imports System.Data.OleDb

Module Модуль_Connect_TaskList
    Public con_excel As New OleDbConnection
    Dim db_com As New OleDbCommand

    Sub Rec(DGV As DataGridView)
        Input_string("UPDATE [TaskList Total] INNER JOIN [time_table] On [TaskList Total].[UIN] = [time_table].[UIN] Set [TaskList Total].[Status] = [time_table].[Status], 
[TaskList Total].[Ispolnitel] = [time_table].[Ispolnitel], [TaskList Total].[DTCS] = [time_table].[DTCS], [TaskList Total].[Internet_tarif] = [time_table].[Internet_tarif], 
[TaskList Total].[Internet_status] = [time_table].[Internet_status], [TaskList Total].[TV_tarif] = [time_table].[TV_tarif], [TaskList Total].[TV_status] = [time_table].[TV_status], 
[TaskList Total].[Povtorki] = [time_table].[Povtorki]")
        Input_string("Insert into [TaskList Total] Select * from time_table where [UIN] not in (Select [UIN] from [TaskList Total])")
    End Sub
End Module