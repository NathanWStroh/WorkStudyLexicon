'***************************
'* Date    Comments
'*           
'* 
'*
'*
'*
'*
'***************************

Imports System.Data.SqlClient
Imports CommonObjects

Public Class Get_Avaliable_Statuses
    Public Function getAvailableStatuses()
        Dim connection As SqlConnection = DataConnection.getDBConnection

        Dim getCommand As New SqlCommand("dbo.ksp_Get_Available_Statuses", connection)
        getCommand.CommandType = CommandType.StoredProcedure

        Dim StatusList As New List(Of Status)

        Try
            connection.Open()
            Dim reader As SqlDataReader = getCommand.ExecuteReader()

            Do While reader.Read
                Dim myStatus As New Status
                myStatus.Status = reader("Status").ToString
                myStatus.Description = reader("Status_Description").ToString

                StatusList.Add(myStatus)
            Loop

        Catch e As Exception
            Throw e
        Finally
            connection.Close()
        End Try
        Return StatusList
    End Function
End Class
