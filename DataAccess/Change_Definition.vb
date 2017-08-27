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

Public Class Change_Definition

    Public Function changeDefintion(ByVal termID As Integer, ByVal newDefinition As String, ByVal authorization As String, ByVal authorizationReason As String)
        Dim connection As SqlConnection = DataConnection.getDBConnection

        Dim insertCommand As New SqlCommand("dbo.ksp_Change_Definition", connection)
        insertCommand.CommandType = CommandType.StoredProcedure
        insertCommand.Parameters.AddWithValue("@termId", termID)
        insertCommand.Parameters.AddWithValue("@newDefinition", newDefinition)
        insertCommand.Parameters.AddWithValue("@authorization", authorization)
        insertCommand.Parameters.AddWithValue("@authorizationReason", authorizationReason)

        Try
            connection.Open()
            Dim count As Integer = insertCommand.ExecuteNonQuery()
            If count > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Throw ex
        Finally
            connection.Close()
        End Try
    End Function
End Class
