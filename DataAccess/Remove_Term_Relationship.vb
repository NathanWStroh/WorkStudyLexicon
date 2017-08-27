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

Public Class Remove_Term_Relationship
    Public Function RemoveTermRelastionship(ByVal termA As Integer, ByVal termB As Integer, ByVal authorization As String, ByVal authorizationDescription As String)
        Dim connection As SqlConnection = DataConnection.getDBConnection

        Dim insertCommand As New SqlCommand("dbo.ksp_Remove_Term_Relationship", connection)
        insertCommand.CommandType = CommandType.StoredProcedure
        insertCommand.Parameters.AddWithValue("@termA", termA)
        insertCommand.Parameters.AddWithValue("@termB", termB)
        insertCommand.Parameters.AddWithValue("@authorization", authorization)
        insertCommand.Parameters.AddWithValue("@authorizationDescription", authorizationDescription)

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
