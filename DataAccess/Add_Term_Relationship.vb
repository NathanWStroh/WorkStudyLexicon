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

Public Class Add_Term_Relationship
    Public Function addTermRelationship(ByVal termA As Term, ByVal termB As Term, ByVal authorization As String, ByVal authorizationDescription As String)
        Dim connection As SqlConnection = DataConnection.getDBConnection

        Dim insertCommand As New SqlCommand("dbo.ksp_Add_Term_Relationship", connection)
        insertCommand.CommandType = CommandType.StoredProcedure
        insertCommand.Parameters.AddWithValue("@termA", termA.TermID)
        insertCommand.Parameters.AddWithValue("@termB", termB.TermID)
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
