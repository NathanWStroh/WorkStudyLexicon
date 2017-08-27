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

Public Class Change_Term_Defeniton_Source

    Public Function changeTermDefintionSource(ByVal termID As Integer, ByVal definitionSource As Integer, ByVal authorization As String, ByVal authorizationDescription As String)
        Dim connection As SqlConnection = DataConnection.getDBConnection

        Dim insertCommand As New SqlCommand("dbo.ksp_Change_Term_Definition_Source", connection)
        insertCommand.CommandType = CommandType.StoredProcedure
        insertCommand.Parameters.AddWithValue("@termId", termID)
        insertCommand.Parameters.AddWithValue("@definitionSource", definitionSource)
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
