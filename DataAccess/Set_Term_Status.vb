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

Public Class Set_Term_Status
    Public Function SetTermStatus(ByVal termId As Integer, ByVal status As String, ByVal authorization As String, ByVal authorizationDescription As String)
        Dim connection As SqlConnection = DataConnection.getDBConnection

        Dim insertCommand As New SqlCommand("dbo.ksp_Set_Term_Status", connection)
        insertCommand.CommandType = CommandType.StoredProcedure
        insertCommand.Parameters.AddWithValue("@termId", termId)
        insertCommand.Parameters.AddWithValue("@status", status)
        insertCommand.Parameters.AddWithValue("@authorization", authorization)
        insertCommand.Parameters.AddWithValue("@authorizationReason", authorizationDescription)

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
