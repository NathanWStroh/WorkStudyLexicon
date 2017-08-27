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

Public Class Add_Term_By_Approver
    Public Function addTermByApprover(ByVal termID As String, ByVal appover As String)
        Dim connection As SqlConnection = DataConnection.getDBConnection

        Dim insertCommand As New SqlCommand("dbo.ksp_Add_Term_By_Approver", connection)
        insertCommand.CommandType = CommandType.StoredProcedure
        insertCommand.Parameters.AddWithValue("@termId", termID)
        insertCommand.Parameters.AddWithValue("@approver", appover)

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
