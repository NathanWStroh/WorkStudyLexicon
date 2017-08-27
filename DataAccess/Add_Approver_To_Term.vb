'***************************
'* Date         Comments
'* 3-17-14      Tested and it works.          
'* 
'*
'*
'*
'*
'***************************

Imports System.Data.SqlClient
Imports CommonObjects

Public Class Add_Approver_To_Term

    Public Function addApproverToTerm(ByVal termID As Integer, ByVal appover As String)
        Dim connection As SqlConnection = DataConnection.getDBConnection

        Dim insertCommand As New SqlCommand("dbo.ksp_Assign_Approver_To_Term", connection)
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
