'***************************
'* Date     Comments
'* 3-17-14  Corrected the inputs and input names of the parameters for the stored procedure so it works correctly.         
'*          Testing and working.
'*
'*
'*
'*
'***************************


Imports System.Data.SqlClient
Imports CommonObjects

Public Class Add_Source

    Public Function addNewSource(ByVal Source As Source, ByVal user As String, ByVal reason As String)
        Dim connection As SqlConnection = DataConnection.getDBConnection

        Dim insertCommand As New SqlCommand("dbo.ksp_Add_Source", connection)
        insertCommand.CommandType = CommandType.StoredProcedure
        insertCommand.Parameters.AddWithValue("@source", Source.SourceName)
        insertCommand.Parameters.AddWithValue("@sourceDescription", Source.Description)
        insertCommand.Parameters.AddWithValue("@authorization", user)
        insertCommand.Parameters.AddWithValue("@authorizationReason", reason)

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
