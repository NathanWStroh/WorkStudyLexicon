Imports CommonObjects
Imports System.Data.SqlClient
Public Class Update_Source
    Private connection As SqlConnection = DataConnection.getDBConnection

    '*************Changes a Source to inactive.
    Public Function Invalidate_Definition_Source(ByVal src As Source, ByVal auth As String, ByVal reason As String)
        Dim insertcommand As New SqlCommand("dbo.ksp_Invalidate_Definition_Source", connection)
        insertcommand.CommandType = CommandType.StoredProcedure
        insertcommand.Parameters.AddWithValue("@definitionSourceId", src.SourceID)
        insertcommand.Parameters.AddWithValue("@authorization", auth)
        insertcommand.Parameters.AddWithValue("@authorizationReason", reason)

        Try
            connection.Open()
            Dim count As Integer = insertcommand.ExecuteNonQuery
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

    '*************Changes a Source's Definition
    Public Function Change_Definition_Source_Name(ByVal src As Source, ByVal auth As String, ByVal reason As String)
        Dim insertcommand As New SqlCommand("dbo.ksp_Change_Definition_Source_Name", connection)
        insertcommand.CommandType = CommandType.StoredProcedure
        insertcommand.Parameters.AddWithValue("@definitionSourceId", src.SourceID)
        insertcommand.Parameters.AddWithValue("@newName", src.SourceName)
        insertcommand.Parameters.AddWithValue("@authorization", auth)
        insertcommand.Parameters.AddWithValue("@authorizationReason", reason)

        Try
            connection.Open()
            Dim count As Integer = insertcommand.ExecuteNonQuery
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

    '*************Changes a Source's Name 
    Public Function Change_Definition_Source_Definition(ByVal src As Source, ByVal auth As String, ByVal reason As String)

        Dim insertcommand As New SqlCommand("dbo.ksp_Change_Definition_Source_Description", connection)
        insertcommand.CommandType = CommandType.StoredProcedure
        insertcommand.Parameters.AddWithValue("@definitionSourceId", src.SourceID)
        insertcommand.Parameters.AddWithValue("@newDescription", src.Description)
        insertcommand.Parameters.AddWithValue("@authorization", auth)
        insertcommand.Parameters.AddWithValue("@authorizationReason", reason)

        Try
            connection.Open()
            Dim count As Integer = insertcommand.ExecuteNonQuery
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
