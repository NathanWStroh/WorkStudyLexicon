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

Public Class Add_Term
    Public Function addTerm(ByVal term As Term, ByVal authorization As String, ByVal addReason As String)
        Dim connection As SqlConnection = DataConnection.getDBConnection

        Dim insertCommand As New SqlCommand("dbo.ksp_Add_Term", connection)
        insertCommand.CommandType = CommandType.StoredProcedure
        insertCommand.Parameters.AddWithValue("@term", term.TermName)
        insertCommand.Parameters.AddWithValue("@definitionSource", term.SourceID)
        insertCommand.Parameters.AddWithValue("@formatNote", term.FormateNote)
        insertCommand.Parameters.AddWithValue("@definition", term.Definition)
        insertCommand.Parameters.AddWithValue("@authorization", authorization)
        insertCommand.Parameters.AddWithValue("@addReason", addReason)

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
