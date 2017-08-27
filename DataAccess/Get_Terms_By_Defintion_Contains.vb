'***************************
'* Date         Comments
'* 3-17-14      Created spelling mistake in Stored Procedure name.           
'* 3-17-14      Working correctly, but uneditable by information given by procedure.
'*
'*
'*
'*
'***************************

Imports System.Data.SqlClient
Imports CommonObjects

Public Class Get_Terms_By_Defintion_Contains
    Public Function getTermsAssignedToApprover(ByVal termText As String)
        Dim connection As SqlConnection = DataConnection.getDBConnection

        Dim getCommand As New SqlCommand("dbo.ksp_Get_Term_By_Definition_Contains", connection)
        getCommand.CommandType = CommandType.StoredProcedure
        getCommand.Parameters.AddWithValue("@termText", termText)

        Dim TermList As New List(Of Term)

        Try
            connection.Open()
            Dim reader As SqlDataReader = getCommand.ExecuteReader()

            Do While reader.Read
                Dim myTerm As New Term

                myTerm.TermID = CInt(reader("Term_ID"))
                myTerm.TermName = reader("Term").ToString
                myTerm.Status = reader("Current_Status").ToString
                myTerm.SourceID = CInt(reader("Index"))
                myTerm.Definition = reader("Definition").ToString

                TermList.Add(myTerm)
            Loop

        Catch ex As Exception
            Throw ex
        End Try

        Return TermList
    End Function
End Class
