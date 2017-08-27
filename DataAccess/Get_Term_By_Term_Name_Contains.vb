'***************************
'* Date         Comments
'* 3-17-14      Created code to properly return a list.          
'* 3-17-14      Created an issue with not being using the right SP.
'* 3-17-14      Changed field names for reader to correct names.
'* 3-17-14      Working correctly, but uneditable by information given by procedure.
'*
'*
'***************************

Imports System.Data.SqlClient
Imports CommonObjects

Public Class Get_Term_By_Term_Name_Contains
    Public Function getTermByTermNameContains(ByVal termName As String)
        Dim connection As SqlConnection = DataConnection.getDBConnection

        Dim getCommand As New SqlCommand("dbo.ksp_Get_Term_By_Term_Name_Contains", connection)
        getCommand.CommandType = CommandType.StoredProcedure
        getCommand.Parameters.AddWithValue("@termText", termName)

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
