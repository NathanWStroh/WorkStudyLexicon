'***************************
'* Date    Comments
'* 3-5-14   Added Correct Syntax and proper names to file.           
'*
'*
'*
'*
'***************************

Imports System.Data.SqlClient
Imports CommonObjects

Public Class Get_Related_Terms_By_Term_Id
    Public Function getRelatedTermsByTermId(ByVal termId As Integer)
        Dim connection As SqlConnection = DataConnection.getDBConnection

        Dim getCommand As New SqlCommand("dbo.ksp_Get_Related_Terms_By_Term_Id", connection)
        getCommand.CommandType = CommandType.StoredProcedure
        getCommand.Parameters.AddWithValue("@termId", termId)

        Dim TermList As New List(Of Term)

        Try
            connection.Open()
            Dim reader As SqlDataReader = getCommand.ExecuteReader()
            Do While reader.Read
                Dim myTerm As New Term

                myTerm.TermID = CInt(reader("Term_ID"))
                myTerm.TermName = reader("Term").ToString
                myTerm.FormateNote = reader("Format_Note").ToString
                myTerm.DefinitionSourceName = reader("Definition_Source").ToString
                myTerm.Status = reader("Current_Status").ToString

                TermList.Add(myTerm)
            Loop

        Catch ex As Exception
            Throw ex
        Finally
            connection.Close()
        End Try
        Return TermList
    End Function
End Class
