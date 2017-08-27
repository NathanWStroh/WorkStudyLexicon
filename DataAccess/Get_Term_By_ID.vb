'***************************
'* Created 3-17-14 by Nathan W Stroh
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

Public Class Get_Term_By_ID
    Private _myTerm As New Term

    Public Function Get_Term_By_ID(ByVal termID As Integer)
        Dim connection As SqlConnection = DataConnection.getDBConnection

        Dim getCommand As New SqlCommand("dbo.ksp_Get_Term_By_ID", connection)
        getCommand.CommandType = CommandType.StoredProcedure

        getCommand.Parameters.AddWithValue("@termId", termID)

        Dim myTerm As Term

        Try
            connection.Open()
            Dim reader As SqlDataReader = getCommand.ExecuteReader()

            Do While reader.Read
                myTerm = New Term
                myTerm.TermID = CInt(reader("Term_ID"))
                myTerm.TermName = reader("Term").ToString
                myTerm.SourceID = CInt(reader("Definition_Source_ID"))
                myTerm.DefinitionSourceName = reader("Definition_Source_Name").ToString
                myTerm.FormateNote = reader("Format_Note").ToString
                myTerm.Definition = reader("Definition").ToString
                myTerm.Status = reader("Status").ToString
                myTerm.Status_Date = reader("Current_Status_Date")

                _myTerm = myTerm
            Loop

        Catch ex As Exception
            Throw ex
        End Try

        Return _myTerm
    End Function
End Class
