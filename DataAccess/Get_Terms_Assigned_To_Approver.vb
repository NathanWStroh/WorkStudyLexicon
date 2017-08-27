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

Public Class Get_Terms_Assigned_To_Approver
    Public Function getTermsAssignedToApprover(ByVal approver As String)
        Dim connection As SqlConnection = DataConnection.getDBConnection

        Dim getCommand As New SqlCommand("dbo.ksp_Get_Terms_Assigned_To_Approver", connection)
        getCommand.CommandType = CommandType.StoredProcedure
        getCommand.Parameters.AddWithValue("@approver", approver)

        Dim termList As New List(Of Term)
        Try
            connection.Open()
            Dim reader As SqlDataReader = getCommand.ExecuteReader()

            Do While reader.Read
                Dim myTerm As New Term

                myTerm.TermID = CInt(reader("Term_ID"))
                myTerm.Status_Date = reader("Added_On").ToString
                myTerm.TermName = reader("Term").ToString
                myTerm.FormateNote = reader("Format_Note").ToString
                myTerm.DefinitionSourceName = reader("Definition_Source_Name").ToString
                myTerm.Definition = reader("Definition").ToString

                termList.Add(myTerm)
            Loop
        Catch ex As Exception
            Throw ex
        Finally
            connection.Close()
        End Try
        Return termList
    End Function
End Class
