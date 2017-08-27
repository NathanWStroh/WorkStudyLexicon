'***************************
'* Date     Comments
'* 3-17-14  Testing and working.       
'* 
'*
'*
'*
'*
'***************************

Imports System.Data.SqlClient
Imports CommonObjects

Public Class Get_All_Terms

    Public ReadOnly rows As New List(Of Dictionary(Of String, String))

    Public Function Get_All_Terms()
        Dim connection As SqlConnection = DataConnection.getDBConnection

        Dim getCommand As New SqlCommand("dbo.ksp_Get_All_Terms", connection)

        Dim TermList As New List(Of Term)

        Try
            connection.Open()
            Dim reader As SqlDataReader = getCommand.ExecuteReader()

            Do While reader.Read
                Dim myTerm As New Term

                myTerm.TermID = CInt(reader("Term_ID"))
                myTerm.TermName = reader("Term").ToString
                myTerm.SourceID = CInt(reader("Definition_Source_ID"))
                myTerm.DefinitionSourceName = reader("Definition_Source_Name").ToString
                myTerm.FormateNote = reader("Format_Note").ToString
                myTerm.Definition = reader("Definition").ToString
                myTerm.Status = reader("Status").ToString
                myTerm.Status_Date = reader("Status_Date").ToString

                TermList.Add(myTerm)
            Loop

        Catch ex As Exception
            Throw ex
        End Try

        Return TermList
    End Function
End Class
