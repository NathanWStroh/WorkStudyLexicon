Imports System.Data.SqlClient
Imports CommonObjects

Public Class Get_Term_by_Status
    Public Function Get_Term_by_Status(ByVal status As String)
        Dim connection As SqlConnection = DataConnection.getDBConnection
        Dim getCommand As New SqlCommand("dbo.ksp_Get_Terms_By_Status", connection)
        getCommand.CommandType = CommandType.StoredProcedure
        getCommand.Parameters.AddWithValue("@status", status)

        Dim termList As New List(Of Term)

        Try
            connection.Open()
            Dim reader As SqlDataReader = getCommand.ExecuteReader()
            Do While reader.Read
                Dim myTerm As New Term
                myTerm.TermID = CInt(reader("Term_ID"))
                myTerm.TermName = reader("Term").ToString
                myTerm.DefinitionSourceName = reader("Definition_Source_Name").ToString
                myTerm.Status = reader("Status").ToString
                myTerm.Status_Date = CDate(reader("Status_Date"))
                myTerm.FormateNote = reader("Format_Note").ToString
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
