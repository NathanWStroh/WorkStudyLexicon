'***************************
'* Date     Comments
'* 3-5-14   Corrected typo in SP that made it unusable.          
'* 
'*
'*
'*
'*
'***************************

Imports System.Data.SqlClient
Imports CommonObjects

Public Class Get_Unrelatd_Terms_By_Term_Id
    Public Function getGetUnrelatedTermsByTermId(ByVal termId As Integer)
        Dim connection As SqlConnection = DataConnection.getDBConnection
        Dim TermList As New List(Of Term)

        Dim insertCommand As New SqlCommand("dbo.ksp_Get_Unrelated_Terms_By_Term_ID", connection)
        insertCommand.CommandType = CommandType.StoredProcedure
        insertCommand.Parameters.AddWithValue("@termId", termId)

        Try
            connection.Open()
            Dim reader As SqlDataReader = insertCommand.ExecuteReader()

            Do While reader.Read
                Dim myTerm As New Term

                myTerm.TermID = reader("Term_ID").ToString
                myTerm.TermName = reader("Term").ToString
                myTerm.Definition = reader("Definition").ToString

                TermList.Add(myTerm)
            Loop

        Catch ex As SqlException
            MsgBox("Database Error. Please contact you system administor. \r\n Error: " & vbCr & ex.ToString)
        Catch ex As DataException
            MsgBox("Connection Error. Please contact The Delevelopement Team")
        Catch ex As Exception
            MsgBox("An Unknown Error has occured. Try again and report the error if it persists: " & vbCr & ex.ToString)
        Finally
            connection.Close()
        End Try
        Return TermList
    End Function
End Class
