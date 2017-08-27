Imports CommonObjects
Imports System.Data.SqlClient

Public Class Get_Changed_To_Term

    Private _historyList As New List(Of TermHistory)

    Public Function Get_Changed_To_Term(ByVal term As Term, ByVal startDate As Date)
        Dim connection As SqlConnection = DataConnection.getDBConnection

        Dim getCommand As New SqlCommand("dbo.ksp_Get_Changed_To_Term", connection)
        getCommand.CommandType = CommandType.StoredProcedure

        getCommand.Parameters.AddWithValue("@termId", term.TermID)
        getCommand.Parameters.AddWithValue("@startDate", startDate)

        Dim myHistory As TermHistory

        Try
            connection.Open()
            Dim reader As SqlDataReader = getCommand.ExecuteReader()

            Do While reader.Read
                myHistory = New TermHistory
                myHistory.LogID = CInt(reader("Log_ID"))
                myHistory.TermID = CInt(reader("Item_Identifier"))
                myHistory.TermName = reader("Term").ToString
                myHistory.ChangeDate = CDate(reader("Change_Date"))
                myHistory.ChangedBy = reader("Changed_BY").ToString
                myHistory.ChangedField = reader("Changed_Field").ToString
                myHistory.ChangeAuth = reader("Change_Authorization").ToString
                myHistory.OldValue = reader("Old_Value").ToString
                myHistory.NewValue = reader("New_Value").ToString

                _historyList.Add(myHistory)
            Loop
        Catch ex As Exception

        End Try

        Return _historyList
    End Function
End Class
