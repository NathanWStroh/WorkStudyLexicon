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

Public Class Get_Definition_Source
    Private _src As New Source
    Private _srcList As New List(Of Source)
    Private connection As SqlConnection = DataConnection.getDBConnection

    Public Function getDefinitionSource()

        Dim getCommand As New SqlCommand("dbo.ksp_Get_Definition_Source", connection)
        getCommand.CommandType = CommandType.StoredProcedure

        Try
            connection.Open()
            Dim reader As SqlDataReader = getCommand.ExecuteReader()
            Dim source As Source
            Do While reader.Read
                source = New Source

                source.SourceID = CInt(reader("Definition_Source_ID"))
                source.SourceName = reader("Definition_Source_Name").ToString
                source.Description = reader("Definition_Source_Description").ToString
                source.InUse = reader("In_Use").ToString
                _srcList.Add(source)

            Loop
        Catch e As SqlException
            Throw e
        Catch e As DataException
            Throw e
        Catch e As Exception
            Throw e
        Finally
            connection.Close()
        End Try

        Return _srcList
    End Function

    Public Function getDefinitionSourcebyID(ByVal srcID As Integer)

        Dim connection As SqlConnection = DataConnection.getDBConnection

        Dim getCommand As New SqlCommand("dbo.ksp_Get_Definition_Source_By_ID", connection)
        getCommand.CommandType = CommandType.StoredProcedure
        getCommand.Parameters.AddWithValue("@definitionSourceID", srcID)

        Try
            connection.Open()
            Dim reader As SqlDataReader = getCommand.ExecuteReader()
            Dim source As Source
            Do While reader.Read
                source = New Source

                source.SourceID = CInt(reader("Definition_Source_ID"))
                source.SourceName = reader("Definition_Source_Name").ToString
                source.Description = reader("Definition_Source_Description")
                source.InUse = reader("In_Use").ToString
                _src = source
            Loop
        Catch e As SqlException
            Throw e
        Catch e As DataException
            Throw e
        Catch e As Exception
            Throw e
        Finally
            connection.Close()
        End Try

        Return _src
    End Function

    Public Function getDefinitionSourceByDescription(ByVal text As String)

        Dim getCommand As New SqlCommand("dbo.ksp_Get_Definition_Source_By_Description_Contains", connection)
        getCommand.CommandType = CommandType.StoredProcedure
        getCommand.Parameters.AddWithValue("@descriptionText", text)

        Try
            connection.Open()
            Dim reader As SqlDataReader = getCommand.ExecuteReader()
            Dim source As Source
            Do While reader.Read
                source = New Source

                source.SourceID = CInt(reader("Definition_Source_ID"))
                source.SourceName = reader("Definition_Source_Name").ToString
                source.Description = reader("Definition_Source_Description")
                source.InUse = reader("In_Use").ToString

                _srcList.Add(source)
            Loop
        Catch e As SqlException
            Throw e
        Catch e As DataException
            Throw e
        Catch e As Exception
            Throw e
        Finally
            connection.Close()
        End Try

        Return _srcList
    End Function

    Public Function getDefinitionSourceByName(ByVal text As String)
        Dim getCommand As New SqlCommand("dbo.ksp_Get_Definition_Source_By_Name", connection)
        getCommand.CommandType = CommandType.StoredProcedure
        getCommand.Parameters.AddWithValue("@name", text)

        Try
            connection.Open()
            Dim reader As SqlDataReader = getCommand.ExecuteReader()
            Dim source As Source
            Do While reader.Read
                source = New Source

                source.SourceID = CInt(reader("Definition_Source_ID"))
                source.SourceName = reader("Definition_Source_Name").ToString
                source.Description = reader("Definition_Source_Description")
                source.InUse = reader("In_Use").ToString

                _srcList.Add(source)
            Loop
        Catch e As SqlException
            Throw e
        Catch e As DataException
            Throw e
        Catch e As Exception
            Throw e
        Finally
            connection.Close()
        End Try

        Return _srcList

    End Function

End Class
