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

Public Class DataConnection
    Public Shared Function getDBConnection() As SqlConnection
        'Will need to change this to get the proper connection.
        'Creates the connection for the DB. Uses Windows ID as access Token.
        Dim connectionString As String = "Data Source=odsdev.kirkwood.edu;Initial Catalog = Lexicon;Integrated Security = True"

        Return New SqlConnection(connectionString)

    End Function
End Class
