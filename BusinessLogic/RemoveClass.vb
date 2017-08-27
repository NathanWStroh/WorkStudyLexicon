'this class handles all removing stored procedures
'***************************
'* Date    Comments
'*           
'* 
'*
'*
'*
'*
'***************************
Imports DataAccess
Imports CommonObjects

Public Class RemoveClass
    Public Sub removeTermRelationship(ByVal termA As Integer, ByVal termB As Integer, ByVal authorization As String, ByVal authorizationDescription As String)
        Dim _removeTermRelationship As New Remove_Term_Relationship
        _removeTermRelationship.RemoveTermRelastionship(termA, termB, authorization, authorizationDescription)
    End Sub


End Class
