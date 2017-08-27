'this class handles all stored procedures that change/update the db
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

Public Class ChangeClass

    Private _updateSrc As New Update_Source

    'change term variables
    Public Sub changeDefinition(ByVal termID As Integer, ByVal newDefinition As String, ByVal authorization As String, ByVal authorizationReason As String)
        Dim _changeDefinition As New Change_Definition
        _changeDefinition.changeDefintion(termID, newDefinition, authorization, authorizationReason)

    End Sub

    Public Sub changeTerm(ByVal termID As Integer, ByVal newTerm As String, ByVal authorization As String, ByVal authorizationReason As String)
        Dim _changeTerm As New Change_Term
        _changeTerm.changeTerm(termID, newTerm, authorization, authorizationReason)

    End Sub

    Public Sub changeTermDefinitionSource(ByVal termID As Integer, ByVal definitionSource As Integer, ByVal authorization As String, ByVal authorizationDescription As String)
        Dim _changeTermDefinitionSource As New Change_Term_Defeniton_Source
        _changeTermDefinitionSource.changeTermDefintionSource(termID, definitionSource, authorization, authorizationDescription)

    End Sub

    Public Sub setTermStatus(ByVal termID As Integer, ByVal status As String, ByVal authorization As String, ByVal authorizationDescription As String)
        Dim _setTermStatus As New Set_Term_Status
        _setTermStatus.SetTermStatus(termID, status, authorization, authorizationDescription)
    End Sub

    'source changes
    Public Function setSourceToInactive(ByVal src As Source, ByVal auth As String, ByVal reason As String)
        Return _updateSrc.Invalidate_Definition_Source(src, auth, reason)
    End Function

    Public Function setNewSourceDef(ByVal src As Source, ByVal auth As String, ByVal reason As String)
        Return _updateSrc.Change_Definition_Source_Definition(src, auth, reason)
    End Function

    Public Function setNewSourceName(ByVal src As Source, ByVal auth As String, ByVal reason As String)
        Return _updateSrc.Change_Definition_Source_Name(src, auth, reason)
    End Function

End Class
