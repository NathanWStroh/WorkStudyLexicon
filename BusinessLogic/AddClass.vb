'AddClass manages all the stored procedures for the business logic layer. 
'***************************
'* Date     Comments
'* 3-5-14   Added Add_Source Section          
'* 
'*
'*
'*
'*
'***************************
Imports DataAccess
Imports CommonObjects

Public Class AddClass
    Public Function addApproverToTerm(ByVal termID As Integer, ByVal appover As String)
        Dim _addApproverToTerm As New Add_Approver_To_Term
        Return _addApproverToTerm.addApproverToTerm(termID, appover)
    End Function

    Public Function addTerm(ByVal term As Term, ByVal authorization As String, ByVal addReason As String)
        Dim _addTerm As New Add_Term
        Return _addTerm.addTerm(term, authorization, addReason)
    End Function

    Public Function addTermByApprover(ByVal termID As Integer, ByVal appover As String)
        Dim _addTermByApprover As New Add_Term_By_Approver
        Return _addTermByApprover.addTermByApprover(termID, appover)
    End Function

    Public Function addTermRelationship(ByVal termA As Term, ByVal termB As Term, ByVal authorization As String, ByVal authorizationDescription As String)
        Dim _addTermRelastionship As New Add_Term_Relationship
        Return _addTermRelastionship.addTermRelationship(termA, termB, authorization, authorizationDescription)
    End Function

    Public Function addSource(ByVal source As Source, ByVal authorization As String, ByVal authorizationReason As String)
        Dim _addSource As New Add_Source
        Return _addSource.addNewSource(source, authorization, authorizationReason)
    End Function
End Class
