'This class handles all stored procedures that have to retrieve information
'***************************
'* Date     Comments
'* 3-17-14  added function for getting term by id.          
'* 
'*
'*
'*
'*
'***************************
Imports DataAccess
Imports CommonObjects

Public Class GetClass

    Public Function getAvaliableStatuses()
        Dim myGetAvaliableStatuses As New Get_Avaliable_Statuses
        Return myGetAvaliableStatuses.getAvailableStatuses()
    End Function

    '***Source BLL***'
    Public Function getDefintionSource()
        Dim gds As New Get_Definition_Source
        Return gds.getDefinitionSource()
    End Function

    Public Function getDefinitionSourceByID(ByVal srcID As Integer)
        Dim gds As New Get_Definition_Source
        Return gds.getDefinitionSourcebyID(srcID)
    End Function

    Public Function getDefinionSrcByName(ByVal text As String)
        Dim gds As New Get_Definition_Source
        Return gds.getDefinitionSourceByName(text)
    End Function

    Public Function getDefinionSrcByDef(ByVal text As String)
        Dim gds As New Get_Definition_Source
        Return gds.getDefinitionSourceByDescription(text)
    End Function

    '************TERM BBL********'
    Public Function getRelatedTermsByTermID(ByVal termId As Integer)
        Dim myGetRelatedTermsByTermID As New Get_Related_Terms_By_Term_Id
        Return myGetRelatedTermsByTermID.getRelatedTermsByTermId(termId)
    End Function

    Public Function getTermByTermNameContains(ByVal termText As String)
        Dim myGetTermNameContains As New Get_Term_By_Term_Name_Contains
        Return myGetTermNameContains.getTermByTermNameContains(termText)
    End Function

    Public Function getTermsAssignedToApprover(ByVal approver As String)
        Dim myGetTermsByApprover As New Get_Terms_Assigned_To_Approver
        Return myGetTermsByApprover.getTermsAssignedToApprover(approver)
    End Function

    Public Function getTermsByDefinitionContains(ByVal termText As String)
        Dim myGetTermsByDefinition As New Get_Terms_By_Defintion_Contains
        Return myGetTermsByDefinition.getTermsAssignedToApprover(termText)
    End Function

    Public Function getTermsByTermNameSoundex(ByVal termText As String)
        Dim myGetTermsBySoundex As New Get_Terms_By_Term_Name_Soundex
        Return myGetTermsBySoundex.getTermsByTermNameSoundex(termText)
    End Function

    Public Function getUnrelatedTermsByTermId(ByVal termId As Integer)
        Dim myGetUnrelatedTermsByTermID As New Get_Unrelatd_Terms_By_Term_Id
        Return myGetUnrelatedTermsByTermID.getGetUnrelatedTermsByTermId(termId)
    End Function

    Public Function getAllTerms()
        Dim myGetAllTerms As New Get_All_Terms
        Return myGetAllTerms.Get_All_Terms
    End Function

    Public Function getTermByID(ByVal termID As Integer)
        Dim myGetTermByID As New Get_Term_By_ID
        Return myGetTermByID.Get_Term_By_ID(termID)

    End Function

    Public Function getTermByStatus(ByVal status As String)
        Dim gtbs As New Get_Term_by_Status
        Return gtbs.Get_Term_by_Status(status)
    End Function

    Public Function getTermHistory(ByVal term As Term, ByVal startDate As Date)
        Dim gctt As New Get_Changed_To_Term
        Return gctt.Get_Changed_To_Term(term, startDate)
    End Function

End Class
