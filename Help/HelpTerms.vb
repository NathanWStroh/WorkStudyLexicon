Public Class HelpTerms
    Private _message As String
    Public Function AddingTerms()

        _message = "When the user selects the Add New Term Option in the Terms menu, an alert will open. From here, all the fields will need to be filled in before the submit button will execute the save feature. The Term ID field and Status Date will be generated automatically."
        _message += vbCrLf + vbCrLf + "Once opened, you must complete or close this field before working on any other windows."
        Return _message
    End Function

    Public Function TheView() As String
        _message = "The main view for Terms can be accessed by the selecting 'View Terms' from the menu window. This window auto-generates all the terms currently stored in the data dictionary."
        _message += vbCrLf + vbCrLf + ""
        Return _message
    End Function

    Public Function Navigation() As String
        _message = "1.	Search By:" + vbCrLf + "The 'Search By' drop down at the top left of the window and has two options (Term Name and Definition). This allows the user to search by either of these two options by typing into the 'Contains' field and clicking 'Search.' "
        _message += vbCrLf + vbCrLf
        _message += "2.	View By" + vbCrLf + "'View By' can be seen above the right-hand corner of the terms list. This option will show only the terms that have the selected status."
        _message += vbCrLf + vbCrLf
        _message += "3.	Update, Attach Approver and Term History" + vbCrLf + "These three buttons will not become active until a term has been selected from the terms list. Once selected, each button will allow you to do as the name suggests."
        _message += vbCrLf + vbCrLf
        _message += "4.	Reset Button" + vbCrLf + "There are four buttons that have not been covered yet. The first of which is the 'Reset' button. This button will: Set the 'View By' to a blank option, Clear all text in the 'Contains' field, and Reset the 'Search By' to Term Name."
        Return _message
    End Function
End Class
