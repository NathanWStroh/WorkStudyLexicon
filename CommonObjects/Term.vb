'***************************
'* Date    Comments
'* 3-5      Changed  line _termID = TermID to _termID = value.          
'* 
'*
'*
'*
'*
'***************************

Public Class Term

    Private _termID As Integer
    Private _termName As String
    Private _sourceID As Integer
    Private _definitionSourceName As String
    Private _formatNote As String
    Private _definition As String
    Private _status As String
    Private _statusDate As String

    Public Property TermID() As Integer
        Get
            TermID = _termID
        End Get
        Set(ByVal value As Integer)
            _termID = value
        End Set
    End Property

    Public Property TermName() As String
        Get
            TermName = _termName
        End Get
        Set(ByVal value As String)
            _termName = value
        End Set
    End Property

    Public Property SourceID() As Integer
        Get
            SourceID = _sourceID
        End Get
        Set(ByVal value As Integer)
            _sourceID = value
        End Set
    End Property

    Public Property DefinitionSourceName() As String
        Get
            DefinitionSourceName = _definitionSourceName
        End Get
        Set(ByVal value As String)
            _definitionSourceName = value
        End Set
    End Property

    Public Property FormateNote() As String
        Get
            FormateNote = _formatNote
        End Get
        Set(ByVal value As String)
            _formatNote = value
        End Set
    End Property

    Public Property Definition() As String
        Get
            Definition = _definition
        End Get
        Set(ByVal value As String)
            _definition = value
        End Set
    End Property

    Public Property Status() As String
        Get
            Status = _status
        End Get
        Set(ByVal value As String)
            _status = value
        End Set
    End Property

    Public Property Status_Date() As String
        Get
            Status_Date = _statusDate
        End Get
        Set(ByVal value As String)
            _statusDate = value
        End Set
    End Property
End Class
