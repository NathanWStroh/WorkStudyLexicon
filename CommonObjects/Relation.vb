'***************************
'* Date    Comments
'*           
'* 
'*
'*
'*
'*
'***************************

Public Class Relation
    Private _termID As Integer
    Private _term As String
    Private _formatNote As String
    Private _definition_Source As String
    Private _currentStatus As String

    Public Sub New()

    End Sub

    Public Property Term_ID() As Integer
        Get
            Term_ID = _termID
        End Get
        Set(ByVal value As Integer)
            _termID = value
        End Set
    End Property

    Public Property Term() As String
        Get
            Term = _term
        End Get
        Set(ByVal value As String)
            _term = value
        End Set
    End Property

    Public Property FormatNote() As String
        Get
            FormatNote = _formatNote
        End Get
        Set(ByVal value As String)
            _formatNote = value
        End Set
    End Property

    Public Property DefintionSource() As String
        Get
            DefintionSource = _definition_Source
        End Get
        Set(ByVal value As String)
            _definition_Source = value
        End Set
    End Property

    Public Property CurrentStatus() As String
        Get
            CurrentStatus = _currentStatus
        End Get
        Set(ByVal value As String)
            _currentStatus = value
        End Set
    End Property
End Class
