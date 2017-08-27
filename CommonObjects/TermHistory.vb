Public Class TermHistory
    Private _LogID As Integer
    Private _termID As Integer
    Private _termName As String
    Private _changeDate As DateTime
    Private _changedBy As String
    Private _changedfield As String
    Private _changeAuth As String
    Private _oldValue As String
    Private _newValue As String

    Public Property LogID() As String
        Get
            LogID = _LogID
        End Get
        Set(ByVal value As String)
            _LogID = value
        End Set
    End Property

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

    Public Property ChangeDate() As DateTime
        Get
            ChangeDate = _changeDate
        End Get
        Set(ByVal value As DateTime)
            _changeDate = value
        End Set
    End Property

    Public Property ChangedBy() As String
        Get
            ChangedBy = _changedBy
        End Get
        Set(ByVal value As String)
            _changedBy = value
        End Set
    End Property

    Public Property ChangedField() As String
        Get
            ChangedField = _changedfield
        End Get
        Set(ByVal value As String)
            _changedfield = value
        End Set
    End Property

    Public Property ChangeAuth() As String
        Get
            ChangeAuth = _changeAuth
        End Get
        Set(ByVal value As String)
            _changeAuth = value
        End Set
    End Property

    Public Property OldValue() As String
        Get
            OldValue = _oldValue
        End Get
        Set(ByVal value As String)
            _oldValue = value
        End Set
    End Property

    Public Property NewValue() As String
        Get
            NewValue = _newValue
        End Get
        Set(ByVal value As String)
            _newValue = value
        End Set
    End Property

End Class
