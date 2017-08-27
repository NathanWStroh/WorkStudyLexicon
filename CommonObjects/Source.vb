'***************************
'* Date     BY  Comments
'*5-13-14   NWS Added In_Use field due to it being added into the database to cause sources to go inactive.
'* 
'*
'*
'*
'*
'***************************

Public Class Source
    Private _sourceID As Integer
    Private _sourceName As String
    Private _sourceDes As String
    Private _inUse As String

    Public Property SourceID() As String
        Get
            SourceID = _sourceID
        End Get
        Set(ByVal value As String)
            _sourceID = value
        End Set
    End Property

    Public Property SourceName() As String
        Get
            SourceName = _sourceName
        End Get
        Set(ByVal value As String)
            _sourceName = value
        End Set
    End Property

    Public Property Description() As String
        Get
            Description = _sourceDes
        End Get
        Set(ByVal value As String)
            _sourceDes = value
        End Set
    End Property

    Public Property InUse() As String
        Get
            InUse = _inUse
        End Get
        Set(ByVal value As String)
            _inUse = value
        End Set
    End Property

End Class
