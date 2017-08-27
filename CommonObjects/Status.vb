'***************************
'* Date    Comments
'*           
'* 
'*
'*
'*
'*
'***************************

Public Class Status

    Private _statusName As String
    Private _statusDes As String

    Public Sub New()

    End Sub

    Public Property Status() As String
        Get
            Status = _statusName
        End Get
        Set(ByVal value As String)
            _statusName = value
        End Set
    End Property

    Public Property Description() As String
        Get
            Description = _statusDes
        End Get
        Set(ByVal value As String)
            _statusDes = value
        End Set
    End Property


End Class
