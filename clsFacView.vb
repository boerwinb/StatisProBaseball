Public Class clsFacView
    Private _action As String
    Private _fname As String
    Private _fvalue As String

    Public Property action() As String
        Get
            Return _action
        End Get
        Set(ByVal value As String)
            _action = value
        End Set
    End Property
    Public Property facField() As String
        Get
            Return _fname
        End Get
        Set(ByVal value As String)
            _fname = value
        End Set
    End Property
    Public Property facValue() As String
        Get
            Return _fvalue
        End Get
        Set(ByVal value As String)
            _fvalue = value
        End Set
    End Property
    Public Sub New()

    End Sub
    Public Sub New(ByVal action As String, ByVal facField As String, ByVal facValue As String)
        _action = action
        _fname = facField
        _fvalue = facValue
    End Sub
End Class
