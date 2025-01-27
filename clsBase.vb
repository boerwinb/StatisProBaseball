Option Strict On
Option Explicit On
Friend Class clsBase
	
    Private _occupied As Boolean
    Private _runnerIndex As Integer
    Private _pitcherIndex As Integer
    Private _isUnearned As Boolean
    Private _autoRunner As Boolean

    Public Property occupied() As Boolean
        Get
            occupied = _occupied
        End Get
        Set(ByVal Value As Boolean)
            _occupied = Value
        End Set
    End Property

    Public Property unearned() As Boolean
        Get
            unearned = _isUnearned
        End Get
        Set(ByVal Value As Boolean)
            _isUnearned = Value
        End Set
    End Property

    Public Property runner() As Integer
        Get
            runner = _runnerIndex
        End Get
        Set(ByVal Value As Integer)
            _runnerIndex = Value
        End Set
    End Property

    Public Property pitcher() As Integer
        Get
            pitcher = _pitcherIndex
        End Get
        Set(ByVal Value As Integer) 'This property represents the pitcher responsible for this runner
            _pitcherIndex = Value
        End Set
    End Property

    Public Property autoRunner() As Boolean
        Get
            autoRunner = _autoRunner
        End Get
        Set(ByVal Value As Boolean)
            _autoRunner = Value
        End Set
    End Property


    Public Sub Clear()
        _runnerIndex = 0
        _pitcherIndex = 0
        _occupied = False
        _isUnearned = False
        _autoRunner = False
    End Sub

    Public Sub New()
        MyBase.New()
        Clear()
    End Sub
End Class