Option Strict On
Option Explicit On
Friend Class clsPitcher
	
    Private _field As String
    Private _throw As String
    Private _pbRange As String
    Private _sr As String
    Private _rr As String
    Private _1Bf As String
    Private _1B7 As String
    Private _1B8 As String
    Private _1B9 As String
    Private _bk As String
    Private _k As String
    Private _w As String
    Private _pb As String
    Private _wp As String
    Private _out As String
    Private _StoR As String
    Private _batCard As String
    Private _player As String
    Private _pitcherSel As Integer
    Private _pitcherNum As Integer
    Private _eOuts As Integer
    Private _available As Boolean
    Private _city As String
    Private _state As String
    Private _age As String
    Private _PitStat As clsPitStat

    Public Property field() As String
        Get
            Field = _field
        End Get
        Set(ByVal Value As String)
            _field = Value
        End Set
    End Property
    Public Property throwField() As String
        Get
            ThrowField = _throw
        End Get
        Set(ByVal Value As String)
            _throw = Value
        End Set
    End Property
    Public Property pbRange() As String
        Get
            PBRange = _pbRange
        End Get
        Set(ByVal Value As String)
            _pbRange = Value
        End Set
    End Property
    Public Property sr() As String
        Get
            SR = _sr
        End Get
        Set(ByVal Value As String)
            _sr = Value
        End Set
    End Property
    Public Property rr() As String
        Get
            RR = _rr
        End Get
        Set(ByVal Value As String)
            _rr = Value
        End Set
    End Property
    Public Property hit1Bf() As String
        Get
            Hit1Bf = _1Bf
        End Get
        Set(ByVal Value As String)
            _1Bf = Value
        End Set
    End Property
    Public Property hit1B7() As String
        Get
            Hit1B7 = _1B7
        End Get
        Set(ByVal Value As String)
            _1B7 = Value
        End Set
    End Property
    Public Property hit1B8() As String
        Get
            Hit1B8 = _1B8
        End Get
        Set(ByVal Value As String)
            _1B8 = Value
        End Set
    End Property
    Public Property hit1B9() As String
        Get
            Hit1B9 = _1B9
        End Get
        Set(ByVal Value As String)
            _1B9 = Value
        End Set
    End Property
    Public Property bk() As String
        Get
            BK = _bk
        End Get
        Set(ByVal Value As String)
            _bk = Value
        End Set
    End Property
    Public Property k() As String
        Get
            K = _k
        End Get
        Set(ByVal Value As String)
            _k = Value
        End Set
    End Property
    Public Property w() As String
        Get
            W = _w
        End Get
        Set(ByVal Value As String)
            _w = Value
        End Set
    End Property
    Public Property pb() As String
        Get
            PB = _pb
        End Get
        Set(ByVal Value As String)
            _pb = Value
        End Set
    End Property
    Public Property wp() As String
        Get
            WP = _wp
        End Get
        Set(ByVal Value As String)
            _wp = Value
        End Set
    End Property
    Public Property out() As String
        Get
            Out = _out
        End Get
        Set(ByVal Value As String)
            _out = Value
        End Set
    End Property
    Public Property StoR() As String
        Get
            StoR = _StoR
        End Get
        Set(ByVal Value As String)
            _StoR = Value
        End Set
    End Property
    Public Property batCard() As String
        Get
            BatCard = _batCard
        End Get
        Set(ByVal Value As String)
            _batCard = Value
        End Set
    End Property
    Public Property player() As String
        Get
            Player = _player
        End Get
        Set(ByVal Value As String)
            _player = Value
        End Set
    End Property
    Public Property pitcherNum() As Integer
        Get
            PitcherNum = _pitcherNum
        End Get
        Set(ByVal Value As Integer)
            _pitcherNum = Value
        End Set
    End Property
    Public Property eOuts() As Integer
        Get
            EOuts = _eOuts
        End Get
        Set(ByVal Value As Integer)
            _eOuts = Value
        End Set
    End Property
    Public Property available() As Boolean
        Get
            Available = _available
        End Get
        Set(ByVal Value As Boolean)
            _available = Value
        End Set
    End Property
    Public Property city() As String
        Get
            Return _city
        End Get
        Set(ByVal value As String)
            _city = value
        End Set
    End Property
    Public Property state() As String
        Get
            Return _state
        End Get
        Set(ByVal value As String)
            _state = value
        End Set
    End Property
    Public Property age() As String
        Get
            Return _age
        End Get
        Set(ByVal value As String)
            _age = value
        End Set
    End Property
    Public Function PitStatPtr() As clsPitStat
        PitStatPtr = _PitStat
    End Function


    Public Function DeterminePOE(ByVal isReliever As Boolean, ByVal isEnteringGame As Boolean) As Integer
        Dim calcPOE As Integer

        If isEnteringGame Then
            If isReliever Then
                calcPOE = IIF(Val(_rr) < 6, 8, 10) - 1
            Else
                calcPOE = IIF(Val(_sr) < 13, 12, 14) - 1
            End If
        Else
            'Entering another inning
            If isReliever Then
                calcPOE = IIF(Val(_rr) < 6, 8, 10) - CInt(Math.Ceiling(_PitStat.ip / 3)) - 1
            Else
                calcPOE = IIF(Val(_sr) < 13, 12, 14) - CInt(Math.Ceiling(_PitStat.ip / 3)) - 1
            End If
        End If
        Return calcPOE
    End Function

    Public Sub Clear()
        _field = ""
        _throw = ""
        _pbRange = ""
        _sr = ""
        _rr = ""
        _1Bf = ""
        _1B7 = ""
        _1B8 = ""
        _1B9 = ""
        _bk = ""
        _k = ""
        _w = ""
        _pb = ""
        _wp = ""
        _out = ""
        _StoR = ""
        _batCard = ""
        _available = True
        _pitcherSel = 0
        _pitcherNum = 0
        _eOuts = 0
        _city = ""
        _state = ""
        _age = ""
        _PitStat.Clear()
    End Sub

    Public Sub New()
        MyBase.New()
        _PitStat = New clsPitStat()
        Clear()
    End Sub
End Class