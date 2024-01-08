Option Strict On
Option Explicit On
Friend Class clsBatter
	
    Private _field As String
    Private _error As String
    Private _throw As String
    Private _obr As String
    Private _sp As String
    Private _hitRun As String
    Private _cd As String
    Private _cdAct As String
    Private _sac As String
    Private _inj As String
    Private _1bf As String
    Private _1b7 As String
    Private _1b8 As String
    Private _1b9 As String
    Private _2b7 As String
    Private _2b8 As String
    Private _2B9 As String
    Private _3B8 As String
    Private _hr As String
    Private _k As String
    Private _w As String
    Private _hpb As String
    Private _out As String
    Private _cht As String
    Private _bd As String
    Private _player As String
    Private _playerIndex As Integer
    Private _position As String
    Private _available As Boolean
    Private _playedLast As Boolean
    Private _hitLast As Boolean
    Private _city As String
    Private _state As String
    Private _age As String
    Private _BatStat As clsBatStat

    Public Property field() As String
        Get
            field = _field
        End Get
        Set(ByVal Value As String)
            _field = Value
        End Set
    End Property
    Public Property errorRating() As String
        Get
            errorRating = _error
        End Get
        Set(ByVal Value As String)
            _error = Value
        End Set
    End Property
    Public Property throwRating() As String
        Get
            throwRating = _throw
        End Get
        Set(ByVal Value As String)
            _throw = Value
        End Set
    End Property
    Public Property obr() As String
        Get
            obr = _obr
        End Get
        Set(ByVal Value As String)
            _obr = Value
        End Set
    End Property
    Public Property sp() As String
        Get
            sp = _sp
        End Get
        Set(ByVal Value As String)
            _sp = Value
        End Set
    End Property
    Public Property hitRun() As String
        Get
            hitRun = _hitRun
        End Get
        Set(ByVal Value As String)
            _hitRun = Value
        End Set
    End Property
    Public Property cd() As String
        Get
            cd = _cd
        End Get
        Set(ByVal Value As String)
            _cd = Value
        End Set
    End Property
    Public Property cdAct() As String
        Get
            cdAct = _cdAct
        End Get
        Set(ByVal Value As String)
            _cdAct = Value
        End Set
    End Property
    Public Property sac() As String
        Get
            sac = _sac
        End Get
        Set(ByVal Value As String)
            _sac = Value
        End Set
    End Property
    Public Property inj() As String
        Get
            inj = _inj
        End Get
        Set(ByVal Value As String)
            _inj = Value
        End Set
    End Property
    Public Property hit1Bf() As String
        Get
            hit1Bf = _1bf
        End Get
        Set(ByVal Value As String)
            _1bf = Value
        End Set
    End Property
    Public Property hit1B7() As String
        Get
            hit1B7 = _1b7
        End Get
        Set(ByVal Value As String)
            _1b7 = Value
        End Set
    End Property
    Public Property hit1B8() As String
        Get
            hit1B8 = _1b8
        End Get
        Set(ByVal Value As String)
            _1b8 = Value
        End Set
    End Property
    Public Property hit1B9() As String
        Get
            hit1B9 = _1b9
        End Get
        Set(ByVal Value As String)
            _1b9 = Value
        End Set
    End Property
    Public Property hit2B7() As String
        Get
            hit2B7 = _2b7
        End Get
        Set(ByVal Value As String)
            _2b7 = Value
        End Set
    End Property
    Public Property hit2B8() As String
        Get
            hit2B8 = _2b8
        End Get
        Set(ByVal Value As String)
            _2b8 = Value
        End Set
    End Property
    Public Property hit2B9() As String
        Get
            hit2B9 = _2B9
        End Get
        Set(ByVal Value As String)
            _2B9 = Value
        End Set
    End Property
    Public Property hit3B8() As String
        Get
            hit3B8 = _3B8
        End Get
        Set(ByVal Value As String)
            _3B8 = Value
        End Set
    End Property
    Public Property hr() As String
        Get
            hr = _HR
        End Get
        Set(ByVal Value As String)
            _HR = Value
        End Set
    End Property
    Public Property k() As String
        Get
            k = _K
        End Get
        Set(ByVal Value As String)
            _K = Value
        End Set
    End Property
    Public Property w() As String
        Get
            w = _W
        End Get
        Set(ByVal Value As String)
            _W = Value
        End Set
    End Property
    Public Property hpb() As String
        Get
            hpb = _HPB
        End Get
        Set(ByVal Value As String)
            _HPB = Value
        End Set
    End Property
    Public Property out() As String
        Get
            out = _Out
        End Get
        Set(ByVal Value As String)
            _Out = Value
        End Set
    End Property
    Public Property cht() As String
        Get
            cht = _Cht
        End Get
        Set(ByVal Value As String)
            _Cht = Value
        End Set
    End Property
    Public Property bd() As String
        Get
            bd = _BD
        End Get
        Set(ByVal Value As String)
            _BD = Value
        End Set
    End Property
    Public Property player() As String
        Get
            player = _Player
        End Get
        Set(ByVal Value As String)
            _Player = Value
        End Set
    End Property
    Public Property playerIndex() As Integer
        Get
            playerIndex = _PlayerIndex
        End Get
        Set(ByVal Value As Integer)
            _PlayerIndex = Value
        End Set
    End Property
    Public Property position() As String
        Get
            position = _Position
        End Get
        Set(ByVal Value As String)
            _position = Value
        End Set
    End Property
    Public Property available() As Boolean
        Get
            available = _Available
        End Get
        Set(ByVal Value As Boolean)
            _Available = Value
        End Set
    End Property
    Public Property playedLast() As Boolean
        Get
            playedLast = _PlayedLast
        End Get
        Set(ByVal Value As Boolean)
            _PlayedLast = Value
        End Set
    End Property
    Public Property hitLast() As Boolean
        Get
            hitLast = _HitLast
        End Get
        Set(ByVal Value As Boolean)
            _HitLast = Value
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
    Public Function BatStatPtr() As clsBatStat
        batStatPtr = _BatStat
    End Function

    Public Sub Clear()
        _field = ""
        _error = ""
        _throw = ""
        _obr = ""
        _sp = ""
        _hitRun = ""
        _cd = ""
        _cdAct = ""
        _sac = ""
        _inj = ""
        _1bf = ""
        _1b7 = ""
        _1b8 = ""
        _1b9 = ""
        _2b7 = ""
        _2b8 = ""
        _2B9 = ""
        _3B8 = ""
        _HR = ""
        _K = ""
        _W = ""
        _HPB = ""
        _Out = ""
        _Cht = ""
        _bd = ""
        _city = ""
        _state = ""
        _age = ""
        _Player = ""
        _PlayerIndex = 0
        _Available = True
        _HitLast = False
        _PlayedLast = False
        _BatStat.Clear()
    End Sub

    Public Sub New()
        MyBase.New()
        _BatStat = New clsBatStat()
        Clear()
    End Sub
End Class