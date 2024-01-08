Option Strict On
Option Explicit On
Friend Class clsPitStat
	
    Private _ip As Integer
    Private _h As Integer
    Private _r As Integer
    Private _er As Integer
    Private _wins As Integer
    Private _l As Integer
    Private _k As Integer
    Private _w As Integer
    Private _saves As Integer
    Private _hr As Integer
    Private _wp As Integer
    Private _hb As Integer
    Private _bk As Integer
    Private _e As Integer
    Private _po As Integer
    Private _a As Integer
    Private _dp As Integer
    Private _starts As Integer
    Private _reliefs As Integer
    Private _cg As Integer
    Private _sho As Integer
    Private _noHit As Integer
    Private _oneHit As Integer
    Private _perfect As Integer
    Private _inactive As Boolean
    Private _gamesInj As Integer

    Public Property ip() As Integer
        Get
            IP = _ip
        End Get
        Set(ByVal Value As Integer)
            _ip = Value
        End Set
    End Property
    Public Property h() As Integer
        Get
            H = _h
        End Get
        Set(ByVal Value As Integer)
            _h = Value
        End Set
    End Property
    Public Property r() As Integer
        Get
            R = _r
        End Get
        Set(ByVal Value As Integer)
            _r = Value
        End Set
    End Property
    Public Property er() As Integer
        Get
            ER = _er
        End Get
        Set(ByVal Value As Integer)
            _er = Value
        End Set
    End Property
    Public Property wins() As Integer
        Get
            Wins = _wins
        End Get
        Set(ByVal Value As Integer)
            _wins = Value
        End Set
    End Property
    Public Property l() As Integer
        Get
            L = _l
        End Get
        Set(ByVal Value As Integer)
            _l = Value
        End Set
    End Property
    Public Property k() As Integer
        Get
            K = _k
        End Get
        Set(ByVal Value As Integer)
            _k = Value
        End Set
    End Property
    Public Property w() As Integer
        Get
            W = _w
        End Get
        Set(ByVal Value As Integer)
            _w = Value
        End Set
    End Property
    Public Property saves() As Integer
        Get
            Saves = _saves
        End Get
        Set(ByVal Value As Integer)
            _saves = Value
        End Set
    End Property
    Public Property hr() As Integer
        Get
            HR = _hr
        End Get
        Set(ByVal Value As Integer)
            _hr = Value
        End Set
    End Property
    Public Property wp() As Integer
        Get
            WP = _wp
        End Get
        Set(ByVal Value As Integer)
            _wp = Value
        End Set
    End Property
    Public Property hb() As Integer
        Get
            HB = _hb
        End Get
        Set(ByVal Value As Integer)
            _hb = Value
        End Set
    End Property
    Public Property bk() As Integer
        Get
            BK = _bk
        End Get
        Set(ByVal Value As Integer)
            _bk = Value
        End Set
    End Property
    Public Property e() As Integer
        Get
            E = _e
        End Get
        Set(ByVal Value As Integer)
            _e = Value
        End Set
    End Property
    Public Property po() As Integer
        Get
            po = _po
        End Get
        Set(ByVal value As Integer)
            _po = value
        End Set
    End Property
    Public Property a() As Integer
        Get
            a = _a
        End Get
        Set(ByVal value As Integer)
            _a = value
        End Set
    End Property
    Public Property dp() As Integer
        Get
            dp = _dp
        End Get
        Set(ByVal value As Integer)
            _dp = value
        End Set
    End Property
    Public Property starts() As Integer
        Get
            Starts = _starts
        End Get
        Set(ByVal Value As Integer)
            _starts = Value
        End Set
    End Property
    Public Property reliefs() As Integer
        Get
            Reliefs = _reliefs
        End Get
        Set(ByVal Value As Integer)
            _reliefs = Value
        End Set
    End Property
    Public Property cg() As Integer
        Get
            CG = _cg
        End Get
        Set(ByVal Value As Integer)
            _cg = Value
        End Set
    End Property
    Public Property sho() As Integer
        Get
            SHO = _sho
        End Get
        Set(ByVal Value As Integer)
            _sho = Value
        End Set
    End Property
    Public Property noHit() As Integer
        Get
            NoHit = _noHit
        End Get
        Set(ByVal Value As Integer)
            _noHit = Value
        End Set
    End Property
    Public Property oneHit() As Integer
        Get
            OneHit = _oneHit
        End Get
        Set(ByVal Value As Integer)
            _oneHit = Value
        End Set
    End Property
    Public Property perfect() As Integer
        Get
            Perfect = _perfect
        End Get
        Set(ByVal Value As Integer)
            _perfect = Value
        End Set
    End Property
    Public Property inactive() As Boolean
        Get
            Inactive = _inactive
        End Get
        Set(ByVal Value As Boolean)
            _inactive = Value
        End Set
    End Property
    Public Property gamesInj() As Integer
        Get
            GamesInj = _gamesInj
        End Get
        Set(ByVal Value As Integer)
            _gamesInj = Value
        End Set
    End Property

    Public Sub Clear()
        _ip = 0
        _h = 0
        _r = 0
        _ER = 0
        _wins = 0
        _l = 0
        _k = 0
        _W = 0
        _saves = 0
        _hr = 0
        _wp = 0
        _hb = 0
        _bk = 0
        _e = 0
        _po = 0
        _a = 0
        _dp = 0
        _starts = 0
        _reliefs = 0
        _cg = 0
        _sho = 0
        _noHit = 0
        _oneHit = 0
        _perfect = 0
        _inactive = False
        _gamesInj = 0
    End Sub

    Public Sub New()
        MyBase.New()
        Clear()
    End Sub
End Class