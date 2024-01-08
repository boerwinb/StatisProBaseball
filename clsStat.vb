Option Strict On
Option Explicit On
Friend Class clsBatStat
	
    Private _Name As String
    Private _AB As Integer
    Private _H As Integer
    Private _D As Integer
    Private _T As Integer
    Private _HR As Integer
    Private _RBI As Integer
    Private _R As Integer
    Private _SB As Integer
    Private _SBA As Integer
    Private _K As Integer
    Private _W As Integer
    Private _HB As Integer
    Private _PB As Integer
    Private _GamesC As Integer
    Private _GamesP As Integer
    Private _Games1B As Integer
    Private _Games2B As Integer
    Private _GamesSS As Integer
    Private _Games3B As Integer
    Private _GamesOF As Integer
    Private _GamesDH As Integer
    Private _Games As Integer
    Private _GameInj As Integer
    Private _SH As Integer
    Private _SF As Integer
    Private _Active As Boolean
    Private _rispAB As Integer
    Private _rispH As Integer
    Private _gidp As Integer
    Private _ibb As Integer

    Private _e As Integer
    Private _a As Integer
    Private _po As Integer
    Private _dp As Integer

    Public Property Name() As String
        Get
            Name = _Name
        End Get
        Set(ByVal Value As String)
            _Name = Value
        End Set
    End Property
    Public Property AB() As Integer
        Get
            AB = _AB
        End Get
        Set(ByVal Value As Integer)
            _AB = Value
        End Set
    End Property
    Public Property H() As Integer
        Get
            H = _H
        End Get
        Set(ByVal Value As Integer)
            _H = Value
        End Set
    End Property
    Public Property D() As Integer
        Get
            D = _D
        End Get
        Set(ByVal Value As Integer)
            _D = Value
        End Set
    End Property
    Public Property T() As Integer
        Get
            T = _T
        End Get
        Set(ByVal Value As Integer)
            _T = Value
        End Set
    End Property
    Public Property HR() As Integer
        Get
            HR = _HR
        End Get
        Set(ByVal Value As Integer)
            _HR = Value
        End Set
    End Property
    Public Property RBI() As Integer
        Get
            RBI = _RBI
        End Get
        Set(ByVal Value As Integer)
            _RBI = Value
        End Set
    End Property
    Public Property R() As Integer
        Get
            R = _R
        End Get
        Set(ByVal Value As Integer)
            _R = Value
        End Set
    End Property
    Public Property SB() As Integer
        Get
            SB = _SB
        End Get
        Set(ByVal Value As Integer)
            _SB = Value
        End Set
    End Property
    Public Property SBA() As Integer
        Get
            SBA = _SBA
        End Get
        Set(ByVal Value As Integer)
            _SBA = Value
        End Set
    End Property
    Public Property E() As Integer
        Get
            E = _E
        End Get
        Set(ByVal Value As Integer)
            _E = Value
        End Set
    End Property
    Public Property K() As Integer
        Get
            K = _K
        End Get
        Set(ByVal Value As Integer)
            _K = Value
        End Set
    End Property
    Public Property W() As Integer
        Get
            W = _W
        End Get
        Set(ByVal Value As Integer)
            _W = Value
        End Set
    End Property
    Public Property HB() As Integer
        Get
            HB = _HB
        End Get
        Set(ByVal Value As Integer)
            _HB = Value
        End Set
    End Property
    Public Property PB() As Integer
        Get
            PB = _PB
        End Get
        Set(ByVal Value As Integer)
            _PB = Value
        End Set
    End Property
    Public Property SH() As Integer
        Get
            SH = _SH
        End Get
        Set(ByVal Value As Integer)
            _SH = Value
        End Set
    End Property
    Public Property SF() As Integer
        Get
            SF = _SF
        End Get
        Set(ByVal Value As Integer)
            _SF = Value
        End Set
    End Property
    Public Property GamesP() As Integer
        Get
            GamesP = _GamesP
        End Get
        Set(ByVal Value As Integer)
            _GamesP = Value
        End Set
    End Property
    Public Property GamesC() As Integer
        Get
            GamesC = _GamesC
        End Get
        Set(ByVal Value As Integer)
            _GamesC = Value
        End Set
    End Property
    Public Property Games1B() As Integer
        Get
            Games1B = _Games1B
        End Get
        Set(ByVal Value As Integer)
            _Games1B = Value
        End Set
    End Property
    Public Property Games2B() As Integer
        Get
            Games2B = _Games2B
        End Get
        Set(ByVal Value As Integer)
            _Games2B = Value
        End Set
    End Property
    Public Property GamesSS() As Integer
        Get
            GamesSS = _GamesSS
        End Get
        Set(ByVal Value As Integer)
            _GamesSS = Value
        End Set
    End Property
    Public Property Games3B() As Integer
        Get
            Games3B = _Games3B
        End Get
        Set(ByVal Value As Integer)
            _Games3B = Value
        End Set
    End Property
    Public Property GamesOF() As Integer
        Get
            GamesOF = _GamesOF
        End Get
        Set(ByVal Value As Integer)
            _GamesOF = Value
        End Set
    End Property
    Public Property GamesDH() As Integer
        Get
            GamesDH = _GamesDH
        End Get
        Set(ByVal Value As Integer)
            _GamesDH = Value
        End Set
    End Property
    Public Property Games() As Integer
        Get
            Games = _Games
        End Get
        Set(ByVal Value As Integer)
            _Games = Value
        End Set
    End Property
    Public Property GamesInj() As Integer
        Get
            GamesInj = _GameInj
        End Get
        Set(ByVal Value As Integer)
            _GameInj = Value
        End Set
    End Property
    Public Property Active() As Boolean
        Get
            Active = _Active
        End Get
        Set(ByVal Value As Boolean)
            _Active = Value
        End Set
    End Property
    Public Property rispAB() As Integer
        Get
            rispAB = _rispAB
        End Get
        Set(ByVal value As Integer)
            _rispAB = value
        End Set
    End Property
    Public Property rispH() As Integer
        Get
            rispH = _rispH
        End Get
        Set(ByVal value As Integer)
            _rispH = value
        End Set
    End Property
    Public Property gidp() As Integer
        Get
            gidp = _gidp
        End Get
        Set(ByVal value As Integer)
            _gidp = value
        End Set
    End Property
    Public Property ibb() As Integer
        Get
            ibb = _ibb
        End Get
        Set(ByVal value As Integer)
            _ibb = value
        End Set
    End Property
    Public Property A() As Integer
        Get
            A = _A
        End Get
        Set(ByVal value As Integer)
            _A = value
        End Set
    End Property
    Public Property DP() As Integer
        Get
            DP = _DP
        End Get
        Set(ByVal value As Integer)
            _DP = value
        End Set
    End Property
    Public Property po() As Integer
        Get
            po = _PO
        End Get
        Set(ByVal value As Integer)
            _PO = value
        End Set
    End Property

    Public Sub Clear()
        _AB = 0
        _H = 0
        _D = 0
        _T = 0
        _HR = 0
        _RBI = 0
        _R = 0
        _SB = 0
        _SBA = 0
        _E = 0
        _K = 0
        _W = 0
        _HB = 0
        _PB = 0
        _SH = 0
        _SF = 0
        _GamesC = 0
        _GamesP = 0
        _Games1B = 0
        _Games2B = 0
        _GamesSS = 0
        _Games3B = 0
        _GamesOF = 0
        _GamesDH = 0
        _Games = 0
        _GameInj = 0
        _Active = True
        _Name = ""
        _rispAB = 0
        _rispH = 0
        _gidp = 0
        _ibb = 0
        _A = 0
        _DP = 0
        _PO = 0
    End Sub

    Public Sub New()
        MyBase.New()
        Clear()
    End Sub
End Class