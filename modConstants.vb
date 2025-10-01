Option Strict On
Option Explicit On

Imports System.Collections.Generic

Module modConstants
	' Team constants
    Public Const conMaxBatters As Integer = 38
    Public Const conMaxPitchers As Integer = 35

    ' Master Team Objects
    Public Home As clsTeam
    Public Visitor As clsTeam
    Public Game As clsGame
    Public Season As clsSeason

    ' Base Objects
    Public FirstBase As clsBase
    Public SecondBase As clsBase
    Public ThirdBase As clsBase
    'Public Plate As clsBase

    ' Required positions
    Public Const conALPos As String = "DH|C|1B|2B|3B|SS|LF|CF|RF|"
    Public Const conNLPos As String = "P|C|1B|2B|3B|SS|LF|CF|RF|"

    Public bolAmericanLeagueRules As Boolean
    Public bol3BatterMinimum As Boolean
    Public bolAutomaticRunner As Boolean
    Public gbolSeason As Boolean
    Public gbolSeasonOver As Boolean
    Public gbolHalfSeason As Boolean
    Public gbolPostSeason As Boolean
    Public gbolNewPS As Boolean
    Public gstrSeason As String
    Public gstrPostSeason As String
    Public bolHomeActive As Boolean
    Public maxRosterSize As Integer
    Public twoWayPlayers As Integer


    Public Const conHome As Integer = 0
    Public Const conVisitor As Integer = 1

    'Post Season Game Types
    Public Const conALWC As String = "ALWC"
    Public Const conNLWC As String = "NLWC"
    Public Const conALWC1 As String = "ALWC1"
    Public Const conALWC2 As String = "ALWC2"
    Public Const conNLWC1 As String = "NLWC1"
    Public Const conNLWC2 As String = "NLWC2"
    Public Const conALDIV1 As String = "ALDIV1"
    Public Const conALDIV2 As String = "ALDIV2"
    Public Const conNLDIV1 As String = "NLDIV1"
    Public Const conNLDIV2 As String = "NLDIV2"
    Public Const conALCS As String = "ALCS"
    Public Const conNLCS As String = "NLCS"
    Public Const conWS As String = "WS"

    'Divisions
    Public Const conALEast As String = "AL EAST"
    Public Const conALWest As String = "AL WEST"
    Public Const conALCentral As String = "AL CENTRAL"
    Public Const conNLEast As String = "NL EAST"
    Public Const conNLWest As String = "NL WEST"
    Public Const conNLCentral As String = "NL CENTRAL"
    Public Const conALBill As String = "AL BILL"
    Public Const conALJohn As String = "AL JOHN"
    Public Const conNLBill As String = "NL BILL"
    Public Const conNLJohn As String = "NL JOHN"

    Public Const REGISTRY_LOCAL_MACHINE As String = "HKEY_LOCAL_MACHINE"

    'Game is run in debug mode
    Public bolDebug As Boolean
    Public gblDebugCardID As Integer

    'Game is run in John mode
    Public bolJohn As Boolean

    Public Const conDefOptPlay As Integer = 6

    'Play
    Public bolNormal As Boolean
    Public bolHitRun As Boolean
    Public bolSteal As Boolean
    Public bolSac As Boolean
    Public bolBunt As Boolean
    Public bolSqueeze As Boolean
    Public bolIntentionalWalk As Boolean
    Public bolInfieldIn As Boolean
    Public bolCornersIn As Boolean
    Public bolGuardLine As Boolean
    Public bolPitchAround1 As Boolean
    Public bolPitchAround2 As Boolean
    Public bolDNS As Boolean 'Do not steal
    Public bolStartGame As Boolean
    Public gbolAutoSteal As Boolean

    'FAC Viewer
    Public lstFacView As New List(Of clsFacView)

    'App.Path
    Public gAppPath As String

    Structure FACCard
        Public PB As String
        Public Random As String
        Public CD As String
        Public ErrorNum As String
        Public InfErrPC As String
        Public InfErr1B As String
        Public InfErr As String
        Public OFErr As String
        Public P As String
        Public LN As String
        Public LP As String
        Public SN As String
        Public SP As String
        Public RN As String
        Public RP As String
        Public Pitch As Boolean

        Public Sub New(ByVal pb As String)
            Me.PB = pb
            Me.Random = ""
            Me.CD = ""
            Me.ErrorNum = ""
            Me.InfErrPC = ""
            Me.InfErr1B = ""
            Me.InfErr = ""
            Me.OFErr = ""
            Me.P = ""
            Me.LN = ""
            Me.LP = ""
            Me.SN = ""
            Me.SP = ""
            Me.RN = ""
            Me.RP = ""
            Me.Pitch = False
        End Sub
    End Structure

    Public Structure Standing
        Dim team As String
        Dim Games As Integer
        Dim Wins As Integer
        Dim Losses As Integer
        Dim Last10Wins As Integer
        Dim Last10Losses As Integer
        Dim Streak As String
    End Structure

    'Enum BatCardLabels
    '    Field = 0
    '    OBR = 1
    '    SP = 2
    '    HitRun = 3
    '    CD = 4
    '    Sac = 5
    '    Inj = 6
    '    H1Bf = 7
    '    H1B7 = 8
    '    H1B8 = 9
    '    H1B9 = 10
    '    H2B7 = 11
    '    H2B8 = 12
    '    H2B9 = 13
    '    H3B8 = 14
    '    HR = 15
    '    K = 16
    '    W = 17
    '    HPB = 18
    '    Out = 19
    '    Cht = 20
    '    BD = 21
    'End Enum

    'Enum PitCardLabels
    '    ThrowField = 0
    '    PBRate = 1
    '    SR = 2
    '    RR = 3
    '    H1Bf = 4
    '    H1B7 = 5
    '    H1B8 = 6
    '    H1B9 = 7
    '    BK = 8
    '    K = 9
    '    W = 10
    '    PB = 11
    '    WP = 12
    '    Out = 13
    '    StartRelief = 14
    'End Enum


    'Position strings
    Public gblPositions(9) As String

    'Other constants
    Public Const conNone As String = "None"
    Public Const conNoPlay As String = "NoPlay"
    Public Const conEjected As String = "Ejected"
    Public Const conInjured As String = "Injured"

    'Result flags
    Public bolAnyBS As Boolean 'Any base situation

    'Count run before 3rd out
    Public bolCountRun As Boolean

    'Give credit for RBI's
    Public bolGiveRBI As Boolean

    'flag for batter reaching on error
    Public gbolSOE As Boolean

    'Game over flag
    Public bolFinal As Boolean

    'Used for updating Play message when invoking frmAdvance
    Public gblUpdateMsg As String

    Public Const conIniFile As String = "season.ini"
    Public Const conALHomeRow As Integer = 1
    Public Const conNLHomeRow As Integer = 16

    Public bolInPlay As Boolean
    Public bolThreadActive As Boolean
    Public lastInvoked As DateTime

    'Used for the description log
    Public gblDescBatter As String

    'Table names
    Public gstrHittingTable As String
    Public gstrPitchingTable As String
    Public gstrStandingsTable As String

    'Data Access
    'Public gAccessConnectStr As String

    Public Enum HITTYPES
        TEXAS_LEAGUER
        BLOOP
        NORMAL
        SMASH
    End Enum

    Public Enum Stuff
        GREAT = 2
        GOOD = 1
        NORMAL = 0
        BAD = -1
        TERRIBLE = -2
    End Enum
End Module