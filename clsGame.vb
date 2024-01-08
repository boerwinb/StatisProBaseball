Option Strict On
Option Explicit On

Imports System.Configuration.configurationManager

Friend Class clsGame
	
    Private _outs As Integer
    Private _inning As Integer
    Private _homeTeamBatting As Boolean
    Private _lastBatterHome As Boolean
    Private _currentBatter As Integer
    Private _fieldPos As Integer
    Private _lineScore(99, 2) As Integer
    Private _result As clsResult
    Private _BTeam As clsTeam
    Private _PTeam As clsTeam
    Private _winner As Integer
    Private _loser As Integer
    Private _save As Integer

    Public Property outs() As Integer
        Get
            outs = _outs
        End Get
        Set(ByVal Value As Integer)
            _outs = Value
        End Set
    End Property
    Public Property inning() As Integer
        Get
            inning = _inning
        End Get
        Set(ByVal Value As Integer)
            _inning = Value
        End Set
    End Property
    Public Property homeTeamBatting() As Boolean
        Get
            homeTeamBatting = _homeTeamBatting
        End Get
        Set(ByVal Value As Boolean)
            _homeTeamBatting = Value
        End Set
    End Property
    Public Property lastBatterHome() As Boolean
        Get
            lastBatterHome = _lastBatterHome
        End Get
        Set(ByVal Value As Boolean)
            _lastBatterHome = Value
        End Set
    End Property
    Public Property currentBatter() As Integer
        Get
            currentBatter = _currentBatter
        End Get
        Set(ByVal Value As Integer)
            _currentBatter = Value
        End Set
    End Property
    Public Property winner() As Integer
        Get
            winner = _winner
        End Get
        Set(ByVal Value As Integer)
            _winner = Value
        End Set
    End Property
    Public Property loser() As Integer
        Get
            loser = _loser
        End Get
        Set(ByVal Value As Integer)
            _loser = Value
        End Set
    End Property
    Public Property save() As Integer
        Get
            save = _save
        End Get
        Set(ByVal Value As Integer)
            _save = Value
        End Set
    End Property
    Public Property currentFieldPos() As Integer
        Get
            currentFieldPos = _fieldPos
        End Get
        Set(ByVal Value As Integer)
            _fieldPos = Value
        End Set
    End Property
    Public Function GetResultPtr() As clsResult
        GetResultPtr = _result
    End Function
    Public Function BTeam() As clsTeam
        BTeam = _BTeam
    End Function
    Public Function PTeam() As clsTeam
        PTeam = _PTeam
    End Function

    Public Sub Clear()
        _fieldPos = 1
        _currentBatter = 1
        _outs = 0
        _inning = 0
        _winner = 0
        _loser = 0
        _save = 0
        _homeTeamBatting = False
        _lastBatterHome = False
        For i As Integer = 1 To 99
            _lineScore(i, 1) = 0
            _lineScore(i, 2) = 0
        Next i
        _BTeam.Clear()
        _PTeam.Clear()
        _result.Clear()
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="inning"></param>
    ''' <param name="homeIndicator">homeIndicator = 1 (Visitor) inthome = 2 (Home)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetLineScore(ByVal inning As Integer, ByVal homeIndicator As Integer) As Integer
        Return _lineScore(inning, homeIndicator)
    End Function

    ''' <summary>
    ''' Updates _linescore array
    ''' </summary>
    ''' <param name="totalRuns"></param>
    ''' <remarks></remarks>
    Public Sub UpdateLineScore(ByVal totalRuns As Integer)
        Dim totalRunsPrev As Integer
        For i As Integer = 1 To _inning - 1
            totalRunsPrev = totalRunsPrev + _lineScore(i, IIF(_homeTeamBatting, 2, 1))
        Next i
        _lineScore(_inning, IIF(_homeTeamBatting, 2, 1)) = totalRuns - totalRunsPrev
    End Sub

    Public Sub New()
        MyBase.New()
        Game = Me
        _result = New clsResult()
        _BTeam = New clsTeam()
        _PTeam = New clsTeam()
        Clear()
    End Sub

    ''' <summary>
    ''' Switches team objects between halves of innings
    ''' </summary>
    ''' <param name="isNewGame"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SwitchTeams(ByRef isNewGame As Boolean) As Boolean
        Dim isEndGame As Boolean
        Dim isPOE As Boolean

        Try
            isPOE = AppSettings("PointsOfEffectiveness").ToString.ToUpper = "ON"
            _homeTeamBatting = Not _homeTeamBatting And Not isNewGame
            If _homeTeamBatting Then
                If _inning = 9 And Home.runs > Visitor.runs Then
                    'End game
                    isEndGame = True
                End If
                _BTeam = Home
                _PTeam = Visitor
            Else
                If _inning >= 9 And Home.runs <> Visitor.runs Then
                    'End game
                    isEndGame = True
                Else
                    _inning += 1
                End If
                _BTeam = Visitor
                _PTeam = Home
            End If
            If Not isEndGame Then
                'refresh order
                _currentBatter = _BTeam.GetPlayerNum(_BTeam.order)
            End If
            _outs = 0
            If isPOE Then
                _BTeam.consecutivePOE = 0
                If _PTeam.GetPitcherPtr(_PTeam.pitcherSel).PitStatPtr.ip > 0 Then
                    If _PTeam.GetPitcherPtr(_PTeam.pitcherSel).PitStatPtr.starts = 1 Then
                        _PTeam.pr = _PTeam.GetPitcherPtr(_PTeam.pitcherSel).DeterminePOE(False, False)
                    Else
                        _PTeam.pr = _PTeam.GetPitcherPtr(_PTeam.pitcherSel).DeterminePOE(True, False)
                    End If
                End If
            End If
            _PTeam.GetPitcherPtr(_PTeam.pitcherSel).eOuts = 0
            System.Diagnostics.Debug.WriteLine("Inning: " & _inning)
        Catch ex As Exception
            Call MsgBox("Error in SwitchTeams. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return isendgame
    End Function

    ''' <summary>
    ''' Determines if the game is over
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GameOver() As Boolean
        Return _homeTeamBatting And _BTeam.Runs > _PTeam.Runs And _inning >= 9
    End Function

    ''' <summary>
    ''' Returns player number of current fielder
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetFielderNum() As Integer
        Dim fielderIndex As Integer

        Try
            For i As Integer = 1 To 9
                If _PTeam.GetBatterPtr(_PTeam.GetPlayerNum(i)).position = gblPositions(_fieldPos) Then
                    fielderIndex = _PTeam.GetPlayerNum(i)
                End If
            Next i
        Catch ex As Exception
            Call MsgBox("Error in GetFielderNum. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return fielderIndex
    End Function

    ''' <summary>
    ''' Returns player number of current fielder
    ''' </summary>
    ''' <param name="positionIndex"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetFielderNumByPos(ByVal positionIndex As Integer) As Integer
        Dim fielderIndex As Integer

        Try
            For i As Integer = 1 To 9
                If _PTeam.GetBatterPtr(_PTeam.GetPlayerNum(i)).position = gblPositions(positionIndex) Then
                    fielderIndex = _PTeam.GetPlayerNum(i)
                End If
            Next i
        Catch ex As Exception
            Call MsgBox("Error in GetFielderNumByPos. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return fielderIndex
    End Function

    ''' <summary>
    ''' Updates the current games stats
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub UpdateStats()
        Dim playerIndex As Integer
        Dim dp As String = ""
        Try
            'Hitting
            With _BTeam.GetBatterPtr(_currentBatter).BatStatPtr
                .AB += _result.ab
                .H += _result.hit
                .D += _result.doubl
                .T += _result.triple
                .HR += _result.hr
                If Game.outs < 3 Or bolGiveRBI Then
                    .RBI += _result.rbi
                End If
                .K += _result.strikeOut
                .W += _result.walk
                .HB += _result.hpb
                .Games = 1
                If SecondBase.occupied Or ThirdBase.occupied Then
                    .rispAB += _result.ab
                    .rispH += _result.hit
                End If
                If _result.ip >= 2 And Left(_result.action, 1) = "G" Then
                    .gidp += 1
                End If
                If bolIntentionalWalk Then
                    .ibb += 1
                End If
                If IsFlyBall(_result.action) And _result.rbi = 1 And Game.outs < 3 Then
                    .SF += 1
                End If
                If _result.description.Contains("SAC") And Game.outs < 3 Then
                    .SH += 1
                End If
            End With
            'Pinch Runners
            If FirstBase.occupied Then
                _BTeam.GetBatterPtr(FirstBase.runner).BatStatPtr.Games = 1
            End If
            If SecondBase.occupied Then
                _BTeam.GetBatterPtr(SecondBase.runner).BatStatPtr.Games = 1
            End If
            If ThirdBase.occupied Then
                _BTeam.GetBatterPtr(ThirdBase.runner).BatStatPtr.Games = 1
            End If
            'Fielding
            For i As Integer = 1 To 9
                With _PTeam.GetBatterPtr(_PTeam.GetPlayerNum(i)).BatStatPtr
                    Select Case _PTeam.GetBatterPtr(_PTeam.GetPlayerNum(i)).position
                        Case "C"
                            .GamesC = 1
                        Case "1B"
                            .Games1B = 1
                        Case "2B"
                            .Games2B = 1
                        Case "3B"
                            .Games3B = 1
                        Case "SS"
                            .GamesSS = 1
                        Case "LF", "CF", "RF"
                            .GamesOF = 1
                        Case "P"
                            .GamesP = 1
                        Case "DH"
                            .GamesDH = 1
                        Case Else
                            Call MsgBox("No position available.", MsgBoxStyle.Exclamation)
                    End Select
                    .Games = 1
                End With
            Next i

            'assists, putouts and double plays
            For Each pos As Char In Game.GetResultPtr.po
                If Val(pos) = 1 Then
                    'Pitcher
                    Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).PitStatPtr.po += 1
                Else
                    playerIndex = Game.GetFielderNumByPos(Val(pos))
                    Game.PTeam.GetBatterPtr(playerIndex).BatStatPtr.po += 1
                End If
                If Game.GetResultPtr.ip >= 2 And Not dp.Contains(pos) Then
                    'give dp credit to anyone with a putout during a double play
                    dp += pos
                End If
            Next
            For Each pos As Char In Game.GetResultPtr.a
                If Val(pos) = 1 Then
                    'Pitcher
                    Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).PitStatPtr.a += 1
                Else
                    playerIndex = Game.GetFielderNumByPos(Val(pos))
                    Game.PTeam.GetBatterPtr(playerIndex).BatStatPtr.A += 1
                End If
                If Game.GetResultPtr.ip >= 2 And Not dp.Contains(pos) Then
                    'give dp credit to anyone with an assist during a double play
                    dp += pos
                End If
            Next
            For Each pos As Char In dp
                If Val(pos) = 1 Then
                    'Pitcher
                    Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).PitStatPtr.dp += 1
                Else
                    playerIndex = Game.GetFielderNumByPos(Val(pos))
                    Game.PTeam.GetBatterPtr(playerIndex).BatStatPtr.DP += 1
                End If
            Next


            'Pitching
            With _PTeam.GetPitcherPtr(_PTeam.pitcherSel).PitStatPtr
                .ip = .ip + _result.ip
                .h = .h + _result.hit
                .k = .k + _result.strikeOut
                .w = .w + _result.walk
                .hr = .hr + _result.hr
                .wp = .wp + _result.wp
                .hb = .hb + _result.hpb
                .bk = .bk + _result.bk
                If _BTeam.order = 1 And _BTeam.runs < 2 And _inning = 1 Then
                    'first batter
                    .starts = 1
                End If
                If .starts = 0 Then
                    .reliefs = 1
                End If
            End With
            'Team
            If _result.ip >= 2 Then
                'double play occurred
                _PTeam.dps += 1
            End If
        Catch ex As Exception
            Call MsgBox("Error in UpdateStats. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' Determines which pitchers get credited with a Win, Loss or Save at end of the game
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CreditWinLossSave()
        Dim homeTeamWonGame As Boolean

        Try
            homeTeamWonGame = Home.Runs > Visitor.Runs
            If homeTeamWonGame Then
                Home.GetPitcherPtr(_winner).PitStatPtr.Wins = 1
                If _save > 0 Then
                    Home.GetPitcherPtr(_save).PitStatPtr.Saves = 1
                ElseIf _winner <> Home.PitcherSel And Home.GetPitcherPtr(Home.PitcherSel).PitStatPtr.IP >= 9 And Home.GetPitcherPtr(Home.PitcherSel).PitStatPtr.ER < 2 Then
                    'if the finishing pitcher pitches well for 3 innings and does not win
                    Home.GetPitcherPtr(Home.PitcherSel).PitStatPtr.Saves = 1
                    _save = Home.PitcherSel
                End If
                Visitor.GetPitcherPtr(_loser).PitStatPtr.L = 1
            Else
                Visitor.GetPitcherPtr(_winner).PitStatPtr.wins = 1
                If _save > 0 Then
                    Visitor.GetPitcherPtr(_save).PitStatPtr.saves = 1
                ElseIf _winner <> Visitor.pitcherSel And Visitor.GetPitcherPtr(Visitor.pitcherSel).PitStatPtr.ip >= 9 And Visitor.GetPitcherPtr(Visitor.pitcherSel).PitStatPtr.er < 2 Then
                    'if the finishing pitcher pitches well for 3 innings and does not win
                    Visitor.GetPitcherPtr(Visitor.pitcherSel).PitStatPtr.saves = 1
                    _save = Visitor.pitcherSel
                End If
                Home.GetPitcherPtr(_loser).PitStatPtr.l = 1
            End If
            If Visitor.PitcherSel = Visitor.Starter Then
                'completegame
                With Visitor.GetPitcherPtr(Visitor.Starter).PitStatPtr
                    .cg = 1
                    If Home.runs = 0 Then
                        .sho = 1
                    End If
                    If Home.Hits = 1 Then
                        .OneHit = 1
                    ElseIf Home.Hits = 0 Then
                        .NoHit = 1
                        If Home.Runs = 0 And Home.Errors = 0 And .W = 0 And .HB = 0 Then
                            .Perfect = 1
                        End If
                    End If
                End With
            End If
            If Home.PitcherSel = Home.Starter Then
                'completegame
                With Home.GetPitcherPtr(Home.Starter).PitStatPtr
                    .cg = 1
                    If Visitor.runs = 0 Then
                        .sho = 1
                    End If
                    If Visitor.Hits = 1 Then
                        .OneHit = 1
                    ElseIf Visitor.Hits = 0 Then
                        .NoHit = 1
                        If Visitor.Runs = 0 And Visitor.Errors = 0 And .W = 0 And .HB = 0 Then
                            .Perfect = 1
                        End If
                    End If
                End With
            End If
        Catch ex As Exception
            Call MsgBox("Error in CreditWinLossSave. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub
End Class