Option Strict On
Option Explicit On

Imports System.Text

Friend Class clsSeason
	
    Private _BatStat As clsBatStat
    Private _PitStat As clsPitStat
    Private _seriesRow As String
    Private _game As Integer

    Private Structure PlayoffTeam
        Dim name As String
        Dim division As String
        Dim league As String
        Dim beginSeries As String
        Dim rsWins As Integer
        Dim dsWins As Integer
        Dim csWins As Integer
        Dim wsWins As Integer
    End Structure

    ''' <summary>
    ''' Load teams by reading the *.scd file
    ''' </summary>
    ''' <param name="season"></param>
    ''' <param name="visitingTeam"></param>
    ''' <param name="homeTeam"></param>
    ''' <remarks></remarks>
    Public Sub LoadTeamsFromSched(ByRef season As String, ByRef visitingTeam As String, ByRef homeTeam As String)
        Dim filSched As Integer
        Dim rowNumber As Integer
        Dim tempValue As Integer
        Dim league As String = ""

        Try
            filSched = FreeFile()
            FileOpen(filSched, gAppPath & "schedule\" & season & ".scd", OpenMode.Input)
            '_SeriesRow = CStr(Val(GetIniString(season, "SchedRow", "", gAppPath & conIniFile)))
            '_Game = Val(GetIniString(season, "Game", "", gAppPath & conIniFile))
            GetGame(season, _game, _seriesRow)
            While Not EOF(filSched) And rowNumber <> CDbl(_SeriesRow)
                rowNumber = rowNumber + 1
                Input(filSched, homeTeam)
                Input(filSched, visitingTeam)
                Input(filSched, tempValue)
                Input(filSched, league)
            End While
            FileClose(filSched)
            'BB New PS Code
            If visitingTeam.ToUpper = "POSTSEASON" Or (visitingTeam.ToUpper = "HALFPOST" And gbolHalfSeason) Then
                gbolPostSeason = True
                Call DeterminePSTeams(homeTeam, visitingTeam)
            Else
                bolAmericanLeagueRules = ((league = "AL") And CInt(season) >= 1973) Or CInt(gstrSeason) >= 2022
            End If
            If gbolPostSeason Then
                gstrHittingTable = "PSHITTINGSTATS"
                gstrPitchingTable = "PSPITCHINGSTATS"
                gstrStandingsTable = "PSSTANDINGS"
            Else
                gstrHittingTable = "HITTINGSTATS"
                gstrPitchingTable = "PITCHINGSTATS"
                gstrStandingsTable = "STANDINGS"
            End If
        Catch ex As Exception
            Call MsgBox("Error in LoadTeamsFromSched. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Public Sub SaveGame(ByRef filBox As Integer, ByVal fullPath As String)
        Try
            Print(filBox, Visitor.teamName & Space(26 - Visitor.teamName.Length))
            PrintLine(filBox, "AB H D T HR RBI R SB K W")

            Call UpdateHitting(Visitor, filBox)
            PrintLine(filBox)
            PrintLine(filBox)
            PrintLine(filBox)
            Print(filBox, Home.teamName & Space(26 - Home.teamName.Length))
            PrintLine(filBox, "AB H D T HR RBI R SB K W")
            Call UpdateHitting(Home, filBox)
            PrintLine(filBox)
            PrintLine(filBox)
            PrintLine(filBox)
            PrintLine(filBox, "IP H ER K W")
            'Update Pitching stats
            Call UpdatePitching(Home, filBox)
            Call UpdatePitching(Visitor, filBox)
            'Update Standings
            If gbolPostSeason Then
                Call UpdatePSStandings(Home, Home.runs > Visitor.runs)
                Call UpdatePSStandings(Visitor, Visitor.runs > Home.runs)
            Else
                Call UpdateStandings(Home, Home.runs > Visitor.runs, (Home.dps), (Visitor.dps))
                Call UpdateStandings(Visitor, Visitor.runs > Home.runs, (Visitor.dps), (Home.dps))
                Call UpdateResults(Home, Visitor, True, fullPath)
                Call UpdateResults(Visitor, Home, False, fullPath)

                'Update Schedule
                UpdateSchedule()
            End If
        Catch ex As Exception
            Call MsgBox("Error in SaveGame. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Public Function BatStatPtr() As clsBatStat
        Return _BatStat
    End Function

    Public Function PitStatPtr() As clsPitStat
        Return _PitStat
    End Function

    Public Property seriesRow() As Integer
        Get
            SeriesRow = CInt(_seriesRow)
        End Get
        Set(ByVal Value As Integer)
            _seriesRow = Value.ToString
        End Set
    End Property

    Public Property game() As Integer
        Get
            Game = _game
        End Get
        Set(ByVal Value As Integer)
            _game = Value
        End Set
    End Property

    Public Sub Clear()
        _BatStat.Clear()
        _PitStat.Clear()
        _SeriesRow = "0"
        _Game = 0
    End Sub

    Public Sub New()
        MyBase.New()
        _BatStat = New clsBatStat()
        _PitStat = New clsPitStat()
        Clear()
    End Sub

    ''' <summary>
    ''' Updates the seasonal hitting stats within the Access database
    ''' </summary>
    ''' <param name="team"></param>
    ''' <param name="fileNumber"></param>
    ''' <remarks></remarks>
    Private Sub UpdateHitting(ByRef team As clsTeam, ByRef fileNumber As Integer)
        Dim dsStat As DataSet = Nothing

        Dim drStat As DataRow = Nothing
        Dim sqlQuery As String = ""
        Dim sqlBuilder As StringBuilder
        Dim sqlFields As String = ""
        Dim sqlValues As String = ""
        Dim tableKey As String
        Dim hittingStreak As Integer
        Dim hittingStreakHigh As Integer
        Dim filDebug As Integer
        Dim league As String

        Dim DataAccess As New clsDataAccess(gstrSeason & "data")

        Try
            If Dir(gAppPath & "debug.bbb") <> Nothing Then
                filDebug = FreeFile()
                FileOpen(filDebug, gAppPath & "debug.out", OpenMode.Append)
                PrintLine(filDebug, "In UpdateHitting")
                FileClose(filDebug)
            End If

            'If gAccessConnectStr = Nothing Then
            '    gAccessConnectStr = GetConnectString()
            'End If
            league = GetDivision(team.teamName).Substring(0, 2)
            'Using accessConn As OleDb.OleDbConnection = New OleDb.OleDbConnection
            '    accessConn.ConnectionString = GetConnectString()
            '    accessConn.Open()

            For i As Integer = 1 To conMaxBatters + IIF(bolAmericanLeagueRules, 0, 10)
                With team.GetBatterPtr(i).BatStatPtr
                    If team.GetBatterPtr(i).player <> Nothing Then
                        tableKey = HandleQuotes(StripChar(team.GetBatterPtr(i).player & team.teamName, " "))
                        sqlQuery = "Select * FROM " & gstrHittingTable & " WHERE playerid = '" & tableKey & "'"
                        dsStat = DataAccess.ExecuteDataSet(sqlQuery)
                        'Update the record
                        If .GamesInj > 0 Then
                            .GamesInj = .GamesInj - 1
                        End If
                        If dsStat.Tables(0).Rows.Count = 0 Then
                            'Add a new record
                            dsStat.Clear()
                            sqlFields = "PLAYERID,NAME,TEAM,AB,HITS,DOUBLES,TRIPLES,HOMERUNS,RBI," & _
                                                "RUNS,SB,SBA,ERRORS,STRIKEOUTS,WALKS,HITBYPITCH," & _
                                                "PASSEDBALLS,INJ,GAMESC,GAMESP,GAMES1B,GAMES2B,GAMESSS," & _
                                                "GAMES3B,GAMESOF,GAMESDH,GAMES,PLAYEDLAST,HITLAST,ACTIVE,HS,HSH," & _
                                                "LEAGUE,SH,SF,RISPAB,RISPH,GIDP,IBB,PO,A,DP"
                            sqlValues = "'" & tableKey & "','" & HandleQuotes((team.GetBatterPtr(i).player)) & "','" & _
                                        team.teamName & "'," & .AB & "," & .H & "," & .D & "," & .T & "," & .HR & "," & _
                                        .RBI & "," & .R & "," & .SB & "," & .SBA & "," & .E & "," & .K & "," & .W & _
                                        "," & .HB & "," & .PB & "," & .GamesInj & "," & .GamesC & "," & .GamesP & "," & _
                                        .Games1B & "," & .Games2B & "," & .GamesSS & "," & .Games3B & "," & .GamesOF & _
                                        "," & .GamesDH & "," & .Games & "," & IIF(.Games > 0, 1, 0) & "," & _
                                        IIF(.H > 0, 1, 0) & ",1," & IIF(.Games > 0 And .H > 0, 1, 0) & "," & _
                                        IIF(.Games > 0 And .H > 0, 1, 0) & ",'" & league & "'," & .SH & "," & .SF & _
                                        "," & .rispAB & "," & .rispH & "," & .gidp & "," & .ibb & "," & .po & "," & _
                                        .A & "," & .DP

                            sqlQuery = "INSERT INTO " & gstrHittingTable & " (" & sqlFields & ") VALUES (" & sqlValues & ")"
                            DataAccess.ExecuteScalar(sqlQuery)
                        Else
                            drStat = dsStat.Tables(0).Rows(0)
                            'Determine Hitting Streaks
                            hittingStreak = 0
                            If .Games > 0 And (.AB > 0 Or .SF > 0) Then
                                hittingStreak = IIF(.H > 0, CInt(drStat.Item("hs")) + 1, 0)
                            Else
                                hittingStreak = CInt(drStat.Item("hs"))
                            End If
                            If hittingStreak > CInt(drStat.Item("hsh")) Then
                                hittingStreakHigh = hittingStreak
                            Else
                                hittingStreakHigh = CInt(drStat.Item("hsh"))
                            End If
                            sqlBuilder = New StringBuilder
                            sqlBuilder.Append("UPDATE " & gstrHittingTable)
                            sqlBuilder.Append(" SET TEAM = '" & team.teamName & "', ")
                            sqlBuilder.Append("NAME = '" & HandleQuotes((team.GetBatterPtr(i).player)) & "', ")
                            sqlBuilder.Append("AB = " & (CInt(drStat.Item("AB")) + .AB).ToString & ", ")
                            sqlBuilder.Append("HITS = " & (CInt(drStat.Item("HITS")) + .H).ToString & ", ")
                            sqlBuilder.Append("DOUBLES = " & (CInt(drStat.Item("DOUBLES")) + .D).ToString & ", ")
                            sqlBuilder.Append("TRIPLES = " & (CInt(drStat.Item("TRIPLES")) + .T).ToString & ", ")
                            sqlBuilder.Append("HOMERUNS = " & (CInt(drStat.Item("HOMERUNS")) + .HR).ToString & ", ")
                            sqlBuilder.Append("RBI = " & (CInt(drStat.Item("RBI")) + .RBI).ToString & ", ")
                            sqlBuilder.Append("RUNS = " & (CInt(drStat.Item("RUNS")) + .R).ToString & ", ")
                            sqlBuilder.Append("SB = " & (CInt(drStat.Item("SB")) + .SB).ToString & ", ")
                            sqlBuilder.Append("SBA = " & (CInt(drStat.Item("SBA")) + .SBA).ToString & ", ")
                            sqlBuilder.Append("ERRORS = " & (CInt(drStat.Item("ERRORS")) + .E).ToString & ", ")
                            sqlBuilder.Append("STRIKEOUTS = " & (CInt(drStat.Item("STRIKEOUTS")) + .K).ToString & ", ")
                            sqlBuilder.Append("WALKS = " & (CInt(drStat.Item("WALKS")) + .W).ToString & ", ")
                            sqlBuilder.Append("HITBYPITCH = " & (CInt(drStat.Item("HITBYPITCH")) + .HB).ToString & ", ")
                            sqlBuilder.Append("PASSEDBALLS = " & (CInt(drStat.Item("PASSEDBALLS")) + .PB).ToString & ", ")
                            sqlBuilder.Append("INJ = " & .GamesInj.ToString & ", ")
                            sqlBuilder.Append("GAMESC = " & (CInt(drStat.Item("GAMESC")) + .GamesC).ToString & ", ")
                            sqlBuilder.Append("GAMESP = " & (CInt(drStat.Item("GAMESP")) + .GamesP).ToString & ", ")
                            sqlBuilder.Append("GAMES1B = " & (CInt(drStat.Item("GAMES1B")) + .Games1B).ToString & ", ")
                            sqlBuilder.Append("GAMES2B = " & (CInt(drStat.Item("GAMES2B")) + .Games2B).ToString & ", ")
                            sqlBuilder.Append("GAMESSS = " & (CInt(drStat.Item("GAMESSS")) + .GamesSS).ToString & ", ")
                            sqlBuilder.Append("GAMES3B = " & (CInt(drStat.Item("GAMES3B")) + .Games3B).ToString & ", ")
                            sqlBuilder.Append("GAMESOF = " & (CInt(drStat.Item("GAMESOF")) + .GamesOF).ToString & ", ")
                            sqlBuilder.Append("GAMESDH = " & (CInt(drStat.Item("GAMESDH")) + .GamesDH).ToString & ", ")
                            sqlBuilder.Append("GAMES = " & (CInt(drStat.Item("GAMES")) + .Games).ToString & ", ")
                            sqlBuilder.Append("PLAYEDLAST = " & IIF(.Games > 0 And .W + .HB + .AB >= 3, 1, 0) & ", ")
                            sqlBuilder.Append("HITLAST = " & IIF(.H > 0, 1, 0) & ", ")
                            sqlBuilder.Append("SH = " & (CInt(drStat.Item("SH")) + .SH).ToString & ", ")
                            sqlBuilder.Append("SF = " & (CInt(drStat.Item("SF")) + .SF).ToString & ", ")
                            sqlBuilder.Append("LEAGUE = '" & league & "', ")
                            sqlBuilder.Append("HS = " & hittingStreak.ToString & ", ")
                            sqlBuilder.Append("HSH = " & hittingStreakHigh.ToString & ", ")
                            sqlBuilder.Append("RISPAB = " & (CInt(drStat.Item("RISPAB")) + .rispAB).ToString & ", ")
                            sqlBuilder.Append("RISPH = " & (CInt(drStat.Item("RISPH")) + .rispH).ToString & ", ")
                            sqlBuilder.Append("GIDP = " & (CInt(drStat.Item("GIDP")) + .gidp).ToString & ", ")
                            sqlBuilder.Append("IBB = " & (CInt(drStat.Item("IBB")) + .ibb).ToString & ", ")
                            sqlBuilder.Append("PO = " & (CInt(drStat.Item("PO")) + .po).ToString & ", ")
                            sqlBuilder.Append("A = " & (CInt(drStat.Item("A")) + .A).ToString & ", ")
                            sqlBuilder.Append("DP = " & (CInt(drStat.Item("DP")) + .DP).ToString & " ")
                            sqlBuilder.Append("WHERE playerid = '" & tableKey & "'")

                            dsStat.Clear()
                            sqlQuery = sqlBuilder.ToString
                            DataAccess.ExecuteScalar(sqlQuery)
                        End If
                        If .Games > 0 Then
                            Print(fileNumber, team.GetBatterPtr(i).player & Space(27 - team.GetBatterPtr(i).player.Length))
                            Print(fileNumber, .AB & Space(1) & .H & Space(1) & .D & Space(1) & .T & Space(2) & .HR & Space(3))
                            PrintLine(fileNumber, .RBI & Space(1) & .R & Space(2) & .SB & Space(1) & .K & Space(1) & .W)
                        End If
                        '"AB H D T HR RBI R SB K W"
                    End If
                End With
            Next i
            'End Using
        Catch ex As Exception
            Call MsgBox("Error in UpdateHitting. " & ex.ToString, MsgBoxStyle.OkOnly)
        Finally
            If Not dsStat Is Nothing Then dsStat.Dispose()
        End Try
    End Sub

    ''' <summary>
    ''' Updates seasonal pitching stats within an Access database
    ''' </summary>
    ''' <param name="team"></param>
    ''' <param name="filBox"></param>
    ''' <remarks></remarks>
    Private Sub UpdatePitching(ByRef team As clsTeam, ByRef filBox As Integer)
        Dim tableKey As String
        Dim dsStat As DataSet = Nothing
        Dim drStat As DataRow = Nothing
        Dim sqlQuery As String = ""
        Dim sqlBuilder As StringBuilder
        Dim sqlFields As String
        Dim sqlValues As String
        Dim filDebug As Integer
        Dim league As String
        Dim DataAccess As New clsDataAccess(gstrSeason & "data")

        Try
            If Dir(gAppPath & "debug.bbb") <> Nothing Then
                filDebug = FreeFile()
                FileOpen(filDebug, gAppPath & "debug.out", OpenMode.Append)
                PrintLine(filDebug, "In UpdatePitching")
                FileClose(filDebug)
            End If

            league = GetDivision(team.teamName).Substring(0, 2)
            'If gAccessConnectStr = Nothing Then
            '    gAccessConnectStr = GetConnectString()
            'End If
            'Using accessConn As OleDb.OleDbConnection = New OleDb.OleDbConnection
            '    accessConn.ConnectionString = GetConnectString()
            '    accessConn.Open()
            For i As Integer = 1 To team.pitchers
                With team.GetPitcherPtr(i).PitStatPtr
                    If team.GetPitcherPtr(i).player <> Nothing Then
                        tableKey = HandleQuotes(StripChar(team.GetPitcherPtr(i).player & team.teamName, " "))
                        sqlQuery = "Select * FROM " & gstrPitchingTable & " WHERE playerid = '" & tableKey & "'"
                        dsStat = DataAccess.ExecuteDataSet(sqlQuery)

                        If dsStat.Tables(0).Rows.Count = 0 Then
                            dsStat.Clear()
                            sqlFields = "PLAYERID,NAME,TEAM,IP,HITS,RUNS,EARNEDRUNS,WINS,LOSSES," & _
                                        "STRIKEOUTS,WALKS,SAVES,HOMERUNS,WP,HB,BK,ERRORS,STARTS," & _
                                        "RELIEFS,CG,SHUTOUTS,ONEHITTERS,NOHITTERS,PERFECT,INACTIVE,ACTIVE,INJ," & _
                                        "LEAGUE,LG1,LG2,LG3"
                            sqlValues = "'" & tableKey & "','" & HandleQuotes(team.GetPitcherPtr(i).player) & "','" & _
                                        team.teamName & "'," & .ip & "," & .h & "," & .r & "," & .er & "," & .wins & "," & .l & _
                                        "," & .k & "," & .w & "," & .saves & "," & .hr & "," & .wp & "," & .hb & "," & .bk & _
                                        "," & .e & "," & .starts & "," & .reliefs & "," & .cg & "," & .sho & "," & .oneHit & _
                                        "," & .noHit & "," & .perfect & "," & IIF(.ip / 3 > 2, 1, 0) & ",1," & .gamesInj & ",'" & _
                                        league & "'," & .ip & ",0,0"
                            sqlQuery = "INSERT INTO " & gstrPitchingTable & " (" & sqlFields & ") VALUES (" & sqlValues & ")"
                            DataAccess.ExecuteScalar(sqlQuery)
                        Else
                            drStat = dsStat.Tables(0).Rows(0)
                            'Update the record
                            If .gamesInj > 0 Then
                                .gamesInj = .gamesInj - 1
                            End If
                            sqlBuilder = New StringBuilder
                            sqlBuilder.Append("UPDATE " & gstrPitchingTable & " SET TEAM = '" & team.teamName & "', ")
                            sqlBuilder.Append("NAME = '" & HandleQuotes(team.GetPitcherPtr(i).player) & "', ")
                            sqlBuilder.Append("IP = " & (CInt(drStat.Item("IP")) + .ip).ToString & ", ")
                            sqlBuilder.Append("HITS = " & (CInt(drStat.Item("HITS")) + .h).ToString & ", ")
                            sqlBuilder.Append("RUNS = " & (CInt(drStat.Item("RUNS")) + .r).ToString & ", ")
                            sqlBuilder.Append("EARNEDRUNS = " & (CInt(drStat.Item("EARNEDRUNS")) + .er).ToString & ", ")
                            sqlBuilder.Append("WINS = " & (CInt(drStat.Item("WINS")) + .wins).ToString & ", ")
                            sqlBuilder.Append("LOSSES = " & (CInt(drStat.Item("LOSSES")) + .l).ToString & ", ")
                            sqlBuilder.Append("STRIKEOUTS = " & (CInt(drStat.Item("STRIKEOUTS")) + .k).ToString & ", ")
                            sqlBuilder.Append("WALKS = " & (CInt(drStat.Item("WALKS")) + .w).ToString & ", ")
                            sqlBuilder.Append("SAVES = " & (CInt(drStat.Item("SAVES")) + .saves).ToString & ", ")
                            sqlBuilder.Append("HOMERUNS = " & (CInt(drStat.Item("HOMERUNS")) + .hr).ToString & ", ")
                            sqlBuilder.Append("WP = " & (CInt(drStat.Item("WP")) + .wp).ToString & ", ")
                            sqlBuilder.Append("HB = " & (CInt(drStat.Item("HB")) + .hb).ToString & ", ")
                            sqlBuilder.Append("BK = " & (CInt(drStat.Item("BK")) + .bk).ToString & ", ")
                            sqlBuilder.Append("ERRORS = " & (CInt(drStat.Item("ERRORS")) + .e).ToString & ", ")
                            sqlBuilder.Append("STARTS = " & (CInt(drStat.Item("STARTS")) + .starts).ToString & ", ")
                            sqlBuilder.Append("RELIEFS = " & (CInt(drStat.Item("RELIEFS")) + .reliefs).ToString & ", ")
                            sqlBuilder.Append("CG = " & (CInt(drStat.Item("CG")) + .cg).ToString & ", ")
                            sqlBuilder.Append("SHUTOUTS = " & (CInt(drStat.Item("SHUTOUTS")) + .sho).ToString & ", ")
                            sqlBuilder.Append("ONEHITTERS = " & (CInt(drStat.Item("ONEHITTERS")) + .oneHit).ToString & ", ")
                            sqlBuilder.Append("NOHITTERS = " & (CInt(drStat.Item("NOHITTERS")) + .noHit).ToString & ", ")
                            sqlBuilder.Append("PERFECT = " & (CInt(drStat.Item("PERFECT")) + .perfect).ToString & ", ")
                            sqlBuilder.Append("LEAGUE = '" & league & "', ")
                            sqlBuilder.Append("INJ = " & .gamesInj & ", ")
                            'sqlBuilder.Append("INACTIVE = " & IIF(.ip / 3 > 2, 1, 0) & ", ")
                            sqlBuilder.Append("INACTIVE = " & IIF(DeterminePitcherInactiveNextGame(.ip, CInt(drStat.Item("LG1")) _
                                                , CInt(drStat.Item("LG2")), CInt(drStat.Item("LG3"))), 1, 0) & ", ")
                            sqlBuilder.Append("LG1 = " & .ip & ", ")
                            sqlBuilder.Append("LG2 = " & CInt(drStat.Item("LG1")) & ", ")
                            sqlBuilder.Append("LG3 = " & CInt(drStat.Item("LG2")) & " ")
                            sqlBuilder.Append("WHERE playerid = '" & tableKey & "'")

                            dsStat.Clear()
                            sqlQuery = sqlBuilder.ToString
                            DataAccess.ExecuteScalar(sqlQuery)
                        End If
                        If .starts > 0 Or .reliefs > 0 Then
                            'IP H ER K W
                            Print(filBox, team.GetPitcherPtr(i).player & Space(27 - team.GetPitcherPtr(i).player.Length))
                            PrintLine(filBox, (.ip / 3).ToString("##.0") & Space(1) & .h & Space(2) & .er & Space(1) & .k & _
                                            Space(1) & .w)
                        End If
                    End If
                End With
            Next i
            'End Using
        Catch ex As Exception
            Call MsgBox("Error in UpdatePitching. sql=" & sqlQuery & " " & ex.ToString, MsgBoxStyle.OkOnly)
        Finally
            If Not dsStat Is Nothing Then dsStat.Dispose()
        End Try
    End Sub

    ''' <summary>
    ''' Determines whether a pitcher will be inactive in the next contest based on recent usage
    ''' ip - Outs in current game
    ''' ipLG1 - Outs in previous game
    ''' ipLG2 - Outs from 2 games ago
    ''' ipLG3 - Outs from 3 games ago
    ''' </summary>
    ''' <remarks></remarks>
    Private Function DeterminePitcherInactiveNextGame(ip As Integer, ipLG1 As Integer, ipLG2 As Integer, ipLG3 As Integer) As Boolean
        If ip > 0 And ipLG1 > 0 And ipLG2 > 0 Then
            'Pitcher pitched in each of the last 3 games including the current one
            Return True
        ElseIf ipLG1 > 0 And ipLG2 > 0 And ipLG3 > 0 And ((ipLG1 + ipLG2 + ipLG3) / 3 > 5) Then
            'Pitcher previously pitched in 3 straight games that exceeded 5 innings total. Need to miss a 2nd game
            Return True
        ElseIf ip > 0 And (ip + ipLG1) / 3 = 4 Then
            'Pitcher pitched 2 innings in each of the last 2 games including the current one
            Return True
        ElseIf ip / 3 > 2 Then
            'Pitcher pitched more than 2 innings in the current game. Original rule.
            Return True
        ElseIf ipLG1 / 3 > 6 Then
            'Pitcher pitched in more than 6 innings in previous game. Needs to miss a 2nd game
            Return True
        ElseIf ipLG2 / 3 > 8 Then
            'Pitcher pitched in more than 8 innings 2 games ago. Needs to miss a 3rd game.
            Return True
        Else
            'Pitcher is available for the next game
            Return False
        End If
    End Function

    ''' <summary>
    ''' Updates team's seasonal results in an access database
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub UpdateResults(ByRef team As clsTeam, ByVal opponent As clsTeam, ByVal isHome As Boolean, _
                                ByVal boxScore As String)
        Dim tableKey As String
        Dim sqlFields As String
        Dim sqlValues As String
        Dim sqlBuilder As StringBuilder
        Dim sqlQuery As String
        Dim dsStat As DataSet = Nothing
        Dim DataAccess As New clsDataAccess(gstrSeason & "data")

        Try
            tableKey = HandleQuotes((team.teamName))

            'If gAccessConnectStr = Nothing Then
            '    gAccessConnectStr = GetConnectString()
            'End If
            'Using accessConn As OleDb.OleDbConnection = New OleDb.OleDbConnection
            '    accessConn.ConnectionString = GetConnectString()
            '    accessConn.Open()
            sqlQuery = "Select * FROM Standings " & "WHERE playerid = '" & tableKey & "'"
            dsStat = DataAccess.ExecuteDataSet(sqlQuery)

            With dsStat.Tables(0).Rows(0)
                'Add a new record
                sqlFields = "TEAM,GAME,OPPONENT,HOME,SCORE,WINLOSS,RECORD,BOXSCORE"
                sqlBuilder = New StringBuilder
                sqlBuilder.Append("'" & tableKey & "',")
                sqlBuilder.Append(CInt(.Item("Wins")) + CInt(.Item("Losses")) & ",'")
                sqlBuilder.Append(HandleQuotes(opponent.teamName) & "',")
                sqlBuilder.Append(IIF(isHome, 1, 0) & ",'")
                sqlBuilder.Append(team.runs.ToString & "-" & opponent.runs.ToString & "','")
                sqlBuilder.Append(IIF(team.runs > opponent.runs, "W", "L") & "','")
                sqlBuilder.Append(.Item("Wins").ToString & "-" & .Item("Losses").ToString & "','")
                sqlBuilder.Append(boxScore & "'")
                sqlValues = sqlBuilder.ToString

                sqlQuery = "INSERT INTO RESULTS (" & sqlFields & ") VALUES (" & sqlValues & ")"
                DataAccess.ExecuteScalar(sqlQuery)
            End With
            'End Using
        Catch ex As Exception
            Call MsgBox("Error in UpdateResults. " & ex.ToString, MsgBoxStyle.OkOnly)
        Finally
            If Not dsStat Is Nothing Then dsStat.Dispose()
        End Try
    End Sub
    ''' <summary>
    ''' Updates standings data with an Access database
    ''' </summary>
    ''' <param name="team"></param>
    ''' <param name="teamWon"></param>
    ''' <param name="dpsFor"></param>
    ''' <param name="dpsAgainst"></param>
    ''' <remarks></remarks>
    Private Sub UpdateStandings(ByRef team As clsTeam, ByRef teamWon As Boolean, ByRef dpsFor As Integer, _
                                ByRef dpsAgainst As Integer)
        Dim tableKey As String
        Dim sqlFields As String
        Dim sqlValues As String
        Dim sqlBuilder As StringBuilder
        Dim sqlQuery As String
        Dim streak As String
        Dim dsStat As DataSet = Nothing
        Dim filDebug As Integer
        Dim DataAccess As New clsDataAccess(gstrSeason & "data")

        Try
            If Dir(gAppPath & "debug.bbb") <> Nothing Then
                filDebug = FreeFile()
                FileOpen(filDebug, gAppPath & "debug.out", OpenMode.Append)
                PrintLine(filDebug, "In UpdateStandings")
                FileClose(filDebug)
            End If

            tableKey = HandleQuotes((team.teamName))

            'If gAccessConnectStr = Nothing Then
            '    gAccessConnectStr = GetConnectString()
            'End If
            'Using accessConn As OleDb.OleDbConnection = New OleDb.OleDbConnection
            '    accessConn.ConnectionString = GetConnectString()
            '    accessConn.Open()
            sqlQuery = "Select * FROM Standings " & "WHERE playerid = '" & tableKey & "'"
            dsStat = DataAccess.ExecuteDataSet(sqlQuery)
            If dsStat.Tables(0).Rows.Count = 0 Then
                dsStat.Clear()
                'Add a new record
                sqlFields = "PLAYERID,WINS,LOSSES,DIVISION,LAST1,STREAK,DP,DPA,LAST2,LAST3," & _
                            "LAST4,LAST5,LAST6,LAST7,LAST8,LAST9,LAST10"

                sqlBuilder = New StringBuilder
                sqlBuilder.Append("'" & tableKey & "',")
                sqlBuilder.Append(IIF(teamWon, 1, 0) & ",")
                sqlBuilder.Append(IIF(teamWon, 0, 1) & ",'")
                sqlBuilder.Append(GetDivision(tableKey) & "','")
                sqlBuilder.Append(IIF(teamWon, "W", "L") & "','")
                sqlBuilder.Append(IIF(teamWon, "W1", "L1") & "',")
                sqlBuilder.Append(dpsFor.ToString & ",")
                sqlBuilder.Append(dpsAgainst.ToString)
                sqlBuilder.Append(",'X','X','X','X','X','X','X','X','X'")
                sqlValues = sqlBuilder.ToString

                sqlQuery = "INSERT INTO STANDINGS (" & sqlFields & ") VALUES (" & sqlValues & ")"
                DataAccess.ExecuteScalar(sqlQuery)
            Else
                With dsStat.Tables(0).Rows(0)
                    streak = .Item("streak").ToString
                    sqlBuilder = New StringBuilder
                    sqlBuilder.Append("UPDATE STANDINGS ")
                    If teamWon Then
                        If streak.Substring(0, 1) = "W" Then
                            streak = "W" & (Val(streak.Substring(1)) + 1).ToString
                        Else
                            streak = "W1"
                        End If
                        sqlBuilder.Append("SET WINS = " & (CInt(.Item("Wins")) + 1).ToString & ", ")
                        sqlBuilder.Append("LAST1 = 'W', ")
                    Else
                        If streak.Substring(0, 1) = "W" Then
                            streak = "L1"
                        Else
                            streak = "L" & (Val(streak.Substring(1)) + 1).ToString
                        End If
                        sqlBuilder.Append("SET LOSSES = " & (CInt(.Item("Losses")) + 1).ToString & ", ")
                        sqlBuilder.Append("LAST1 = 'L', ")
                    End If
                    sqlBuilder.Append("STREAK = '" & streak & "', ")
                    sqlBuilder.Append("DP = " & (CInt(.Item("dp")) + dpsFor).ToString & ", ")
                    sqlBuilder.Append("DPA = " & (CInt(.Item("dpa")) + dpsAgainst).ToString & ", ")
                    sqlBuilder.Append("LAST10 = '" & .Item("last9").ToString & "', ")
                    sqlBuilder.Append("LAST9 = '" & .Item("last8").ToString & "', ")
                    sqlBuilder.Append("LAST8 = '" & .Item("last7").ToString & "', ")
                    sqlBuilder.Append("LAST7 = '" & .Item("last6").ToString & "', ")
                    sqlBuilder.Append("LAST6 = '" & .Item("last5").ToString & "', ")
                    sqlBuilder.Append("LAST5 = '" & .Item("last4").ToString & "', ")
                    sqlBuilder.Append("LAST4 = '" & .Item("last3").ToString & "', ")
                    sqlBuilder.Append("LAST3 = '" & .Item("last2").ToString & "', ")
                    sqlBuilder.Append("LAST2 = '" & .Item("last1").ToString & "' ")
                    sqlBuilder.Append("WHERE playerid = '" & tableKey & "'")
                End With
                sqlQuery = sqlBuilder.ToString
                DataAccess.ExecuteScalar(sqlQuery)
            End If
            'End Using
        Catch ex As Exception
            Call MsgBox("Error in UpdateStandings. " & ex.ToString, MsgBoxStyle.OkOnly)
        Finally
            If Not dsStat Is Nothing Then dsStat.Dispose()
        End Try
    End Sub

    ''' <summary>
    ''' Updates post-season standings data with an Access database
    ''' </summary>
    ''' <param name="team"></param>
    ''' <param name="teamWon"></param>
    ''' <remarks></remarks>
    Private Sub UpdatePSStandings(ByRef team As clsTeam, ByRef teamWon As Boolean)
        Dim tableKey As String
        Dim sqlFields As String = ""
        Dim sqlValues As String = ""
        Dim sqlQuery As String = ""
        Dim dsStat As DataSet = Nothing
        Dim filDebug As Integer
        Dim DataAccess As New clsDataAccess(gstrSeason & "data")

        If Dir(gAppPath & "debug.bbb") <> Nothing Then
            filDebug = FreeFile()
            FileOpen(filDebug, gAppPath & "debug.out", OpenMode.Append)
            PrintLine(filDebug, "In UpdatePSStandings")
            FileClose(filDebug)
        End If

        Try
            If Not teamWon Then Exit Sub
            tableKey = team.teamName

            'If gAccessConnectStr = Nothing Then
            '    gAccessConnectStr = GetConnectString()
            'End If
            'Using accessConn As OleDb.OleDbConnection = New OleDb.OleDbConnection
            '    accessConn.ConnectionString = GetConnectString()
            '    accessConn.Open()
            sqlQuery = "Select * FROM PSStandings " & "WHERE teamid = '" & tableKey & "'"
            dsStat = DataAccess.ExecuteDataSet(sqlQuery)
            If dsStat.Tables(0).Rows.Count = 0 Then
                dsStat.Clear()
                'Add a new record
                Select Case gstrPostSeason
                    Case conALDIV1, conALDIV2, conNLDIV1, conNLDIV2
                        sqlFields = "TEAMID,DSWINS,LEAGUE,DIVISION"
                    Case conALCS, conNLCS
                        sqlFields = "TEAMID,CSWINS,LEAGUE,DIVISION"
                    Case conWS
                        sqlFields = "TEAMID,WSWINS,LEAGUE,DIVISION"
                End Select
                sqlValues = "'" & tableKey & "'," & IIF(teamWon, 1, 0) & ",'" & GetDivision(tableKey).Substring(0, 2) & "','" & GetDivision(tableKey) & "'"
                sqlQuery = "INSERT INTO PSSTANDINGS (" & sqlFields & ") VALUES (" & sqlValues & ")"
            Else
                With dsStat.Tables(0).Rows(0)
                    Select Case gstrPostSeason
                        Case conALWC
                            sqlQuery = "UPDATE PSSTANDINGS SET beginseries = '" & _
                                    conALDIV1 & "' " & "WHERE teamid = '" & tableKey & "'"
                        Case conNLWC
                            sqlQuery = "UPDATE PSSTANDINGS SET beginseries = '" & _
                                    conNLDIV1 & "' " & "WHERE teamid = '" & tableKey & "'"
                        Case conALDIV1, conALDIV2, conNLDIV1, conNLDIV2
                            sqlQuery = "UPDATE PSSTANDINGS SET dswins = " & _
                                    (CInt(.Item("dswins")) + 1).ToString & " " & "WHERE teamid = '" & tableKey & "'"
                        Case conALCS, conNLCS
                            sqlQuery = "UPDATE PSSTANDINGS SET cswins = " & _
                                    (CInt(.Item("cswins")) + 1).ToString & " " & "WHERE teamid = '" & tableKey & "'"
                        Case conWS
                            sqlQuery = "UPDATE PSSTANDINGS SET wswins = " & _
                                    (CInt(.Item("wswins")) + 1).ToString & " " & "WHERE teamid = '" & tableKey & "'"
                    End Select
                End With
            End If
            DataAccess.ExecuteScalar(sqlQuery)
            'End Using
        Catch ex As Exception
            Call MsgBox("Error in UpdatePSStandings. Team = " & team.teamName & " " & ex.ToString, MsgBoxStyle.OkOnly)
        Finally
            If Not dsStat Is Nothing Then dsStat.Dispose()
        End Try
    End Sub

    ''' <summary>
    ''' Updates the current schedule within the registry
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub UpdateSchedule()
        Dim filSched As Integer
        Dim seriesSize As Integer
        Dim rowNumber As Integer
        Dim tempValue As String = ""

        Try
            filSched = FreeFile()
            FileOpen(filSched, gAppPath & "schedule\" & gstrSeason & ".scd", OpenMode.Input)
            While Not EOF(filSched) And rowNumber <> CDbl(_seriesRow)
                rowNumber = rowNumber + 1
                Input(filSched, tempValue)
                Input(filSched, tempValue)
                Input(filSched, seriesSize)
                Input(filSched, tempValue)
            End While
            If _game + 1 <= seriesSize Then
                'increment game, otherwise move to next series
                _game = _game + 1
            Else
                _seriesRow = (Convert.ToInt32(_seriesRow) + 1).ToString
                _game = 1
            End If
            FileClose(filSched)
            SetGame(gstrSeason, _game.ToString, _seriesRow)

        Catch ex As Exception
            Call MsgBox("Error in UpdateSchedule " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' Determines which teams are to participate in the postseason
    ''' </summary>
    ''' <param name="homeTeam">out parameter</param>
    ''' <param name="visitingTeam">out parameter</param>
    ''' <remarks></remarks>
    Private Sub DeterminePSTeams(ByRef homeTeam As String, ByRef visitingTeam As String)
        Dim dsStandings As DataSet = Nothing
        Dim sqlFields As String
        Dim sqlValues As String
        Dim playoffTeam(12) As clsSeason.PlayoffTeam
        Dim sqlQuery As String = ""
        Dim tempValue As Integer
        Dim tempValue2 As Integer
        Dim gamesPlayed As Integer
        Dim isALWestHome As Boolean
        Dim isALHome As Boolean
        Dim numberOfTeams As Integer
        Dim filDebug As Integer
        Dim csWinsNeeded As Integer = 0
        Dim teamIndex As Integer
        Dim DataAccess As New clsDataAccess(gstrSeason & "data")

        If Dir(gAppPath & "debug.bbb") <> Nothing Then
            filDebug = FreeFile()
            FileOpen(filDebug, gAppPath & "debug.out", OpenMode.Append)
            PrintLine(filDebug, "In DeterminePSTeams")
            FileClose(filDebug)
        End If

        Try
            gbolNewPS = False

            'If gAccessConnectStr = Nothing Then
            '    gAccessConnectStr = GetConnectString()
            'End If
            'Using accessConn As OleDb.OleDbConnection = New OleDb.OleDbConnection
            '    accessConn.ConnectionString = GetConnectString()
            '    accessConn.Open()
            Select Case Val(gstrSeason)
                Case 1969 To 1993
                    If Val(gstrSeason) >= 1985 Then
                        'CS is 7 games and WS is 7 games
                        csWinsNeeded = 4
                    Else
                        'CS is 5 games and WS is 7 games
                        csWinsNeeded = 3
                    End If
                    numberOfTeams = 4

                    isALWestHome = Val(gstrSeason) Mod 2 > 0
                    sqlQuery = "Select Division, League, CSWins, TeamID " & "FROM PSStandings"
                    dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                    If dsStandings.Tables(0).Rows.Count = 0 Then
                        'No record, new playoffs
                        gbolNewPS = True
                        sqlQuery = "Select Wins, PlayerID, Division From Standings WHERE " & _
                                    "Division = 'AL EAST' ORDER BY WINS DESC LIMIT 1"
                        dsStandings.Clear()
                        dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                        With dsStandings.Tables(0).Rows(0)
                            playoffTeam(1).name = .Item("PlayerID").ToString
                            playoffTeam(1).division = .Item("Division").ToString
                            playoffTeam(1).rsWins = CInt(.Item("Wins"))
                        End With
                        sqlQuery = "Select Wins, PlayerID, Division From Standings WHERE " & _
                                    "Division = 'AL WEST' ORDER BY WINS DESC LIMIT 1"
                        dsStandings.Clear()
                        dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                        With dsStandings.Tables(0).Rows(0)
                            playoffTeam(2).name = .Item("PlayerID").ToString
                            playoffTeam(2).division = .Item("Division").ToString
                            playoffTeam(2).rsWins = CInt(.Item("Wins"))
                        End With
                        sqlQuery = "Select Wins, PlayerID, Division, Wins From Standings WHERE " & _
                                    "Division = 'NL EAST' ORDER BY WINS DESC LIMIT 1"
                        dsStandings.Clear()
                        dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                        With dsStandings.Tables(0).Rows(0)
                            playoffTeam(3).name = .Item("PlayerID").ToString
                            playoffTeam(3).division = .Item("Division").ToString
                            playoffTeam(3).rsWins = CInt(.Item("Wins"))
                        End With
                        sqlQuery = "Select Wins, PlayerID, Division, Wins From Standings WHERE " & _
                                    "Division = 'NL WEST' ORDER BY WINS DESC LIMIT 1"
                        dsStandings.Clear()
                        dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                        With dsStandings.Tables(0).Rows(0)
                            playoffTeam(4).name = .Item("PlayerID").ToString
                            playoffTeam(4).division = .Item("Division").ToString
                            playoffTeam(4).rsWins = CInt(.Item("Wins"))
                        End With
                        dsStandings.Clear()
                        If isALWestHome Then
                            homeTeam = playoffTeam(1).name
                            visitingTeam = playoffTeam(2).name
                        Else
                            homeTeam = playoffTeam(2).name
                            visitingTeam = playoffTeam(1).name
                        End If
                        For i As Integer = 1 To 4
                            sqlFields = "TEAMID,RSWINS,DSWINS,CSWINS,WSWINS,DIVISION,LEAGUE"
                            sqlValues = "'" & playoffTeam(i).name & "'," & playoffTeam(i).rsWins & ",0,0,0,'" & playoffTeam(i).division & "','" & IIF(i < 3, "AL", "NL") & "'"
                            sqlQuery = "INSERT INTO PSSTANDINGS (" & sqlFields & ") VALUES (" & sqlValues & ")"
                            DataAccess.ExecuteScalar(sqlQuery)
                        Next i
                        gstrPostSeason = conALCS
                        bolAmericanLeagueRules = True
                        _game = 1
                    Else
                        sqlQuery = "Select * FROM PSStandings WHERE League = 'AL' " & _
                                    "AND CSWins = " & csWinsNeeded.ToString
                        dsStandings.Clear()
                        dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                        If dsStandings.Tables(0).Rows.Count = 0 Then
                            'In middle of ALCS
                            sqlQuery = "Select * FROM PSStandings WHERE League = 'AL'"
                            dsStandings.Clear()
                            dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                            teamIndex = 1
                            For Each dr As DataRow In dsStandings.Tables(0).Rows
                                playoffTeam(teamIndex).name = dr.Item("TeamID").ToString
                                playoffTeam(teamIndex).division = dr.Item("Division").ToString
                                gamesPlayed += CInt(dr.Item("CSWins"))
                                teamIndex += 1
                            Next dr
                            If gamesPlayed = 0 Then
                                gbolNewPS = True
                            End If
                            dsStandings.Clear()
                            Select Case gamesPlayed
                                Case 0, 1, 5, 6
                                    If isALWestHome Then
                                        If playoffTeam(1).division = "AL EAST" Then
                                            homeTeam = playoffTeam(1).name
                                            visitingTeam = playoffTeam(2).name
                                        Else
                                            homeTeam = playoffTeam(2).name
                                            visitingTeam = playoffTeam(1).name
                                        End If
                                    Else
                                        If playoffTeam(1).division = "AL EAST" Then
                                            homeTeam = playoffTeam(2).name
                                            visitingTeam = playoffTeam(1).name
                                        Else
                                            homeTeam = playoffTeam(1).name
                                            visitingTeam = playoffTeam(2).name
                                        End If
                                    End If
                                Case 2, 3, 4
                                    If isALWestHome Then
                                        If playoffTeam(1).division = "AL EAST" Then
                                            homeTeam = playoffTeam(2).name
                                            visitingTeam = playoffTeam(1).name
                                        Else
                                            homeTeam = playoffTeam(1).name
                                            visitingTeam = playoffTeam(2).name
                                        End If
                                    Else
                                        If playoffTeam(1).division = "AL EAST" Then
                                            homeTeam = playoffTeam(1).name
                                            visitingTeam = playoffTeam(2).name
                                        Else
                                            homeTeam = playoffTeam(2).name
                                            visitingTeam = playoffTeam(1).name
                                        End If
                                    End If
                            End Select
                            gstrPostSeason = conALCS
                            bolAmericanLeagueRules = True
                            _game = gamesPlayed + 1
                        Else
                            sqlQuery = "Select * FROM PSStandings WHERE League = 'NL'" & "AND CSWins = " & csWinsNeeded.ToString
                            dsStandings.Clear()
                            dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                            If dsStandings.Tables(0).Rows.Count = 0 Then
                                'In middle of NLCS
                                sqlQuery = "Select * FROM PSStandings WHERE League = 'NL'"
                                dsStandings.Clear()
                                dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                teamIndex = 3
                                For Each dr As DataRow In dsStandings.Tables(0).Rows
                                    playoffTeam(teamIndex).name = dr.Item("TeamID").ToString
                                    playoffTeam(teamIndex).division = dr.Item("Division").ToString
                                    gamesPlayed = gamesPlayed + CInt(dr.Item("CSWins"))
                                    teamIndex += 1
                                Next dr
                                If gamesPlayed = 0 Then
                                    gbolNewPS = True
                                End If
                                dsStandings.Clear()
                                Select Case gamesPlayed
                                    Case 0, 1, 5, 6
                                        If isALWestHome Then
                                            If playoffTeam(3).division = "NL EAST" Then
                                                homeTeam = playoffTeam(3).name
                                                visitingTeam = playoffTeam(4).name
                                            Else
                                                homeTeam = playoffTeam(4).name
                                                visitingTeam = playoffTeam(3).name
                                            End If
                                        Else
                                            If playoffTeam(3).division = "NL EAST" Then
                                                homeTeam = playoffTeam(4).name
                                                visitingTeam = playoffTeam(3).name
                                            Else
                                                homeTeam = playoffTeam(3).name
                                                visitingTeam = playoffTeam(4).name
                                            End If
                                        End If
                                    Case 2, 3, 4
                                        If isALWestHome Then
                                            If playoffTeam(3).division = "NL EAST" Then
                                                homeTeam = playoffTeam(4).name
                                                visitingTeam = playoffTeam(3).name
                                            Else
                                                homeTeam = playoffTeam(3).name
                                                visitingTeam = playoffTeam(4).name
                                            End If
                                        Else
                                            If playoffTeam(3).division = "NL EAST" Then
                                                homeTeam = playoffTeam(3).name
                                                visitingTeam = playoffTeam(4).name
                                            Else
                                                homeTeam = playoffTeam(4).name
                                                visitingTeam = playoffTeam(3).name
                                            End If
                                        End If
                                End Select
                                gstrPostSeason = conNLCS
                                bolAmericanLeagueRules = False
                                _game = gamesPlayed + 1
                            Else
                                sqlQuery = "Select * FROM PSStandings WHERE WSWins = 4"
                                dsStandings.Clear()
                                dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                If dsStandings.Tables(0).Rows.Count = 0 Then
                                    'In World Series
                                    sqlQuery = "Select * FROM PSStandings WHERE CSWins = " & csWinsNeeded.ToString
                                    dsStandings.Clear()
                                    dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                    teamIndex = 1
                                    For Each dr As DataRow In dsStandings.Tables(0).Rows
                                        playoffTeam(teamIndex).name = dr.Item("TeamID").ToString
                                        playoffTeam(teamIndex).league = dr.Item("League").ToString
                                        gamesPlayed += CInt(dr.Item("WSWins"))
                                        teamIndex += 1
                                    Next dr
                                    dsStandings.Clear()
                                    Select Case gamesPlayed
                                        Case 0, 1, 5, 6
                                            If isALWestHome Then
                                                bolAmericanLeagueRules = True
                                                If playoffTeam(1).league = "AL" Then
                                                    homeTeam = playoffTeam(1).name
                                                    visitingTeam = playoffTeam(2).name
                                                Else
                                                    homeTeam = playoffTeam(2).name
                                                    visitingTeam = playoffTeam(1).name
                                                End If
                                            Else
                                                bolAmericanLeagueRules = False
                                                If playoffTeam(1).league = "AL" Then
                                                    homeTeam = playoffTeam(2).name
                                                    visitingTeam = playoffTeam(1).name
                                                Else
                                                    homeTeam = playoffTeam(1).name
                                                    visitingTeam = playoffTeam(2).name
                                                End If
                                            End If
                                        Case 2, 3, 4
                                            If isALWestHome Then
                                                bolAmericanLeagueRules = False
                                                If playoffTeam(1).league = "AL" Then
                                                    homeTeam = playoffTeam(2).name
                                                    visitingTeam = playoffTeam(1).name
                                                Else
                                                    homeTeam = playoffTeam(1).name
                                                    visitingTeam = playoffTeam(2).name
                                                End If
                                            Else
                                                bolAmericanLeagueRules = True
                                                If playoffTeam(1).league = "AL" Then
                                                    homeTeam = playoffTeam(1).name
                                                    visitingTeam = playoffTeam(2).name
                                                Else
                                                    homeTeam = playoffTeam(2).name
                                                    visitingTeam = playoffTeam(1).name
                                                End If
                                            End If
                                    End Select
                                    gstrPostSeason = conWS
                                    _game = gamesPlayed + 1
                                Else
                                    'Season is over
                                    With dsStandings.Tables(0).Rows(0)
                                        Call MsgBox("Season is over. " & .Item("TeamID").ToString & _
                                                        " is the World Champion.")
                                    End With
                                    gbolSeasonOver = True
                                    homeTeam = ""
                                    visitingTeam = ""
                                End If
                            End If
                        End If
                    End If

                Case Is > 2021
                    'WC is 3 games, DS is 5 games, CS and WS are 7 games
                    bolAmericanLeagueRules = True
                    numberOfTeams = 12
                    sqlQuery = "Select Division, League, DSWins, TeamID " & "FROM PSStandings"
                    dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                    If dsStandings.Tables(0).Rows.Count = 0 Then
                        'No record, new playoffs
                        gbolNewPS = True
                        sqlQuery = "Select Wins, PlayerID, Division From Standings WHERE " & "Division = 'AL EAST' ORDER BY WINS DESC LIMIT 1"
                        dsStandings.Clear()
                        dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                        With dsStandings.Tables(0).Rows(0)
                            playoffTeam(1).name = .Item("PlayerID").ToString
                            playoffTeam(1).division = .Item("Division").ToString
                            playoffTeam(1).rsWins = CInt(.Item("Wins"))
                        End With
                        sqlQuery = "Select Wins, PlayerID, Division From Standings WHERE " & "Division = 'AL WEST' ORDER BY WINS DESC LIMIT 1"
                        dsStandings.Clear()
                        dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                        With dsStandings.Tables(0).Rows(0)
                            playoffTeam(2).name = .Item("PlayerID").ToString
                            playoffTeam(2).division = .Item("Division").ToString
                            playoffTeam(2).rsWins = CInt(.Item("Wins"))
                        End With
                        sqlQuery = "Select Wins, PlayerID, Division, Wins From Standings WHERE " & "Division = 'AL CENTRAL' ORDER BY WINS DESC LIMIT 1"
                        dsStandings.Clear()
                        dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                        With dsStandings.Tables(0).Rows(0)
                            playoffTeam(3).name = .Item("PlayerID").ToString
                            playoffTeam(3).division = .Item("Division").ToString
                            playoffTeam(3).rsWins = CInt(.Item("Wins"))
                        End With
                        sqlQuery = "Select Wins, PlayerID, Division, Wins From Standings WHERE " & "Division like 'AL%' AND " & "PlayerID Not IN ('" & playoffTeam(1).name & "','" & playoffTeam(2).name & "','" & playoffTeam(3).name & "') ORDER BY WINS DESC LIMIT 3"
                        dsStandings.Clear()
                        dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                        With dsStandings.Tables(0).Rows(0)
                            playoffTeam(4).name = .Item("PlayerID").ToString
                            playoffTeam(4).division = .Item("Division").ToString
                            playoffTeam(4).rsWins = CInt(.Item("Wins"))
                        End With
                        With dsStandings.Tables(0).Rows(1)
                            playoffTeam(5).name = .Item("PlayerID").ToString
                            playoffTeam(5).division = .Item("Division").ToString
                            playoffTeam(5).rsWins = CInt(.Item("Wins"))
                        End With
                        With dsStandings.Tables(0).Rows(2)
                            playoffTeam(6).name = .Item("PlayerID").ToString
                            playoffTeam(6).division = .Item("Division").ToString
                            playoffTeam(6).rsWins = CInt(.Item("Wins"))
                        End With

                        sqlQuery = "Select Wins, PlayerID, Division From Standings WHERE " & "Division = 'NL EAST' ORDER BY WINS DESC LIMIT 1"
                        dsStandings.Clear()
                        dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                        With dsStandings.Tables(0).Rows(0)
                            playoffTeam(7).name = .Item("PlayerID").ToString
                            playoffTeam(7).division = .Item("Division").ToString
                            playoffTeam(7).rsWins = CInt(.Item("Wins"))
                        End With
                        sqlQuery = "Select Wins, PlayerID, Division From Standings WHERE " & "Division = 'NL WEST' ORDER BY WINS DESC LIMIT 1"
                        dsStandings.Clear()
                        dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                        With dsStandings.Tables(0).Rows(0)
                            playoffTeam(8).name = .Item("PlayerID").ToString
                            playoffTeam(8).division = .Item("Division").ToString
                            playoffTeam(8).rsWins = CInt(.Item("Wins"))
                        End With
                        sqlQuery = "Select Wins, PlayerID, Division, Wins From Standings WHERE " & "Division = 'NL CENTRAL' ORDER BY WINS DESC LIMIT 1"
                        dsStandings.Clear()
                        dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                        With dsStandings.Tables(0).Rows(0)
                            playoffTeam(9).name = .Item("PlayerID").ToString
                            playoffTeam(9).division = .Item("Division").ToString
                            playoffTeam(9).rsWins = CInt(.Item("Wins"))
                        End With
                        sqlQuery = "Select Wins, PlayerID, Division, Wins From Standings WHERE " & "Division like 'NL%' AND " & "PlayerID Not IN ('" & playoffTeam(5).name & "','" & playoffTeam(6).name & "','" & playoffTeam(7).name & "') ORDER BY WINS DESC LIMIT 3"
                        dsStandings.Clear()
                        dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                        With dsStandings.Tables(0).Rows(0)
                            playoffTeam(10).name = .Item("PlayerID").ToString
                            playoffTeam(10).division = .Item("Division").ToString
                            playoffTeam(10).rsWins = CInt(.Item("Wins"))
                        End With
                        With dsStandings.Tables(0).Rows(1)
                            playoffTeam(11).name = .Item("PlayerID").ToString
                            playoffTeam(11).division = .Item("Division").ToString
                            playoffTeam(11).rsWins = CInt(.Item("Wins"))
                        End With
                        With dsStandings.Tables(0).Rows(2)
                            playoffTeam(12).name = .Item("PlayerID").ToString
                            playoffTeam(12).division = .Item("Division").ToString
                            playoffTeam(12).rsWins = CInt(.Item("Wins"))
                        End With

                        dsStandings.Clear()

                        'pick the weakest AL div winner
                        tempValue = 0 'index of chosen team
                        playoffTeam(tempValue).rsWins = 200
                        For i As Integer = 1 To 3
                            If playoffTeam(i).rsWins <= playoffTeam(tempValue).rsWins Then
                                playoffTeam(tempValue).beginSeries = ""
                                playoffTeam(i).beginSeries = conALWC1
                                tempValue = i
                            Else
                                playoffTeam(i).beginSeries = ""
                            End If
                        Next i

                        'pick the weakest AL wildcard
                        tempValue2 = 0 'index of chosen team
                        playoffTeam(tempValue2).rsWins = 200
                        For i As Integer = 4 To 6
                            If playoffTeam(i).rsWins <= playoffTeam(tempValue).rsWins Then
                                playoffTeam(tempValue).beginSeries = conALWC2
                                playoffTeam(i).beginSeries = conALWC1
                                tempValue = i
                            Else
                                playoffTeam(i).beginSeries = conALWC2
                            End If
                        Next i

                        'Set home/visitor for 1st playoff game (AL)
                        homeTeam = playoffTeam(tempValue).name
                        visitingTeam = playoffTeam(tempValue2).name

                        'pick the weakest NL div winner
                        tempValue = 0 'index of chosen team
                        playoffTeam(tempValue).rsWins = 200
                        For i As Integer = 7 To 9
                            If playoffTeam(i).rsWins <= playoffTeam(tempValue).rsWins Then
                                playoffTeam(tempValue).beginSeries = ""
                                playoffTeam(i).beginSeries = conNLWC1
                                tempValue = i
                            Else
                                playoffTeam(i).beginSeries = ""
                            End If
                        Next i

                        'pick the weakest NL wildcard
                        tempValue2 = 0 'index of chosen team
                        playoffTeam(tempValue2).rsWins = 200
                        For i As Integer = 10 To 12
                            If playoffTeam(i).rsWins <= playoffTeam(tempValue).rsWins Then
                                playoffTeam(tempValue).beginSeries = conNLWC2
                                playoffTeam(i).beginSeries = conNLWC1
                                tempValue = i
                            Else
                                playoffTeam(i).beginSeries = conNLWC2
                            End If
                        Next i

                        'Pick the top seeds in each league
                        'American League
                        tempValue = 0
                        playoffTeam(tempValue).rsWins = 0
                        For i As Integer = 1 To 3
                            If playoffTeam(i).beginSeries <> conALWC1 Then 'skip over div winner wildcard participant
                                If playoffTeam(i).rsWins >= playoffTeam(tempValue).rsWins Then
                                    playoffTeam(tempValue).beginSeries = conALDIV2
                                    playoffTeam(i).beginSeries = conALDIV1
                                    tempValue = i
                                Else
                                    playoffTeam(i).beginSeries = conALDIV2
                                End If
                            End If
                        Next i

                        'National League
                        tempValue = 0
                        playoffTeam(tempValue).rsWins = 0
                        For i As Integer = 7 To 9
                            If playoffTeam(i).beginSeries <> conNLWC1 Then 'skip over div winner wildcard participant
                                If playoffTeam(i).rsWins >= playoffTeam(tempValue).rsWins Then
                                    playoffTeam(tempValue).beginSeries = conNLDIV2
                                    playoffTeam(i).beginSeries = conNLDIV1
                                    tempValue = i
                                Else
                                    playoffTeam(i).beginSeries = conNLDIV2
                                End If
                            End If
                        Next i

                        For i As Integer = 1 To numberOfTeams
                            sqlFields = "TEAMID,RSWINS,WCWINS,DSWINS,CSWINS,WSWINS,DIVISION,LEAGUE,BEGINSERIES"
                            sqlValues = "'" & playoffTeam(i).name & "'," & playoffTeam(i).rsWins & ",0,0,0,0,'" & playoffTeam(i).division & "','" & IIF(i < 7, "AL", "NL") & "','" & playoffTeam(i).beginSeries & "'"
                            sqlQuery = "INSERT INTO PSSTANDINGS (" & sqlFields & ") VALUES (" & sqlValues & ")"
                            DataAccess.ExecuteScalar(sqlQuery)
                        Next i
                        gstrPostSeason = conALWC1
                        _game = 1
                    Else
                        sqlQuery = "Select * FROM PSStandings WHERE beginseries = '" & conALWC1 & "' " & "AND WCWins = 2"
                        dsStandings.Clear()
                        dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                        If dsStandings.Tables(0).Rows.Count = 0 Then
                            'In middle of ALWC1
                            sqlQuery = "Select * FROM PSStandings WHERE beginseries = '" & conALWC1 & "'"
                            dsStandings.Clear()
                            dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                            teamIndex = 1
                            For Each dr As DataRow In dsStandings.Tables(0).Rows
                                playoffTeam(teamIndex).name = dr.Item("TeamID").ToString
                                playoffTeam(teamIndex).division = dr.Item("Division").ToString
                                playoffTeam(teamIndex).rsWins = CInt(dr.Item("rswins"))
                                gamesPlayed = gamesPlayed + CInt(dr.Item("WCWins"))
                                teamIndex += 1
                            Next dr
                            If gamesPlayed = 0 Then
                                gbolNewPS = True
                            End If
                            dsStandings.Clear()
                            '1st row with ALWC1 is the home team (div winner) for all 3
                            homeTeam = playoffTeam(1).name
                            visitingTeam = playoffTeam(2).name

                            gstrPostSeason = conALWC1
                            _game = gamesPlayed + 1
                        Else
                            sqlQuery = "Select * FROM PSStandings WHERE beginseries = '" & conALWC2 & "' " & "AND WCWins = 2"
                            dsStandings.Clear()
                            dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                            If dsStandings.Tables(0).Rows.Count = 0 Then
                                'In middle of ALWC2
                                sqlQuery = "Select * FROM PSStandings WHERE beginseries = '" & conALWC2 & "'"
                                dsStandings.Clear()
                                dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                teamIndex = 1
                                For Each dr As DataRow In dsStandings.Tables(0).Rows
                                    playoffTeam(teamIndex).name = dr.Item("TeamID").ToString
                                    playoffTeam(teamIndex).division = dr.Item("Division").ToString
                                    playoffTeam(teamIndex).rsWins = CInt(dr.Item("rswins"))
                                    gamesPlayed = gamesPlayed + CInt(dr.Item("WCWins"))
                                    teamIndex += 1
                                Next dr
                                If gamesPlayed = 0 Then
                                    gbolNewPS = True
                                End If
                                dsStandings.Clear()
                                '1st row with ALWC2 is the home team (div winner) for all 3
                                homeTeam = playoffTeam(1).name
                                visitingTeam = playoffTeam(2).name

                                gstrPostSeason = conALWC2
                                _game = gamesPlayed + 1
                            Else
                                sqlQuery = "Select * FROM PSStandings WHERE beginseries = '" & conNLWC1 & "' " & "AND WCWins = 2"
                                dsStandings.Clear()
                                dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                If dsStandings.Tables(0).Rows.Count = 0 Then
                                    'In middle of NLWC1
                                    sqlQuery = "Select * FROM PSStandings WHERE beginseries = '" & conNLWC1 & "'"
                                    dsStandings.Clear()
                                    dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                    teamIndex = 1
                                    For Each dr As DataRow In dsStandings.Tables(0).Rows
                                        playoffTeam(teamIndex).name = dr.Item("TeamID").ToString
                                        playoffTeam(teamIndex).division = dr.Item("Division").ToString
                                        playoffTeam(teamIndex).rsWins = CInt(dr.Item("rswins"))
                                        gamesPlayed = gamesPlayed + CInt(dr.Item("WCWins"))
                                        teamIndex += 1
                                    Next dr
                                    If gamesPlayed = 0 Then
                                        gbolNewPS = True
                                    End If
                                    dsStandings.Clear()
                                    '1st row with NLWC1 is the home team (div winner) for all 3
                                    homeTeam = playoffTeam(1).name
                                    visitingTeam = playoffTeam(2).name

                                    gstrPostSeason = conNLWC1
                                    _game = gamesPlayed + 1
                                Else
                                    sqlQuery = "Select * FROM PSStandings WHERE beginseries = '" & conNLWC2 & "' " & "AND WCWins = 2"
                                    dsStandings.Clear()
                                    dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                    If dsStandings.Tables(0).Rows.Count = 0 Then
                                        'In middle of NLWC1
                                        sqlQuery = "Select * FROM PSStandings WHERE beginseries = '" & conNLWC2 & "'"
                                        dsStandings.Clear()
                                        dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                        teamIndex = 1
                                        For Each dr As DataRow In dsStandings.Tables(0).Rows
                                            playoffTeam(teamIndex).name = dr.Item("TeamID").ToString
                                            playoffTeam(teamIndex).division = dr.Item("Division").ToString
                                            playoffTeam(teamIndex).rsWins = CInt(dr.Item("rswins"))
                                            gamesPlayed = gamesPlayed + CInt(dr.Item("WCWins"))
                                            teamIndex += 1
                                        Next dr
                                        If gamesPlayed = 0 Then
                                            gbolNewPS = True
                                        End If
                                        dsStandings.Clear()
                                        '1st row with NLWC2 is the home team (div winner) for all 3
                                        homeTeam = playoffTeam(1).name
                                        visitingTeam = playoffTeam(2).name

                                        gstrPostSeason = conNLWC2
                                        _game = gamesPlayed + 1
                                    Else
                                        sqlQuery = "Select * FROM PSStandings WHERE beginseries = '" & conALDIV1 & "' " & "AND DSWins = 3"
                                        dsStandings.Clear()
                                        dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                        If dsStandings.Tables(0).Rows.Count = 0 Then
                                            'In middle of ALDS1
                                            'AL Div top seed plays the ALWC1 winner
                                            sqlQuery = "Select * FROM PSStandings WHERE beginseries = '" & conALDIV1 & "'" & _
                                                " OR (beginseries = '" & conALWC1 & "' and WCWins = 2)"
                                            dsStandings.Clear()
                                            dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                            teamIndex = 1
                                            For Each dr As DataRow In dsStandings.Tables(0).Rows
                                                playoffTeam(teamIndex).name = dr.Item("TeamID").ToString
                                                playoffTeam(teamIndex).division = dr.Item("Division").ToString
                                                playoffTeam(teamIndex).rsWins = CInt(dr.Item("rswins"))
                                                gamesPlayed = gamesPlayed + CInt(dr.Item("DSWins"))
                                                teamIndex += 1
                                            Next dr
                                            If gamesPlayed = 0 Then
                                                gbolNewPS = True
                                            End If
                                            dsStandings.Clear()
                                            Select Case gamesPlayed
                                                Case 2, 3
                                                    If playoffTeam(1).beginSeries = conALDIV1 Then
                                                        homeTeam = playoffTeam(2).name
                                                        visitingTeam = playoffTeam(1).name
                                                    Else
                                                        homeTeam = playoffTeam(1).name
                                                        visitingTeam = playoffTeam(2).name
                                                    End If
                                                Case 0, 1, 4
                                                    If playoffTeam(1).beginSeries = conALDIV1 Then
                                                        homeTeam = playoffTeam(1).name
                                                        visitingTeam = playoffTeam(2).name
                                                    Else
                                                        homeTeam = playoffTeam(2).name
                                                        visitingTeam = playoffTeam(1).name
                                                    End If
                                            End Select

                                            gstrPostSeason = conALDIV1
                                            _game = gamesPlayed + 1
                                        Else
                                            sqlQuery = "Select * FROM PSStandings WHERE beginseries = '" & conALDIV2 & "' " & "AND DSWins = 3"
                                            dsStandings.Clear()
                                            dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                            If dsStandings.Tables(0).Rows.Count = 0 Then
                                                'In middle of ALDS2
                                                'AL #2 seed plays the ALWC2 winner
                                                sqlQuery = "Select * FROM PSStandings WHERE beginseries = '" & conALDIV2 & "'" & _
                                                    " OR (beginseries = '" & conALWC2 & "' and WCWins = 2)"
                                                dsStandings.Clear()
                                                dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                                teamIndex = 1
                                                For Each dr As DataRow In dsStandings.Tables(0).Rows
                                                    playoffTeam(teamIndex).name = dr.Item("TeamID").ToString
                                                    playoffTeam(teamIndex).division = dr.Item("Division").ToString
                                                    playoffTeam(teamIndex).rsWins = CInt(dr.Item("rswins"))
                                                    gamesPlayed = gamesPlayed + CInt(dr.Item("DSWins"))
                                                    teamIndex += 1
                                                Next dr
                                                If gamesPlayed = 0 Then
                                                    gbolNewPS = True
                                                End If
                                                dsStandings.Clear()
                                                Select Case gamesPlayed
                                                    Case 2, 3
                                                        If playoffTeam(1).beginSeries = conALDIV2 Then
                                                            homeTeam = playoffTeam(2).name
                                                            visitingTeam = playoffTeam(1).name
                                                        Else
                                                            homeTeam = playoffTeam(1).name
                                                            visitingTeam = playoffTeam(2).name
                                                        End If
                                                    Case 0, 1, 4
                                                        If playoffTeam(1).beginSeries = conALDIV2 Then
                                                            homeTeam = playoffTeam(1).name
                                                            visitingTeam = playoffTeam(2).name
                                                        Else
                                                            homeTeam = playoffTeam(2).name
                                                            visitingTeam = playoffTeam(1).name
                                                        End If
                                                End Select

                                                gstrPostSeason = conALDIV2
                                                _game = gamesPlayed + 1
                                            Else
                                                sqlQuery = "Select * FROM PSStandings WHERE beginseries = '" & conNLDIV1 & "' " & "AND DSWins = 3"
                                                dsStandings.Clear()
                                                dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                                If dsStandings.Tables(0).Rows.Count = 0 Then
                                                    'In middle of NLDS1
                                                    'NL Div top seed plays the NLWC1 winner
                                                    sqlQuery = "Select * FROM PSStandings WHERE beginseries = '" & conNLDIV1 & "'" & _
                                                        " OR (beginseries = '" & conNLWC1 & "' and WCWins = 2)"
                                                    dsStandings.Clear()
                                                    dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                                    teamIndex = 1
                                                    For Each dr As DataRow In dsStandings.Tables(0).Rows
                                                        playoffTeam(teamIndex).name = dr.Item("TeamID").ToString
                                                        playoffTeam(teamIndex).division = dr.Item("Division").ToString
                                                        playoffTeam(teamIndex).rsWins = CInt(dr.Item("rswins"))
                                                        gamesPlayed = gamesPlayed + CInt(dr.Item("DSWins"))
                                                        teamIndex += 1
                                                    Next dr
                                                    If gamesPlayed = 0 Then
                                                        gbolNewPS = True
                                                    End If
                                                    dsStandings.Clear()
                                                    Select Case gamesPlayed
                                                        Case 2, 3
                                                            If playoffTeam(1).beginSeries = conNLDIV1 Then
                                                                homeTeam = playoffTeam(2).name
                                                                visitingTeam = playoffTeam(1).name
                                                            Else
                                                                homeTeam = playoffTeam(1).name
                                                                visitingTeam = playoffTeam(2).name
                                                            End If
                                                        Case 0, 1, 4
                                                            If playoffTeam(1).beginSeries = conNLDIV1 Then
                                                                homeTeam = playoffTeam(1).name
                                                                visitingTeam = playoffTeam(2).name
                                                            Else
                                                                homeTeam = playoffTeam(2).name
                                                                visitingTeam = playoffTeam(1).name
                                                            End If
                                                    End Select

                                                    gstrPostSeason = conNLDIV1
                                                    _game = gamesPlayed + 1
                                                Else
                                                    sqlQuery = "Select * FROM PSStandings WHERE beginseries = '" & conNLDIV2 & "' " & "AND DSWins = 3"
                                                    dsStandings.Clear()
                                                    dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                                    If dsStandings.Tables(0).Rows.Count = 0 Then
                                                        'In middle of NLDS2
                                                        'NL #2 seed plays the NLWC2 winner
                                                        sqlQuery = "Select * FROM PSStandings WHERE beginseries = '" & conNLDIV2 & "'" & _
                                                            " OR (beginseries = '" & conNLWC2 & "' and WCWins = 2)"
                                                        dsStandings.Clear()
                                                        dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                                        teamIndex = 1
                                                        For Each dr As DataRow In dsStandings.Tables(0).Rows
                                                            playoffTeam(teamIndex).name = dr.Item("TeamID").ToString
                                                            playoffTeam(teamIndex).division = dr.Item("Division").ToString
                                                            playoffTeam(teamIndex).rsWins = CInt(dr.Item("rswins"))
                                                            gamesPlayed = gamesPlayed + CInt(dr.Item("DSWins"))
                                                            teamIndex += 1
                                                        Next dr
                                                        If gamesPlayed = 0 Then
                                                            gbolNewPS = True
                                                        End If
                                                        dsStandings.Clear()
                                                        Select Case gamesPlayed
                                                            Case 2, 3
                                                                If playoffTeam(1).beginSeries = conNLDIV2 Then
                                                                    homeTeam = playoffTeam(2).name
                                                                    visitingTeam = playoffTeam(1).name
                                                                Else
                                                                    homeTeam = playoffTeam(1).name
                                                                    visitingTeam = playoffTeam(2).name
                                                                End If
                                                            Case 0, 1, 4
                                                                If playoffTeam(1).beginSeries = conNLDIV2 Then
                                                                    homeTeam = playoffTeam(1).name
                                                                    visitingTeam = playoffTeam(2).name
                                                                Else
                                                                    homeTeam = playoffTeam(2).name
                                                                    visitingTeam = playoffTeam(1).name
                                                                End If
                                                        End Select

                                                        gstrPostSeason = conNLDIV2
                                                        _game = gamesPlayed + 1
                                                    Else
                                                        sqlQuery = "Select * FROM PSStandings WHERE league = 'AL' " & "AND CSWins = 4"
                                                        dsStandings.Clear()
                                                        dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                                        If dsStandings.Tables(0).Rows.Count = 0 Then
                                                            'In middle of ALCS
                                                            sqlQuery = "Select * FROM PSStandings WHERE league = 'AL' " & "AND DSWins = 3"
                                                            dsStandings.Clear()
                                                            dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                                            teamIndex = 1
                                                            For Each dr As DataRow In dsStandings.Tables(0).Rows
                                                                playoffTeam(teamIndex).name = dr.Item("TeamID").ToString
                                                                playoffTeam(teamIndex).division = dr.Item("Division").ToString
                                                                playoffTeam(teamIndex).rsWins = CInt(dr.Item("rswins"))
                                                                gamesPlayed = gamesPlayed + CInt(dr.Item("CSWins"))
                                                                teamIndex += 1
                                                            Next dr
                                                            If gamesPlayed = 0 Then
                                                                gbolNewPS = True
                                                            End If
                                                            dsStandings.Clear()
                                                            Select Case gamesPlayed
                                                                Case 0, 1, 5, 6
                                                                    If playoffTeam(1).rsWins > playoffTeam(2).rsWins Then
                                                                        homeTeam = playoffTeam(1).name
                                                                        visitingTeam = playoffTeam(2).name
                                                                    Else
                                                                        homeTeam = playoffTeam(2).name
                                                                        visitingTeam = playoffTeam(1).name
                                                                    End If
                                                                Case 2, 3, 4
                                                                    If playoffTeam(1).rsWins > playoffTeam(2).rsWins Then
                                                                        homeTeam = playoffTeam(2).name
                                                                        visitingTeam = playoffTeam(1).name
                                                                    Else
                                                                        homeTeam = playoffTeam(1).name
                                                                        visitingTeam = playoffTeam(2).name
                                                                    End If
                                                            End Select
                                                            gstrPostSeason = conALCS
                                                            _game = gamesPlayed + 1
                                                        Else
                                                            sqlQuery = "Select * FROM PSStandings WHERE league = 'NL' " & "AND CSWins = 4"
                                                            dsStandings.Clear()
                                                            dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                                            If dsStandings.Tables(0).Rows.Count = 0 Then
                                                                'In middle of NLCS
                                                                sqlQuery = "Select * FROM PSStandings WHERE league = 'NL' " & "AND DSWins = 3"
                                                                dsStandings.Clear()
                                                                dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                                                teamIndex = 1
                                                                For Each dr As DataRow In dsStandings.Tables(0).Rows
                                                                    playoffTeam(teamIndex).name = dr.Item("TeamID").ToString
                                                                    playoffTeam(teamIndex).division = dr.Item("Division").ToString
                                                                    playoffTeam(teamIndex).rsWins = CInt(dr.Item("rswins"))
                                                                    gamesPlayed = gamesPlayed + CInt(dr.Item("CSWins"))
                                                                    teamIndex += 1
                                                                Next dr
                                                                If gamesPlayed = 0 Then
                                                                    gbolNewPS = True
                                                                End If
                                                                dsStandings.Clear()
                                                                Select Case gamesPlayed
                                                                    Case 0, 1, 5, 6
                                                                        If playoffTeam(1).rsWins > playoffTeam(2).rsWins Then
                                                                            homeTeam = playoffTeam(1).name
                                                                            visitingTeam = playoffTeam(2).name
                                                                        Else
                                                                            homeTeam = playoffTeam(2).name
                                                                            visitingTeam = playoffTeam(1).name
                                                                        End If
                                                                    Case 2, 3, 4
                                                                        If playoffTeam(1).rsWins > playoffTeam(2).rsWins Then
                                                                            homeTeam = playoffTeam(2).name
                                                                            visitingTeam = playoffTeam(1).name
                                                                        Else
                                                                            homeTeam = playoffTeam(1).name
                                                                            visitingTeam = playoffTeam(2).name
                                                                        End If
                                                                End Select
                                                                gstrPostSeason = conNLCS
                                                                _game = gamesPlayed + 1
                                                            Else
                                                                sqlQuery = "Select * FROM PSStandings WHERE WSWins = 4"
                                                                dsStandings.Clear()
                                                                dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                                                If dsStandings.Tables(0).Rows.Count = 0 Then
                                                                    'In World Series
                                                                    sqlQuery = "Select * FROM PSStandings WHERE CSWins = 4"
                                                                    dsStandings.Clear()
                                                                    dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                                                    teamIndex = 1
                                                                    For Each dr As DataRow In dsStandings.Tables(0).Rows
                                                                        playoffTeam(teamIndex).name = dr.Item("TeamID").ToString
                                                                        playoffTeam(teamIndex).league = dr.Item("League").ToString
                                                                        playoffTeam(teamIndex).rsWins = CInt(dr.Item("rswins"))
                                                                        gamesPlayed = gamesPlayed + CInt(dr.Item("WSWins"))
                                                                        teamIndex += 1
                                                                    Next dr
                                                                    If Val(gstrSeason) > 2016 Then
                                                                        'After 2016, best record is home
                                                                        If playoffTeam(1).rsWins > playoffTeam(2).rsWins Then
                                                                            isALHome = playoffTeam(1).league = "AL"
                                                                        Else
                                                                            isALHome = playoffTeam(2).league = "AL"
                                                                        End If
                                                                    Else
                                                                        'Up until 2016, AL Team has home advantage in even years
                                                                        isALHome = Val(gstrSeason) Mod 2 = 0
                                                                    End If

                                                                    dsStandings.Clear()
                                                                    Select Case gamesPlayed
                                                                        Case 0, 1, 5, 6
                                                                            If isALHome Then
                                                                                If playoffTeam(1).league = "AL" Then
                                                                                    homeTeam = playoffTeam(1).name
                                                                                    visitingTeam = playoffTeam(2).name
                                                                                Else
                                                                                    homeTeam = playoffTeam(2).name
                                                                                    visitingTeam = playoffTeam(1).name
                                                                                End If
                                                                            Else
                                                                                If playoffTeam(1).league = "AL" Then
                                                                                    homeTeam = playoffTeam(2).name
                                                                                    visitingTeam = playoffTeam(1).name
                                                                                Else
                                                                                    homeTeam = playoffTeam(1).name
                                                                                    visitingTeam = playoffTeam(2).name
                                                                                End If
                                                                            End If
                                                                        Case 2, 3, 4
                                                                            If isALHome Then
                                                                                If playoffTeam(1).league = "AL" Then
                                                                                    homeTeam = playoffTeam(2).name
                                                                                    visitingTeam = playoffTeam(1).name
                                                                                Else
                                                                                    homeTeam = playoffTeam(1).name
                                                                                    visitingTeam = playoffTeam(2).name
                                                                                End If
                                                                            Else
                                                                                If playoffTeam(1).league = "AL" Then
                                                                                    homeTeam = playoffTeam(1).name
                                                                                    visitingTeam = playoffTeam(2).name
                                                                                Else
                                                                                    homeTeam = playoffTeam(2).name
                                                                                    visitingTeam = playoffTeam(1).name
                                                                                End If
                                                                            End If
                                                                    End Select
                                                                    gstrPostSeason = conWS
                                                                    _game = gamesPlayed + 1
                                                                Else
                                                                    'Season is over
                                                                    With dsStandings.Tables(0).Rows(0)
                                                                        Call MsgBox("Season is over. " & .Item("TeamID").ToString & " is the World Champion.")
                                                                    End With
                                                                    gbolSeasonOver = True
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If

                Case 1994 To 2021
                    'DS is 5 games, CS and WS are 7 games
                    numberOfTeams = IIF(Val(gstrSeason) >= 2012, 10, 8)
                    sqlQuery = "Select Division, League, DSWins, TeamID " & "FROM PSStandings"
                    dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                    If dsStandings.Tables(0).Rows.Count = 0 Then
                        'No record, new playoffs
                        gbolNewPS = True
                        sqlQuery = "Select Wins, PlayerID, Division From Standings WHERE " & "Division = 'AL EAST' ORDER BY WINS DESC LIMIT 1"
                        dsStandings.Clear()
                        dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                        With dsStandings.Tables(0).Rows(0)
                            playoffTeam(1).name = .Item("PlayerID").ToString
                            playoffTeam(1).division = .Item("Division").ToString
                            playoffTeam(1).rsWins = CInt(.Item("Wins"))
                        End With
                        sqlQuery = "Select Wins, PlayerID, Division From Standings WHERE " & "Division = 'AL WEST' ORDER BY WINS DESC LIMIT 1"
                        dsStandings.Clear()
                        dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                        With dsStandings.Tables(0).Rows(0)
                            playoffTeam(2).name = .Item("PlayerID").ToString
                            playoffTeam(2).division = .Item("Division").ToString
                            playoffTeam(2).rsWins = CInt(.Item("Wins"))
                        End With
                        sqlQuery = "Select Wins, PlayerID, Division, Wins From Standings WHERE " & "Division = 'AL CENTRAL' ORDER BY WINS DESC LIMIT 1"
                        dsStandings.Clear()
                        dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                        With dsStandings.Tables(0).Rows(0)
                            playoffTeam(3).name = .Item("PlayerID").ToString
                            playoffTeam(3).division = .Item("Division").ToString
                            playoffTeam(3).rsWins = CInt(.Item("Wins"))
                        End With
                        sqlQuery = "Select Wins, PlayerID, Division, Wins From Standings WHERE " & "Division like 'AL%' AND " & "PlayerID Not IN ('" & playoffTeam(1).name & "','" & playoffTeam(2).name & "','" & playoffTeam(3).name & "') ORDER BY WINS DESC LIMIT 2"
                        dsStandings.Clear()
                        dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                        With dsStandings.Tables(0).Rows(0)
                            playoffTeam(4).name = .Item("PlayerID").ToString
                            playoffTeam(4).division = .Item("Division").ToString
                            playoffTeam(4).rsWins = CInt(.Item("Wins"))
                        End With
                        If Val(gstrSeason) >= 2012 Then
                            'set bonus AL wildcard
                            With dsStandings.Tables(0).Rows(1)
                                playoffTeam(9).name = .Item("PlayerID").ToString
                                playoffTeam(9).division = .Item("Division").ToString
                                playoffTeam(9).rsWins = CInt(.Item("Wins"))
                            End With
                        End If
                        sqlQuery = "Select Wins, PlayerID, Division From Standings WHERE " & "Division = 'NL EAST' ORDER BY WINS DESC LIMIT 1"
                        dsStandings.Clear()
                        dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                        With dsStandings.Tables(0).Rows(0)
                            playoffTeam(5).name = .Item("PlayerID").ToString
                            playoffTeam(5).division = .Item("Division").ToString
                            playoffTeam(5).rsWins = CInt(.Item("Wins"))
                        End With
                        sqlQuery = "Select Wins, PlayerID, Division From Standings WHERE " & "Division = 'NL WEST' ORDER BY WINS DESC LIMIT 1"
                        dsStandings.Clear()
                        dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                        With dsStandings.Tables(0).Rows(0)
                            playoffTeam(6).name = .Item("PlayerID").ToString
                            playoffTeam(6).division = .Item("Division").ToString
                            playoffTeam(6).rsWins = CInt(.Item("Wins"))
                        End With
                        sqlQuery = "Select Wins, PlayerID, Division, Wins From Standings WHERE " & "Division = 'NL CENTRAL' ORDER BY WINS DESC LIMIT 1"
                        dsStandings.Clear()
                        dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                        With dsStandings.Tables(0).Rows(0)
                            playoffTeam(7).name = .Item("PlayerID").ToString
                            playoffTeam(7).division = .Item("Division").ToString
                            playoffTeam(7).rsWins = CInt(.Item("Wins"))
                        End With
                        sqlQuery = "Select Wins, PlayerID, Division, Wins From Standings WHERE " & "Division like 'NL%' AND " & "PlayerID Not IN ('" & playoffTeam(5).name & "','" & playoffTeam(6).name & "','" & playoffTeam(7).name & "') ORDER BY WINS DESC LIMIT 2"
                        dsStandings.Clear()
                        dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                        With dsStandings.Tables(0).Rows(0)
                            playoffTeam(8).name = .Item("PlayerID").ToString
                            playoffTeam(8).division = .Item("Division").ToString
                            playoffTeam(8).rsWins = CInt(.Item("Wins"))
                        End With
                        If Val(gstrSeason) >= 2012 Then
                            'set bonus NL wildcard
                            With dsStandings.Tables(0).Rows(1)
                                playoffTeam(10).name = .Item("PlayerID").ToString
                                playoffTeam(10).division = .Item("Division").ToString
                                playoffTeam(10).rsWins = CInt(.Item("Wins"))
                            End With
                        End If
                        dsStandings.Clear()
                        tempValue = 0
                        playoffTeam(tempValue).rsWins = 0
                        'playoffTeam(1).beginSeries = conALDIV1
                        'pick the strongest team for ALDIV1, to play the weakest
                        For i As Integer = 1 To 3
                            If playoffTeam(i).rsWins >= playoffTeam(tempValue).rsWins And _
                                    playoffTeam(i).division <> playoffTeam(4).division Then
                                playoffTeam(tempValue).beginSeries = conALDIV2
                                playoffTeam(i).beginSeries = conALDIV1
                                tempValue = i
                            Else
                                playoffTeam(i).beginSeries = conALDIV2
                            End If
                        Next i
                        If Val(gstrSeason) >= 2012 Then
                            playoffTeam(4).beginSeries = conALWC
                            playoffTeam(9).beginSeries = conALWC
                        Else
                            playoffTeam(4).beginSeries = conALDIV1
                        End If

                        If Val(gstrSeason) >= 2012 Then
                            homeTeam = playoffTeam(4).name '1 game WC host
                            visitingTeam = playoffTeam(9).name
                        Else
                            homeTeam = playoffTeam(tempValue).name
                            visitingTeam = playoffTeam(4).name 'wildcard team
                        End If
                        'homeTeam = playoffTeam(4).name 'wildcard team, or 1 game WC host
                        'visitingTeam = IIF(Val(gstrSeason) >= 2012, playoffTeam(9).name, playoffTeam(tempValue).name)
                        tempValue = 0
                        playoffTeam(tempValue).rsWins = 0
                        'playoffTeam(tempValue).beginSeries = conNLDIV1
                        'pick the strongest team for NLDIV1, to play the weakest
                        For i As Integer = 5 To 7
                            If playoffTeam(i).rsWins >= playoffTeam(tempValue).rsWins And _
                                    playoffTeam(i).division <> playoffTeam(8).division Then
                                playoffTeam(tempValue).beginSeries = conNLDIV2
                                playoffTeam(i).beginSeries = conNLDIV1
                                tempValue = i
                            Else
                                playoffTeam(i).beginSeries = conNLDIV2
                            End If
                        Next i

                        If Val(gstrSeason) >= 2012 Then
                            'Add 1 game wildcard
                            playoffTeam(8).beginSeries = conNLWC
                            playoffTeam(10).beginSeries = conNLWC
                        Else
                            playoffTeam(8).beginSeries = conNLDIV1
                        End If

                        For i As Integer = 1 To numberOfTeams
                            sqlFields = "TEAMID,RSWINS,DSWINS,CSWINS,WSWINS,DIVISION,LEAGUE,BEGINSERIES"
                            sqlValues = "'" & playoffTeam(i).name & "'," & playoffTeam(i).rsWins & ",0,0,0,'" & playoffTeam(i).division & "','" & IIF(i < 5 Or i = 9, "AL", "NL") & "','" & playoffTeam(i).beginSeries & "'"
                            sqlQuery = "INSERT INTO PSSTANDINGS (" & sqlFields & ") VALUES (" & sqlValues & ")"
                            DataAccess.ExecuteScalar(sqlQuery)
                        Next i
                        gstrPostSeason = IIF(Val(gstrSeason) >= 2012, conALWC, conALDIV1)
                        bolAmericanLeagueRules = True
                        _game = 1
                    Else
                        sqlQuery = "Select * FROM PSStandings WHERE beginseries = '" & conALDIV1 & "' " & "AND DSWins = 3"
                        dsStandings.Clear()
                        dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                        If dsStandings.Tables(0).Rows.Count = 0 Then
                            'check for wildcard in progress
                            If Val(gstrSeason) >= 2012 Then
                                sqlQuery = "Select * FROM PSStandings WHERE beginseries = '" & conALWC & "'"
                                dsStandings.Clear()
                                dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                If dsStandings.Tables(0).Rows.Count = 2 Then
                                    'ALWC is in progress
                                    teamIndex = 1
                                    For Each dr As DataRow In dsStandings.Tables(0).Rows
                                        playoffTeam(teamIndex).name = dr.Item("TeamID").ToString
                                        playoffTeam(teamIndex).division = dr.Item("Division").ToString
                                        playoffTeam(teamIndex).rsWins = CInt(dr.Item("rswins"))
                                        gamesPlayed = gamesPlayed + CInt(dr.Item("DSWins"))
                                        teamIndex += 1
                                    Next dr
                                    gbolNewPS = True
                                    gstrPostSeason = conALWC
                                    bolAmericanLeagueRules = True
                                Else
                                    'Check for NL WC
                                    sqlQuery = "Select * FROM PSStandings WHERE beginseries = '" & conNLWC & "'"
                                    dsStandings.Clear()
                                    dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                    If dsStandings.Tables(0).Rows.Count = 2 Then
                                        'NLWC is in progress
                                        teamIndex = 1
                                        For Each dr As DataRow In dsStandings.Tables(0).Rows
                                            playoffTeam(teamIndex).name = dr.Item("TeamID").ToString
                                            playoffTeam(teamIndex).division = dr.Item("Division").ToString
                                            playoffTeam(teamIndex).rsWins = CInt(dr.Item("rswins"))
                                            gamesPlayed = gamesPlayed + CInt(dr.Item("DSWins"))
                                            teamIndex += 1
                                        Next dr
                                        gbolNewPS = True
                                        gstrPostSeason = conNLWC
                                        bolAmericanLeagueRules = False
                                    End If
                                End If
                                If gstrPostSeason = conALWC Or gstrPostSeason = conNLWC Then
                                    If playoffTeam(1).rsWins > playoffTeam(2).rsWins Then
                                        homeTeam = playoffTeam(1).name
                                        visitingTeam = playoffTeam(2).name
                                    Else
                                        homeTeam = playoffTeam(2).name
                                        visitingTeam = playoffTeam(1).name
                                    End If
                                End If
                            End If
                            If Not (gstrPostSeason = conALWC Or gstrPostSeason = conNLWC) Then
                                'In middle of ALDIV1
                                sqlQuery = "Select * FROM PSStandings WHERE beginseries = '" & conALDIV1 & "'"
                                dsStandings.Clear()
                                dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                teamIndex = 1
                                For Each dr As DataRow In dsStandings.Tables(0).Rows
                                    playoffTeam(teamIndex).name = dr.Item("TeamID").ToString
                                    playoffTeam(teamIndex).division = dr.Item("Division").ToString
                                    playoffTeam(teamIndex).rsWins = CInt(dr.Item("rswins"))
                                    gamesPlayed = gamesPlayed + CInt(dr.Item("DSWins"))
                                    teamIndex += 1
                                Next dr
                                If gamesPlayed = 0 Then
                                    gbolNewPS = True
                                End If
                                dsStandings.Clear()
                                Select Case gamesPlayed
                                    Case 2, 3
                                        If playoffTeam(1).rsWins > playoffTeam(2).rsWins Then
                                            homeTeam = playoffTeam(2).name
                                            visitingTeam = playoffTeam(1).name
                                        Else
                                            homeTeam = playoffTeam(1).name
                                            visitingTeam = playoffTeam(2).name
                                        End If
                                    Case 0, 1, 4
                                        If playoffTeam(1).rsWins > playoffTeam(2).rsWins Then
                                            homeTeam = playoffTeam(1).name
                                            visitingTeam = playoffTeam(2).name
                                        Else
                                            homeTeam = playoffTeam(2).name
                                            visitingTeam = playoffTeam(1).name
                                        End If
                                End Select
                                gstrPostSeason = conALDIV1
                                bolAmericanLeagueRules = True
                                _game = gamesPlayed + 1
                            End If
                        Else
                            sqlQuery = "Select * FROM PSStandings WHERE beginseries = '" & conALDIV2 & "' " & "AND DSWins = 3"
                            dsStandings.Clear()
                            dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                            If dsStandings.Tables(0).Rows.Count = 0 Then
                                'In middle of ALDIV2
                                sqlQuery = "Select * FROM PSStandings WHERE beginseries = '" & conALDIV2 & "'"
                                dsStandings.Clear()
                                dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                teamIndex = 3
                                For Each dr As DataRow In dsStandings.Tables(0).Rows
                                    playoffTeam(teamIndex).name = dr.Item("TeamID").ToString
                                    playoffTeam(teamIndex).division = dr.Item("Division").ToString
                                    playoffTeam(teamIndex).rsWins = CInt(dr.Item("rswins"))
                                    gamesPlayed = gamesPlayed + CInt(dr.Item("DSWins"))
                                    teamIndex += 1
                                Next dr
                                If gamesPlayed = 0 Then
                                    gbolNewPS = True
                                End If
                                dsStandings.Clear()
                                Select Case gamesPlayed
                                    Case 2, 3
                                        If playoffTeam(3).rsWins > playoffTeam(4).rsWins Then
                                            homeTeam = playoffTeam(4).name
                                            visitingTeam = playoffTeam(3).name
                                        Else
                                            homeTeam = playoffTeam(3).name
                                            visitingTeam = playoffTeam(4).name
                                        End If
                                    Case 0, 1, 4
                                        If playoffTeam(3).rsWins > playoffTeam(4).rsWins Then
                                            homeTeam = playoffTeam(3).name
                                            visitingTeam = playoffTeam(4).name
                                        Else
                                            homeTeam = playoffTeam(4).name
                                            visitingTeam = playoffTeam(3).name
                                        End If
                                End Select
                                gstrPostSeason = conALDIV2
                                bolAmericanLeagueRules = True
                                _game = gamesPlayed + 1
                            Else
                                sqlQuery = "Select * FROM PSStandings WHERE beginseries = '" & conNLDIV1 & "' " & "AND DSWins = 3"
                                dsStandings.Clear()
                                dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                If dsStandings.Tables(0).Rows.Count = 0 Then
                                    'In middle of NLDIV1
                                    sqlQuery = "Select * FROM PSStandings WHERE beginseries = '" & conNLDIV1 & "'"
                                    dsStandings.Clear()
                                    dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                    teamIndex = 5
                                    For Each dr As DataRow In dsStandings.Tables(0).Rows
                                        playoffTeam(teamIndex).name = dr.Item("TeamID").ToString
                                        playoffTeam(teamIndex).division = dr.Item("Division").ToString
                                        playoffTeam(teamIndex).rsWins = CInt(dr.Item("rswins"))
                                        gamesPlayed = gamesPlayed + CInt(dr.Item("DSWins"))
                                        teamIndex += 1
                                    Next dr
                                    If gamesPlayed = 0 Then
                                        gbolNewPS = True
                                    End If
                                    dsStandings.Clear()
                                    Select Case gamesPlayed
                                        Case 2, 3
                                            If playoffTeam(5).rsWins > playoffTeam(6).rsWins Then
                                                homeTeam = playoffTeam(6).name
                                                visitingTeam = playoffTeam(5).name
                                            Else
                                                homeTeam = playoffTeam(5).name
                                                visitingTeam = playoffTeam(6).name
                                            End If
                                        Case 0, 1, 4
                                            If playoffTeam(5).rsWins > playoffTeam(6).rsWins Then
                                                homeTeam = playoffTeam(5).name
                                                visitingTeam = playoffTeam(6).name
                                            Else
                                                homeTeam = playoffTeam(6).name
                                                visitingTeam = playoffTeam(5).name
                                            End If
                                    End Select
                                    gstrPostSeason = conNLDIV1
                                    bolAmericanLeagueRules = False
                                    _game = gamesPlayed + 1
                                Else
                                    sqlQuery = "Select * FROM PSStandings WHERE beginseries = '" & conNLDIV2 & "' " & "AND DSWins = 3"
                                    dsStandings.Clear()
                                    dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                    If dsStandings.Tables(0).Rows.Count = 0 Then
                                        'In middle of NLDIV2
                                        sqlQuery = "Select * FROM PSStandings WHERE beginseries = '" & conNLDIV2 & "'"
                                        dsStandings.Clear()
                                        dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                        teamIndex = 7
                                        For Each dr As DataRow In dsStandings.Tables(0).Rows
                                            playoffTeam(teamIndex).name = dr.Item("TeamID").ToString
                                            playoffTeam(teamIndex).division = dr.Item("Division").ToString
                                            playoffTeam(teamIndex).rsWins = CInt(dr.Item("rswins"))
                                            gamesPlayed = gamesPlayed + CInt(dr.Item("DSWins"))
                                            teamIndex += 1
                                        Next dr
                                        If gamesPlayed = 0 Then
                                            gbolNewPS = True
                                        End If
                                        dsStandings.Clear()
                                        Select Case gamesPlayed
                                            Case 2, 3
                                                If playoffTeam(7).rsWins > playoffTeam(8).rsWins Then
                                                    homeTeam = playoffTeam(8).name
                                                    visitingTeam = playoffTeam(7).name
                                                Else
                                                    homeTeam = playoffTeam(7).name
                                                    visitingTeam = playoffTeam(8).name
                                                End If
                                            Case 0, 1, 4
                                                If playoffTeam(7).rsWins > playoffTeam(8).rsWins Then
                                                    homeTeam = playoffTeam(7).name
                                                    visitingTeam = playoffTeam(8).name
                                                Else
                                                    homeTeam = playoffTeam(8).name
                                                    visitingTeam = playoffTeam(7).name
                                                End If
                                        End Select
                                        gstrPostSeason = conNLDIV2
                                        bolAmericanLeagueRules = False
                                        _game = gamesPlayed + 1
                                    Else
                                        sqlQuery = "Select * FROM PSStandings WHERE league = 'AL' " & "AND CSWins = 4"
                                        dsStandings.Clear()
                                        dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                        If dsStandings.Tables(0).Rows.Count = 0 Then
                                            'In middle of ALCS
                                            sqlQuery = "Select * FROM PSStandings WHERE league = 'AL' " & "AND DSWins = 3"
                                            dsStandings.Clear()
                                            dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                            teamIndex = 1
                                            For Each dr As DataRow In dsStandings.Tables(0).Rows
                                                playoffTeam(teamIndex).name = dr.Item("TeamID").ToString
                                                playoffTeam(teamIndex).division = dr.Item("Division").ToString
                                                playoffTeam(teamIndex).rsWins = CInt(dr.Item("rswins"))
                                                gamesPlayed = gamesPlayed + CInt(dr.Item("CSWins"))
                                                teamIndex += 1
                                            Next dr
                                            If gamesPlayed = 0 Then
                                                gbolNewPS = True
                                            End If
                                            dsStandings.Clear()
                                            Select Case gamesPlayed
                                                Case 0, 1, 5, 6
                                                    If playoffTeam(1).rsWins > playoffTeam(2).rsWins Then
                                                        homeTeam = playoffTeam(1).name
                                                        visitingTeam = playoffTeam(2).name
                                                    Else
                                                        homeTeam = playoffTeam(2).name
                                                        visitingTeam = playoffTeam(1).name
                                                    End If
                                                Case 2, 3, 4
                                                    If playoffTeam(1).rsWins > playoffTeam(2).rsWins Then
                                                        homeTeam = playoffTeam(2).name
                                                        visitingTeam = playoffTeam(1).name
                                                    Else
                                                        homeTeam = playoffTeam(1).name
                                                        visitingTeam = playoffTeam(2).name
                                                    End If
                                            End Select
                                            gstrPostSeason = conALCS
                                            bolAmericanLeagueRules = True
                                            _game = gamesPlayed + 1
                                        Else
                                            sqlQuery = "Select * FROM PSStandings WHERE league = 'NL' " & "AND CSWins = 4"
                                            dsStandings.Clear()
                                            dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                            If dsStandings.Tables(0).Rows.Count = 0 Then
                                                'In middle of NLCS
                                                sqlQuery = "Select * FROM PSStandings WHERE league = 'NL' " & "AND DSWins = 3"
                                                dsStandings.Clear()
                                                dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                                teamIndex = 1
                                                For Each dr As DataRow In dsStandings.Tables(0).Rows
                                                    playoffTeam(teamIndex).name = dr.Item("TeamID").ToString
                                                    playoffTeam(teamIndex).division = dr.Item("Division").ToString
                                                    playoffTeam(teamIndex).rsWins = CInt(dr.Item("rswins"))
                                                    gamesPlayed = gamesPlayed + CInt(dr.Item("CSWins"))
                                                    teamIndex += 1
                                                Next dr
                                                If gamesPlayed = 0 Then
                                                    gbolNewPS = True
                                                End If
                                                dsStandings.Clear()
                                                Select Case gamesPlayed
                                                    Case 0, 1, 5, 6
                                                        If playoffTeam(1).rsWins > playoffTeam(2).rsWins Then
                                                            homeTeam = playoffTeam(1).name
                                                            visitingTeam = playoffTeam(2).name
                                                        Else
                                                            homeTeam = playoffTeam(2).name
                                                            visitingTeam = playoffTeam(1).name
                                                        End If
                                                    Case 2, 3, 4
                                                        If playoffTeam(1).rsWins > playoffTeam(2).rsWins Then
                                                            homeTeam = playoffTeam(2).name
                                                            visitingTeam = playoffTeam(1).name
                                                        Else
                                                            homeTeam = playoffTeam(1).name
                                                            visitingTeam = playoffTeam(2).name
                                                        End If
                                                End Select
                                                gstrPostSeason = conNLCS
                                                bolAmericanLeagueRules = False
                                                _game = gamesPlayed + 1
                                            Else
                                                sqlQuery = "Select * FROM PSStandings WHERE WSWins = 4"
                                                dsStandings.Clear()
                                                dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                                If dsStandings.Tables(0).Rows.Count = 0 Then
                                                    'In World Series
                                                    sqlQuery = "Select * FROM PSStandings WHERE CSWins = 4"
                                                    dsStandings.Clear()
                                                    dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                                                    teamIndex = 1
                                                    For Each dr As DataRow In dsStandings.Tables(0).Rows
                                                        playoffTeam(teamIndex).name = dr.Item("TeamID").ToString
                                                        playoffTeam(teamIndex).league = dr.Item("League").ToString
                                                        playoffTeam(teamIndex).rsWins = CInt(dr.Item("rswins"))
                                                        gamesPlayed = gamesPlayed + CInt(dr.Item("WSWins"))
                                                        teamIndex += 1
                                                    Next dr
                                                    If Val(gstrSeason) > 2016 Then
                                                        'After 2016, best record is home
                                                        If playoffTeam(1).rsWins > playoffTeam(2).rsWins Then
                                                            isALHome = playoffTeam(1).league = "AL"
                                                        Else
                                                            isALHome = playoffTeam(2).league = "AL"
                                                        End If
                                                    Else
                                                        'Up until 2016, AL Team has home advantage in even years
                                                        isALHome = Val(gstrSeason) Mod 2 = 0
                                                    End If

                                                    dsStandings.Clear()
                                                    Select Case gamesPlayed
                                                        Case 0, 1, 5, 6
                                                            If isALHome Then
                                                                bolAmericanLeagueRules = True
                                                                If playoffTeam(1).league = "AL" Then
                                                                    homeTeam = playoffTeam(1).name
                                                                    visitingTeam = playoffTeam(2).name
                                                                Else
                                                                    homeTeam = playoffTeam(2).name
                                                                    visitingTeam = playoffTeam(1).name
                                                                End If
                                                            Else
                                                                bolAmericanLeagueRules = False
                                                                If playoffTeam(1).league = "AL" Then
                                                                    homeTeam = playoffTeam(2).name
                                                                    visitingTeam = playoffTeam(1).name
                                                                Else
                                                                    homeTeam = playoffTeam(1).name
                                                                    visitingTeam = playoffTeam(2).name
                                                                End If
                                                            End If
                                                        Case 2, 3, 4
                                                            If isALHome Then
                                                                bolAmericanLeagueRules = False
                                                                If playoffTeam(1).league = "AL" Then
                                                                    homeTeam = playoffTeam(2).name
                                                                    visitingTeam = playoffTeam(1).name
                                                                Else
                                                                    homeTeam = playoffTeam(1).name
                                                                    visitingTeam = playoffTeam(2).name
                                                                End If
                                                            Else
                                                                bolAmericanLeagueRules = True
                                                                If playoffTeam(1).league = "AL" Then
                                                                    homeTeam = playoffTeam(1).name
                                                                    visitingTeam = playoffTeam(2).name
                                                                Else
                                                                    homeTeam = playoffTeam(2).name
                                                                    visitingTeam = playoffTeam(1).name
                                                                End If
                                                            End If
                                                    End Select
                                                    gstrPostSeason = conWS
                                                    _game = gamesPlayed + 1
                                                Else
                                                    'Season is over
                                                    With dsStandings.Tables(0).Rows(0)
                                                        Call MsgBox("Season is over. " & .Item("TeamID").ToString & " is the World Champion.")
                                                    End With
                                                    gbolSeasonOver = True
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Case Is < 1969
                    numberOfTeams = 2
                    bolAmericanLeagueRules = False
                    sqlQuery = "Select Division, League, CSWins, TeamID " & "FROM PSStandings"
                    dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                    If dsStandings.Tables(0).Rows.Count = 0 Then
                        'No record, new playoffs
                        gbolNewPS = True
                        sqlQuery = "Select Wins, PlayerID, Division From Standings WHERE " & "Division = 'AL' ORDER BY WINS DESC LIMIT 1"
                        dsStandings.Clear()
                        dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                        With dsStandings.Tables(0).Rows(0)
                            playoffTeam(1).name = .Item("PlayerID").ToString
                            playoffTeam(1).division = .Item("Division").ToString
                            playoffTeam(1).rsWins = CInt(.Item("Wins"))
                        End With
                        sqlQuery = "Select Wins, PlayerID, Division From Standings WHERE " & "Division = 'NL' ORDER BY WINS DESC LIMIT 1"
                        dsStandings.Clear()
                        dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                        With dsStandings.Tables(0).Rows(0)
                            playoffTeam(2).name = .Item("PlayerID").ToString
                            playoffTeam(2).division = .Item("Division").ToString
                            playoffTeam(2).rsWins = CInt(.Item("Wins"))
                        End With
                        dsStandings.Clear()
                        If CInt(gstrSeason) Mod 2 > 0 Then
                            homeTeam = playoffTeam(1).name
                            visitingTeam = playoffTeam(2).name
                        Else
                            homeTeam = playoffTeam(2).name
                            visitingTeam = playoffTeam(1).name
                        End If
                        For i As Integer = 1 To 2
                            sqlFields = "TEAMID,RSWINS,DSWINS,CSWINS,WSWINS,DIVISION,LEAGUE"
                            sqlValues = "'" & playoffTeam(i).name & "'," & playoffTeam(i).rsWins & ",0,0,0,'" & playoffTeam(i).division & "','" & playoffTeam(i).division & "'"
                            sqlQuery = "INSERT INTO PSSTANDINGS (" & sqlFields & ") VALUES (" & sqlValues & ")"
                            DataAccess.ExecuteScalar(sqlQuery)
                        Next i
                        gstrPostSeason = conWS
                        _game = 1
                    Else
                        sqlQuery = "Select * FROM PSStandings WHERE WSWins = 4"
                        dsStandings.Clear()
                        dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                        If dsStandings.Tables(0).Rows.Count = 0 Then
                            'In World Series
                            sqlQuery = "Select * FROM PSStandings"
                            dsStandings.Clear()
                            dsStandings = DataAccess.ExecuteDataSet(sqlQuery)
                            teamIndex = 1
                            For Each dr As DataRow In dsStandings.Tables(0).Rows
                                playoffTeam(teamIndex).name = dr.Item("TeamID").ToString
                                playoffTeam(teamIndex).league = dr.Item("League").ToString
                                gamesPlayed = gamesPlayed + CInt(dr.Item("WSWins"))
                                teamIndex += 1
                            Next dr
                            dsStandings.Clear()
                            Select Case gamesPlayed
                                Case 0, 1, 5, 6
                                    If CInt(gstrSeason) Mod 2 > 0 Then
                                        If playoffTeam(1).league = "AL" Then
                                            homeTeam = playoffTeam(1).name
                                            visitingTeam = playoffTeam(2).name
                                        Else
                                            homeTeam = playoffTeam(2).name
                                            visitingTeam = playoffTeam(1).name
                                        End If
                                    Else
                                        If playoffTeam(1).division = "AL" Then
                                            homeTeam = playoffTeam(2).name
                                            visitingTeam = playoffTeam(1).name
                                        Else
                                            homeTeam = playoffTeam(1).name
                                            visitingTeam = playoffTeam(2).name
                                        End If
                                    End If
                                Case 2, 3, 4
                                    If CInt(gstrSeason) Mod 2 > 0 Then
                                        If playoffTeam(1).league = "AL" Then
                                            homeTeam = playoffTeam(2).name
                                            visitingTeam = playoffTeam(1).name
                                        Else
                                            homeTeam = playoffTeam(1).name
                                            visitingTeam = playoffTeam(2).name
                                        End If
                                    Else
                                        If playoffTeam(1).league = "AL" Then
                                            homeTeam = playoffTeam(1).name
                                            visitingTeam = playoffTeam(2).name
                                        Else
                                            homeTeam = playoffTeam(2).name
                                            visitingTeam = playoffTeam(1).name
                                        End If
                                    End If
                            End Select
                            gstrPostSeason = conWS
                            _game = gamesPlayed + 1
                        Else
                            'Season is over
                            With dsStandings.Tables(0).Rows(0)
                                Call MsgBox("Season is over. " & .Item("TeamID").ToString & _
                                                " is the World Champion.")
                            End With
                            gbolSeasonOver = True
                            homeTeam = ""
                            visitingTeam = ""
                        End If
                    End If
            End Select

            'Set table names

            If gbolPostSeason Then
                gstrHittingTable = "PSHITTINGSTATS"
                gstrPitchingTable = "PSPITCHINGSTATS"
                gstrStandingsTable = "PSSTANDINGS"
            Else
                gstrHittingTable = "HITTINGSTATS"
                gstrPitchingTable = "PITCHINGSTATS"
                gstrStandingsTable = "STANDINGS"
            End If
            'End Using
        Catch ex As Exception
            Call MsgBox("Error in DeterminePSTeams " & ex.ToString, MsgBoxStyle.OkOnly)
        Finally
            If Not dsStandings Is Nothing Then dsStandings.Dispose()
        End Try
    End Sub
End Class