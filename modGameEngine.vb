Option Strict On
Option Explicit On

Imports System.Configuration.configurationManager

Module modGameEngine
    ''' <summary>
    ''' main method for determining outcome of play
    ''' </summary>
    ''' <param name="baseAdvance"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetResult(ByRef baseAdvance As Boolean, ByRef baseAdvanceFO As Boolean) As String
        Dim FACCard As FACCard
        Dim playResult As String = ""
        Dim hitType As String = ""
        Dim pbValue As Integer
        Dim isError As Boolean
        Dim suppDescription As String = ""
        Dim isLvsR As Boolean
        Dim isPOE As Boolean

        Try
            isPOE = AppSettings("PointsOfEffectiveness").ToString.ToUpper = "ON"
            bolCountRun = False
            bolGiveRBI = False
            bolAnyBS = False
            gbolSOE = False

            Game.GetResultPtr.description = ""
            If bolNormal Then
                If CheckAutoSteal(playResult) Then
                    Return playResult
                Else
                    'resume normal play
                    FACCard = GetFAC("playType", "PB")
                End If
                gbolAutoSteal = False

                If IsNumeric(FACCard.PB) Then
                    'normal play. Not a Z, CD or BD play.
                    pbValue = CInt(Val(FACCard.PB))

                    FACCard = GetFAC("Normal", "Random")

                    'apply lefty/right effect
                    isLvsR = CheckLeftyRighty(playResult, FACCard)
                    If isLvsR Then
                        Call Game.GetResultPtr.ChartLookup(playResult)
                    End If

                    'apply Pitch Around effect
                    If bolPitchAround1 Then
                        pbValue -= 1
                    End If
                    If bolPitchAround2 Then
                        pbValue -= 2
                    End If
                    If pbValue < 2 Then
                        pbValue = 2
                    End If

                    If InRange((Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).pbRange), pbValue) And Not isLvsR And _
                                                    (Game.PTeam.pr > 0 Or isPOE) Then
                        'Refer to pitcher
                        GetResultFromPitcherCard(playResult, FACCard, baseAdvance, suppDescription)
                        CheckExtraBaseHitEffectPitcher(playResult, FACCard, baseAdvance, suppDescription)
                        'Check for high outcome variance
                        If gbolSeason Then
                            CheckExcessiveVariancePitcher(playResult, FACCard, baseAdvance, suppDescription)
                        End If
                    ElseIf Not isLvsR Then
                        GetResultFromHitterCard(playResult, FACCard, baseAdvance, suppDescription)
                        CheckExtraBaseHitEffectBatter(playResult, FACCard, baseAdvance, suppDescription)
                        'Check for high outcome variance
                        If gbolSeason Then
                            CheckExcessiveVarianceBatter(playResult, FACCard, baseAdvance, suppDescription)
                        End If
                    End If


                    If playResult.IndexOf("*") > -1 Then
                        'Check error
                        FACCard = GetFAC("checkError", "Error")
                        isError = CheckForError(playResult, FACCard, baseAdvance)
                    Else
                        'line out or fly out. Set current fielder
                        Game.currentFieldPos = CInt(Val(Right(playResult, 1)))
                    End If

                    playResult = playResult.Trim
                    If playResult.IndexOf(conNoPlay) > -1 Then
                        Return conNoPlay
                    End If

                    If Not isError And Not isLvsR Then
                        'Get normal play result
                        hitType = Left(playResult, 2)
                        If hitType = "1B" Or hitType = "2B" Or hitType = "3B" Then
                            'Determine if base situation yields a base advance opportunity
                            baseAdvance = baseAdvance And ((hitType = "1B" And (FirstBase.occupied Or SecondBase.occupied)) Or _
                                    (hitType = "2B" And FirstBase.occupied))
                            playResult = hitType & BaseSituation()
                            Call Game.GetResultPtr.ChartLookup(playResult)
                            hitType = Left(playResult, 2)
                            If hitType = "1B" Or hitType = "2B" Or hitType = "3B" Then
                                'hit type may have changed during chart lookup
                                Game.GetResultPtr.description += suppDescription
                            End If
                        ElseIf playResult.Substring(playResult.Length - 1) = "*" Then
                            playResult = playResult.Substring(0, playResult.Length - 1) & BaseSituation()
                            Call Game.GetResultPtr.ChartLookup(playResult)
                        Else
                            playResult += BaseSituation()
                            Call Game.GetResultPtr.ChartLookup(playResult)
                        End If

                        If Game.GetResultPtr.resBatter = conDefOptPlay Then
                            HandleDefensiveOptionPlay(FACCard)
                        End If

                        HandleOutExceptions(playResult)

                        If AppSettings("FlyBallAdvance").ToString.ToUpper = "ON" Then
                            If playResult.Substring(0, 1) = "F" And (Game.currentFieldPos >= 7 And Game.currentFieldPos <= 9) And _
                                    Game.outs < 2 Then
                                'Determine if there is an opportunity for an outfield flyout advance. Basically if there is a runner that did 
                                'not move, they can still advance
                                If (ThirdBase.occupied And Game.GetResultPtr.resThird = 3) Or _
                                    (SecondBase.occupied And Game.GetResultPtr.resSecond = 2) Or _
                                    (FirstBase.occupied And Game.GetResultPtr.resFirst = 1) Then
                                    baseAdvanceFO = True
                                End If
                            End If
                        End If
                    End If
                    CheckForOutStretchingHit(playResult, baseAdvance)
                Else
                    'Handle special situations (Z,CD,BD)
                    playResult = FACCard.PB.Trim
                    Select Case playResult
                        Case "BD"
                            HandleBD(playResult, FACCard)
                        Case "CD"
                            HandleCD(playResult, FACCard)
                        Case "Z"
                            HandleZPlay(playResult, FACCard)
                            If playResult.Length >= 4 AndAlso playResult.Substring(0, 4) = "Zfld" Then
                                HandleZPlayFielding(playResult, FACCard)
                            End If

                            If Game.GetResultPtr.description.IndexOf("Result:E") > -1 Then
                                HandleZPlayErrors()
                            End If

                            If Game.GetResultPtr.description.IndexOf(conEjected) > -1 Then
                                If AppSettings("PlayerEjections").ToString.ToUpper = "ON" Then
                                    HandlePlayerEjections()
                                Else
                                    playResult = conNoPlay
                                End If
                            End If

                            If Game.GetResultPtr.description.IndexOf(conInjured) > -1 Then
                                FACCard = GetFAC("injCheck", "Random")
                                If CInt(FACCard.Random) > 28 Then
                                    'Only valid just 25% of the time. Too many in game injuries
                                    playResult = conNoPlay
                                Else
                                    HandlePlayerInjuries()
                                End If
                            End If

                            If Game.GetResultPtr.description.Trim = Nothing Then
                                playResult = conNoPlay
                            End If
                    End Select
                End If
            ElseIf bolHitRun Then
                FACCard = GetFAC("hitRun", "Random")
                HandleHitAndRun(playResult, FACCard)
            ElseIf bolSteal Then
                FACCard = GetFAC("steal", "Random")
                HandleBaseStealAttempt(playResult, FACCard)
            ElseIf bolSac Then
                FACCard = GetFAC("sac", "Random")
                HandleSacrifice(playResult, FACCard)
                'TO DO Handle Squeeze
                'Remember, apply add 10 rule for Corners In for Squeeze
            ElseIf bolIntentionalWalk Then
                playResult = "W" & BaseSituation()
                Call Game.GetResultPtr.ChartLookup(playResult)
            ElseIf bolBunt Then
                FACCard = GetFAC("bunt", "Random")
                HandleBunt(playResult, FACCard)
            ElseIf bolSqueeze Then
                FACCard = GetFAC("squeeze", "Random")
                HandleSqueeze(playResult, FACCard)
            Else
                Call MsgBox("Error, play not selected!", MsgBoxStyle.Exclamation)
            End If
        Catch ex As Exception
            Call MsgBox("GetResult. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return playResult
    End Function

    ''' <summary>
    ''' determine if situation warrants an auto steal
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function AutoStealContext2nd() As Boolean
        Dim runLead As Integer
        Dim hasContext As Boolean
        runLead = Game.BTeam.runs - Game.PTeam.runs
        If FirstBase.occupied Then
            hasContext = FirstBase.occupied And Not SecondBase.occupied And runLead <= 1 And runLead >= -2 And _
                        Game.BTeam.GetBatterPtr(FirstBase.runner).sp <> "E" And Not bolDNS
        End If
        Return hasContext
    End Function

    ''' <summary>
    ''' determine if situation warrants an auto steal
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function AutoStealContext3rd() As Boolean
        Dim runLead As Integer
        Dim hasContext As Boolean
        runLead = Game.BTeam.runs - Game.PTeam.runs
        If SecondBase.occupied Then
            hasContext = SecondBase.occupied And Not ThirdBase.occupied And runLead <= 1 And runLead >= -2 And _
                        "CDE".IndexOf(Game.BTeam.GetBatterPtr(SecondBase.runner).sp) = -1 And Not bolDNS And Game.outs < 2
        End If
        Return hasContext
    End Function

    ''' <summary>
    ''' if a player is injured, determine how many games they are out based on their injury rating
    ''' </summary>
    ''' <param name="rating"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DetermineInjGames(ByRef rating As Integer) As Integer
        Dim FACCard As FACCard
        Dim injuredGames As Integer

        Try
            If AppSettings("Injuries").ToString.ToUpper <> "ON" Then
                Return 0
            End If
            FACCard = GetFAC("injGames", "Random")
            Select Case CInt(rating.ToString)
                Case 0
                    injuredGames = 0
                Case 1
                    injuredGames = 1
                Case 2
                    injuredGames = CInt(FACCard.Random.Substring(0, 1))
                    If injuredGames > 5 Then
                        injuredGames = 5
                    End If
                Case 3
                    injuredGames = CInt(FACCard.Random.Substring(0, 1))
                Case 4
                    injuredGames = CInt(FACCard.Random.Substring(0, 1)) + CInt(FACCard.Random.Substring(FACCard.Random.Length - 1))
                    If injuredGames > 10 Then
                        injuredGames = 10
                    End If
                Case 5
                    injuredGames = CInt(FACCard.Random.Substring(0, 1)) + CInt(FACCard.Random.Substring(FACCard.Random.Length - 1))
                Case 6
                    injuredGames = CInt(FACCard.Random)
                    If injuredGames > 20 Then
                        injuredGames = 20
                    End If
                Case 7
                    injuredGames = CInt(FACCard.Random)
                    If injuredGames > 30 Then
                        injuredGames = 30
                    End If
                Case 8
                    injuredGames = CInt(FACCard.Random)
            End Select
        Catch ex As Exception
            Call MsgBox("DetermineInjGames. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return injuredGames
    End Function

    ''' <summary>
    ''' loads batting data from csv file
    ''' </summary>
    ''' <param name="Team"></param>
    ''' <remarks></remarks>
    Public Sub LoadBattingCard(ByRef Team As clsTeam)
        Dim filInput As Integer
        Dim cardID As String
        Dim battingCardId As String = ""
        Dim fieldLine As String = ""
        Dim obrRating As String = ""
        Dim spRating As String = ""
        Dim hitrunRating As String = ""
        Dim cdRating As String = ""
        Dim sacRating As String = ""
        Dim injRating As String = ""
        Dim hit1bfRating As String = ""
        Dim hit1b7Rating As String = ""
        Dim hit1b8Rating As String = ""
        Dim hit1b9Rating As String = ""
        Dim hit2b7Rating As String = ""
        Dim hit2b8Rating As String = ""
        Dim hit2b9Rating As String = ""
        Dim hit3b8Rating As String = ""
        Dim hrRating As String = ""
        Dim kRating As String = ""
        Dim walkRating As String = ""
        Dim hpbRating As String = ""
        Dim outRating As String = ""
        Dim chtRating As String = ""
        Dim bdRating As String = ""
        Dim tempValue As String
        Dim fileName As String = ""
        Dim useTeamCard As Boolean
        Dim forceALBattingCard As Boolean

        Try
            forceALBattingCard = gstrPostSeason = conWS AndAlso Team.teamName = Visitor.teamName AndAlso _
                            Not bolAmericanLeagueRules AndAlso CInt(gstrSeason) < 1997 AndAlso _
                            CInt(gstrSeason) >= 1973
            useTeamCard = AppSettings("UseTeamBattingCardForPitchers").ToString.ToUpper = "ON" And Not forceALBattingCard
            If useTeamCard Then
                fileName = gAppPath & "teams\" & gstrSeason & "\teampitchbat.csv"
            Else
                fileName = gAppPath & "pitchbat.csv"
            End If
            With Team.GetBatterPtr(Team.hitters + Team.nthPitcherUsed)
                If Team.pitcherSel = 0 Then
                    'Initialize to first pitcher
                    .player = "Pitcher"
                    .position = "P"
                    .cht = "P"
                Else
                    If useTeamCard Then
                        cardID = Team.id
                    Else
                        cardID = Team.GetPitcherPtr(Team.pitcherSel).batCard
                    End If

                    filInput = FreeFile()
                    FileOpen(filInput, fileName, OpenMode.Input)
                    tempValue = LineInput(filInput) 'Clear Header
                    While Not EOF(filInput)
                        Input(filInput, battingCardId)
                        Input(filInput, fieldLine)
                        Input(filInput, obrRating)
                        Input(filInput, spRating)
                        Input(filInput, hitrunRating)
                        Input(filInput, cdRating)
                        Input(filInput, sacRating)
                        Input(filInput, injRating)
                        Input(filInput, hit1bfRating)
                        Input(filInput, hit1b7Rating)
                        Input(filInput, hit1b8Rating)
                        Input(filInput, hit1b9Rating)
                        Input(filInput, hit2b7Rating)
                        Input(filInput, hit2b8Rating)
                        Input(filInput, hit2b9Rating)
                        Input(filInput, hit3b8Rating)
                        Input(filInput, hrRating)
                        Input(filInput, kRating)
                        Input(filInput, walkRating)
                        Input(filInput, hpbRating)
                        Input(filInput, outRating)
                        Input(filInput, chtRating)
                        Input(filInput, bdRating)
                        Input(filInput, "")
                        Input(filInput, "")
                        Input(filInput, "")

                        If battingCardId = cardID Then
                            .position = "P"
                            .player = Team.GetPitcherPtr(Team.pitcherSel).player
                            .field = Team.GetPitcherPtr(Team.pitcherSel).field
                            .obr = obrRating
                            .sp = spRating
                            .hitRun = hitrunRating
                            .cd = GetCD((Team.GetPitcherPtr(Team.pitcherSel).field))
                            .sac = sacRating
                            .inj = injRating
                            .hit1Bf = hit1bfRating
                            .hit1B7 = hit1b7Rating
                            .hit1B8 = hit1b8Rating
                            .hit1B9 = hit1b9Rating
                            .hit2B7 = hit2b7Rating
                            .hit2B8 = hit2b8Rating
                            .hit2B9 = hit2b9Rating
                            .hit3B8 = hit3b8Rating
                            .hr = hrRating
                            .k = kRating
                            .w = walkRating
                            .hpb = hpbRating
                            .out = outRating
                            .cht = chtRating
                            .bd = bdRating
                            .playerIndex = Team.pitcherSel
                        End If
                    End While
                    FileClose(filInput)
                End If
            End With
        Catch ex As Exception
            Call MsgBox("LoadBattingCard. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' load player data from team .csv files
    ''' </summary>
    ''' <param name="Team"></param>
    ''' <remarks></remarks>
    Public Sub LoadData(ByRef Team As clsTeam)
        Dim fileName As String = ""
        Dim filFileNum As Integer
        Dim batterIndex As Integer
        Dim pitcherIndex As Integer
        Dim tempValue As String
        Dim playerName As String = ""
        Dim fieldLine As String = ""
        Dim obrRating As String = ""
        Dim spRating As String = ""
        Dim hitrunRating As String = ""
        Dim cdRating As String = ""
        Dim sacRating As String = ""
        Dim injRating As String = ""
        Dim hit1bfRating As String = ""
        Dim hit1b7Rating As String = ""
        Dim hit1b8Rating As String = ""
        Dim hit1b9Rating As String = ""
        Dim hit2b7Rating As String = ""
        Dim hit2b8Rating As String = ""
        Dim hit2b9Rating As String = ""
        Dim hit3b8Rating As String = ""
        Dim hrRating As String = ""
        Dim strikeoutRating As String = ""
        Dim walkRating As String = ""
        Dim hpbRating As String = ""
        Dim outRating As String = ""
        Dim chtRating As String = ""
        Dim bdRating As String = ""
        Dim city As String = ""
        Dim state As String = ""
        Dim dob As String = ""
        'Pitching data
        Dim pitcherPBRating As String = ""
        Dim startRating As String = ""
        Dim reliefRating As String = ""
        Dim bkRating As String = ""
        Dim pbRating As String = ""
        Dim wpRating As String = ""
        Dim startReliefLine As String = ""
        Dim throwingArm As String = ""
        Dim battingCard As String = ""
        Dim fullTeamName As String = ""
        Dim ds As DataSet = Nothing
        Dim dr As DataRow = Nothing
        Dim sqlQuery As String = ""
        Dim tableKey As String
        Dim filDebug As Integer
        Dim DataAccess As New clsDataAccess(gstrSeason & "data")

        Try
            If Dir(gAppPath & "debug.bbb") <> Nothing Then
                filDebug = FreeFile()
                FileOpen(filDebug, gAppPath & "debug.out", OpenMode.Append)
                PrintLine(filDebug, "In LoadData")
                FileClose(filDebug)
            End If

            batterIndex = 1
            fileName = GetShortTeamTranslation(Team.teamName) & "bat.csv"
            filFileNum = FreeFile()
            'If gbolSeason Then
            FileOpen(filFileNum, gAppPath & "teams\" & gstrSeason & "\" & Team.teamName & "\" & fileName, OpenMode.Input)
            'Else
            'FileOpen(filFileNum, gAppPath & "teams\1985\" & Team.teamName & "\" & fileName, OpenMode.Input)
            'End If
            tempValue = LineInput(filFileNum) 'Clear Header
            'If gbolSeason Then
            '    If gAccessConnectStr = Nothing Then
            '        gAccessConnectStr = GetConnectString()
            '    End If
            'End If
            'Using accessConn As OleDb.OleDbConnection = New OleDb.OleDbConnection
            '    accessConn.ConnectionString = GetConnectString()
            '    accessConn.Open()
            While Not EOF(filFileNum)
                Input(filFileNum, playerName)
                Input(filFileNum, fieldLine)
                Input(filFileNum, obrRating)
                Input(filFileNum, spRating)
                Input(filFileNum, hitrunRating)
                Input(filFileNum, cdRating)
                Input(filFileNum, sacRating)
                Input(filFileNum, injRating)
                Input(filFileNum, hit1bfRating)
                Input(filFileNum, hit1b7Rating)
                Input(filFileNum, hit1b8Rating)
                Input(filFileNum, hit1b9Rating)
                Input(filFileNum, hit2b7Rating)
                Input(filFileNum, hit2b8Rating)
                Input(filFileNum, hit2b9Rating)
                Input(filFileNum, hit3b8Rating)
                Input(filFileNum, hrRating)
                Input(filFileNum, strikeoutRating)
                Input(filFileNum, walkRating)
                Input(filFileNum, hpbRating)
                Input(filFileNum, outRating)
                Input(filFileNum, chtRating)
                Input(filFileNum, bdRating)
                Input(filFileNum, city)
                Input(filFileNum, state)
                Input(filFileNum, dob)
                If playerName.Substring(playerName.Length - 1) <> "*" Then '* denotes inactive
                    With Team.GetBatterPtr(batterIndex)
                        .player = playerName
                        .field = fieldLine
                        .obr = obrRating
                        .sp = spRating
                        .hitRun = hitrunRating
                        .cd = cdRating
                        .sac = sacRating
                        .inj = injRating
                        .hit1Bf = hit1bfRating
                        .hit1B7 = hit1b7Rating
                        .hit1B8 = hit1b8Rating
                        .hit1B9 = hit1b9Rating
                        .hit2B7 = hit2b7Rating
                        .hit2B8 = hit2b8Rating
                        .hit2B9 = hit2b9Rating
                        .hit3B8 = hit3b8Rating
                        .hr = hrRating
                        .k = strikeoutRating
                        .w = walkRating
                        .hpb = hpbRating
                        .out = outRating
                        .cht = chtRating
                        .bd = bdRating
                        .city = city
                        .state = state
                        .age = Team.GetPlayerAge(dob)
                        If gbolSeason Then
                            tableKey = HandleQuotes(StripChar(.player & Team.teamName, " "))
                            sqlQuery = "SELECT * FROM " & gstrHittingTable & " WHERE playerid = '" & tableKey & "'"
                            ds = DataAccess.ExecuteDataSet(sqlQuery)
                            If ds.Tables(0).Rows.Count > 0 Then
                                dr = ds.Tables(0).Rows(0)
                                .BatStatPtr.GamesInj = CInt(dr.Item("Inj").ToString)
                                .BatStatPtr.Active = CBool(dr.Item("Active"))
                                .playedLast = CBool(dr.Item("PlayedLast"))
                                .hitLast = CBool(dr.Item("HitLast"))
                                If Not .BatStatPtr.Active Then
                                    Team.inactiveHitters += 1
                                End If
                            End If
                            ds.Clear()
                        End If
                        .playerIndex = batterIndex
                    End With
                    Team.hitters = batterIndex
                    batterIndex += 1
                End If
            End While
            FileClose(filFileNum)

            'Load Pitching
            pitcherIndex = 0
            fileName = GetShortTeamTranslation(Team.teamName) & "pitch.csv"
            Team.id = fileName.Substring(0, 3)
            filFileNum = FreeFile()
            'If gbolSeason Then
            FileOpen(filFileNum, gAppPath & "teams\" & gstrSeason & "\" & Team.teamName & "\" & fileName, OpenMode.Input)
            'Else
            'FileOpen(filFileNum, gAppPath & "teams\1985\" & Team.teamName & "\" & fileName, OpenMode.Input)
            'End If
            tempValue = LineInput(filFileNum) 'Clear Header
            While Not EOF(filFileNum)
                Input(filFileNum, playerName)
                Input(filFileNum, fieldLine)
                Input(filFileNum, pitcherPBRating)
                Input(filFileNum, startRating)
                Input(filFileNum, reliefRating)
                Input(filFileNum, hit1bfRating)
                Input(filFileNum, hit1b7Rating)
                Input(filFileNum, hit1b8Rating)
                Input(filFileNum, hit1b9Rating)
                Input(filFileNum, bkRating)
                Input(filFileNum, strikeoutRating)
                Input(filFileNum, walkRating)
                Input(filFileNum, pbRating)
                Input(filFileNum, wpRating)
                Input(filFileNum, outRating)
                Input(filFileNum, startReliefLine)
                Input(filFileNum, throwingArm)
                Input(filFileNum, battingCard)
                Input(filFileNum, fullTeamName)
                Input(filFileNum, city)
                Input(filFileNum, state)
                Input(filFileNum, dob)
                If playerName.Substring(playerName.Length - 1) <> "*" Then '* denotes inactive
                    pitcherIndex += 1
                    With Team.GetPitcherPtr(pitcherIndex)
                        .player = playerName
                        .field = fieldLine
                        .throwField = throwingArm
                        .pbRange = pitcherPBRating
                        .sr = startRating
                        .rr = reliefRating
                        .hit1Bf = hit1bfRating
                        .hit1B7 = hit1b7Rating
                        .hit1B8 = hit1b8Rating
                        .hit1B9 = hit1b9Rating
                        .bk = bkRating
                        .k = strikeoutRating
                        .w = walkRating
                        .pb = pbRating
                        .wp = wpRating
                        .out = outRating
                        .StoR = startReliefLine
                        .pitcherNum = pitcherIndex
                        .batCard = battingCard
                        .city = city
                        .state = state
                        .age = Team.GetPlayerAge(dob)
                        If gbolSeason Then
                            tableKey = HandleQuotes(StripChar(.player & Team.teamName, " "))
                            sqlQuery = "SELECT * FROM " & gstrPitchingTable & " WHERE playerid = '" & tableKey & "'"
                            ds = DataAccess.ExecuteDataSet(sqlQuery)
                            If ds.Tables(0).Rows.Count > 0 Then
                                dr = ds.Tables(0).Rows(0)
                                If Not CBool(dr.Item("Inactive")) And CBool(dr.Item("Active")) Then
                                    .available = True
                                    .PitStatPtr.inactive = False
                                Else
                                    .available = False
                                    .PitStatPtr.inactive = True
                                    Team.inactivePitchers += 1
                                End If
                            End If
                            ds.Clear()
                        End If
                    End With
                    Team.pitchers = pitcherIndex
                End If
            End While
            FileClose(filFileNum)
            'End Using
            'load statistics
        Catch ex As Exception
            Call MsgBox("Error in LoadData. " & sqlQuery & " " & ex.ToString, MsgBoxStyle.OkOnly)
            Call MsgBox("File = " & fileName & ". Team = " & Team.teamName & ". Season = " & gstrSeason, MsgBoxStyle.OkOnly)
        Finally
            If Not ds Is Nothing Then ds.Dispose()
        End Try
    End Sub

    Public Sub NewGame()
        bolStartGame = True
        Call Game.SwitchTeams(True)
    End Sub

    ''' <summary>
    ''' updates the game objects between plays
    ''' </summary>
    ''' <param name="switchTeams"></param>
    ''' <remarks></remarks>
    Public Sub UpdateGame(ByRef switchTeams As Boolean)
        Dim numOuts As Integer
        Dim playResult As String
        Dim numRuns As Integer
        Dim numOutsOld As Integer
        Dim numErrors As Integer
        Dim tempErrors As Integer
        Dim numHits As Integer
        Dim isFC As Boolean
        Dim arrPitchers(3) As Integer
        Dim arrUnearned(3) As Boolean
        Dim runnerIndex As Integer
        Dim isPOE As Boolean

        Try
            isPOE = AppSettings("PointsOfEffectiveness").ToString.ToUpper = "ON"

            'Calculate outs, IPs
            playResult = Game.GetResultPtr.resBatter.ToString
            playResult += Game.GetResultPtr.resFirst.ToString
            playResult += Game.GetResultPtr.resSecond.ToString
            playResult += Game.GetResultPtr.resThird.ToString
            numOutsOld = Game.outs
            For i As Integer = 0 To playResult.Length - 1
                If playResult.Substring(i, 1) = "0" Then
                    Game.outs += 1
                End If
            Next i
            If Game.outs > 3 Then Game.outs = 3
            Game.GetResultPtr.ip = Game.outs - numOutsOld
            'Count errors
            For i As Integer = 1 To Game.PTeam.hitters
                numErrors += Game.PTeam.GetBatterPtr(i).BatStatPtr.E
            Next i
            For i As Integer = 1 To Game.PTeam.pitchers
                numErrors += Game.PTeam.GetPitcherPtr(i).PitStatPtr.e
            Next i
            'Count hits
            numHits = Game.GetResultPtr.hit
            'Update numRuns and Update description for those who scored
            If Game.outs < 3 Or bolCountRun Then
                For i As Integer = playResult.Length - 1 To 0 Step -1
                    If playResult.Substring(i, 1) = "4" Then
                        numRuns += 1
                        Select Case i
                            Case 0
                                Game.BTeam.GetBatterPtr(Game.currentBatter).BatStatPtr.R += 1
                                With Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel)
                                    .PitStatPtr.r += 1
                                    If Not gbolSOE And ((.eOuts + Game.outs) < 3 Or (.eOuts = 0 And bolCountRun)) Then
                                        .PitStatPtr.er += 1
                                    End If
                                End With
                                If Game.PTeam.pr > 0 And Not isPOE Then
                                    'reduce PR by 1
                                    Game.PTeam.pr = Game.PTeam.pr - 1
                                End If
                                Call MarkWinLoss(Game.PTeam.pitcherSel, numRuns)
                            Case 1
                                Game.BTeam.GetBatterPtr(FirstBase.runner).BatStatPtr.R += 1
                                With Game.PTeam.GetPitcherPtr(FirstBase.pitcher)
                                    .PitStatPtr.r += 1
                                    If Not FirstBase.unearned And ((.eOuts + Game.outs) < 3 Or (.eOuts = 0 And bolCountRun)) Then
                                        .PitStatPtr.er += 1
                                    End If
                                End With
                                If Game.PTeam.pr > 0 And FirstBase.pitcher = Game.PTeam.pitcherSel And Not isPOE Then
                                    'reduce PR by 1
                                    Game.PTeam.pr = Game.PTeam.pr - 1
                                End If
                                gblUpdateMsg = gblUpdateMsg & Space(1) & Game.BTeam.GetBatterPtr(FirstBase.runner).player & " will score."
                                Call MarkWinLoss((FirstBase.pitcher), numRuns)
                            Case 2
                                Game.BTeam.GetBatterPtr(SecondBase.runner).BatStatPtr.R += 1
                                With Game.PTeam.GetPitcherPtr(SecondBase.pitcher)
                                    .PitStatPtr.r += 1
                                    If Not SecondBase.unearned And ((.eOuts + Game.outs) < 3 Or (.eOuts = 0 And bolCountRun)) Then
                                        .PitStatPtr.er += 1
                                    End If
                                End With
                                If Game.PTeam.pr > 0 And SecondBase.pitcher = Game.PTeam.pitcherSel And Not isPOE Then
                                    'reduce PR by 1
                                    Game.PTeam.pr = Game.PTeam.pr - 1
                                End If
                                gblUpdateMsg = gblUpdateMsg & Space(1) & Game.BTeam.GetBatterPtr(SecondBase.runner).player & " will score."
                                Call MarkWinLoss((SecondBase.pitcher), numRuns)
                            Case 3
                                Game.BTeam.GetBatterPtr(ThirdBase.runner).BatStatPtr.R += 1
                                With Game.PTeam.GetPitcherPtr(ThirdBase.pitcher)
                                    .PitStatPtr.r += 1
                                    If Not ThirdBase.unearned And ((.eOuts + Game.outs) < 3 Or (.eOuts = 0 And bolCountRun)) Then
                                        .PitStatPtr.er += 1
                                    End If
                                End With
                                If Game.PTeam.pr > 0 And ThirdBase.pitcher = Game.PTeam.pitcherSel And Not isPOE Then
                                    'reduce PR by 1
                                    Game.PTeam.pr = Game.PTeam.pr - 1
                                End If
                                gblUpdateMsg = gblUpdateMsg & Space(1) & Game.BTeam.GetBatterPtr(ThirdBase.runner).player & " will score."
                                Call MarkWinLoss((ThirdBase.pitcher), numRuns)
                        End Select
                    End If
                Next i
            End If
            'Update remaining stats
            Game.UpdateStats()
            Game.BTeam.runs += numRuns
            Game.BTeam.hits += numHits
            tempErrors = numErrors - Game.PTeam.errors 'New errors
            Game.PTeam.errors = numErrors
            Call Game.UpdateLineScore(Game.BTeam.runs)
            If Game.GetResultPtr.resBatter <> 5 Then
                Game.BTeam.UpdateLineUp()
            End If
            'Update PR
            With Game.GetResultPtr
                If Not bolIntentionalWalk Then
                    If isPOE Then
                        DeterminePrFromPOE()
                    Else
                        Game.PTeam.pr = Game.PTeam.pr - .hit - .walk - .hpb - .wp - tempErrors
                    End If
                End If
            End With
            If Game.PTeam.pr < 0 Then Game.PTeam.pr = 0
            'Update bases
            If Not bolAnyBS Then
                Game.lastBatterHome = Game.homeTeamBatting
                If Game.outs = 3 Or Game.GameOver Then
                    switchTeams = True
                    FirstBase.Clear()
                    SecondBase.Clear()
                    ThirdBase.Clear()
                    'Switch sides or game over
                    bolFinal = Game.SwitchTeams(False) Or Game.GameOver
                Else
                    With Game.GetResultPtr
                        If .resBatter = 0 Then numOuts += 1
                        If .resFirst = 0 Then numOuts += 1
                        If .resSecond = 0 Then numOuts += 1
                        If .resThird = 0 Then numOuts += 1
                        'Was there a fielder's choice?
                        isFC = .hit = 0 And numOuts = 1 And .resBatter = 1
                    End With
                    If isFC Then
                        'Runner responsibility does not change!
                        If ThirdBase.occupied Then
                            runnerIndex = runnerIndex + 1
                            arrPitchers(runnerIndex) = ThirdBase.pitcher
                            arrUnearned(runnerIndex) = ThirdBase.unearned
                        End If
                        If SecondBase.occupied Then
                            runnerIndex += 1
                            arrPitchers(runnerIndex) = SecondBase.pitcher
                            arrUnearned(runnerIndex) = SecondBase.unearned
                        End If
                        If FirstBase.occupied Then
                            runnerIndex += 1
                            arrPitchers(runnerIndex) = FirstBase.pitcher
                            arrUnearned(runnerIndex) = FirstBase.unearned
                        End If
                    End If
                    If ThirdBase.occupied Then
                        If Not isFC And Game.GetResultPtr.resThird = 0 And ThirdBase.unearned Then
                            'Unearned runner is wiped out with an out
                            With Game.PTeam.GetPitcherPtr(ThirdBase.pitcher)
                                .eOuts = .eOuts - 1
                            End With
                        End If
                    End If
                    If SecondBase.occupied Then
                        If Not isFC And Game.GetResultPtr.resSecond = 0 And SecondBase.unearned Then
                            'Unearned runner is wiped out with an out
                            With Game.PTeam.GetPitcherPtr(SecondBase.pitcher)
                                .eOuts = .eOuts - 1
                            End With
                        End If
                        If Game.GetResultPtr.resSecond = 3 Then
                            ThirdBase.runner = SecondBase.runner
                            ThirdBase.pitcher = SecondBase.pitcher
                            ThirdBase.unearned = SecondBase.unearned
                        End If
                    End If
                    If FirstBase.occupied Then
                        If Not isFC And Game.GetResultPtr.resFirst = 0 And FirstBase.unearned Then
                            'Unearned runner is wiped out with an out
                            With Game.PTeam.GetPitcherPtr(FirstBase.pitcher)
                                .eOuts = .eOuts - 1
                            End With
                        End If
                        Select Case Game.GetResultPtr.resFirst
                            Case 2
                                SecondBase.runner = FirstBase.runner
                                SecondBase.pitcher = FirstBase.pitcher
                                SecondBase.unearned = FirstBase.unearned
                            Case 3
                                ThirdBase.runner = FirstBase.runner
                                ThirdBase.pitcher = FirstBase.pitcher
                                ThirdBase.unearned = FirstBase.unearned
                        End Select
                    End If
                    Select Case Game.GetResultPtr.resBatter
                        Case 1
                            FirstBase.runner = Game.currentBatter
                            FirstBase.pitcher = Game.PTeam.pitcherSel
                            FirstBase.unearned = gbolSOE
                        Case 2
                            SecondBase.runner = Game.currentBatter
                            SecondBase.pitcher = Game.PTeam.pitcherSel
                            SecondBase.unearned = gbolSOE
                        Case 3
                            ThirdBase.runner = Game.currentBatter
                            ThirdBase.pitcher = Game.PTeam.pitcherSel
                            ThirdBase.unearned = gbolSOE
                    End Select
                    'update frmmain
                    FirstBase.occupied = playResult.IndexOf("1") > -1
                    SecondBase.occupied = playResult.IndexOf("2") > -1
                    ThirdBase.occupied = playResult.IndexOf("3") > -1
                    If isFC Then
                        'restore old runner assignments
                        runnerIndex = 0
                        If ThirdBase.occupied Then
                            runnerIndex += 1
                            ThirdBase.pitcher = arrPitchers(runnerIndex)
                            ThirdBase.unearned = arrUnearned(runnerIndex)
                        End If
                        If SecondBase.occupied Then
                            runnerIndex += 1
                            SecondBase.pitcher = arrPitchers(runnerIndex)
                            SecondBase.unearned = arrUnearned(runnerIndex)
                        End If
                        If FirstBase.occupied Then
                            runnerIndex += 1
                            FirstBase.pitcher = arrPitchers(runnerIndex)
                            FirstBase.unearned = arrUnearned(runnerIndex)
                        End If
                    End If
                End If
            End If
            'Update batter
            Game.currentBatter = Game.BTeam.GetPlayerNum(Game.BTeam.order)
            'Update team objects
            If Game.homeTeamBatting Then
                Home = Game.BTeam
                Visitor = Game.PTeam
            Else
                Home = Game.PTeam
                Visitor = Game.BTeam
            End If
        Catch ex As Exception
            Call MsgBox("Error in UpdateGame. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub DeterminePrFromPOE()
        Dim poe As Integer
        Dim pbHigh As Integer
        Dim pbRange As String = Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).pbRange

        With Game.GetResultPtr
            If .hit + .walk + .hpb > 0 Then
                poe = .hit + .doubl + .triple + .hr * 2 + .walk + .hpb + Game.PTeam.consecutivePOE
                Game.PTeam.consecutivePOE += 1
                'determine effect on PB Range
                If poe > Game.PTeam.pr Then
                    'reduce the PB by the number of points in which POE exceeds PR
                    pbHigh = CInt(pbRange.Substring(2))
                    pbHigh -= (poe - Game.PTeam.pr)
                    If pbHigh < 2 Then
                        pbHigh = 2
                    End If
                    pbRange = "2-" & pbHigh.ToString
                    If Game.PTeam.pr < 0 Then
                        Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).pbRange = "2-" & pbHigh.ToString
                    End If
                    Game.PTeam.pr = 0
                Else
                    Game.PTeam.pr -= poe
                End If
            ElseIf .resBatter <> 5 Then
                Game.PTeam.consecutivePOE = 0
            End If
        End With
    End Sub

    ''' <summary>
    ''' determines CD (clutch defense) rating of fielder
    ''' </summary>
    ''' <param name="fieldLine"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetCD(ByRef fieldLine As String) As String
        Dim cdRating As String = ""
        Dim index As Integer
        index = fieldLine.IndexOf("CD")
        If index > -1 Then
            cdRating = fieldLine.Substring(index + 2, 1)
        Else
            Call MsgBox("Missing CD number on pitcher!", MsgBoxStyle.Critical)
        End If
        Return cdRating
    End Function

    ''' <summary>
    ''' determines number of outs in inning
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetOuts() As Integer
        Dim playResult As String
        Dim totalOuts As Integer

        Try
            'Calculate outs, IPs
            playResult = Game.GetResultPtr.resBatter.ToString
            playResult += Game.GetResultPtr.resFirst.ToString
            playResult += Game.GetResultPtr.resSecond.ToString
            playResult += Game.GetResultPtr.resThird.ToString
            For i As Integer = 0 To playResult.Length - 1
                If playResult.Substring(i, 1) = "0" Then
                    totalOuts += 1
                End If
            Next i
        Catch ex As Exception
            Call MsgBox("Error in GetOuts. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return totalOuts
    End Function

    ''' <summary>
    ''' fetch results and set the Game object with them
    ''' </summary>
    ''' <param name="tmpResultptr"></param>
    ''' <remarks></remarks>
    Private Sub InsertStats(ByRef tmpResultptr As clsResult)
        With Game.GetResultPtr
            .ab = tmpResultptr.ab
            .hit = tmpResultptr.hit
            .doubl = tmpResultptr.doubl
            .triple = tmpResultptr.triple
            .hr = tmpResultptr.hr
            .rbi = tmpResultptr.rbi
            .strikeOut = tmpResultptr.strikeOut
            .walk = tmpResultptr.walk
            .hpb = tmpResultptr.hpb
            .wp = tmpResultptr.wp
            .bk = tmpResultptr.bk
        End With
    End Sub

    ''' <summary>
    ''' sets the pitcher(s) of record
    ''' </summary>
    ''' <param name="pitcherIndex"></param>
    ''' <param name="numRuns"></param>
    ''' <remarks></remarks>
    Private Sub MarkWinLoss(ByRef pitcherIndex As Integer, ByRef numRuns As Integer)
        Dim isTie As Boolean

        'Determine original situation
        isTie = Game.BTeam.runs + numRuns - 1 = Game.PTeam.runs
        If isTie Then
            'batting team is now taking the lead. Credit win and/or loss
            'Loss
            Game.loser = pitcherIndex
            'Win
            Game.winner = Game.BTeam.pitcherSel
        Else
            'Determine new situation
            isTie = Game.BTeam.runs + numRuns = Game.PTeam.runs
            If isTie Then
                'reset loser/winner flag, otherwise leave alone
                Game.loser = 0
                Game.winner = 0
                Game.save = 0
            End If
        End If
    End Sub

    ''' <summary>
    ''' gets the division based on the team name and season
    ''' </summary>
    ''' <param name="teamName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDivision(ByRef teamName As String) As String
        Dim divisionName As String = ""
        Dim ds As DataSet = Nothing
        Dim query As String
        Dim teamID As String
        Dim DataAccess As New clsDataAccess("lahman")

        Try
            If gstrSeason <> "1982" Then
                teamID = GetShortTeamTranslation(teamName)
                query = "SELECT lgid,divid FROM teams WHERE teamid = '" & teamID & "' AND yearid = " & _
                                gstrSeason
                ds = DataAccess.ExecuteDataSet(query)
                If ds.Tables(0).Rows.Count > 0 Then
                    With ds.Tables(0).Rows(0)
                        If CInt(gstrSeason) >= 1969 Then
                            'Divisional Format
                            divisionName = .Item("lgid").ToString & " " & _
                                                GetDivisionType(.Item("divid").ToString)
                        Else
                            'Old league format
                            divisionName = .Item("lgid").ToString
                        End If
                    End With
                End If
            Else
                'Bill and John's private season
                Select Case teamName.ToUpper
                    Case "BALTIMORE"
                        divisionName = conALJohn
                    Case "CLEVELAND"
                        divisionName = conALBill
                    Case "NEW YORK (A)"
                        divisionName = conALBill
                    Case "DETROIT"
                        divisionName = conALJohn
                    Case "MILWAUKEE"
                        divisionName = conALJohn
                    Case "TORONTO"
                        divisionName = conALJohn
                    Case "BOSTON"
                        divisionName = conALBill
                    Case "CHICAGO (A)"
                        divisionName = conALJohn
                    Case "OAKLAND"
                        divisionName = conALBill
                    Case "KANSAS CITY"
                        divisionName = conALJohn
                    Case "MINNESOTA"
                        divisionName = conALBill
                    Case "CALIFORNIA"
                        divisionName = conALBill
                    Case "SEATTLE"
                        divisionName = conALBill
                    Case "TEXAS"
                        divisionName = conALJohn
                    Case "ST LOUIS"
                        divisionName = conNLBill
                    Case "NEW YORK (N)"
                        divisionName = conNLBill
                    Case "CHICAGO (N)"
                        divisionName = conNLJohn
                    Case "PITTSBURGH"
                        divisionName = conNLJohn
                    Case "PHILADELPHIA"
                        divisionName = conNLBill
                    Case "MONTREAL"
                        divisionName = conNLJohn
                    Case "LOS ANGELES (N)"
                        divisionName = conNLBill
                    Case "SAN FRANCISCO"
                        divisionName = conNLJohn
                    Case "SAN DIEGO"
                        divisionName = conNLBill
                    Case "CINCINNATI"
                        divisionName = conNLBill
                    Case "ATLANTA"
                        divisionName = conNLJohn
                    Case "HOUSTON"
                        divisionName = conNLJohn
                End Select
            End If
        Catch ex As Exception
            Call MsgBox("Error in GetDivision" & ex.ToString, MsgBoxStyle.OkOnly)
        Finally
            If Not ds Is Nothing Then ds.Dispose()
        End Try
        Return divisionName
    End Function

    ''' <summary>
    ''' determines the highest number on the batter card that will result in a hit. This is mostly determined
    ''' the hit range for the hit and run option
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function HighHitNum(ByVal offset As Integer) As Integer
        Dim highestHitNumberOnCard As Integer
        With Game.BTeam.GetBatterPtr(Game.currentBatter)
            If .hr.Trim <> Nothing Then
                highestHitNumberOnCard = CInt(Val(.hr.Substring(.hr.Length - 2)))
            ElseIf .hit3B8.Trim <> Nothing Then
                highestHitNumberOnCard = CInt(Val(.hit3B8.Substring(.hit3B8.Length - 2)))
            ElseIf .hit2B9.Trim <> Nothing Then
                highestHitNumberOnCard = CInt(Val(.hit2B9.Substring(.hit2B9.Length - 2)))
            ElseIf .hit2B8.Trim <> Nothing Then
                highestHitNumberOnCard = CInt(Val(.hit2B8.Substring(.hit2B8.Length - 2)))
            ElseIf .hit2B7.Trim <> Nothing Then
                highestHitNumberOnCard = CInt(Val(.hit2B7.Substring(.hit2B7.Length - 2)))
            ElseIf .hit1B9.Trim <> Nothing Then
                highestHitNumberOnCard = CInt(Val(.hit1B9.Substring(.hit1B9.Length - 2)))
            ElseIf .hit1B8.Trim <> Nothing Then
                highestHitNumberOnCard = CInt(Val(.hit1B8.Substring(.hit1B8.Length - 2)))
            Else
                highestHitNumberOnCard = 11
            End If
            For i As Integer = 1 To offset
                highestHitNumberOnCard -= 1
                If highestHitNumberOnCard Mod 10 = 0 Then
                    highestHitNumberOnCard -= 2
                End If
            Next i
            If highestHitNumberOnCard < 11 Then highestHitNumberOnCard = 11
        End With
        Return highestHitNumberOnCard
    End Function

    ''' <summary>
    ''' determines the highest single (base hit) number on hitter card
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function HighSingleNum() As Integer
        Dim highestSingleNumberOnCard As Integer
        With Game.BTeam.GetBatterPtr(Game.currentBatter)
            If .hit1B9.Trim <> Nothing Then
                highestSingleNumberOnCard = CInt(Val(.hit1B9.Substring(.hit1B9.Length - 2)))
            ElseIf .hit1B8.Trim <> Nothing Then
                highestSingleNumberOnCard = CInt(Val(.hit1B8.Substring(.hit1B8.Length - 2)))
            ElseIf .hit1B7.Trim <> Nothing Then
                highestSingleNumberOnCard = CInt(Val(.hit1B7.Substring(.hit1B7.Length - 2)))
            Else
                highestSingleNumberOnCard = 0
            End If
        End With
        Return highestSingleNumberOnCard
    End Function

    ''' <summary>
    ''' determines the result if the Squeeze option is chosen
    ''' </summary>
    ''' <param name="playResult"></param>
    ''' <param name="facCard"></param>
    ''' <remarks></remarks>
    Private Sub HandleSqueeze(ByRef playResult As String, ByRef facCard As FACCard)
        Dim stringPosition As Integer
        Dim fieldPos As Integer

        Try
            playResult = "SQZ"
            If bolCornersIn Or bolInfieldIn Then
                'add 10 to random number
                If Val(facCard.Random) + 10 > 88 Then
                    facCard.Random = "88"
                Else
                    facCard.Random = (Val(facCard.Random) + 10).ToString
                End If
            End If
            If Game.BTeam.GetBatterPtr(Game.currentBatter).sac = "AA" Then
                Select Case CInt(facCard.Random)
                    Case 11 To 28
                        playResult += "1"
                    Case 31 To 32
                        playResult += "2"
                    Case 33 To 38
                        playResult += "3"
                    Case 41 To 43
                        playResult += "4"
                    Case 44 To 46
                        playResult += "5"
                    Case 47 To 52
                        playResult += "6"
                    Case 53 To 54
                        playResult += "7"
                    Case 55 To 71
                        playResult += "8"
                    Case 72 To 88
                        playResult += "9"
                    Case Else
                        Call MsgBox("Error in Squeeze lookup!", MsgBoxStyle.Exclamation)
                        playResult = conNoPlay
                End Select
            ElseIf Game.BTeam.GetBatterPtr(Game.currentBatter).sac = "BB" Then
                Select Case CInt(facCard.Random)
                    Case 11 To 26
                        playResult += "1"
                    Case 27 To 28
                        playResult += "2"
                    Case 31 To 38
                        playResult += "3"
                    Case 41 To 43
                        playResult += "4"
                    Case 44 To 46
                        playResult += "5"
                    Case 47 To 52
                        playResult += "6"
                    Case 53 To 57
                        playResult += "7"
                    Case 58 To 71
                        playResult += "8"
                    Case 72 To 88
                        playResult += "9"
                    Case Else
                        Call MsgBox("Error in Squeeze lookup!", MsgBoxStyle.Exclamation)
                        playResult = conNoPlay
                End Select
            ElseIf Game.BTeam.GetBatterPtr(Game.currentBatter).sac = "CC" Then
                Select Case CInt(facCard.Random)
                    Case 11 To 21
                        playResult += "1"
                    Case 22 To 25
                        playResult += "2"
                    Case 26 To 38
                        playResult += "3"
                    Case 41 To 43
                        playResult += "4"
                    Case 44 To 46
                        playResult += "5"
                    Case 47 To 52
                        playResult += "6"
                    Case 53 To 63
                        playResult += "7"
                    Case 64 To 71
                        playResult += "8"
                    Case 72 To 88
                        playResult += "9"
                    Case Else
                        Call MsgBox("Error in Squeeze lookup!", MsgBoxStyle.Exclamation)
                        playResult = conNoPlay
                End Select
            Else
                Select Case CInt(facCard.Random)
                    Case 11 To 16
                        playResult += "1"
                    Case 17 To 22
                        playResult += "2"
                    Case 23 To 38
                        playResult += "3"
                    Case 41 To 43
                        playResult += "4"
                    Case 44 To 46
                        playResult += "5"
                    Case 47 To 52
                        playResult += "6"
                    Case 53 To 71
                        playResult += "7"
                    Case 72 To 78
                        playResult += "8"
                    Case 81 To 88
                        playResult += "9"
                    Case Else
                        Call MsgBox("Error in Squeeze lookup!", MsgBoxStyle.Exclamation)
                        playResult = conNoPlay
                End Select
            End If
            If playResult <> conNoPlay Then
                playResult += BaseSituation()
                Call Game.GetResultPtr.ChartLookup(playResult)
                stringPosition = Game.GetResultPtr.description.IndexOf("Result:")
                If stringPosition > -1 Then
                    fieldPos = CInt(Val(Game.GetResultPtr.description.Substring(Game.GetResultPtr.description.Length - 1)))
                    If fieldPos = 1 Then
                        'Pitcher
                        Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).PitStatPtr.e += 1
                    Else
                        fieldPos = Game.GetFielderNumByPos(fieldPos)
                        Game.PTeam.GetBatterPtr(fieldPos).BatStatPtr.E += 1
                    End If
                    Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).eOuts += 1
                    gbolSOE = True
                End If
            End If
        Catch ex As Exception
            Call MsgBox("Error in HandleSqueeze. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub
    ''' <summary>
    ''' determines the result if the SAC option is chosen
    ''' </summary>
    ''' <param name="playResult"></param>
    ''' <param name="facCard"></param>
    ''' <remarks></remarks>
    Private Sub HandleSAC(ByRef playResult As String, ByRef facCard As String)
        Try
            If bolCornersIn Or bolInfieldIn Then
                'add 10 to random number
                If Val(facCard) + 10 > 88 Then
                    facCard = "88"
                Else
                    facCard = (Val(facCard) + 10).ToString
                End If
            End If
            With Game.BTeam.GetBatterPtr(Game.currentBatter)
                If .sac = "AA" Then
                    Select Case CInt(facCard)
                        Case 11 To 18
                            If .obr = "A" Or .obr = "B" Then
                                playResult += "5" & BaseSituation()
                            Else
                                playResult += "1" & BaseSituation()
                            End If
                        Case 21 To 27
                            playResult += "2" & BaseSituation()
                        Case 28 To 37
                            playResult += "3" & BaseSituation()
                        Case 38 To 65
                            playResult += "4" & BaseSituation()
                        Case 66 To 67
                            playResult += "5" & BaseSituation()
                        Case 68
                            playResult += "6" & BaseSituation()
                        Case 71
                            playResult += "7" & BaseSituation()
                        Case 72
                            playResult += "8" & BaseSituation()
                        Case 73
                            playResult += "9" & BaseSituation()
                        Case 74
                            playResult += "10" & BaseSituation()
                        Case 75
                            playResult += "11" & BaseSituation()
                        Case 76 To 77
                            playResult += "12" & BaseSituation()
                        Case 78 To 88
                            playResult += "13" & "A"
                            If .cht = "P" Then
                                playResult += "P" & BaseSituation()
                            Else
                                'Foul ball
                                playResult = conNoPlay
                            End If
                        Case Else
                            Call MsgBox("Error in SAC lookup!", MsgBoxStyle.Exclamation)
                            playResult = conNoPlay
                    End Select
                ElseIf .sac = "BB" Or .sac = "CC" Then
                    'Incentivize more successful bunting
                    Select Case CInt(facCard)
                        Case 11 To 17
                            If .obr = "A" Or .obr = "B" Then
                                playResult += "5" & BaseSituation()
                            Else
                                playResult += "1" & BaseSituation()
                            End If
                        Case 18 To 25
                            playResult += "2" & BaseSituation()
                        Case 26 To 34
                            playResult += "3" & BaseSituation()
                        Case 35 To 55
                            playResult += "4" & BaseSituation()
                        Case 56
                            playResult += "5" & BaseSituation()
                        Case 57 To 68
                            playResult += "6" & BaseSituation()
                        Case 71
                            playResult += "7" & BaseSituation()
                        Case 72
                            playResult += "8" & BaseSituation()
                        Case 73
                            playResult += "9" & BaseSituation()
                        Case 74
                            playResult += "10" & BaseSituation()
                        Case 75 To 77
                            playResult += "11" & BaseSituation()
                        Case 78 To 83
                            playResult += "12" & BaseSituation()
                        Case 84 To 88
                            playResult += "13" & "A"
                            If .cht = "P" Then
                                playResult += "P" & BaseSituation()
                            Else
                                playResult = conNoPlay
                            End If
                        Case Else
                            Call MsgBox("Error in SAC lookup!", MsgBoxStyle.Exclamation)
                            playResult = conNoPlay
                    End Select
                    'ElseIf .sac = "CC" Then
                    '    Select Case CInt(facCard)
                    '        Case 11 To 15
                    '            If .obr = "A" Or .obr = "B" Then
                    '                playResult += "5" & BaseSituation()
                    '            Else
                    '                playResult += "1" & BaseSituation()
                    '            End If
                    '        Case 16 To 23
                    '            playResult += "2" & BaseSituation()
                    '        Case 24 To 31
                    '            playResult += "3" & BaseSituation()
                    '        Case 32 To 48
                    '            playResult += "4" & BaseSituation()
                    '        Case 51 To 68
                    '            playResult += "6" & BaseSituation()
                    '        Case 71
                    '            playResult += "7" & BaseSituation()
                    '        Case 72
                    '            playResult += "8" & BaseSituation()
                    '        Case 73
                    '            playResult += "9" & BaseSituation()
                    '        Case 74
                    '            playResult += "10" & BaseSituation()
                    '        Case 75 To 82
                    '            playResult += "11" & BaseSituation()
                    '        Case 83 To 85
                    '            playResult += "12" & BaseSituation()
                    '        Case 86 To 88
                    '            playResult += "13" & "A"
                    '            If .cht = "P" Then
                    '                playResult += "P" & BaseSituation()
                    '            Else
                    '                playResult = conNoPlay
                    '            End If
                    '        Case Else
                    '            Call MsgBox("Error in SAC lookup!", MsgBoxStyle.Exclamation)
                    '            playResult = conNoPlay
                    '    End Select
                ElseIf .sac = "DD" Then
                    Select Case CInt(facCard)
                        Case 11 To 13
                            If .obr = "A" Or .obr = "B" Then
                                playResult += "5" & BaseSituation()
                            Else
                                playResult += "1" & BaseSituation()
                            End If
                        Case 14 To 18
                            playResult += "2" & BaseSituation()
                        Case 21 To 24
                            playResult += "3" & BaseSituation()
                        Case 25 To 41
                            playResult += "4" & BaseSituation()
                        Case 42 To 68
                            playResult += "6" & BaseSituation()
                        Case 71
                            playResult += "7" & BaseSituation()
                        Case 72
                            playResult += "8" & BaseSituation()
                        Case 73
                            playResult += "9" & BaseSituation()
                        Case 74
                            playResult += "10" & BaseSituation()
                        Case 75 To 82
                            playResult += "11" & BaseSituation()
                        Case 83 To 86
                            playResult += "12" & BaseSituation()
                        Case 87 To 88
                            playResult += "13" & "A"
                            If .cht = "P" Then
                                playResult += "P" & BaseSituation()
                            Else
                                playResult = conNoPlay
                            End If
                        Case Else
                            Call MsgBox("Error in SAC lookup!", MsgBoxStyle.Exclamation)
                            playResult = conNoPlay
                    End Select
                End If
            End With
        Catch ex As Exception
            Call MsgBox("Error in HandleSAC. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' determines whether runner steals on their own.
    ''' </summary>
    ''' <param name="playResult"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckAutoSteal(ByRef playResult As String) As Boolean
        Dim hasAutoSteal As Boolean
        Dim facCard As FACCard

        Try
            If Not gbolAutoSteal Then Return False
            If AutoStealContext2nd() Then
                facCard = GetFAC("autoSteal", "Random")
                Select Case Game.BTeam.GetBatterPtr(FirstBase.runner).sp
                    Case "A", "Y", "Z"
                        hasAutoSteal = CInt(facCard.Random) < 41
                    Case "B"
                        hasAutoSteal = CInt(facCard.Random) < 36
                    Case "C"
                        hasAutoSteal = CInt(facCard.Random) < 32
                    Case "D"
                        hasAutoSteal = CInt(facCard.Random) < 21
                End Select
                If hasAutoSteal Then
                    facCard = GetFAC("stealSuccess", "Random")
                    Select Case CInt(facCard.Random)
                        Case 11 To 48
                            'Stolen Base
                            playResult = "SBR111-15" & BaseSituation()
                            Game.BTeam.GetBatterPtr(FirstBase.runner).BatStatPtr.SB += 1
                            Game.BTeam.GetBatterPtr(FirstBase.runner).BatStatPtr.SBA += 1
                        Case 51 To 54
                            If Game.BTeam.GetBatterPtr(FirstBase.runner).sp = "D" Then
                                'Caught Stealing (Lower success rate of SP D)
                                playResult = "SBR148-52" & BaseSituation()
                                Game.BTeam.GetBatterPtr(FirstBase.runner).BatStatPtr.SBA += 1
                            Else
                                'Stolen Base
                                playResult = "SBR111-15" & BaseSituation()
                                Game.BTeam.GetBatterPtr(FirstBase.runner).BatStatPtr.SB += 1
                                Game.BTeam.GetBatterPtr(FirstBase.runner).BatStatPtr.SBA += 1
                            End If
                        Case 62 To 65
                            'Stolen Base
                            playResult += String.Format("SBR1{0}{1}", facCard.Random, BaseSituation())
                            Game.BTeam.GetBatterPtr(FirstBase.runner).BatStatPtr.SB += 1
                            Game.BTeam.GetBatterPtr(FirstBase.runner).BatStatPtr.SBA += 1
                        Case 71 To 72
                            playResult += String.Format("SBR1{0}{1}", facCard.Random, BaseSituation())
                        Case 55 To 56
                            If Game.BTeam.GetBatterPtr(FirstBase.runner).sp = "C" Or _
                                       Game.BTeam.GetBatterPtr(FirstBase.runner).sp = "D" Then
                                'Caught Stealing 
                                playResult = "SBR148-52" & BaseSituation()
                                Game.BTeam.GetBatterPtr(FirstBase.runner).BatStatPtr.SBA += 1
                            Else
                                'Stolen Base (increase the success of all SP ratings, except C,D
                                playResult = "SBR111-15" & BaseSituation()
                                Game.BTeam.GetBatterPtr(FirstBase.runner).BatStatPtr.SB += 1
                                Game.BTeam.GetBatterPtr(FirstBase.runner).BatStatPtr.SBA += 1
                            End If
                        Case 57 To 61, 66 To 68, 73 To 88
                            'Caught Stealing
                            playResult = "SBR148-52" & BaseSituation()
                            Game.BTeam.GetBatterPtr(FirstBase.runner).BatStatPtr.SBA += 1
                    End Select
                    'If CInt(facCard.Random) < 63 Then
                    '    'Stolen Base
                    '    playResult = "SBR111-15" & BaseSituation()
                    '    Game.BTeam.GetBatterPtr(FirstBase.runner).BatStatPtr.SB += 1
                    '    Game.BTeam.GetBatterPtr(FirstBase.runner).BatStatPtr.SBA += 1
                    'Else
                    '    'Caught Stealing
                    '    playResult = "SBR148-52" & BaseSituation()
                    '    Game.BTeam.GetBatterPtr(FirstBase.runner).BatStatPtr.SBA += 1
                    'End If
                End If
            ElseIf AutoStealContext3rd() Then
                facCard = GetFAC("autoSteal", "Random")
                Select Case Game.BTeam.GetBatterPtr(SecondBase.runner).sp
                    Case "A", "Y", "Z"
                        hasAutoSteal = CInt(facCard.Random) < 16
                    Case "B"
                        hasAutoSteal = CInt(facCard.Random) < 14
                End Select
                If hasAutoSteal Then
                    facCard = GetFAC("stealSuccess", "Random")
                    Select Case CInt(facCard.Random)
                        Case 11 To 54
                            'Stolen Base
                            playResult = "SBR211-14" & BaseSituation()
                            Game.BTeam.GetBatterPtr(SecondBase.runner).BatStatPtr.SB += 1
                            Game.BTeam.GetBatterPtr(SecondBase.runner).BatStatPtr.SBA += 1
                        Case 62 To 64
                            'Stolen Base
                            playResult += String.Format("SBR2{0}{1}", facCard.Random, BaseSituation())
                            Game.BTeam.GetBatterPtr(SecondBase.runner).BatStatPtr.SB += 1
                            Game.BTeam.GetBatterPtr(SecondBase.runner).BatStatPtr.SBA += 1
                        Case 65
                            playResult += String.Format("SBR2{0}{1}", facCard.Random, BaseSituation())
                        Case 71 To 72
                            playResult = "SBR271-72" & BaseSituation()
                        Case 55 To 61, 66 To 68, 73 To 88
                            'Caught Stealing
                            playResult = "SBR242-52" & BaseSituation()
                            Game.BTeam.GetBatterPtr(SecondBase.runner).BatStatPtr.SBA += 1
                    End Select
                    'If CInt(facCard.Random) < 63 Then
                    '    'Stolen Base
                    '    playResult = "SBR211-14" & BaseSituation()
                    '    Game.BTeam.GetBatterPtr(SecondBase.runner).BatStatPtr.SB += 1
                    '    Game.BTeam.GetBatterPtr(SecondBase.runner).BatStatPtr.SBA += 1
                    'Else
                    '    'Caught Stealing
                    '    playResult = "SBR242-52" & BaseSituation()
                    '    Game.BTeam.GetBatterPtr(SecondBase.runner).BatStatPtr.SBA += 1
                    'End If
                End If
            End If
            If hasAutoSteal And playResult <> conNoPlay Then
                Dim stringPosition As Integer
                Dim tempIntegerValue As Integer

                Game.GetResultPtr.ChartLookup(playResult)
                'Handle fielding errors
                stringPosition = Game.GetResultPtr.description.IndexOf("Result:E")
                If stringPosition > -1 And playResult <> conNoPlay Then
                    tempIntegerValue = CInt(Val(Game.GetResultPtr.description.Substring(Game.GetResultPtr.description.Length - 1)))
                    If tempIntegerValue = 1 Then
                        'Pitcher
                        Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).PitStatPtr.e += 1
                    Else
                        tempIntegerValue = Game.GetFielderNumByPos(tempIntegerValue)
                        Game.PTeam.GetBatterPtr(tempIntegerValue).BatStatPtr.E += 1
                    End If
                    Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).eOuts += 1
                    Select Case playResult.Substring(3, 1)
                        Case "1"
                            FirstBase.unearned = True
                        Case "2"
                            SecondBase.unearned = True
                        Case "3"
                            ThirdBase.unearned = True
                    End Select
                End If
                'Cleanup description
                stringPosition = Game.GetResultPtr.description.IndexOf("Result:")
                If stringPosition > 0 Then
                    Game.GetResultPtr.description = Game.GetResultPtr.description.Substring(0, stringPosition).Trim
                End If
            End If

        Catch ex As Exception
            Call MsgBox("Error in CheckAutoSteal. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return hasAutoSteal
    End Function

    ''' <summary>
    ''' handles the effects of lefty vs lefty, lefty vs righty between pitcher and batter, and vice versa.
    ''' </summary>
    ''' <param name="playResult"></param>
    ''' <param name="FACCard"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckLeftyRighty(ByRef playResult As String, ByRef FACCard As FACCard) As Boolean
        Dim battingSide As String
        Dim isLvsR As Boolean

        Try
            battingSide = Game.BTeam.GetBatterPtr(Game.currentBatter).cht.Substring(0, 1)
            If battingSide = "P" Then 'Pitcher
                battingSide = Game.BTeam.GetPitcherPtr(Game.BTeam.pitcherSel).throwField
            End If
            'Ignore switch hitters only
            If Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).throwField.Substring(0, 1) = "L" Then
                If battingSide = "L" Then
                    Select Case FACCard.Random
                        Case "11", "12", "13"
                            isLvsR = True
                            playResult = "LL11-13" & BaseSituation()
                        Case "14"
                            isLvsR = True
                            playResult = "LL14" & BaseSituation()
                        Case "15"
                            isLvsR = True
                            playResult = "LL15" & BaseSituation()
                    End Select
                ElseIf battingSide = "R" Then  'Ignore switch hitters
                    Select Case FACCard.Random
                        Case "87"
                            isLvsR = True
                            playResult = "LR87" & BaseSituation()
                        Case "88"
                            isLvsR = True
                            playResult = "LR88" & BaseSituation()
                    End Select
                End If
            Else 'Righty pitcher
                If battingSide = "L" Then
                    Select Case FACCard.Random
                        Case "87"
                            isLvsR = True
                            playResult = "RL87" & BaseSituation()
                        Case "88"
                            isLvsR = True
                            playResult = "RL88" & BaseSituation()
                    End Select
                ElseIf battingSide = "R" Then
                    Select Case FACCard.Random
                        Case "11"
                            isLvsR = True
                            playResult = "RR11" & BaseSituation()
                        Case "12"
                            isLvsR = True
                            playResult = "RR12" & BaseSituation()
                    End Select
                End If
            End If
        Catch ex As Exception
            Call MsgBox("Error in CheckLeftyRighty. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return isLvsR
    End Function

    ''' <summary>
    ''' determines result based on the pitcher's card
    ''' </summary>
    ''' <param name="playResult"></param>
    ''' <param name="FACCard"></param>
    ''' <param name="baseAdvance"></param>
    ''' <param name="suppDescription"></param>
    ''' <remarks></remarks>
    Private Sub GetResultFromPitcherCard(ByRef playResult As String, ByRef FACCard As FACCard, ByRef baseAdvance As Boolean, _
                    ByRef suppDescription As String)
        Dim walkRange As String
        Dim lastNumberBeforeWRange As Integer
        Try
            With Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel)
                If bolPitchAround1 Then
                    lastNumberBeforeWRange = GetLastNumberOfPreviousRange(Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel))
                    'add 8 walk numbers to pitcher card
                    walkRange = IncreaseRange(.w, 8, lastNumberBeforeWRange)
                ElseIf bolPitchAround2 Then
                    lastNumberBeforeWRange = GetLastNumberOfPreviousRange(Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel))
                    'add 16 walk numbers to pitcher card
                    walkRange = IncreaseRange(.w, 16, lastNumberBeforeWRange)
                Else
                    walkRange = .w
                End If

                If InRange(walkRange, CInt(FACCard.Random)) Then
                    'always check walks first because of the pitch around effect
                    playResult = "W"
                ElseIf InRange(.out, CInt(FACCard.Random)) Then
                    FACCard = GetFAC("checkOut", Game.BTeam.GetBatterPtr(Game.currentBatter).cht)
                    Select Case Game.BTeam.GetBatterPtr(Game.currentBatter).cht
                        Case "P"
                            playResult = FACCard.P
                        Case "LN"
                            playResult = FACCard.LN
                        Case "LP"
                            playResult = FACCard.LP
                        Case "SN"
                            playResult = FACCard.SN
                        Case "SP"
                            playResult = FACCard.SP
                        Case "RN"
                            playResult = FACCard.RN
                        Case "RP"
                            playResult = FACCard.RP
                        Case Else
                            Call MsgBox("Cannot find .Cht for batter.", MsgBoxStyle.OkOnly)
                    End Select
                ElseIf InRange(.hit1Bf, CInt(FACCard.Random)) Then
                    playResult = "1Bf"
                    suppDescription = " through the infield. "
                    Game.currentFieldPos = 6
                ElseIf InRange(.hit1B7, CInt(FACCard.Random)) Then
                    playResult = "1B7"
                    baseAdvance = True
                    suppDescription = " to left. "
                    Game.currentFieldPos = 7
                ElseIf InRange(.hit1B8, CInt(FACCard.Random)) Then
                    playResult = "1B8"
                    baseAdvance = True
                    suppDescription = " to center. "
                    Game.currentFieldPos = 8
                ElseIf InRange(.hit1B9, CInt(FACCard.Random)) Then
                    playResult = "1B9"
                    baseAdvance = True
                    suppDescription = " to right. "
                    Game.currentFieldPos = 9
                ElseIf InRange(.bk, CInt(FACCard.Random)) Then
                    FACCard = GetFAC("balk", "Pitch")
                    If FACCard.Pitch And BaseSituation() <> "000" Then
                        playResult = "BK"
                    Else
                        playResult = conNoPlay
                    End If
                ElseIf InRange(.k, CInt(FACCard.Random)) Then
                    playResult = "K"
                ElseIf InRange(.pb, CInt(FACCard.Random)) Then
                    FACCard = GetFAC("passedBall", "Pitch")
                    If FACCard.Pitch And BaseSituation() <> "000" Then
                        playResult = "PB"
                        With Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(2))
                            .BatStatPtr.PB += 1
                        End With
                    Else
                        playResult = conNoPlay
                    End If
                ElseIf InRange(.wp, CInt(FACCard.Random)) Then
                    FACCard = GetFAC("wildPitch", "Pitch")
                    If FACCard.Pitch And BaseSituation() <> "000" Then
                        playResult = "WP"
                    Else
                        playResult = conNoPlay
                    End If
                Else
                    Call MsgBox("Random number " & FACCard.Random & " is not in batter's ranges.", MsgBoxStyle.OkOnly)
                End If
            End With
        Catch ex As Exception
            Call MsgBox("Error in GetResultFromPitcherCard. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Function GetLastNumberOfPreviousRange(ByVal pitcherCard As clsPitcher) As Integer
        If Right(pitcherCard.k, 2) <> Nothing Then
            Return CInt(Right(pitcherCard.k, 2))
        ElseIf Right(pitcherCard.bk, 2) <> Nothing Then
            Return CInt(Right(pitcherCard.bk, 2))
        ElseIf Right(pitcherCard.hit1B9, 2) <> Nothing Then
            Return CInt(Right(pitcherCard.hit1B9, 2))
        ElseIf Right(pitcherCard.hit1B8, 2) <> Nothing Then
            Return CInt(Right(pitcherCard.hit1B8, 2))
        ElseIf Right(pitcherCard.hit1B7, 2) <> Nothing Then
            Return CInt(Right(pitcherCard.hit1B7, 2))
        ElseIf Right(pitcherCard.hit1Bf, 2) <> Nothing Then
            Return CInt(Right(pitcherCard.hit1Bf, 2))
        Else
            Return 0
        End If
    End Function

    ''' <summary>
    ''' determine result based on the hitter's card
    ''' </summary>
    ''' <param name="playResult"></param>
    ''' <param name="FACCard"></param>
    ''' <param name="baseAdvance"></param>
    ''' <param name="suppDescription"></param>
    ''' <remarks></remarks>
    Private Sub GetResultFromHitterCard(ByRef playResult As String, ByRef FACCard As FACCard, ByRef baseAdvance As Boolean, _
                    ByRef suppDescription As String)
        Try
            With Game.BTeam.GetBatterPtr(Game.currentBatter)
                If InRange(.out, CInt(FACCard.Random)) Then
                    FACCard = GetFAC("checkOut", .cht)
                    Select Case .cht
                        Case "P"
                            playResult = FACCard.P
                        Case "LN"
                            playResult = FACCard.LN
                        Case "LP"
                            playResult = FACCard.LP
                        Case "SN"
                            playResult = FACCard.SN
                        Case "SP"
                            playResult = FACCard.SP
                        Case "RN"
                            playResult = FACCard.RN
                        Case "RP"
                            playResult = FACCard.RP
                        Case Else
                            Call MsgBox("Cannot find .Cht for batter.", MsgBoxStyle.OkOnly)
                    End Select
                ElseIf InRange(.hit1Bf, CInt(FACCard.Random)) Then
                    playResult = "1Bf*"
                    suppDescription = " through the infield. "
                ElseIf InRange(.hit1B7, CInt(FACCard.Random)) Then
                    playResult = "1B7*"
                    baseAdvance = True
                    suppDescription = " to left. "
                ElseIf InRange(.hit1B8, CInt(FACCard.Random)) Then
                    playResult = "1B8*"
                    baseAdvance = True
                    suppDescription = " to center. "
                ElseIf InRange(.hit1B9, CInt(FACCard.Random)) Then
                    playResult = "1B9*"
                    baseAdvance = True
                    suppDescription = " to right. "
                ElseIf InRange(.hit2B7, CInt(FACCard.Random)) Then
                    playResult = "2B7*"
                    baseAdvance = True
                    suppDescription = " to left. "
                ElseIf InRange(.hit2B8, CInt(FACCard.Random)) Then
                    playResult = "2B8*"
                    baseAdvance = True
                    suppDescription = " to center. "
                ElseIf InRange(.hit2B9, CInt(FACCard.Random)) Then
                    playResult = "2B9*"
                    baseAdvance = True
                    suppDescription = " to right. "
                ElseIf InRange(.hit3B8, CInt(FACCard.Random)) Then
                    playResult = "3B8*"
                ElseIf InRange(.hr, CInt(FACCard.Random)) Then
                    playResult = "HR"
                ElseIf InRange(.k, CInt(FACCard.Random)) Then
                    If Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).pbRange = "2-5" AndAlso _
                               Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).k = "" AndAlso _
                               SwitchFromK() Then
                        'reverse strikeout to out if 2-5 pitcher has no K range
                        FACCard = GetFAC("checkOut", .cht)
                        Select Case .cht
                            Case "P"
                                playResult = FACCard.P
                            Case "LN"
                                playResult = FACCard.LN
                            Case "LP"
                                playResult = FACCard.LP
                            Case "SN"
                                playResult = FACCard.SN
                            Case "SP"
                                playResult = FACCard.SP
                            Case "RN"
                                playResult = FACCard.RN
                            Case "RP"
                                playResult = FACCard.RP
                            Case Else
                                Call MsgBox("Cannot find .Cht for batter.", MsgBoxStyle.OkOnly)
                        End Select
                    Else
                        playResult = "K"
                    End If
                ElseIf InRange(.w, CInt(FACCard.Random)) Then
                    playResult = "W"
                ElseIf InRange(.hpb, CInt(FACCard.Random)) Then
                    playResult = "HPB"
                Else
                    Call MsgBox("Random number " & FACCard.Random & " is not in batter's ranges.", MsgBoxStyle.OkOnly)
                End If
            End With
        Catch ex As Exception
            Call MsgBox("Error in GetResultFromHitterCard. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' determines if an error occurred based on the fielder rating. If error occurs the method also looks up the error
    ''' and determines the outcome
    ''' </summary>
    ''' <param name="playResult"></param>
    ''' <param name="FACCard"></param>
    ''' <param name="baseAdvance"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckForError(ByRef playResult As String, ByRef FACCard As FACCard, ByRef baseAdvance As Boolean) As Boolean
        Dim stringPosition As Integer
        Dim isError As Boolean
        Dim playResultOld As String = ""
        Dim errorOuts As Integer
        Dim tmpResultPtr As clsResult
        Dim chargeError As Boolean
        Dim tempValue As String
        Dim tempIntegerValue As Integer

        Try
            'Get position
            If playResult = "G31A*" Then
                Game.currentFieldPos = 3
            Else
                stringPosition = playResult.IndexOf("*") - 1
                While stringPosition > -1 And Not IsNumeric(playResult.Substring(stringPosition, 1))
                    'Handle GnA* outs
                    stringPosition = stringPosition - 1
                End While
                If stringPosition <> -1 Then
                    Game.currentFieldPos = CInt(Val(playResult.Substring(stringPosition, 1)))
                End If
            End If
            If FACCard.ErrorNum <> conNone Then
                For i As Integer = 1 To 9
                    If Game.PTeam.GetBatterPtr(Game.PTeam.GetPlayerNum(i)).position = _
                                                gblPositions(Game.currentFieldPos) Then
                        isError = Val(Game.PTeam.GetBatterPtr(Game.PTeam.GetPlayerNum(i)).errorRating) >= _
                                                Val(FACCard.ErrorNum)
                        If i < 7 And playResult.Substring(0, 1) = "F" And (BaseSituation() = "111" Or BaseSituation() = "011") And _
                                Game.outs < 2 Then
                            'Infield fly rule. batter automatically out.
                            isError = False
                        End If
                    End If
                Next i
                If isError Then
                    If playResult.Trim.Substring(playResult.Trim.Length - 1) = "*" Then
                        playResultOld = playResult.Substring(0, playResult.Trim.Length - 1) & BaseSituation()
                    Else
                        playResultOld = playResult.Trim & BaseSituation()
                    End If
                    If playResult.Length >= 2 AndAlso playResult.Substring(0, 2) = "F2" Then
                        playResult = "F2ERR"
                    ElseIf playResult.Substring(0, 1) = "F" Then
                        playResult = IIF(Game.currentFieldPos >= 7, "FOERR", "FIERR")
                    ElseIf playResult.Substring(0, 1) = "L" Then
                        playResult = "LOERR"
                    Else
                        'Normal errors
                        'Check error code
                        Select Case Game.currentFieldPos
                            Case 1 To 2
                                FACCard = GetFAC("errorType", "InfErrPC")
                                playResult = "E" & FACCard.InfErrPC
                            Case 3
                                FACCard = GetFAC("errorType", "InfErr1B")
                                playResult = "E" & FACCard.InfErr1B
                            Case 4 To 6
                                FACCard = GetFAC("errorType", "InfErr")
                                playResult = "E" & FACCard.InfErr
                            Case 7 To 9
                                If playResult.Length >= 2 Then
                                    playResultOld = playResult.Substring(0, 2) & BaseSituation()
                                Else
                                    playResultOld = playResult & BaseSituation()
                                End If
                                FACCard = GetFAC("errorType", "OFErr")
                                playResult = "E" & FACCard.OFErr
                            Case Else
                                Call MsgBox("Invalid position number.", MsgBoxStyle.OkOnly)
                                Return False
                        End Select
                    End If
                    playResult += BaseSituation()
                    Call Game.GetResultPtr.ChartLookup(playResultOld)
                    errorOuts = GetOuts() 'This is the number of outs had there not been an error

                    'Save original result for statistical purposes
                    'Error results do not offer statistics

                    tmpResultPtr = New clsResult
                    tmpResultPtr.ab = Game.GetResultPtr.ab
                    tmpResultPtr.hit = Game.GetResultPtr.hit
                    tmpResultPtr.doubl = Game.GetResultPtr.doubl
                    tmpResultPtr.triple = Game.GetResultPtr.triple
                    tmpResultPtr.hr = Game.GetResultPtr.hr
                    tmpResultPtr.rbi = Game.GetResultPtr.rbi
                    tmpResultPtr.strikeOut = Game.GetResultPtr.strikeOut
                    tmpResultPtr.walk = Game.GetResultPtr.walk
                    tmpResultPtr.hpb = Game.GetResultPtr.hpb
                    tmpResultPtr.wp = Game.GetResultPtr.wp
                    tmpResultPtr.bk = Game.GetResultPtr.bk
                    Call Game.GetResultPtr.ChartLookup(playResult)
                    Call InsertStats(tmpResultPtr)

                    'Define exception error contexts
                    chargeError = True
                    With Game.BTeam.GetBatterPtr(Game.currentBatter)
                        Select Case playResult.Substring(0, 5)
                            Case "E2000"
                                If .obr = "A" Or .obr = "B" Then
                                    Game.GetResultPtr.resBatter = 2
                                    Game.GetResultPtr.description = "Error on " & gblPositions(Game.currentFieldPos) & ". Batter safe at second."
                                Else
                                    Game.GetResultPtr.description = "Error on " & gblPositions(Game.currentFieldPos) & ". Batter safe at first."
                                End If
                            Case "E2001", "E2010", "E2110"
                                'Modify description only
                                Game.GetResultPtr.description = "Error on " & gblPositions(Game.currentFieldPos) & _
                                                            Game.GetResultPtr.description.Substring(8)
                            Case "E4000"
                                Call Game.GetResultPtr.ChartLookup(playResultOld)
                                tempValue = playResultOld.Substring(0, 2)
                                If tempValue = "1B" Then
                                    Game.GetResultPtr.resBatter = 2
                                    Game.GetResultPtr.description += ". Error on " & _
                                                        gblPositions(Game.currentFieldPos) & ". Batter to second"
                                ElseIf tempValue = "2B" Then
                                    If .obr = "A" Then
                                        Game.GetResultPtr.resBatter = 4
                                        Game.GetResultPtr.description += _
                                                ". Error on " & gblPositions(Game.currentFieldPos) & ". Batter scores"
                                    Else
                                        Game.GetResultPtr.resBatter = 3
                                        Game.GetResultPtr.description += _
                                                ". Error on " & gblPositions(Game.currentFieldPos) & ". Batter to third"
                                    End If
                                Else '3B
                                    Game.GetResultPtr.resBatter = 4
                                    Game.GetResultPtr.description += _
                                                ". Error on " & gblPositions(Game.currentFieldPos) & ". Batter scores"
                                End If
                            Case "E5000"
                                Call Game.GetResultPtr.ChartLookup(playResultOld)
                                If .obr <> "D" And .obr <> "E" Then 'else no error
                                    tempValue = playResultOld.Substring(0, 2)
                                    If tempValue = "1B" Then
                                        Game.GetResultPtr.resBatter = 2
                                        Game.GetResultPtr.description += _
                                                    ". Error on " & gblPositions(Game.currentFieldPos) & ". Batter to second"
                                    ElseIf tempValue = "2B" Then
                                        Game.GetResultPtr.resBatter = 3
                                        Game.GetResultPtr.description += _
                                                    ". Error on " & gblPositions(Game.currentFieldPos) & ". Batter to third"
                                    Else '3B
                                        Game.GetResultPtr.resBatter = 4
                                        Game.GetResultPtr.description += _
                                                    ". Error on " & gblPositions(Game.currentFieldPos) & ". Batter scores"
                                    End If
                                Else
                                    isError = False
                                    chargeError = False
                                    playResult = playResultOld
                                End If
                            Case "E4001"
                                Call Game.GetResultPtr.ChartLookup(playResultOld)
                                tempValue = playResultOld.Substring(0, 2)
                                Game.GetResultPtr.resFirst = 4
                                If tempValue = "1B" Then
                                    Game.GetResultPtr.resBatter = 2
                                    Game.GetResultPtr.description += _
                                            ". Error on " & gblPositions(Game.currentFieldPos) & ". Batter to second and runner scores"
                                ElseIf tempValue = "2B" Then
                                    Game.GetResultPtr.resBatter = 3
                                    Game.GetResultPtr.description += _
                                            ". Error on " & gblPositions(Game.currentFieldPos) & ". Batter to third and runner scores"
                                Else '3B
                                    Game.GetResultPtr.resBatter = 4
                                    Game.GetResultPtr.description += _
                                            ". Error on " & gblPositions(Game.currentFieldPos) & ". Batter and runner scores"
                                End If
                            Case "E5001"
                                'No error
                                playResult = playResultOld
                                Call Game.GetResultPtr.ChartLookup(playResult)
                                chargeError = False
                                isError = False
                            Case "E1010"
                                If Game.BTeam.GetBatterPtr(SecondBase.runner).obr = "A" And Game.outs = 2 Then
                                    Game.GetResultPtr.resSecond = 4
                                    Game.GetResultPtr.description = "Error on " & gblPositions(Game.currentFieldPos) & ". Batter safe and runner scores"
                                Else
                                    Game.GetResultPtr.description = "Error on " & gblPositions(Game.currentFieldPos) & ". Batter safe and runner to third"
                                End If
                            Case "E3010"
                                If Game.outs = 2 Then
                                    Game.GetResultPtr.resSecond = 3
                                    Game.GetResultPtr.description = "Error on " & gblPositions(Game.currentFieldPos) & ". Batter safe and runner to third"
                                Else
                                    Game.GetResultPtr.description = "Error on " & gblPositions(Game.currentFieldPos) & ". Batter safe and runner holds"
                                End If
                            Case "E4010"
                                Call Game.GetResultPtr.ChartLookup(playResultOld)
                                tempValue = playResultOld.Substring(0, 2)
                                Game.GetResultPtr.resSecond = 4
                                If tempValue = "1B" Then
                                    Game.GetResultPtr.resBatter = 2
                                    Game.GetResultPtr.description += _
                                            ". Error on " & gblPositions(Game.currentFieldPos) & _
                                            ". Batter to second and runner scores"
                                ElseIf tempValue = "2B" Then
                                    Game.GetResultPtr.resBatter = 3
                                    Game.GetResultPtr.description += _
                                            ". Error on " & gblPositions(Game.currentFieldPos) & _
                                            ". Batter to third and runner scores"
                                Else '3B
                                    Game.GetResultPtr.resBatter = 4
                                    Game.GetResultPtr.description += _
                                            ". Error on " & gblPositions(Game.currentFieldPos) & _
                                            ". Batter and runner scores"
                                End If
                            Case "E5010"
                                Call Game.GetResultPtr.ChartLookup(playResultOld)
                                tempValue = playResultOld.Substring(0, 2)
                                Game.GetResultPtr.resSecond = 4
                                If Game.PTeam.GetBatterPtr(Game.GetFielderNum).throwRating = "5" Then
                                    bolCountRun = True
                                    bolGiveRBI = True
                                    Game.GetResultPtr.resBatter = 0
                                    Game.GetResultPtr.description += _
                                            ". Error on " & gblPositions(Game.currentFieldPos) & _
                                            ". Batter is thrown out attempting extra base, but the runner scores"
                                ElseIf tempValue = "1B" Then
                                    Game.GetResultPtr.resBatter = 2
                                    Game.GetResultPtr.description += _
                                            ". Error on " & gblPositions(Game.currentFieldPos) & _
                                            ". Batter to second and runner scores"
                                ElseIf tempValue = "2B" Then
                                    Game.GetResultPtr.resBatter = 3
                                    Game.GetResultPtr.description += _
                                            ". Error on " & gblPositions(Game.currentFieldPos) & _
                                            ". Batter to third and runner scores"
                                Else '3B
                                    Game.GetResultPtr.resBatter = 4
                                    Game.GetResultPtr.description += _
                                            ". Error on " & gblPositions(Game.currentFieldPos) & _
                                            ". Batter and runner scores"
                                End If
                            Case "E3100"
                                If Game.outs = 2 Then
                                    Game.GetResultPtr.resThird = 4
                                    Game.GetResultPtr.description = "Error on " & gblPositions(Game.currentFieldPos) & ". Batter safe and runner scores"
                                Else
                                    Game.GetResultPtr.description = "Error on " & gblPositions(Game.currentFieldPos) & ". Batter safe and runner holds"
                                End If
                            Case "E4100"
                                Call Game.GetResultPtr.ChartLookup(playResultOld)
                                tempValue = playResultOld.Substring(0, 2)
                                Game.GetResultPtr.resThird = 4
                                If tempValue = "1B" Then
                                    Game.GetResultPtr.resBatter = 3
                                    Game.GetResultPtr.description += _
                                            ". Error on " & gblPositions(Game.currentFieldPos) & _
                                            ". Batter to third and runner scores"
                                Else '2B,3B
                                    Game.GetResultPtr.resBatter = 4
                                    Game.GetResultPtr.description += _
                                            ". Error on " & gblPositions(Game.currentFieldPos) & _
                                            ". Batter and runner scores"
                                End If
                            Case "E5100"
                                Call Game.GetResultPtr.ChartLookup(playResultOld)
                                tempValue = playResultOld.Substring(0, 2)
                                Game.GetResultPtr.resThird = 4
                                If .obr = "A" Or .obr = "B" Then
                                    If tempValue = "1B" Then
                                        Game.GetResultPtr.resBatter = 2
                                        Game.GetResultPtr.description += _
                                                ". Error on " & gblPositions(Game.currentFieldPos) & _
                                                ". Batter to second and runner scores"
                                    ElseIf tempValue = "2B" Then
                                        Game.GetResultPtr.resBatter = 3
                                        Game.GetResultPtr.description += _
                                                ". Error on " & gblPositions(Game.currentFieldPos) & _
                                                ". Batter to third and runner scores"
                                    Else
                                        Game.GetResultPtr.resBatter = 4
                                        Game.GetResultPtr.description += _
                                                ". Error on " & gblPositions(Game.currentFieldPos) & _
                                                ". Batter and runner scores"
                                    End If
                                Else
                                    bolCountRun = True
                                    bolGiveRBI = True
                                    Game.GetResultPtr.resBatter = 0
                                    Game.GetResultPtr.description += _
                                            ". Error on " & gblPositions(Game.currentFieldPos) & _
                                            ". Batter is thrown out attempting extra base, but the runner scores"
                                End If
                            Case "E4011"
                                Call Game.GetResultPtr.ChartLookup(playResultOld)
                                tempValue = playResultOld.Substring(0, 2)
                                Game.GetResultPtr.resFirst = 4
                                Game.GetResultPtr.resSecond = 4
                                If tempValue = "1B" Then
                                    Game.GetResultPtr.resBatter = 2
                                    Game.GetResultPtr.description += _
                                            ". Error on " & gblPositions(Game.currentFieldPos) & _
                                            ". Batter to second and both runners score"
                                ElseIf tempValue = "2B" Then
                                    Game.GetResultPtr.resBatter = 3
                                    Game.GetResultPtr.description += _
                                            ". Error on " & gblPositions(Game.currentFieldPos) & _
                                            ". Batter to third and both runners score"
                                Else
                                    Game.GetResultPtr.resBatter = 4
                                    Game.GetResultPtr.description += _
                                            ". Error on " & gblPositions(Game.currentFieldPos) & _
                                            ". Batter and runners score"
                                End If
                            Case "E5011"
                                'No error
                                playResult = playResultOld
                                Call Game.GetResultPtr.ChartLookup(playResult)
                                chargeError = False
                                isError = False
                            Case "E3110"
                                If Game.outs = 2 Then
                                    Game.GetResultPtr.resThird = 4
                                    Game.GetResultPtr.resSecond = 3
                                    Game.GetResultPtr.description = "Error on " & gblPositions(Game.currentFieldPos) & ". Batter safe and runners advance"
                                Else
                                    Game.GetResultPtr.description = "Error on " & gblPositions(Game.currentFieldPos) & ". Batter safe and runners hold"
                                End If
                            Case "E4110"
                                Call Game.GetResultPtr.ChartLookup(playResultOld)
                                Game.GetResultPtr.resThird = 4
                                Game.GetResultPtr.resSecond = 4
                                Game.GetResultPtr.resBatter = 4
                                Game.GetResultPtr.description += ". Error on " & _
                                        gblPositions(Game.currentFieldPos) & ". Batter and runners score"
                            Case "E5110"
                                'No error
                                playResult = playResultOld
                                Call Game.GetResultPtr.ChartLookup(playResult)
                                chargeError = False
                                isError = False
                            Case "E2101"
                                Game.GetResultPtr.resThird = 4
                                If Game.BTeam.GetBatterPtr(FirstBase.runner).obr = "E" Then
                                    Game.GetResultPtr.resFirst = 3
                                    Game.GetResultPtr.description = "Error on " & gblPositions(Game.currentFieldPos) & ". Batter to second and runners advance two"
                                Else
                                    Game.GetResultPtr.resFirst = 4
                                    Game.GetResultPtr.description = "Error on " & gblPositions(Game.currentFieldPos) & ". Batter to second and both runners score"
                                End If
                            Case "E3101"
                                If Game.outs = 2 Then
                                    Game.GetResultPtr.resThird = 4
                                    Game.GetResultPtr.resFirst = 2
                                    Game.GetResultPtr.description = "Error on " & gblPositions(Game.currentFieldPos) & ". Batter safe and runners advance"
                                Else
                                    Game.GetResultPtr.resThird = 3
                                    Game.GetResultPtr.resFirst = 2
                                    Game.GetResultPtr.description = "Error on " & gblPositions(Game.currentFieldPos) & ". Batter safe and runner on third holds"
                                End If
                            Case "E4101"
                                Call Game.GetResultPtr.ChartLookup(playResultOld)
                                tempValue = playResultOld.Substring(0, 2)
                                Game.GetResultPtr.resThird = 4
                                Game.GetResultPtr.description += ". Error on " & _
                                        gblPositions(Game.currentFieldPos) & _
                                        ". Batter and baserunners advance one extra base"
                                If tempValue = "1B" Then
                                    Game.GetResultPtr.resBatter = 2
                                    Game.GetResultPtr.resFirst = 3
                                ElseIf tempValue = "2B" Then
                                    Game.GetResultPtr.resBatter = 3
                                    Game.GetResultPtr.resFirst = 4
                                Else
                                    Game.GetResultPtr.resBatter = 4
                                    Game.GetResultPtr.resFirst = 4
                                End If
                            Case "E5101"
                                Call Game.GetResultPtr.ChartLookup(playResultOld)
                                tempValue = playResultOld.Substring(0, 2)
                                Game.GetResultPtr.resThird = 4
                                Game.GetResultPtr.description += ". Error on " & _
                                        gblPositions(Game.currentFieldPos) & _
                                        ". Batter holds to hit and baserunners advance one extra base"
                                If tempValue = "1B" Then
                                    Game.GetResultPtr.resFirst = 3
                                Else
                                    Game.GetResultPtr.resFirst = 4
                                End If
                            Case "E1111"
                                If (Game.BTeam.GetBatterPtr(SecondBase.runner).obr = "A" Or _
                                        Game.BTeam.GetBatterPtr(SecondBase.runner).obr = "B") And Game.outs = 2 Then
                                    Game.GetResultPtr.resSecond = 4
                                    Game.GetResultPtr.description = "Error on " & gblPositions(Game.currentFieldPos) & ". Batter safe and runners on second and third score"
                                Else
                                    Game.GetResultPtr.description = "Error on " & gblPositions(Game.currentFieldPos) & ". Batter safe and runners advance"
                                End If
                            Case "E2111"
                                If (Game.BTeam.GetBatterPtr(FirstBase.runner).obr = "A" Or _
                                        Game.BTeam.GetBatterPtr(FirstBase.runner).obr = "B") And Game.outs = 2 Then
                                    Game.GetResultPtr.resFirst = 4
                                    Game.GetResultPtr.description = "Error on " & gblPositions(Game.currentFieldPos) & ". Batter to second and all runners score"
                                Else
                                    Game.GetResultPtr.description = "Error on " & gblPositions(Game.currentFieldPos) & ". Batter to second and runners advance two"
                                End If
                            Case "E4111"
                                'No error
                                playResult = playResultOld
                                Call Game.GetResultPtr.ChartLookup(playResult)
                                chargeError = False
                                isError = False
                            Case "E5111"
                                Call Game.GetResultPtr.ChartLookup(playResultOld)
                                If Game.PTeam.GetBatterPtr(Game.GetFielderNum).cdAct = "0" Then
                                    Game.GetResultPtr.resThird = 4
                                    Game.GetResultPtr.resSecond = 4
                                    Game.GetResultPtr.resFirst = 4
                                    Game.GetResultPtr.resBatter = 4
                                    Game.GetResultPtr.description += ". Error on " & _
                                            gblPositions(Game.currentFieldPos) & ". Batter and runners score"
                                Else
                                    playResult = playResultOld
                                    chargeError = False
                                    isError = False
                                End If
                            Case "F2ERR"
                                chargeError = True
                                isError = True
                                playResult = conNoPlay
                            Case "FIERR", "LOERR"
                                If Game.outs < 2 Then
                                    Game.GetResultPtr.description += " Runners advance if forced to. E" & _
                                                Game.currentFieldPos
                                Else
                                    Game.GetResultPtr.description += " Runners advance. E" & Game.currentFieldPos
                                    If BaseSituation() = "010" Then
                                        Game.GetResultPtr.resSecond = 3
                                    ElseIf BaseSituation() = "100" Or BaseSituation() = "101" Then
                                        Game.GetResultPtr.resThird = 4
                                    ElseIf BaseSituation() = "110" Then
                                        Game.GetResultPtr.resThird = 4
                                        Game.GetResultPtr.resSecond = 3
                                    End If
                                End If
                            Case "FOERR"
                                If Game.outs < 2 Then
                                    Game.GetResultPtr.description += " Batter safe and runners advance. E" & _
                                            Game.currentFieldPos
                                Else
                                    Game.GetResultPtr.description += " Batter safe and runners advance two. E" & _
                                            Game.currentFieldPos
                                    If BaseSituation() = "001" Or BaseSituation() = "101" Then
                                        Game.GetResultPtr.resFirst = 3
                                    ElseIf BaseSituation() = "010" Or BaseSituation() = "110" Then
                                        Game.GetResultPtr.resSecond = 4
                                    ElseIf BaseSituation() = "011" Or BaseSituation() = "111" Then
                                        Game.GetResultPtr.resFirst = 3
                                        Game.GetResultPtr.resSecond = 4
                                    End If
                                End If
                            Case Else
                                If Game.GetResultPtr.description.ToUpper.IndexOf("ERROR ON.") > -1 Then
                                    Game.GetResultPtr.description = "Error on " & gblPositions(Game.currentFieldPos) & Game.GetResultPtr.description.Substring(8)
                                End If
                        End Select
                        'reset base advance flag if error occurred
                        baseAdvance = baseAdvance And Not isError

                        If chargeError Then
                            'Charge error
                            tempIntegerValue = Game.currentFieldPos
                            If tempIntegerValue = 1 Then
                                'Pitcher
                                Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).PitStatPtr.e += 1
                            Else
                                tempIntegerValue = Game.GetFielderNumByPos(tempIntegerValue)
                                Game.PTeam.GetBatterPtr(tempIntegerValue).BatStatPtr.E += 1
                            End If
                            Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).eOuts += errorOuts
                            gbolSOE = True
                        End If
                    End With
                End If
            End If
        Catch ex As Exception
            Call MsgBox("Error in CheckForError. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return isError
    End Function

    ''' <summary>
    ''' Defensive Option Play is a situation where the defense has a decision whether to throw an aggressive runner,
    ''' or take the sure out at first.
    ''' </summary>
    ''' <param name="FACCard"></param>
    ''' <remarks></remarks>
    Private Sub HandleDefensiveOptionPlay(ByRef FACCard As FACCard)
        Dim rangeDOP As String = ""
        Try
            Select Case Game.BTeam.GetBatterPtr(ThirdBase.runner).obr
                Case "A"
                    rangeDOP = "11-48"
                Case "B"
                    rangeDOP = "11-42"
                Case "C"
                    rangeDOP = "11-35"
                Case "D"
                    rangeDOP = "11-32"
                Case "E"
                    rangeDOP = "11-28"
            End Select
            With Game.GetResultPtr
                If MsgBox("Defensive Option Play. Will the runner on third try to score? Chances are " & _
                        rangeDOP, MsgBoxStyle.YesNo, Game.GetResultPtr.description) = MsgBoxResult.Yes Then
                    If MsgBox("Will the defense throw home?", MsgBoxStyle.YesNo, _
                            Game.GetResultPtr.description) = MsgBoxResult.Yes Then
                        FACCard = GetFAC("defOptPlay", "Random")
                        If InRange(rangeDOP, CInt(FACCard.Random)) Then
                            'All safe
                            .resBatter = 1
                            .resFirst = IIF(FirstBase.occupied, 2, 5)
                            .resSecond = IIF(SecondBase.occupied, 3, 5)
                            .resThird = 4 'Third is always occupied for defensive option play
                            .description += ". Here comes the throw to home and he's ... SAFE!"
                            .rbi = 1
                            .ip = 0
                            .run = 1
                        Else
                            'out at home, but batter safe runners advance
                            .resBatter = 1
                            .resFirst = IIF(FirstBase.occupied, 2, 5)
                            .resSecond = IIF(SecondBase.occupied, 3, 5)
                            .resThird = 0 'Third is always occupied for defensive option play
                            .description += ". Here comes the throw to home and he's ... OUT AT THE PLATE!"
                            .rbi = 0
                            .ip = 1
                            .run = 0
                            .a = Game.currentFieldPos.ToString
                        End If
                    Else
                        'Runner on third scores, batter is out, runners advance
                        .resBatter = 0
                        .resFirst = IIF(FirstBase.occupied, 2, 5)
                        .resSecond = IIF(SecondBase.occupied, 3, 5)
                        .resThird = 4 'Third is always occupied for defensive option play
                        .description += " and the throw is to first for the sure out."
                        .rbi = 1
                        .ip = 1
                        .run = 1
                        .a = IIF(Game.currentFieldPos.ToString = "3", "", Game.currentFieldPos.ToString)
                        .po = "3"
                    End If
                Else
                    'Batter out, runners hold unless forced
                    .resBatter = 0
                    .resFirst = IIF(FirstBase.occupied, 2, 5)
                    .resSecond = IIF(SecondBase.occupied, 2, 5)
                    .resThird = 3 'Third is always occupied for defensive option play
                    .description += ". The runner is checked at third and the throw is to first for the out."
                    .rbi = 0
                    .ip = 1
                    .run = 0
                    .a = IIF(Game.currentFieldPos.ToString = "3", "", Game.currentFieldPos.ToString)
                    .po = "3"
                End If
            End With
        Catch ex As Exception
            Call MsgBox("Error in HandleDefensiveOptionPlay. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' Handle conditional situations on normal outs. For example, a fast runner advancing on an out where
    ''' normally a runner would hold. Another example would be an exceptional fielder doubling up a runner on line out, 
    ''' where normal fielders cannot.
    ''' </summary>
    ''' <param name="playResult"></param>
    ''' <remarks></remarks>
    Private Sub HandleOutExceptions(ByRef playResult As String)
        Dim stringPosition As Integer
        Dim tempValue As String

        Try
            Select Case playResult
                Case "F2101"
                    If Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(2)).cdAct = "0" Then
                        'No play
                        Game.GetResultPtr.description = ""
                        playResult = conNoPlay
                    Else
                        'Play stands
                        stringPosition = Game.GetResultPtr.description.IndexOf("Cnd:")
                        If stringPosition > -1 Then
                            Game.GetResultPtr.description = Game.GetResultPtr.description.Substring(0, stringPosition)
                        End If
                    End If
                Case "FD7001", "FD7101"
                    If Game.BTeam.GetBatterPtr(FirstBase.runner).obr = "A" And Game.outs < 2 Then
                        Game.GetResultPtr.resFirst = 2
                        Game.GetResultPtr.description += ". Runner on first advances"
                    End If
                Case "FD8001", "FD8101", "FD9111"
                    tempValue = Game.BTeam.GetBatterPtr(FirstBase.runner).obr
                    If (tempValue = "A" Or tempValue = "B") And Game.outs < 2 Then
                        Game.GetResultPtr.resFirst = 2
                        Game.GetResultPtr.description += ". Runner on first advances"
                    End If
                Case "FD9001"
                    tempValue = Game.BTeam.GetBatterPtr(FirstBase.runner).obr
                    If (tempValue = "A" Or tempValue = "B" Or tempValue = "C") And Game.outs < 2 Then
                        Game.GetResultPtr.resFirst = 2
                        Game.GetResultPtr.description += ". Runner on first advances"
                    End If
                Case "F8010", "F9011", "FD7011"
                    If Game.BTeam.GetBatterPtr(SecondBase.runner).obr = "A" And Game.outs < 2 Then
                        Game.GetResultPtr.resSecond = 3
                        Game.GetResultPtr.description += ". Runner on second advances"
                    End If
                Case "F9010", "FD7010", "FD8111", "FD7110"
                    tempValue = Game.BTeam.GetBatterPtr(SecondBase.runner).obr
                    If (tempValue = "A" Or tempValue = "B") And Game.outs < 2 Then
                        Game.GetResultPtr.resSecond = 3
                        Game.GetResultPtr.description += ". Runner on second advances"
                    End If
                Case "FD8010", "FD8011", "FD8110"
                    tempValue = Game.BTeam.GetBatterPtr(SecondBase.runner).obr
                    If (tempValue = "A" Or tempValue = "B" Or tempValue = "C") And Game.outs < 2 Then
                        Game.GetResultPtr.resSecond = 3
                        Game.GetResultPtr.description += ". Runner on second advances"
                    End If
                Case "F4110", "F7110", "Gx3100BA", "G1A100IN", "G3110BA", "G5110BA", "Gx3110BA", "G1A110IN", "G1A110BA", "G3A110IN"
                    tempValue = Game.BTeam.GetBatterPtr(ThirdBase.runner).obr
                    If (tempValue = "A" Or tempValue = "B") And Game.outs < 2 Then
                        Game.GetResultPtr.resThird = 4
                        Game.GetResultPtr.description += ". Runner on third scores"
                        If playResult.Substring(0, 1) = "F" Then
                            Game.GetResultPtr.ab = 0 'SAC Fly
                        End If
                        Game.GetResultPtr.rbi = 1
                        Game.GetResultPtr.run = 1
                    End If
                Case "F7111", "F8111", "F7100", "F8100", "F9100", "F9111", "F8110", "Gx4100BA", "Gx5100BA", "G4110BA", "G6110BA", "Gx4110BA"
                    tempValue = Game.BTeam.GetBatterPtr(ThirdBase.runner).obr
                    If (tempValue = "A" Or tempValue = "B" Or tempValue = "C") And Game.outs < 2 Then
                        Game.GetResultPtr.resThird = 4
                        Game.GetResultPtr.description += ". Runner on third scores"
                        If playResult.Substring(0, 1) = "F" Then
                            Game.GetResultPtr.ab = 0 'SAC Fly
                        End If
                        Game.GetResultPtr.rbi = 1
                        Game.GetResultPtr.run = 1
                    Else
                        Game.GetResultPtr.description += ". Runners hold"
                    End If
                Case "F7101", "F8101", "F9101", "F9110", "Gx6100BA"
                    tempValue = Game.BTeam.GetBatterPtr(ThirdBase.runner).obr
                    If (tempValue = "A" Or tempValue = "B" Or tempValue = "C") And Game.outs < 2 Then
                        Game.GetResultPtr.resThird = 4
                        Game.GetResultPtr.description += ". Runner on third scores"
                        If playResult.Substring(0, 1) = "F" Then
                            Game.GetResultPtr.ab = 0 'SAC Fly
                        End If
                        Game.GetResultPtr.rbi = 1
                        Game.GetResultPtr.run = 1
                    End If
                Case "G31A100IN", "G31A100BA"
                    tempValue = Game.BTeam.GetBatterPtr(ThirdBase.runner).obr
                    If (tempValue <> "E") And Game.outs < 2 Then
                        Game.GetResultPtr.resThird = 4
                        Game.GetResultPtr.description += ". Runner on third scores"
                        Game.GetResultPtr.rbi = 1
                        Game.GetResultPtr.run = 1
                    End If
                Case "FD9101"
                    If (Game.BTeam.GetBatterPtr(FirstBase.runner).obr <> "E") And Game.outs < 2 Then
                        Game.GetResultPtr.resFirst = 2
                        Game.GetResultPtr.description += ". Runner on first advances"
                    End If
                Case "G1001", "G5001"
                    If Game.outs < 2 Then
                        If Game.BTeam.GetBatterPtr(Game.currentBatter).obr = "A" Then
                            Game.GetResultPtr.resBatter = 1
                            Game.GetResultPtr.description += " SAFE"
                            Game.GetResultPtr.po = "4"
                            Game.GetResultPtr.a = Game.currentFieldPos.ToString
                            Game.currentFieldPos.ToString()
                        Else
                            Game.GetResultPtr.description += " DOUBLE PLAY"
                        End If
                    End If
                Case "G3101BA"
                    If Game.outs < 2 Then
                        If Game.BTeam.GetBatterPtr(Game.currentBatter).obr = "A" Then
                            Game.GetResultPtr.resBatter = 1
                            Game.GetResultPtr.description += " SAFE"
                            Game.GetResultPtr.po = "6"
                            Game.GetResultPtr.a = Game.currentFieldPos.ToString
                        Else
                            Game.GetResultPtr.description += " DOUBLE PLAY"
                        End If
                    End If
                Case "G2001", "G3001"
                    If Game.outs < 2 Then
                        tempValue = Game.BTeam.GetBatterPtr(Game.currentBatter).obr
                        If tempValue = "A" Or tempValue = "B" Then
                            Game.GetResultPtr.resBatter = 1
                            Game.GetResultPtr.description += " SAFE"
                            Game.GetResultPtr.po = "6"
                            Game.GetResultPtr.a = Game.currentFieldPos.ToString
                        Else
                            Game.GetResultPtr.description += " DOUBLE PLAY"
                        End If
                    End If
                Case "G1010", "Gx1010"
                    If Game.BTeam.GetBatterPtr(SecondBase.runner).obr = "A" And Game.outs < 2 Then
                        Game.GetResultPtr.resSecond = 3
                        Game.GetResultPtr.description += ". Runner to third"
                    End If
                Case "G3111IN"
                    If Game.outs < 2 Then
                        If Game.BTeam.GetBatterPtr(Game.currentBatter).obr <> "E" Then
                            Game.GetResultPtr.description += " is safe at first"
                        Else
                            Game.GetResultPtr.resBatter = 0
                            Game.GetResultPtr.description += " is too slow and doubled off at first"
                            Game.GetResultPtr.po = "23"
                            Game.GetResultPtr.a = "23"
                        End If
                    End If
                Case "G1A111IN", "G1A111BA"
                    If Game.outs < 2 Then
                        If GetCD((Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).field)) = "0" Then
                            Game.GetResultPtr.description += " Batter safe"
                        Else
                            Game.GetResultPtr.resBatter = 0
                            Game.GetResultPtr.description += " Batter is doubled off by slick fielding catcher"
                            Game.GetResultPtr.po = "23"
                            Game.GetResultPtr.a = "12"
                        End If
                    End If
                Case "L3001", "L3101"
                    If Game.outs < 2 And Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(3)).cdAct = "2" Then
                        Game.GetResultPtr.resFirst = 0
                        Game.GetResultPtr.description = _
                                "Line out to first. Runner on first is doubled off on a great play by " & _
                                Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(3)).player
                        Game.GetResultPtr.po = "33"
                    End If
                Case "L4010"
                    If Game.outs < 2 And Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(4)).cdAct = "2" Then
                        Game.GetResultPtr.resSecond = 0
                        Game.GetResultPtr.description = _
                                "Line out to second. Runner on second is doubled off on a great play by " & _
                                Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(4)).player
                        Game.GetResultPtr.po = "44"
                    End If
                Case "L5100"
                    If Game.outs < 2 And Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(5)).cdAct = "2" Then
                        Game.GetResultPtr.resThird = 0
                        Game.GetResultPtr.description = _
                                "Line out to third. Runner on third is doubled off on a great play by " & _
                                Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(5)).player
                        Game.GetResultPtr.po = "55"
                    End If
                Case "L1011"
                    If Game.outs < 2 And GetCD((Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).field)) = "2" Then
                        Game.GetResultPtr.resFirst = 0
                        Game.GetResultPtr.description = _
                                "Line out to pitcher. Runner on first is doubled off on a great play by " & _
                                Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).player
                        Game.GetResultPtr.po = "13"
                        Game.GetResultPtr.a = "1"
                    End If
                Case "L6011"
                    If Game.outs < 2 And Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(6)).cdAct = "2" Then
                        Game.GetResultPtr.resSecond = 0
                        Game.GetResultPtr.description = _
                                "Line out to short. Runner on second is doubled off on a great play by " & _
                                Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(6)).player
                        Game.GetResultPtr.po = "66"
                    End If
                Case "L1111"
                    If Game.outs < 2 And GetCD((Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).field)) = "2" Then
                        Game.GetResultPtr.resThird = 0
                        Game.GetResultPtr.description = _
                                "Line out to pitcher. Runner on third is doubled off on a great play by " & _
                                Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).player
                        Game.GetResultPtr.po = "15"
                        Game.GetResultPtr.a = "1"
                    End If
            End Select
        Catch ex As Exception
            Call MsgBox("Error in HandleOutExceptions. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' Handle BD (clutch hitting) situations
    ''' </summary>
    ''' <param name="playResult"></param>
    ''' <param name="FACCard"></param>
    ''' <remarks></remarks>
    Private Sub HandleBD(ByRef playResult As String, ByRef FACCard As FACCard)
        Dim fullBD As Boolean

        Try
            fullBD = AppSettings("FullBD").ToString.ToUpper = "ON"
            With Game.BTeam.GetBatterPtr(Game.currentBatter)
                FACCard = GetFAC("checkBD", "Random")
                If .bd = "0" Or BaseSituation() = "000" Or (CInt(FACCard.Random) > 16 And Not fullBD) Then
                    'Allow only 10% of BDs
                    playResult = conNoPlay
                Else
                    FACCard = GetFAC("bd", "Random")
                    If .bd = "1" Then
                        Select Case CInt(FACCard.Random)
                            Case 11 To 24
                                playResult += "1"
                            Case 25 To 26
                                playResult += "2"
                            Case 27
                                playResult += "3"
                            Case Else
                                playResult = conNoPlay
                        End Select
                    ElseIf .bd = "2" Then
                        Select Case CInt(FACCard.Random)
                            Case 11 To 28
                                playResult += "1"
                            Case 31 To 34
                                playResult += "2"
                            Case 35 To 37
                                playResult += "3"
                            Case Else
                                playResult = conNoPlay
                        End Select
                    ElseIf .bd = "3" Then
                        Select Case CInt(FACCard.Random)
                            Case 11 To 42
                                playResult += "1"
                            Case 43 To 44
                                playResult += "2"
                            Case 45 To 48
                                playResult += "3"
                            Case Else
                                playResult = conNoPlay
                        End Select
                    End If
                    If playResult <> conNoPlay Then
                        playResult += BaseSituation()
                        Call Game.GetResultPtr.ChartLookup(playResult)
                    End If
                End If
            End With
        Catch ex As Exception
            Call MsgBox("Error in HandleBD. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' Handle CD (clutch defense) situations
    ''' </summary>
    ''' <param name="playResult"></param>
    ''' <param name="FACCard"></param>
    ''' <remarks></remarks>
    Private Sub HandleCD(ByRef playResult As String, ByRef FACCard As FACCard)
        Dim cdValue As Integer
        Dim tempValue As String = ""
        Dim stringPosition As Integer
        Dim fullCD As Boolean

        Try
            fullCD = AppSettings("FullCD").ToString.ToUpper = "ON"
            FACCard = GetFAC("cdCheck", "Random")
            If BaseSituation.Substring(1) = "00" Or (CInt(FACCard.Random) > 16 And Not fullCD) Then
                'Only valid with men on 1st or 2nd, and even then, just 10% of the time
                playResult = conNoPlay
            Else
                'Get Position
                FACCard = GetFAC("cdPosition", "CD")
                For i As Integer = 1 To 9
                    If FACCard.CD = gblPositions(i) Then
                        Game.currentFieldPos = i
                    End If
                    If FACCard.CD = "P" Then
                        cdValue = CInt(Val(GetCD((Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).field))))
                    Else
                        With Game.PTeam.GetBatterPtr(Game.PTeam.GetPlayerNum(i))
                            If .position = FACCard.CD Then
                                If .cd.IndexOf(.position) > -1 Or (.position.Substring(.position.Length - 1) = _
                                        "F" And .cd.Substring(.cd.Length - 1) = "F") Then
                                    cdValue = CInt(Val(.cd))
                                End If
                            End If
                        End With
                    End If
                Next i
                Select Case Game.currentFieldPos
                    Case 1, 3 To 6
                        tempValue = "INF"
                    Case 2
                        tempValue = "C"
                    Case 7 To 9
                        tempValue = "OF"
                End Select
                'Determine result
                FACCard = GetFAC("cd", "Random")
                If cdValue = 0 Then
                    Select Case CInt(FACCard.Random)
                        Case 11 To 21
                            playResult += "1" & tempValue
                        Case 22 To 32
                            playResult += IIF(tempValue = "C", "1", "2") & tempValue
                        Case Else
                            playResult = conNoPlay
                    End Select
                ElseIf cdValue = 1 Then
                    Select Case CInt(FACCard.Random)
                        Case 11 To 42
                            playResult += "1" & tempValue
                        Case 43 To 55
                            playResult += IIF(tempValue = "C", "1", "2") & tempValue
                        Case Else
                            playResult = conNoPlay
                    End Select
                Else
                    Select Case CInt(FACCard.Random)
                        Case 11 To 56
                            playResult += "1" & tempValue
                        Case 57 To 78
                            playResult += IIF(tempValue = "C", "1", "2") & tempValue
                        Case Else
                            playResult = conNoPlay
                    End Select
                End If
                'set fielding stats
                Select Case playResult
                    Case "CD1INF"
                        If Game.outs < 2 Then
                            'unassisted DP
                            Game.GetResultPtr.po += Game.currentFieldPos.ToString
                        Else
                            Game.GetResultPtr.po = Game.currentFieldPos.ToString
                        End If
                    Case "CD1C", "CD1OF", "CD2OF"
                        If Game.outs < 2 Then
                            Game.GetResultPtr.po += Game.currentFieldPos.ToString
                            Game.GetResultPtr.a = Game.currentFieldPos.ToString
                        Else
                            Game.GetResultPtr.po = Game.currentFieldPos.ToString
                        End If
                    Case "CD2INF"
                        If Game.currentFieldPos.ToString <> Game.GetResultPtr.po Then
                            Game.GetResultPtr.a = Game.currentFieldPos.ToString
                        End If
                End Select

                If playResult <> conNoPlay Then
                    playResult += BaseSituation()
                    Call Game.GetResultPtr.ChartLookup(playResult)
                    stringPosition = Game.GetResultPtr.description.IndexOf("Cnd:")
                    If stringPosition > -1 Then
                        tempValue = Game.GetResultPtr.description.Substring(Game.GetResultPtr.description.Length - 1)
                        If SecondBase.occupied Then
                            If Game.BTeam.GetBatterPtr(SecondBase.runner).obr = "A" Then
                                'Runner is not doubled off
                                Game.GetResultPtr.resSecond = 2
                                Game.GetResultPtr.po = Game.currentFieldPos.ToString
                                Game.GetResultPtr.a = ""
                            End If
                        Else 'First is occupied
                            If Game.BTeam.GetBatterPtr(FirstBase.runner).obr = "A" Then
                                'Runner is not doubled off
                                Game.GetResultPtr.resFirst = 1
                                Game.GetResultPtr.po = Game.currentFieldPos.ToString
                                Game.GetResultPtr.a = ""
                            End If
                        End If
                    End If
                End If
                'Modify description
                stringPosition = Game.GetResultPtr.description.ToUpper.IndexOf("LINE OUT TO FIELDER")
                If stringPosition > -1 Then
                    Game.GetResultPtr.description = "Line out to " & gblPositions(Game.currentFieldPos) & _
                            Game.GetResultPtr.description.Substring(stringPosition + 19)
                End If
                stringPosition = Game.GetResultPtr.description.ToUpper.IndexOf("GROUND OUT TO FIELDER")
                If stringPosition > -1 Then
                    Game.GetResultPtr.description = "Ground out to " & gblPositions(Game.currentFieldPos) & _
                            Game.GetResultPtr.description.Substring(stringPosition + 21)
                End If
            End If
        Catch ex As Exception
            Call MsgBox("Error in HandleCD. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' handle Z play (unusual plays)
    ''' </summary>
    ''' <param name="playResult"></param>
    ''' <param name="FACCard"></param>
    ''' <remarks></remarks>
    Private Sub HandleZPlay(ByRef playResult As String, ByRef FACCard As FACCard)
        Dim stringPosition As Integer

        Try
            FACCard = GetFAC("zPlay", "Random")
            Select Case CInt(FACCard.Random)
                Case 11 To 13, 15 To 16
                    playResult += FACCard.Random & "A"
                    bolAnyBS = True
                Case 14, 17 To 18, 27, 31 To 44
                    playResult += FACCard.Random
                Case 21 To 26
                    playResult = conNoPlay
                Case 28
                    'Catcher drops 3 strike
                    If Game.outs = 2 Or Not FirstBase.occupied Then
                        playResult += FACCard.Random
                    Else
                        playResult = conNoPlay
                    End If
                Case 45 To 78
                    'Z play fielding
                    FACCard = GetFAC("zField", "Random")
                    playResult += "fld"
                    Select Case CInt(FACCard.Random)
                        Case 11 To 14
                            playResult += "11-14"
                        Case 15 To 18
                            playResult += "15-18"
                        Case 21 To 24
                            playResult += "21-24"
                        Case 25 To 28
                            playResult += "25-28"
                        Case 31 To 34
                            playResult += "31-34"
                        Case 35 To 38
                            playResult += "35-38"
                        Case 41 To 44
                            playResult += "41-44"
                        Case 45 To 48
                            playResult += "45-48"
                        Case 51 To 54
                            playResult += "51-54"
                        Case 55 To 63
                            playResult += "55-63"
                        Case 64 To 74
                            playResult += "64-74"
                        Case 75 To 83
                            playResult += "75-83"
                        Case 84 To 87
                            playResult += "84-87"
                        Case 88
                            playResult += "88"
                        Case Else
                            playResult = conNoPlay
                            Call MsgBox("Cannot lookup Z play.", MsgBoxStyle.Exclamation)
                    End Select
                Case 81 To 84
                    'Used to be in Z Play injury range below. Effort to reduce in game injuries.
                    playResult = conNoPlay
                Case 85 To 88
                    'Z play injury
                    FACCard = GetFAC("zInj", "Random")
                    playResult += "inj"
                    Select Case CInt(FACCard.Random)
                        Case 11 To 12
                            playResult += "11-12" & "A"
                            bolAnyBS = True
                        Case 13 To 14
                            playResult += "13-14" & "A"
                            bolAnyBS = True
                        Case 15 To 16
                            playResult += "15-16" & "A"
                            bolAnyBS = True
                        Case 17 To 18
                            playResult += "17-18" & "A"
                            bolAnyBS = True
                        Case 21 To 22
                            playResult += "21-22" & "A"
                            bolAnyBS = True
                        Case 23 To 25
                            playResult += "23-25"
                        Case 26 To 28
                            playResult += "26-28"
                        Case 31 To 33
                            playResult += "31-33"
                        Case 34
                            playResult += FACCard.Random
                        Case 48
                            playResult += FACCard.Random & "A"
                            bolAnyBS = True
                        Case 35 To 36
                            playResult += "35-36" & "A"
                            bolAnyBS = True
                        Case 37 To 41
                            playResult += "37-41"
                        Case 42 To 44
                            playResult += "42-44"
                        Case 45 To 47
                            playResult += "45-47"
                        Case 51 To 53
                            playResult += "51-53" & "A"
                            bolAnyBS = True
                        Case 54 To 58
                            playResult += "54-58"
                        Case 61 To 65
                            playResult += "61-65"
                        Case 66 To 68
                            playResult += "66-68"
                        Case 71 To 73
                            playResult += "71-73"
                        Case 74 To 76
                            playResult += "74-76"
                        Case 77 To 81
                            playResult += "77-81"
                        Case 82 To 84
                            playResult += "82-84" & "A"
                            bolAnyBS = True
                        Case 85 To 88
                            playResult += "85-88"
                        Case Else
                            playResult = conNoPlay
                            Call MsgBox("Cannot lookup Z play.", MsgBoxStyle.Exclamation)
                    End Select
                Case Else
                    playResult = conNoPlay
                    Call MsgBox("Cannot lookup Z play.", MsgBoxStyle.Exclamation)
            End Select
            If playResult.Substring(playResult.Length - 1) <> "A" Then 'A means all base situations
                playResult += BaseSituation()
            End If
            Call Game.GetResultPtr.ChartLookup(playResult)
            Select Case playResult
                Case "Z27101"
                    bolCountRun = True
                    bolGiveRBI = True
                Case "Z27111"
                    bolCountRun = True
                    bolGiveRBI = True
                    Game.GetResultPtr.resSecond = 3
                Case "Z32001", "Z32101"
                    'Handle Stolen Base
                    Game.BTeam.GetBatterPtr(FirstBase.runner).BatStatPtr.SB += 1
                    Game.BTeam.GetBatterPtr(FirstBase.runner).BatStatPtr.SBA += 1
                    stringPosition = Game.GetResultPtr.description.IndexOf("Result:SB")
                    If stringPosition > -1 Then
                        Game.GetResultPtr.description = Game.GetResultPtr.description.Substring(0, stringPosition)
                    End If
                Case "Z33001", "Z33101", "Z34001", "Z35001", "Z35011", "Z35101", "Z35111"
                    'Handle Caught Stealing
                    Game.BTeam.GetBatterPtr(FirstBase.runner).BatStatPtr.SBA += 1
                    stringPosition = Game.GetResultPtr.description.IndexOf("Result:CS")
                    If stringPosition > -1 Then
                        Game.GetResultPtr.description = Game.GetResultPtr.description.Substring(0, stringPosition)
                    End If
                Case "Z34101"
                    bolCountRun = True
                    'Handle Caught Stealing
                    Game.BTeam.GetBatterPtr(FirstBase.runner).BatStatPtr.SBA += 1
                    stringPosition = Game.GetResultPtr.description.IndexOf("Result:CS")
                    If stringPosition > -1 Then
                        Game.GetResultPtr.description = Game.GetResultPtr.description.Substring(0, stringPosition)
                    End If
                Case "Z35101", "Z35101"
                    bolCountRun = False
                    'Handle Caught Stealing
                    Game.BTeam.GetBatterPtr(FirstBase.runner).BatStatPtr.SBA += 1
                    stringPosition = Game.GetResultPtr.description.IndexOf("Result:CS")
                    If stringPosition > -1 Then
                        Game.GetResultPtr.description = Game.GetResultPtr.description.Substring(0, stringPosition)
                    End If
                Case "Z36100", "Z36101", "Z36110", "Z36111"
                    'Handle Caught Stealing from third
                    Game.BTeam.GetBatterPtr(ThirdBase.runner).BatStatPtr.SBA += 1
                    stringPosition = Game.GetResultPtr.description.IndexOf("Result:CS")
                    If stringPosition > -1 Then
                        Game.GetResultPtr.description = Game.GetResultPtr.description.Substring(0, stringPosition)
                    End If
                Case "Z38010", "Z38011", "Z38100", "Z38101", "Z38110", "Z38111", "Z41111"
                    bolCountRun = True
                    bolGiveRBI = True
            End Select
        Catch ex As Exception
            Call MsgBox("Error in HandleZPlay. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' Handles Z Play fielding chart
    ''' </summary>
    ''' <param name="playResult"></param>
    ''' <param name="FACCard"></param>
    ''' <remarks></remarks>
    Private Sub HandleZPlayFielding(ByRef playResult As String, ByRef FACCard As FACCard)
        Dim stringPosition As Integer
        Dim fielderIndex As Integer
        Dim baseRunners As Integer
        Dim cdRating As String = ""

        Try
            Select Case CInt(FACCard.Random)
                Case 11 To 34
                    If FirstBase.occupied Then
                        stringPosition = Game.GetResultPtr.description.IndexOf("Cnd:CD")
                        If stringPosition > -1 Then
                            Game.currentFieldPos = CInt(Game.GetResultPtr.description.Substring(stringPosition + 6, 1))
                            If Game.currentFieldPos.ToString = "1" Then
                                'Pitcher
                                cdRating = GetCD((Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).field))
                            Else
                                cdRating = Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(Game.currentFieldPos)).cdAct
                            End If
                            If Game.outs < 2 And cdRating <> "0" Then
                                Game.GetResultPtr.resFirst = 0
                                Game.GetResultPtr.description = Game.GetResultPtr.description.Substring(0, _
                                        stringPosition - 1) & "Throws to second for one. Relay to first..DOUBLE PLAY"
                                If Game.currentFieldPos < 5 Then
                                    'grounder to first or second
                                    Game.GetResultPtr.po = "63"
                                    Game.GetResultPtr.a = Game.currentFieldPos.ToString & "6"
                                Else
                                    'grounder to short or third
                                    Game.GetResultPtr.po = "43"
                                    Game.GetResultPtr.a = Game.currentFieldPos.ToString & "4"
                                End If
                            Else
                                Game.GetResultPtr.description = Game.GetResultPtr.description.Substring(0, _
                                        stringPosition - 1) & "Throws to first for the safe out. Runners advance"
                            End If
                        End If
                    Else
                        playResult = conNoPlay
                    End If
                Case 35 To 54
                    stringPosition = Game.GetResultPtr.description.IndexOf("Cnd:")
                    Game.currentFieldPos = CInt(Val(Game.GetResultPtr.description.Substring(Game.GetResultPtr.description.Length - 1)))
                    fielderIndex = Game.GetFielderNumByPos(Game.currentFieldPos)
                    If Val(Game.PTeam.GetBatterPtr(fielderIndex).errorRating) >= _
                                Val(Game.GetResultPtr.description.Substring(stringPosition + 5, 1)) Then
                        Call Game.GetResultPtr.ChartLookup("1B" & BaseSituation())
                        Game.currentFieldPos = 6
                        Game.GetResultPtr.description = "Single through infield. Runners advance"
                        Game.GetResultPtr.a = ""
                        Game.GetResultPtr.po = ""
                    Else
                        Game.GetResultPtr.description = Game.GetResultPtr.description.Substring(0, stringPosition)
                    End If
                Case 55 To 83
                    stringPosition = Game.GetResultPtr.description.IndexOf("Cnd:")
                    Game.currentFieldPos = CInt(Val(Game.GetResultPtr.description.Substring(Game.GetResultPtr.description.Length - 1)))
                    fielderIndex = Game.GetFielderNumByPos(Game.currentFieldPos)
                    If Val(Game.PTeam.GetBatterPtr(fielderIndex).errorRating) >= _
                                Val(Game.GetResultPtr.description.Substring(stringPosition + 5, 1)) Then
                        Call Game.GetResultPtr.ChartLookup("2B" & BaseSituation())
                        Game.GetResultPtr.description = "Double to " & gblPositions(Game.currentFieldPos) & _
                                ". All runners score"
                        If Game.GetResultPtr.resFirst <> 5 Then Game.GetResultPtr.resFirst = 4
                        Game.GetResultPtr.a = ""
                        Game.GetResultPtr.po = ""
                    Else
                        Game.GetResultPtr.description = Game.GetResultPtr.description.Substring(0, stringPosition)
                    End If
                Case 84 To 87
                    Game.currentFieldPos = 2
                    fielderIndex = Game.GetFielderNumByPos(Game.currentFieldPos)
                    stringPosition = Game.GetResultPtr.description.IndexOf("Cnd:")
                    If Val(Game.PTeam.GetBatterPtr(fielderIndex).errorRating) >= _
                                Val(Game.GetResultPtr.description.Substring(stringPosition + 5, 1)) Then
                        playResult = conNoPlay
                    Else
                        Game.GetResultPtr.description = Game.GetResultPtr.description.Substring(0, stringPosition)
                    End If
                Case 88
                    'Triple play
                    If Game.outs = 0 Then
                        For i As Integer = 0 To BaseSituation.Length - 1
                            baseRunners += CInt(BaseSituation.Substring(i, 1))
                        Next i
                        If baseRunners >= 2 Then '2 or more on base
                            If BaseSituation() = "110" Then
                                Game.GetResultPtr.description += " The throw to third, TRIPLE PLAY!!!"
                                Game.GetResultPtr.resThird = 0
                                Game.GetResultPtr.po += "5"
                                Game.GetResultPtr.a += "4"
                            Else
                                Game.GetResultPtr.description += " The throw to first, TRIPLE PLAY!!!"
                                Game.GetResultPtr.resFirst = 0
                                Game.GetResultPtr.po += "3"
                                If BaseSituation() = "011" Then
                                    Game.GetResultPtr.a += "4"
                                Else
                                    Game.GetResultPtr.a += "5"
                                End If
                            End If
                        End If
                    End If
            End Select
        Catch ex As Exception
            Call MsgBox("Error in HandleZPlayFielding. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' handle errors if they are part of the Z play result
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub HandleZPlayErrors()
        Dim tempIntegerValue As Integer

        Try
            tempIntegerValue = CInt(Val(Game.GetResultPtr.description.Substring(Game.GetResultPtr.description.Length - 1)))
            If tempIntegerValue = 1 Then
                'Pitcher
                Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).PitStatPtr.e += 1
            Else
                tempIntegerValue = Game.GetFielderNumByPos(tempIntegerValue)
                Game.PTeam.GetBatterPtr(tempIntegerValue).BatStatPtr.E = Game.PTeam.GetBatterPtr(tempIntegerValue).BatStatPtr.E + 1
            End If
            Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).eOuts += 1
            gbolSOE = True
        Catch ex As Exception
            Call MsgBox("Error in HandleZPlayErrors. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' call this method if a player is ejected as a result of a Z play
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub HandlePlayerEjections()
        Dim tempValue As String
        Dim stringPosition As Integer
        Dim fieldPosition As String = ""
        Dim enoughPlayers As Boolean

        Try
            stringPosition = Game.GetResultPtr.description.IndexOf(conEjected)
            tempValue = Game.GetResultPtr.description.Substring(stringPosition + conEjected.Length + 1)
            If tempValue.IndexOf(" ") > -1 Then
                'Eliminate anything after the first space
                tempValue = GetStrBefore(tempValue, " ")
            End If
            Do
                stringPosition = tempValue.IndexOf("/")
                If stringPosition > -1 Then
                    fieldPosition = GetStrBefore(tempValue, "/")
                Else
                    fieldPosition = tempValue
                End If
                If fieldPosition <> "Runner" And fieldPosition <> "Batter" And fieldPosition <> "Pitcher" Then
                    enoughPlayers = GetIncidents(tempValue, "/") + 1 <= Game.PTeam.GetNumAvailablePlayers
                End If
                If enoughPlayers Then
                    Select Case fieldPosition
                        Case "Center"
                            Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(8)).available = False
                            Game.GetResultPtr.description += ". " & _
                                    Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(8)).player & _
                                    " has been ejected"
                        Case "Catcher"
                            Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(2)).available = False
                            Game.GetResultPtr.description += ". " & _
                                    Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(2)).player & _
                                    " has been ejected"
                        Case "Short"
                            Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(6)).available = False
                            Game.GetResultPtr.description += ". " & _
                                    Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(6)).player & _
                                    " has been ejected"
                        Case "Third"
                            Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(5)).available = False
                            Game.GetResultPtr.description += ". " & _
                                    Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(5)).player & _
                                    " has been ejected"
                        Case "Second"
                            Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(4)).available = False
                            Game.GetResultPtr.description += ". " & _
                                    Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(4)).player & _
                                    " has been ejected"
                        Case "First"
                            Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(3)).available = False
                            Game.GetResultPtr.description += ". " & _
                                    Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(3)).player & _
                                    " has been ejected"
                        Case "Left"
                            Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(7)).available = False
                            Game.GetResultPtr.description += ". " & _
                                    Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(7)).player & _
                                    " has been ejected"
                        Case "Right"
                            Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(9)).available = False
                            Game.GetResultPtr.description += ". " & _
                                    Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(9)).player & _
                                    " has been ejected"
                        Case "Batter"
                            If Game.BTeam.GetBatterPtr(Game.currentBatter).position = "P" Then
                                If Game.BTeam.GetNumAvailablePitchers > 0 Then
                                    Game.BTeam.GetPitcherPtr(Game.BTeam.pitcherSel).available = False
                                    Game.BTeam.pr = 0
                                    Game.BTeam.pitchChange = True
                                    Game.GetResultPtr.description += ". " & _
                                            Game.BTeam.GetPitcherPtr(Game.BTeam.pitcherSel).player & _
                                            " has been ejected"
                                End If
                            Else
                                If GetIncidents(tempValue, "/") + 1 <= Game.BTeam.GetNumAvailablePlayers Then
                                    Game.BTeam.GetBatterPtr(Game.currentBatter).available = False
                                    Game.GetResultPtr.description += ". " & _
                                            Game.BTeam.GetBatterPtr(Game.currentBatter).player & _
                                            " has been ejected"
                                End If
                            End If
                        Case "Runner"
                            If Game.BTeam.GetBatterPtr(FirstBase.runner).position = "P" Then
                                If Game.BTeam.GetNumAvailablePitchers > 0 Then
                                    Game.BTeam.GetPitcherPtr(Game.BTeam.pitcherSel).available = False
                                    Game.BTeam.pr = 0
                                    Game.BTeam.pitchChange = True
                                    Game.GetResultPtr.description += ". " & _
                                            Game.BTeam.GetPitcherPtr(Game.BTeam.pitcherSel).player & _
                                            " has been ejected"
                                End If
                            Else
                                If GetIncidents(tempValue, "/") + 1 <= Game.BTeam.GetNumAvailablePlayers Then
                                    Game.BTeam.GetBatterPtr(FirstBase.runner).available = False
                                    Game.GetResultPtr.description += ". " & _
                                            Game.BTeam.GetBatterPtr(FirstBase.runner).player & _
                                            " has been ejected"
                                End If
                            End If
                        Case "Pitcher"
                            If Game.PTeam.GetNumAvailablePitchers > 0 Then
                                Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).available = False
                                Game.PTeam.pr = 0
                                Game.PTeam.pitchChange = True
                                Game.GetResultPtr.description += ". " & _
                                        Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).player & _
                                        " has been ejected"
                            End If
                        Case Else
                            Call MsgBox("Cannot find ejected Postion: " & fieldPosition, MsgBoxStyle.OkOnly)
                    End Select
                End If
                If stringPosition > -1 Then
                    tempValue = tempValue.Substring(stringPosition + 1)
                End If
            Loop While stringPosition > -1
        Catch ex As Exception
            Call MsgBox("Error in HandlePlayerEjections. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' Handles player injuries as a result of a Z play.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub HandlePlayerInjuries()
        Dim tempValue As String
        Dim indexPosition As Integer
        Dim fieldPosition As String = ""
        Dim tempBoolValue As Boolean

        Try
            'Add injury code here. Initially add one to the number of Injury games.
            indexPosition = Game.GetResultPtr.description.IndexOf(conInjured)
            tempValue = Game.GetResultPtr.description.Substring(indexPosition + conInjured.Length + 1)
            If tempValue.IndexOf(" ") > -1 Then
                'Eliminate anything after the first space
                tempValue = GetStrBefore(tempValue, " ")
            End If
            Do
                indexPosition = tempValue.IndexOf("/")
                If indexPosition > -1 Then
                    fieldPosition = GetStrBefore(tempValue, "/")
                Else
                    fieldPosition = tempValue
                End If
                If fieldPosition <> "Batter" Then
                    tempBoolValue = GetIncidents(tempValue, "/") + 1 > Game.PTeam.GetNumAvailablePlayers
                End If
                If Not tempBoolValue Then 'If enough position players are available to replace
                    Select Case fieldPosition
                        Case "Center"
                            With Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(8))
                                .available = False
                                .BatStatPtr.GamesInj = DetermineInjGames(CInt(.inj))
                                Game.GetResultPtr.description += ". " & .player & " is out for " & _
                                        .BatStatPtr.GamesInj & " games"
                                '.BatStatPtr.GamesInj += 1 'This is to account for the inj game decrement at end of game
                            End With
                        Case "Catcher"
                            With Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(2))
                                .available = False
                                .BatStatPtr.GamesInj = DetermineInjGames(CInt(.inj))
                                Game.GetResultPtr.description += ". " & .player & " is out for " & _
                                        .BatStatPtr.GamesInj & " games"
                                '.BatStatPtr.GamesInj += 1 'This is to account for the inj game decrement at end of game
                            End With
                        Case "Short"
                            With Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(6))
                                .available = False
                                .BatStatPtr.GamesInj = DetermineInjGames(CInt(.inj))
                                Game.GetResultPtr.description += ". " & .player & " is out for " & _
                                        .BatStatPtr.GamesInj & " games"
                                '.BatStatPtr.GamesInj += 1 'This is to account for the inj game decrement at end of game
                            End With
                        Case "Third"
                            With Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(5))
                                .available = False
                                .BatStatPtr.GamesInj = DetermineInjGames(CInt(.inj))
                                Game.GetResultPtr.description += ". " & .player & " is out for " & _
                                        .BatStatPtr.GamesInj & " games"
                                '.BatStatPtr.GamesInj += 1 'This is to account for the inj game decrement at end of game
                            End With
                        Case "Second"
                            With Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(4))
                                .available = False
                                .BatStatPtr.GamesInj = DetermineInjGames(CInt(.inj))
                                Game.GetResultPtr.description += ". " & .player & " is out for " & _
                                        .BatStatPtr.GamesInj & " games"
                                '.BatStatPtr.GamesInj += 1 'This is to account for the inj game decrement at end of game
                            End With
                        Case "First"
                            With Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(3))
                                .available = False
                                .BatStatPtr.GamesInj = DetermineInjGames(CInt(.inj))
                                Game.GetResultPtr.description += ". " & .player & " is out for " & _
                                        .BatStatPtr.GamesInj & " games"
                                '.BatStatPtr.GamesInj += 1 'This is to account for the inj game decrement at end of game
                            End With
                        Case "Left"
                            With Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(7))
                                .available = False
                                .BatStatPtr.GamesInj = DetermineInjGames(CInt(.inj))
                                Game.GetResultPtr.description += ". " & .player & " is out for " & _
                                        .BatStatPtr.GamesInj & " games"
                                '.BatStatPtr.GamesInj += 1 'This is to account for the inj game decrement at end of game
                            End With
                        Case "Right"
                            With Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(9))
                                .available = False
                                .BatStatPtr.GamesInj = DetermineInjGames(CInt(.inj))
                                Game.GetResultPtr.description += ". " & .player & " is out for " & _
                                        .BatStatPtr.GamesInj & " games"
                                '.BatStatPtr.GamesInj += 1 'This is to account for the inj game decrement at end of game
                            End With
                        Case "Batter"
                            If Game.BTeam.GetBatterPtr(Game.currentBatter).position <> "P" Then
                                'Pitcher cannot be injured
                                If GetIncidents(tempValue, "/") + 1 <= Game.BTeam.GetNumAvailablePlayers Then
                                    With Game.BTeam.GetBatterPtr(Game.currentBatter)
                                        .available = False
                                        .BatStatPtr.GamesInj = DetermineInjGames(CInt(.inj))
                                        Game.GetResultPtr.description += ". " & .player & " is out for " & _
                                                .BatStatPtr.GamesInj & " games"
                                        '.BatStatPtr.GamesInj += 1 'This is to account for the inj game decrement at end of game
                                    End With
                                End If
                            End If
                        Case Else
                            Call MsgBox("Cannot find injured Postion: " & fieldPosition, MsgBoxStyle.OkOnly)
                    End Select
                End If
                If indexPosition > -1 Then
                    tempValue = tempValue.Substring(indexPosition + 1)
                End If
            Loop While indexPosition > -1
        Catch ex As Exception
            Call MsgBox("Error in HandlePlayerInjuries. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' handles hit and run situation
    ''' </summary>
    ''' <param name="playResult"></param>
    ''' <param name="FACCard"></param>
    ''' <remarks></remarks>
    Private Sub HandleHitAndRun(ByRef playResult As String, ByRef FACCard As FACCard)
        Dim highHitValue As Integer
        Dim lower2bValue As Integer
        Dim tempValue As String = ""

        Try
            playResult = "HAR"
            With Game.BTeam.GetBatterPtr(Game.currentBatter)
                If .hitRun = "0" Then
                    highHitValue = HighHitNum(2)
                    If CInt(FACCard.Random) < 26 Then
                        If highHitValue < 25 Then
                            'reduce hitting benefits of batter
                            lower2bValue = highHitValue - 1
                            If lower2bValue Mod 10 = 0 Then
                                lower2bValue = highHitValue - 3
                            End If
                            Select Case CInt(FACCard.Random)
                                Case 11 To lower2bValue - 1
                                    playResult += "1"
                                Case lower2bValue To highHitValue
                                    playResult += "2"
                                Case highHitValue + 1 To 25
                                    playResult += "5"
                            End Select
                        Else
                            Select Case CInt(FACCard.Random)
                                Case 11 To 21
                                    playResult += "1"
                                Case 22 To 24
                                    playResult += "2"
                                Case 25
                                    playResult += "3"
                            End Select
                        End If
                    Else
                        Select Case CInt(FACCard.Random)
                            Case 26 To 34
                                playResult += "4"
                            Case 35 To 42
                                playResult += "5"
                            Case 43 To 51
                                playResult += "6"
                            Case 52 To 61
                                playResult += "7"
                            Case 62
                                playResult += "8"
                            Case 63 To 67
                                playResult += "12"
                            Case 68 To 74
                                playResult += "9"
                            Case 75 To 81
                                playResult += "10"
                            Case 82 To 88
                                playResult += "11"
                            Case Else
                                playResult = conNoPlay
                                Call MsgBox("Cannot lookup Hit and Run play.", MsgBoxStyle.Exclamation)
                        End Select
                    End If
                ElseIf .hitRun = "1" Then
                    highHitValue = HighHitNum(2)
                    If CInt(FACCard.Random) < 36 Then
                        If highHitValue < 35 Then
                            'reduce hitting benefits of batter
                            lower2bValue = highHitValue - 1
                            If lower2bValue Mod 10 = 0 Then
                                lower2bValue = highHitValue - 3
                            End If
                            Select Case CInt(FACCard.Random)
                                Case 11 To lower2bValue - 1
                                    playResult += "1"
                                Case lower2bValue To highHitValue
                                    playResult += "2"
                                Case highHitValue + 1 To 35
                                    playResult += "5"
                            End Select
                        Else
                            Select Case CInt(FACCard.Random)
                                Case 11 To 28
                                    playResult += "1"
                                Case 31 To 34
                                    playResult += "2"
                                Case 35
                                    playResult += "3"
                            End Select
                        End If
                    Else
                        Select Case CInt(FACCard.Random)
                            Case 36
                                playResult += "4"
                            Case 37 To 42
                                playResult += "5"
                            Case 43 To 48
                                playResult += "6"
                            Case 51 To 56
                                playResult += "7"
                            Case 57 To 61
                                playResult += "8"
                            Case 62
                                playResult += "12"
                            Case 63 To 64
                                playResult += "9"
                            Case 65 To 74
                                playResult += "10"
                            Case 75 To 88
                                playResult += "11"
                            Case Else
                                playResult = conNoPlay
                                Call MsgBox("Cannot lookup Hit and Run play.", MsgBoxStyle.Exclamation)
                        End Select
                    End If
                Else
                    highHitValue = HighHitNum(2)
                    If CInt(FACCard.Random) < 43 Then
                        If highHitValue < 42 Then
                            'reduce hitting benefits of batter
                            lower2bValue = highHitValue - 1
                            If lower2bValue Mod 10 = 0 Then
                                lower2bValue = highHitValue - 3
                            End If
                            Select Case CInt(FACCard.Random)
                                Case 11 To lower2bValue - 1
                                    playResult += "1"
                                Case lower2bValue To highHitValue
                                    playResult += "2"
                                Case highHitValue + 1 To 42
                                    playResult += "5"
                            End Select
                        Else
                            Select Case CInt(FACCard.Random)
                                Case 11 To 34
                                    playResult += "1"
                                Case 35 To 41
                                    playResult += "2"
                                Case 42
                                    playResult += "3"
                            End Select
                        End If
                    Else
                        Select Case CInt(FACCard.Random)
                            Case 43 To 46
                                playResult += "4"
                            Case 47 To 52
                                playResult += "5"
                            Case 53 To 61
                                playResult += "6"
                            Case 62 To 64
                                playResult += "7"
                            Case 65 To 67
                                playResult += "8"
                            Case 68 To 71
                                playResult += "12"
                            Case 72 To 75
                                playResult += "9"
                            Case 76 To 83
                                playResult += "10"
                            Case 84 To 88
                                playResult += "11"
                            Case Else
                                playResult = conNoPlay
                                Call MsgBox("Cannot lookup Hit and Run play.", MsgBoxStyle.Exclamation)
                        End Select

                    End If
                    End If
                    'Correction for pitchers 2-7 or greater
                    tempValue = Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).pbRange
                    If Game.PTeam.pr > 0 Then
                        If Val(tempValue.Substring(tempValue.Length - 1)) >= 7 Or _
                                    Val(tempValue.Substring(tempValue.Length - 1)) = 0 Then
                            If playResult.Substring(playResult.Length - 1) = "2" Or playResult.Substring(playResult.Length - 1) = "3" Then
                                playResult = "HAR5" 'G4, runners advance
                            End If
                        End If
                    End If
                    playResult += BaseSituation()
                    Call Game.GetResultPtr.ChartLookup(playResult)
                Select Case playResult.Substring(0, 4)
                    Case "HAR4"
                        tempValue = Game.BTeam.GetBatterPtr(FirstBase.runner).sp
                        If "ABYZ".IndexOf(tempValue) > -1 Then
                            Game.GetResultPtr.resFirst = 2
                            Game.GetResultPtr.description = "Runner steals"
                            Game.BTeam.GetBatterPtr(FirstBase.runner).BatStatPtr.SB += 1
                            Game.BTeam.GetBatterPtr(FirstBase.runner).BatStatPtr.SBA += 1
                            Game.GetResultPtr.po = ""
                            Game.GetResultPtr.a = ""
                        Else
                            Game.GetResultPtr.description = "Runner out stealing"
                            Game.BTeam.GetBatterPtr(FirstBase.runner).BatStatPtr.SBA += 1
                        End If
                    Case "HAR8"
                        tempValue = Game.BTeam.GetBatterPtr(FirstBase.runner).sp
                        If (tempValue = "D" Or tempValue = "E") Then
                            Game.GetResultPtr.resFirst = 1
                            playResult = conNoPlay
                            Game.GetResultPtr.description = ""
                        Else
                            Game.GetResultPtr.resFirst = 2
                            Game.GetResultPtr.description = "Runner steals"
                            Game.BTeam.GetBatterPtr(FirstBase.runner).BatStatPtr.SB += 1
                            Game.BTeam.GetBatterPtr(FirstBase.runner).BatStatPtr.SBA += 1
                        End If
                End Select
            End With
        Catch ex As Exception
            Call MsgBox("Error in HandleHitAndRun. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' Handles base stealing situations
    ''' </summary>
    ''' <param name="playResult"></param>
    ''' <param name="FACCard"></param>
    ''' <remarks></remarks>
    Private Sub HandleBaseStealAttempt(ByRef playResult As String, ByRef FACCard As FACCard)
        Dim isDblStl1 As Boolean
        Dim isDblStl2 As Boolean
        Dim isDblStl3 As Boolean
        Dim isSB As Boolean
        Dim isCS As Boolean
        Dim speedThrowRating As String = ""
        Dim stringPosition As Integer
        Dim tempIntegerValue As Integer

        Try
            playResult = "SBR"
            Select Case BaseSituation()
                Case "000"
                    Call MsgBox("No one on base!", MsgBoxStyle.Exclamation)
                    playResult = conNoPlay
                Case "001"
                    playResult += "1"
                Case "010"
                    playResult += "2"
                Case "011"
                    If MsgBox("Are you attempting a double steal?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                        If MsgBox("Attempt to get runner stealing third?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                            isDblStl1 = True
                            playResult += "2"
                        Else
                            isDblStl2 = True
                            playResult += "1"
                        End If
                    Else
                        playResult += "2"
                    End If
                Case "100"
                    playResult += "3"
                Case "101"
                    If MsgBox("Are you attempting to steal second only?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                        playResult += "1"
                    Else
                        If MsgBox("Are you attempting a double steal", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                            If MsgBox("Attempt to get runner stealing home?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                                isDblStl1 = True
                                playResult += "3"
                            Else
                                isDblStl3 = True
                                playResult += "1"
                            End If
                        Else
                            'Stealing home
                            playResult += "3"
                        End If
                    End If
                Case "110"
                    If MsgBox("Are you attempting a double steal", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                        If MsgBox("Attempt to get runner stealing home?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                            isDblStl2 = True
                            playResult += "3"
                        Else
                            isDblStl3 = True
                            playResult += "2"
                        End If
                    Else
                        'Stealing home
                        playResult += "3"
                    End If
                Case "111"
                    If MsgBox("Are you attempting a triple steal", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                        If MsgBox("Attempt to get runner stealing home?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                            isDblStl1 = True
                            isDblStl2 = True
                            playResult += "3"
                        ElseIf MsgBox("Attempt to get runner stealing third?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                            isDblStl1 = True
                            isDblStl3 = True
                            playResult += "2"
                        Else
                            isDblStl2 = True
                            isDblStl3 = True
                            playResult += "1"
                        End If
                    Else
                        'Stealing home
                        playResult += "3"
                    End If
            End Select
            If playResult = "SBR1" Then
                'Stealing 2nd base
                Select Case CInt(FACCard.Random)
                    Case 11 To 12
                        playResult += "11-15" & BaseSituation()
                    Case 13 To 15
                        playResult += "16-23" & BaseSituation()
                    Case 16 To 18
                        playResult += "24-32" & BaseSituation()
                    Case 21 To 23
                        playResult += "33-47" & BaseSituation()
                    Case 24 To 35
                        If "YZ".IndexOf(Game.BTeam.GetBatterPtr(FirstBase.runner).sp) > -1 Then
                            playResult += "33-47" & BaseSituation()
                        Else
                            playResult = conNoPlay
                        End If
                    Case 36 To 38
                        If Game.BTeam.GetBatterPtr(FirstBase.runner).sp = "Z" Then
                            playResult += "33-47" & BaseSituation()
                        Else
                            playResult = conNoPlay
                        End If
                    Case 41 To 47
                        playResult = conNoPlay
                    Case 48 To 51
                        playResult += "48-52" & BaseSituation()
                    Case 52
                        playResult = conNoPlay
                    Case 53 To 54
                        playResult += "53-55" & BaseSituation()
                    Case 55
                        playResult = conNoPlay
                    Case 56 To 57
                        playResult += "56-61" & BaseSituation()
                    Case 58 To 61
                        playResult = conNoPlay
                    Case 62 To 65
                        isSB = True
                        playResult += FACCard.Random & BaseSituation()
                    Case 66
                        playResult = conNoPlay
                    Case 67
                        playResult += "66-68" & BaseSituation()
                        'Picked off
                        isDblStl2 = False
                        isDblStl3 = False
                    Case 68
                        If "YZ".IndexOf(Game.BTeam.GetBatterPtr(FirstBase.runner).sp) > -1 Then
                            playResult += "66-68" & BaseSituation()
                            'Picked off
                            isDblStl2 = False
                            isDblStl3 = False
                        Else
                            playResult = conNoPlay
                        End If
                    Case 71 To 72
                        playResult += FACCard.Random & BaseSituation()
                    Case 73 To 88
                        playResult = conNoPlay
                    Case Else
                        playResult = conNoPlay
                        Call MsgBox("Cannot lookup Base Stealing play.", MsgBoxStyle.Exclamation)
                End Select
            ElseIf playResult = "SBR2" Then
                'Stealing 3rd base
                Select Case CInt(FACCard.Random)
                    Case 11 To 12
                        playResult += "11-14" & BaseSituation()
                    Case 13 To 14
                        playResult += "15-17" & BaseSituation()
                    Case 15 To 17
                        playResult += "18-28" & BaseSituation()
                    Case 18 To 22
                        playResult += "31-41" & BaseSituation()
                    Case 23 To 32
                        If "YZ".IndexOf(Game.BTeam.GetBatterPtr(SecondBase.runner).sp) > -1 Then
                            playResult += "31-41" & BaseSituation()
                        Else
                            playResult = conNoPlay
                        End If
                    Case 33 To 35
                        If Game.BTeam.GetBatterPtr(SecondBase.runner).sp = "Z" Then
                            playResult += "31-41" & BaseSituation()
                        Else
                            playResult = conNoPlay
                        End If
                    Case 36 To 41
                        playResult = conNoPlay
                    Case 42 To 43
                        playResult += "42-52" & BaseSituation()
                    Case 44
                        If "YZ".IndexOf(Game.BTeam.GetBatterPtr(SecondBase.runner).sp) > -1 Then
                            playResult += "42-52" & BaseSituation()
                        Else
                            playResult = conNoPlay
                        End If
                    Case 45 To 52
                        playResult = conNoPlay
                    Case 53 To 54
                        playResult += "53-55" & BaseSituation()
                    Case 55
                        If "YZ".IndexOf(Game.BTeam.GetBatterPtr(SecondBase.runner).sp) > -1 Then
                            playResult += "53-55" & BaseSituation()
                        Else
                            playResult = conNoPlay
                        End If
                    Case 56 To 57
                        playResult += "56-61" & BaseSituation()
                    Case 58
                        If "YZ".IndexOf(Game.BTeam.GetBatterPtr(SecondBase.runner).sp) > -1 Then
                            playResult += "56-61" & BaseSituation()
                        Else
                            playResult = conNoPlay
                        End If
                    Case 61, 63
                        playResult = conNoPlay
                    Case 62, 64 To 65
                        playResult += FACCard.Random & BaseSituation()
                    Case 66
                        playResult += "66-68" & BaseSituation()
                        'Picked off
                        isDblStl1 = False
                        isDblStl3 = False
                    Case 67
                        If "YZ".IndexOf(Game.BTeam.GetBatterPtr(SecondBase.runner).sp) > -1 Then
                            playResult += "66-68" & BaseSituation()
                            'Picked off
                            isDblStl1 = False
                            isDblStl3 = False
                        Else
                            playResult = conNoPlay
                        End If
                    Case 68
                        playResult = conNoPlay
                    Case 71 To 72
                        playResult += "71-72" & BaseSituation()
                    Case 73 To 88
                        playResult = conNoPlay
                    Case Else
                        playResult = conNoPlay
                        Call MsgBox("Cannot lookup Base Stealing play.", MsgBoxStyle.Exclamation)
                End Select
            ElseIf playResult = "SBR3" Then
                'Stealing home
                Select Case CInt(FACCard.Random)
                    Case 11 To 13
                        playResult += "11-13" & BaseSituation()
                    Case 14 To 15
                        playResult += "14-15" & BaseSituation()
                    Case 16 To 18
                        playResult += "16-18" & BaseSituation()
                    Case 21 To 25
                        playResult += "21-26" & BaseSituation()
                    Case 26
                        If "YZ".IndexOf(Game.BTeam.GetBatterPtr(ThirdBase.runner).sp) > -1 Then
                            playResult += "21-26" & BaseSituation()
                        Else
                            playResult = conNoPlay
                        End If
                    Case 27 To 37
                        'playResult += "27-37" & "A"
                        playResult = conNoPlay
                    Case 38 To 43
                        playResult += "38-43" & BaseSituation()
                    Case 44 To 52
                        'playResult += "44-52" & "A"
                        playResult = conNoPlay
                    Case 53
                        playResult += FACCard.Random & BaseSituation()
                        'Picked off
                        isDblStl1 = False
                        isDblStl2 = False
                    Case 62
                        playResult += FACCard.Random & BaseSituation()
                    Case 54 To 55
                        playResult += "54-55" & BaseSituation()
                        'Picked off
                        isDblStl1 = False
                        isDblStl2 = False
                    Case 56 To 57
                        playResult += "56-57" & BaseSituation()
                    Case 58 To 61
                        playResult += "58-61" & BaseSituation()
                    Case 63 To 66
                        If "YZ".IndexOf(Game.BTeam.GetBatterPtr(ThirdBase.runner).sp) > -1 Then
                            playResult += "21-26" & BaseSituation()
                        Else
                            playResult = conNoPlay
                        End If
                    Case 67 To 71
                        If Game.BTeam.GetBatterPtr(ThirdBase.runner).sp = "Z" Then
                            playResult += "21-26" & BaseSituation()
                        Else
                            playResult = conNoPlay
                        End If
                    Case 72 To 88
                        'playresult += "63-88" & "A"
                        playResult = conNoPlay
                    Case Else
                        playResult = conNoPlay
                        Call MsgBox("Cannot lookup Base Stealing play.", MsgBoxStyle.Exclamation)
                End Select
            End If
            If playResult <> conNoPlay Then
                Call Game.GetResultPtr.ChartLookup(playResult)
                'Determine conditionals
                If Game.GetResultPtr.description.IndexOf("Cnd:SP") > -1 Then
                    speedThrowRating = Game.GetResultPtr.description.Trim
                    speedThrowRating = speedThrowRating.Substring(speedThrowRating.Length - 1)
                    Select Case playResult.Substring(3, 1)
                        Case "1"
                            isSB = Asc(Game.BTeam.GetBatterPtr(FirstBase.runner).sp) <= Asc(speedThrowRating) Or _
                                    "YZ".IndexOf(Game.BTeam.GetBatterPtr(FirstBase.runner).sp) > -1
                        Case "2"
                            isSB = Asc(Game.BTeam.GetBatterPtr(SecondBase.runner).sp) <= Asc(speedThrowRating) Or _
                                    "YZ".IndexOf(Game.BTeam.GetBatterPtr(SecondBase.runner).sp) > -1
                        Case "3"
                            isSB = Asc(Game.BTeam.GetBatterPtr(ThirdBase.runner).sp) <= Asc(speedThrowRating) Or _
                                    "YZ".IndexOf(Game.BTeam.GetBatterPtr(ThirdBase.runner).sp) > -1
                        Case Else
                            Call MsgBox("Error determining the base stealing situation. String = " & playResult, MsgBoxStyle.Critical)
                    End Select
                    If Not isSB Then playResult = conNoPlay
                End If
                If Game.GetResultPtr.description.IndexOf("Cnd:T") > -1 And playResult <> conNoPlay Then
                    speedThrowRating = Game.GetResultPtr.description.Trim
                    speedThrowRating = speedThrowRating.Substring(speedThrowRating.Length - 1)
                    isCS = Asc(Game.PTeam.GetBatterPtr(Game.GetFielderNumByPos(2)).throwRating) <= Asc(speedThrowRating)
                    isSB = Not isCS
                    If isCS Then
                        Select Case playResult.Substring(3, 1)
                            Case "1"
                                Game.GetResultPtr.resFirst = 0
                            Case "2"
                                Game.GetResultPtr.resSecond = 0
                            Case "3"
                                Game.GetResultPtr.resThird = 0
                            Case Else
                                Call MsgBox("Error determining the base stealing situation. String = " & playResult, MsgBoxStyle.Critical)
                        End Select
                    End If
                End If
                'Handle double steals
                If (Game.outs < 2 Or Game.GetResultPtr.description.IndexOf("Result:CS") = -1) And playResult <> conNoPlay Then
                    If isDblStl1 Then
                        Game.GetResultPtr.resFirst = 2
                        Game.BTeam.GetBatterPtr(FirstBase.runner).BatStatPtr.SB += 1
                        Game.BTeam.GetBatterPtr(FirstBase.runner).BatStatPtr.SBA += 1
                    End If
                    If isDblStl2 Then
                        If Game.GetResultPtr.resSecond <> 4 Then
                            Game.GetResultPtr.resSecond = 3
                        End If
                        Game.BTeam.GetBatterPtr(SecondBase.runner).BatStatPtr.SB += 1
                        Game.BTeam.GetBatterPtr(SecondBase.runner).BatStatPtr.SBA += 1
                    End If
                    If isDblStl3 Then
                        Game.GetResultPtr.resThird = 4
                        Game.BTeam.GetBatterPtr(ThirdBase.runner).BatStatPtr.SB += 1
                        Game.BTeam.GetBatterPtr(ThirdBase.runner).BatStatPtr.SBA += 1
                    End If
                End If
                'Handle Caught Stealing
                If (Game.GetResultPtr.description.IndexOf("Result:CS") > -1 Or isCS) And playResult <> conNoPlay Then
                    Select Case playResult.Substring(3, 1)
                        Case "1"
                            Game.BTeam.GetBatterPtr(FirstBase.runner).BatStatPtr.SBA += 1
                            Game.GetResultPtr.a = "2"
                            Game.GetResultPtr.po = "6"
                        Case "2"
                            Game.BTeam.GetBatterPtr(SecondBase.runner).BatStatPtr.SBA += 1
                            Game.GetResultPtr.a = "2"
                            Game.GetResultPtr.po = "5"
                        Case "3"
                            Game.BTeam.GetBatterPtr(ThirdBase.runner).BatStatPtr.SBA += 1
                            Game.GetResultPtr.a = "1"
                            Game.GetResultPtr.po = "2"
                        Case Else
                            Call MsgBox("Error determining the runner caught stealing. String = " & playResult, MsgBoxStyle.Critical)
                    End Select
                End If
                'Handle Stolen Base
                If (Game.GetResultPtr.description.IndexOf("Result:SB") > -1 Or isSB) And playResult <> conNoPlay Then
                    Select Case playResult.Substring(3, 1)
                        Case "1"
                            Game.BTeam.GetBatterPtr(FirstBase.runner).BatStatPtr.SB += 1
                            Game.BTeam.GetBatterPtr(FirstBase.runner).BatStatPtr.SBA += 1
                        Case "2"
                            Game.BTeam.GetBatterPtr(SecondBase.runner).BatStatPtr.SB += 1
                            Game.BTeam.GetBatterPtr(SecondBase.runner).BatStatPtr.SBA += 1
                        Case "3"
                            Game.BTeam.GetBatterPtr(ThirdBase.runner).BatStatPtr.SB += 1
                            Game.BTeam.GetBatterPtr(ThirdBase.runner).BatStatPtr.SBA += 1
                        Case Else
                            Call MsgBox("Error determining the runner caught stealing. String = " & playResult, MsgBoxStyle.Critical)
                    End Select
                End If
                'Handle fielding errors
                stringPosition = Game.GetResultPtr.description.IndexOf("Result:E")
                If stringPosition > -1 And playResult <> conNoPlay Then
                    tempIntegerValue = CInt(Val(Game.GetResultPtr.description.Substring(Game.GetResultPtr.description.Length - 1)))
                    If tempIntegerValue = 1 Then
                        'Pitcher
                        Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).PitStatPtr.e += 1
                    Else
                        tempIntegerValue = Game.GetFielderNumByPos(tempIntegerValue)
                        Game.PTeam.GetBatterPtr(tempIntegerValue).BatStatPtr.E += 1
                    End If
                    Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).eOuts += 1
                    Select Case playResult.Substring(3, 1)
                        Case "1"
                            FirstBase.unearned = True
                        Case "2"
                            SecondBase.unearned = True
                        Case "3"
                            ThirdBase.unearned = True
                    End Select
                End If
                'Cleanup description
                stringPosition = Game.GetResultPtr.description.IndexOf("Result:")
                If stringPosition > 0 Then
                    Game.GetResultPtr.description = Game.GetResultPtr.description.Substring(0, stringPosition).Trim
                End If
                stringPosition = Game.GetResultPtr.description.IndexOf("Cnd:")
                If stringPosition > -1 Then
                    If isSB Then
                        Game.GetResultPtr.description = "Runner steals"
                    ElseIf isCS Then
                        Game.GetResultPtr.description = "Runner caught stealing"
                    Else
                        Game.GetResultPtr.description = ""
                    End If
                End If
            End If
        Catch ex As Exception
            Call MsgBox("Error in HandleBaseStealAttempt. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' handles Sacrifice situations
    ''' </summary>
    ''' <param name="playResult"></param>
    ''' <param name="FACCard"></param>
    ''' <remarks></remarks>
    Private Sub HandleSacrifice(ByRef playResult As String, ByRef FACCard As FACCard)
        Dim stringPosition As Integer
        Dim tempIntegerValue As Integer
        Try
            playResult = "SAC"
            Call HandleSAC(playResult, FACCard.Random)
            If playResult <> conNoPlay Then
                Call Game.GetResultPtr.ChartLookup(playResult)
                stringPosition = Game.GetResultPtr.description.IndexOf("Result:")
                If stringPosition > -1 Then
                    tempIntegerValue = CInt(Val(Game.GetResultPtr.description.Substring(Game.GetResultPtr.description.Length - 1)))
                    If tempIntegerValue = 1 Then
                        'Pitcher
                        Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).PitStatPtr.e += 1
                    Else
                        tempIntegerValue = Game.GetFielderNumByPos(tempIntegerValue)
                        Game.PTeam.GetBatterPtr(tempIntegerValue).BatStatPtr.E += 1
                    End If
                    Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).eOuts += 1
                    gbolSOE = True
                End If
            End If
        Catch ex As Exception
            Call MsgBox("Error in HandleSacrifice. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' handles bunt for base-hit situation
    ''' </summary>
    ''' <param name="playResult"></param>
    ''' <param name="FACCard"></param>
    ''' <remarks></remarks>
    Private Sub HandleBunt(ByRef playResult As String, ByRef FACCard As FACCard)
        Dim stringPosition As Integer
        Dim tempIntegerValue As Integer

        Try
            playResult = "BUNT"
            If bolCornersIn Or bolInfieldIn Then
                'add 10 to random number
                If Val(FACCard.Random) + 10 > 88 Then
                    FACCard.Random = "88"
                Else
                    FACCard.Random = (Val(FACCard.Random) + 10).ToString
                End If
            End If
            Select Case CInt(FACCard.Random)
                Case 11 To 28
                    playResult += "1"
                Case 31 To 35
                    playResult += "2"
                Case 36 To 42
                    playResult += "3"
                Case 43 To 48
                    playResult += "4"
                Case 51 To 57
                    playResult += "5"
                Case 58 To 64
                    playResult += "6"
                Case 65
                    playResult += "7"
                Case 66
                    playResult += "8"
                Case 67
                    playResult += "9"
                Case 68
                    playResult += "10"
                Case 71
                    playResult += "11"
                Case 72 To 88
                    playResult += "12"
                Case Else
                    Call MsgBox("Error in Bunt lookup!", MsgBoxStyle.Exclamation)
                    playResult = conNoPlay
            End Select
            If playResult <> conNoPlay Then
                Call Game.GetResultPtr.ChartLookup(playResult)
                stringPosition = Game.GetResultPtr.description.IndexOf("Result:")
                If stringPosition > -1 Then
                    tempIntegerValue = CInt(Val(Game.GetResultPtr.description.Substring(Game.GetResultPtr.description.Length - 1)))
                    If tempIntegerValue = 1 Then
                        'Pitcher
                        Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).PitStatPtr.e += 1
                    Else
                        tempIntegerValue = Game.GetFielderNumByPos(tempIntegerValue)
                        Game.PTeam.GetBatterPtr(tempIntegerValue).BatStatPtr.E += 1
                    End If
                    Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).eOuts += 1
                    gbolSOE = True
                End If
            End If
        Catch ex As Exception
            Call MsgBox("Error in HandleBunt. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' If result is an extra base hit, and the pitcher is a 2-6, a certain percentage (see UseOtherCard()) of these should
    ''' become a single, because 2-6 pitchers give up too many extra base hits under the original system
    ''' </summary>
    ''' <param name="playResult"></param>
    ''' <param name="FACCard"></param>
    ''' <param name="baseAdvance"></param>
    ''' <param name="suppDescription"></param>
    ''' <remarks></remarks>
    Private Sub CheckExtraBaseHitEffectBatter(ByRef playResult As String, ByRef FACCard As FACCard, ByRef baseAdvance As Boolean, _
                    ByRef suppDescription As String)
        Dim singleHitNum As Integer = HighSingleNum()
        Dim numIncidents As Integer

        If "2B|3B|HR".IndexOf(Left(playResult, 2)) >= 0 AndAlso _
                    Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).pbRange = "2-6" AndAlso _
                    UseOtherCard("2-6") AndAlso _
                    singleHitNum <> 0 Then
            numIncidents = GetRangeIncidents(11, singleHitNum)
            FACCard.Random = SelectIncidentInRange(11, singleHitNum, FetchRandomNumber(numIncidents), "RevPB").ToString
            If FACCard.Random = "0" Then
                Call MsgBox("Random number of 0 is about to get processed from CheckExtraBaseHitEffectBatter()")
            End If
            playResult = ""
            baseAdvance = False
            suppDescription = ""
            GetResultFromHitterCard(playResult, FACCard, baseAdvance, suppDescription)
        End If
    End Sub

    ''' <summary>
    ''' This routine checks for excessive variance in the result compared to the hitting card. If an excessive outcome is found, lets 
    ''' assign an underrepresented outcome
    ''' </summary>
    ''' <param name="playResult"></param>
    ''' <param name="FACCard"></param>
    ''' <param name="baseAdvance"></param>
    ''' <param name="suppDescription"></param>
    ''' <remarks></remarks>
    Public Sub CheckExcessiveVarianceBatter(ByRef playResult As String, ByRef FACCard As FACCard, ByRef baseAdvance As Boolean, _
                    ByRef suppDescription As String)
        Dim currentResult As String
        Dim lowVariance As String = ""
        Dim highVariance As String = ""
        Dim rndNum As Integer
        Dim numIncidents As Integer
        Dim lowRange As Integer
        Dim highRange As Integer

        If "1B|2B|3B|HR|HP".IndexOf(Left(playResult, 2)) >= 0 Then
            currentResult = playResult.Substring(0, 2)
        ElseIf "K|W".IndexOf(Left(playResult, 1)) >= 0 Then
            currentResult = playResult.Substring(0, 1)
        Else
            currentResult = "OUT"
        End If

        GetBatterVariances(lowVariance, highVariance)
        If currentResult = highVariance And lowVariance <> "" Then
            Randomize()
            rndNum = Convert.ToInt32(4 * Rnd() + 1)
            If rndNum = 1 Then
                With Game.BTeam.GetBatterPtr(Game.currentBatter)
                    'Only enabled 25% of the time. Move this up later after bug checking
                    Select Case lowVariance
                        Case "1B"
                            lowRange = 11
                            highRange = HighSingleNum()
                            numIncidents = GetRangeIncidents(lowRange, highRange)
                        Case "2B"
                            numIncidents = DetermineCardNumbers(.hit2B8, lowRange, highRange)
                        Case "3B"
                            numIncidents = DetermineCardNumbers(.hit3B8, lowRange, highRange)
                        Case "HR"
                            numIncidents = DetermineCardNumbers(.hr, lowRange, highRange)
                        Case "HP"
                            numIncidents = DetermineCardNumbers(.hpb, lowRange, highRange)
                        Case "K"
                            numIncidents = DetermineCardNumbers(.k, lowRange, highRange)
                        Case "W"
                            numIncidents = DetermineCardNumbers(.w, lowRange, highRange)
                        Case "OUT"
                            numIncidents = DetermineCardNumbers(.out, lowRange, highRange)
                    End Select
                    If numIncidents = 0 Then Exit Sub
                    FACCard.Random = SelectIncidentInRange(lowRange, highRange, FetchRandomNumber(numIncidents), "HiVar").ToString
                    playResult = ""
                    baseAdvance = False
                    suppDescription = ""
                    GetResultFromHitterCard(playResult, FACCard, baseAdvance, suppDescription)
                End With
            End If
        End If
    End Sub

    ''' <summary>
    ''' This routine checks for excessive variance in the result compared to the pitching card. If an excessive outcome is found, lets 
    ''' assign an underrepresented outcome
    ''' </summary>
    ''' <param name="playResult"></param>
    ''' <param name="FACCard"></param>
    ''' <param name="baseAdvance"></param>
    ''' <param name="suppDescription"></param>
    ''' <remarks></remarks>
    Public Sub CheckExcessiveVariancePitcher(ByRef playResult As String, ByRef FACCard As FACCard, ByRef baseAdvance As Boolean, _
                    ByRef suppDescription As String)
        Dim currentResult As String
        Dim lowVariance As String = ""
        Dim highVariance As String = ""
        Dim rndNum As Integer
        Dim numIncidents As Integer
        Dim lowRange As Integer
        Dim highRange As Integer

        If "1B".IndexOf(Left(playResult, 2)) >= 0 Then
            currentResult = playResult.Substring(0, 2)
        ElseIf "K|W".IndexOf(Left(playResult, 1)) >= 0 Then
            currentResult = playResult.Substring(0, 1)
        Else
            currentResult = "OUT"
        End If

        GetPitcherVariances(lowVariance, highVariance)
        If currentResult = highVariance And lowVariance <> "" Then
            Randomize()
            rndNum = Convert.ToInt32(4 * Rnd() + 1)
            If rndNum = 1 Then
                With Game.BTeam.GetBatterPtr(Game.currentBatter)
                    'Only enabled 25% of the time. Move this up later after bug checking
                    Select Case lowVariance
                        Case "1B"
                            lowRange = 11
                            highRange = HighSingleNum()
                            numIncidents = GetRangeIncidents(lowRange, highRange)
                        Case "K"
                            numIncidents = DetermineCardNumbers(.k, lowRange, highRange)
                        Case "W"
                            numIncidents = DetermineCardNumbers(.w, lowRange, highRange)
                        Case "OUT"
                            numIncidents = DetermineCardNumbers(.out, lowRange, highRange)
                    End Select
                    If numIncidents = 0 Then Exit Sub
                    FACCard.Random = SelectIncidentInRange(lowRange, highRange, FetchRandomNumber(numIncidents), "HiVar").ToString
                    playResult = ""
                    baseAdvance = False
                    suppDescription = ""
                    GetResultFromHitterCard(playResult, FACCard, baseAdvance, suppDescription)
                End With
            End If
        End If
    End Sub

    ''' <summary>
    ''' Determine the high and low variances for the batter, including 1B|2B|3B|HR|HP|K|W|OUT
    ''' </summary>
    ''' <param name="lowVariance"></param>
    ''' <param name="highVariance"></param>
    ''' <remarks></remarks>
    Public Sub GetBatterVariances(ByRef lowVariance As String, ByRef highVariance As String)
        Dim lowVar As Double = 0
        Dim highVar As Double = 0
        Dim variance As Double = 0
        Dim cardNumCurrent As Integer
        Dim cardNumProjected As Double
        Dim cumProjected As Double
        Dim tableKey As String = ""
        Dim sqlQuery As String = ""
        Dim dsStat As DataSet = Nothing
        Dim drStat As DataRow = Nothing
        Dim DataAccess As New clsDataAccess(gstrSeason & "data")
        Dim baseHits As Integer
        Dim singles As Integer
        Dim doubles As Integer
        Dim triples As Integer
        Dim homeruns As Integer
        Dim strikeouts As Integer
        Dim walks As Integer
        Dim hbp As Integer
        'Dim outs As Integer
        Dim atbats As Integer
        Dim intwalks As Integer
        Dim sacflies As Integer
        Dim plateAppearances As Integer

        With Game.BTeam.GetBatterPtr(Game.currentBatter)
            tableKey = StripChar(.player & Game.BTeam.teamName, " ")
            tableKey = HandleQuotes(tableKey)
            sqlQuery = "SELECT * FROM " & gstrHittingTable & " WHERE playerid = '" & tableKey & "'"
            dsStat = DataAccess.ExecuteDataSet(sqlQuery)

            If dsStat.Tables(0).Rows.Count = 0 Then
                atbats = .BatStatPtr.AB
                baseHits = .BatStatPtr.H
                doubles = .BatStatPtr.D
                triples = .BatStatPtr.T
                homeruns = .BatStatPtr.HR
                strikeouts = .BatStatPtr.K
                walks = .BatStatPtr.W
                hbp = .BatStatPtr.HB
                intwalks = .BatStatPtr.ibb
                sacflies = .BatStatPtr.SF
            Else
                drStat = dsStat.Tables(0).Rows(0)
                atbats = .BatStatPtr.AB + CInt(drStat.Item("AB"))
                baseHits = .BatStatPtr.H + CInt(drStat.Item("Hits"))
                doubles = .BatStatPtr.D + CInt(drStat.Item("Doubles"))
                triples = .BatStatPtr.T + CInt(drStat.Item("Triples"))
                homeRuns = .BatStatPtr.HR + CInt(drStat.Item("HomeRuns"))
                strikeouts = .BatStatPtr.K + CInt(drStat.Item("Strikeouts"))
                walks = .BatStatPtr.W + CInt(drStat.Item("Walks"))
                hbp = .BatStatPtr.HB + CInt(drStat.Item("HitByPitch"))
                intwalks = .BatStatPtr.ibb + CInt(drStat.Item("IBB"))
                sacflies = .BatStatPtr.SF + CInt(drStat.Item("SF"))
            End If
            singles = baseHits - doubles - triples - homeruns
            'outs = atbats - baseHits - strikeouts + sacflies
            plateAppearances = atbats + walks - intwalks + hbp + sacflies
            If plateAppearances < 20 Then
                Exit Sub
            End If

            'Determine variances for card outcomes 1B|2B|3B|HR|HP|K|W|OUT
            cardNumCurrent = DetermineCardNumbers(.hit1Bf, 0, 0) + DetermineCardNumbers(.hit1B7, 0, 0) + DetermineCardNumbers(.hit1B8, 0, 0) + DetermineCardNumbers(.hit1B9, 0, 0)
            cardNumProjected = singles * 128 / plateAppearances - Game.BTeam.adjSingle
            cumProjected += cardNumProjected
            variance = cardNumProjected - cardNumCurrent
            If variance > highVar Then
                highVar = variance
                highVariance = "1B"
            ElseIf variance < lowVar Then
                lowVar = variance
                lowVariance = "1B"
            End If

            cardNumCurrent = DetermineCardNumbers(.hit2B7, 0, 0) + DetermineCardNumbers(.hit2B8, 0, 0) + DetermineCardNumbers(.hit2B9, 0, 0)
            cardNumProjected = doubles * 128 / plateAppearances
            cumProjected += cardNumProjected
            variance = cardNumProjected - cardNumCurrent
            If variance > highVar Then
                highVar = variance
                highVariance = "2B"
            ElseIf variance < lowVar Then
                lowVar = variance
                lowVariance = "2B"
            End If

            cardNumCurrent = DetermineCardNumbers(.hit3B8, 0, 0)
            cardNumProjected = triples * 128 / plateAppearances
            cumProjected += cardNumProjected
            variance = cardNumProjected - cardNumCurrent
            If variance > highVar Then
                highVar = variance
                highVariance = "3B"
            ElseIf variance < lowVar Then
                lowVar = variance
                lowVariance = "3B"
            End If

            cardNumCurrent = DetermineCardNumbers(.hr, 0, 0)
            cardNumProjected = homeruns * 128 / plateAppearances
            cumProjected += cardNumProjected
            variance = cardNumProjected - cardNumCurrent
            If variance > highVar Then
                highVar = variance
                highVariance = "HR"
            ElseIf variance < lowVar Then
                lowVar = variance
                lowVariance = "HR"
            End If

            cardNumCurrent = DetermineCardNumbers(.hpb, 0, 0)
            cardNumProjected = hbp * 128 / plateAppearances
            cumProjected += cardNumProjected
            variance = cardNumProjected - cardNumCurrent
            If variance > highVar Then
                highVar = variance
                highVariance = "HP"
            ElseIf variance < lowVar Then
                lowVar = variance
                lowVariance = "HP"
            End If

            cardNumCurrent = DetermineCardNumbers(.k, 0, 0)
            cardNumProjected = strikeouts * 128 / plateAppearances - Game.BTeam.adjK
            cumProjected += cardNumProjected
            variance = cardNumProjected - cardNumCurrent
            If variance > highVar Then
                highVar = variance
                highVariance = "K"
            ElseIf variance < lowVar Then
                lowVar = variance
                lowVariance = "K"
            End If

            cardNumCurrent = DetermineCardNumbers(.w, 0, 0)
            cardNumProjected = walks * 128 / plateAppearances - Game.BTeam.adjBB
            cumProjected += cardNumProjected
            variance = cardNumProjected - cardNumCurrent
            If variance > highVar Then
                highVar = variance
                highVariance = "W"
            ElseIf variance < lowVar Then
                lowVar = variance
                lowVariance = "W"
            End If

            cardNumCurrent = DetermineCardNumbers(.out, 0, 0)
            'projected Out numbers are the remaining card numbers out of 64
            cardNumProjected = 64 - cumProjected
            variance = cardNumProjected - cardNumCurrent
            If variance > highVar Then
                highVar = variance
                highVariance = "OUT"
            ElseIf variance < lowVar Then
                lowVar = variance
                lowVariance = "OUT"
            End If
        End With
    End Sub

    ''' <summary>
    ''' Determine the high and low variances for the pitcher, including 1B|K|W|OUT
    ''' </summary>
    ''' <param name="lowVariance"></param>
    ''' <param name="highVariance"></param>
    ''' <remarks></remarks>
    Public Sub GetPitcherVariances(ByRef lowVariance As String, ByRef highVariance As String)
        Dim lowVar As Double = 0
        Dim highVar As Double = 0
        Dim variance As Double = 0
        Dim cardNumCurrent As Integer
        Dim cardNumProjected As Double
        Dim cumProjected As Double
        Dim tableKey As String = ""
        Dim sqlQuery As String = ""
        Dim dsStat As DataSet = Nothing
        Dim drStat As DataRow = Nothing
        Dim DataAccess As New clsDataAccess(gstrSeason & "data")
        Dim totalOuts As Integer
        Dim baseHits As Integer
        Dim strikeouts As Integer
        Dim walks As Integer
        Dim outcomes As Integer

        With Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel)
            tableKey = StripChar(.player & Game.PTeam.teamName, " ")
            tableKey = HandleQuotes(tableKey)
            sqlQuery = "SELECT * FROM " & gstrPitchingTable & " WHERE playerid = '" & tableKey & "'"
            dsStat = DataAccess.ExecuteDataSet(sqlQuery)

            If dsStat.Tables(0).Rows.Count = 0 Then
                totalOuts = .PitStatPtr.ip
                baseHits = .PitStatPtr.h
                strikeouts = .PitStatPtr.k
                walks = .PitStatPtr.w
            Else
                drStat = dsStat.Tables(0).Rows(0)
                totalOuts = .PitStatPtr.ip + CInt(drStat.Item("IP"))
                baseHits = .PitStatPtr.h + CInt(drStat.Item("Hits"))
                strikeouts = .PitStatPtr.k + CInt(drStat.Item("Strikeouts"))
                walks = .PitStatPtr.w + CInt(drStat.Item("Walks"))
            End If
            outcomes = CInt(totalOuts * (26 / 27)) + walks + baseHits
            If outcomes < 20 Then
                Exit Sub
            End If

            'Determine variances for card outcomes 1B|K|W|OUT
            cardNumCurrent = DetermineCardNumbers(.hit1Bf, 0, 0) + DetermineCardNumbers(.hit1B7, 0, 0) + DetermineCardNumbers(.hit1B8, 0, 0) + DetermineCardNumbers(.hit1B9, 0, 0)
            cardNumProjected = (baseHits * Game.PTeam.singlePct) * 128 / outcomes - Game.PTeam.adjSingle
            Select Case .pb
                'make adjustments by pb rating
                Case "2-9"
                    cardNumProjected = CInt((0.334 * cardNumProjected + 8.81) * (62 / 64))
                Case "2-8"
                    cardNumProjected = CInt((0.556 * cardNumProjected + 6.07) * (62 / 64))
                Case "2-7"
                    cardNumProjected = CInt((0.833 * cardNumProjected + 1.88) * (62 / 64))
                Case "2-6"
                    cardNumProjected = CInt(((cardNumProjected - 3.68) / 0.833) * (62 / 64))
                Case "2-5"
                    cardNumProjected = CInt(((cardNumProjected - 7.89) / 0.556) * (62 / 64))
            End Select
            cumProjected += cardNumProjected
            variance = cardNumProjected - cardNumCurrent
            If variance > highVar Then
                highVar = variance
                highVariance = "1B"
            ElseIf variance < lowVar Then
                lowVar = variance
                lowVariance = "1B"
            End If

            cardNumCurrent = DetermineCardNumbers(.k, 0, 0)
            cardNumProjected = strikeouts * 128 / outcomes - Game.PTeam.adjK
            Select Case .pb
                'make adjustments by pb rating
                Case "2-9"
                    cardNumProjected = CInt((0.334 * cardNumProjected + 6.62) * (62 / 64))
                Case "2-8"
                    cardNumProjected = CInt((0.556 * cardNumProjected + 4.48) * (62 / 64))
                Case "2-7"
                    cardNumProjected = CInt((0.833 * cardNumProjected + 0.99) * (62 / 64))
                Case "2-6"
                    cardNumProjected = CInt(((cardNumProjected - 2.82) / 0.833) * (62 / 64))
                Case "2-5"
                    cardNumProjected = CInt(((cardNumProjected - 4.63) / 0.556) * (62 / 64))
            End Select
            cumProjected += cardNumProjected
            variance = cardNumProjected - cardNumCurrent
            If variance > highVar Then
                highVar = variance
                highVariance = "K"
            ElseIf variance < lowVar Then
                lowVar = variance
                lowVariance = "K"
            End If

            cardNumCurrent = DetermineCardNumbers(.w, 0, 0)
            cardNumProjected = walks * 128 / outcomes - Game.PTeam.adjBB
            Select Case .pb
                'make adjustments by pb rating
                Case "2-9"
                    cardNumProjected = CInt((0.334 * cardNumProjected + 3.03) * (62 / 64))
                Case "2-8"
                    cardNumProjected = CInt((0.556 * cardNumProjected + 2.09) * (62 / 64))
                Case "2-7"
                    cardNumProjected = CInt((0.833 * cardNumProjected + 0.93) * (62 / 64))
                Case "2-6"
                    cardNumProjected = CInt(((cardNumProjected - 0.45) / 0.833) * (62 / 64))
                Case "2-5"
                    cardNumProjected = CInt(((cardNumProjected - 1.18) / 0.556) * (62 / 64))
            End Select
            cumProjected += cardNumProjected
            variance = cardNumProjected - cardNumCurrent
            If variance > highVar Then
                highVar = variance
                highVariance = "W"
            ElseIf variance < lowVar Then
                lowVar = variance
                lowVariance = "W"
            End If

            cardNumCurrent = DetermineCardNumbers(.out, 0, 0)
            'projected Out numbers are the remaining card numbers out of 64
            cardNumProjected = 64 - cumProjected - 2 'assumes 2 numbers for BK, WP, PB
            variance = cardNumProjected - cardNumCurrent
            If variance > highVar Then
                highVar = variance
                highVariance = "OUT"
            ElseIf variance < lowVar Then
                lowVar = variance
                lowVariance = "OUT"
            End If
        End With
    End Sub

    ''' <summary>
    ''' If result is a single on the pitcher's card , and the pitcher is not a 2-6, a certain percentage (see UseOtherCard())
    ''' of these should refer to the batting, because non 2-6 pitchers give up too few extra base hits under the original 
    ''' system
    ''' </summary>
    ''' <param name="playResult"></param>
    ''' <param name="FACCard"></param>
    ''' <param name="baseAdvance"></param>
    ''' <param name="suppDescription"></param>
    ''' <remarks></remarks>
    Private Sub CheckExtraBaseHitEffectPitcher(ByRef playResult As String, ByRef FACCard As FACCard, ByRef baseAdvance As Boolean, _
                    ByRef suppDescription As String)
        Dim highestHitNumberOnBatter As Integer = HighHitNum(0)
        Dim numIncidents As Integer

        If playResult.Contains("1B") AndAlso _
                    Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).pbRange <> "2-6" AndAlso _
                    UseOtherCard(Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).pbRange) AndAlso _
                    highestHitNumberOnBatter <> 0 Then
            numIncidents = GetRangeIncidents(11, highestHitNumberOnBatter)
            FACCard.Random = SelectIncidentInRange(11, highestHitNumberOnBatter, FetchRandomNumber(numIncidents), "RevPB").ToString
            If FACCard.Random = "0" Then
                Call MsgBox("Random number of 0 is about to get processed from CheckExtraBaseHitEffectPitcher()")
            End If
            playResult = ""
            baseAdvance = False
            suppDescription = ""
            GetResultFromHitterCard(playResult, FACCard, baseAdvance, suppDescription)
        End If
    End Sub

    Private Function GetDivisionType(ByVal abbDivision As String) As String
        Select Case abbDivision
            Case "C"
                Return "CENTRAL"
            Case "E"
                Return "EAST"
            Case "W"
                Return "WEST"
        End Select
        Return ""
    End Function

    Public Function IsFlyBall(ByVal result As String) As Boolean
        Select Case Left(result, 2)
            Case "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "FD"
                Return True
            Case Else
                Return False
        End Select
    End Function

    ''' <summary>
    ''' Determines whether runner is thrown out stretching hit to an extra base. Chances are 1 out of 130. Do not
    ''' attempt if it is a baserunner advance situation, or if it is the 9th inning, the team is losing and runner will
    ''' not represent the tying run.
    ''' </summary>
    ''' <param name="result"></param>
    ''' <param name="baserunnerAdvance"></param>
    ''' <remarks></remarks>
    Public Sub CheckForOutStretchingHit(ByVal result As String, ByVal baserunnerAdvance As Boolean)
        Dim runLead As Integer
        Dim runsScoredOnHit As Integer
        Dim hitType As String
        Dim rndNum As Integer

        If baserunnerAdvance Then
            Exit Sub
        End If
        If Left(result, 3).ToUpper = "1BF" Then
            Exit Sub
        End If
        hitType = Left(result, 2)
        If hitType <> "3B" And hitType <> "2B" And hitType <> "1B" And hitType <> "LR" And hitType <> "RL" Then
            'This routine only applies to base hits
            Exit Sub
        End If
        If hitType = "LR" Or hitType = "RL" Then
            'handle lefty/righty singles
            If Right(BaseSituation, 2) <> "00" Then
                'there is a runner(s) on 1st or 2nd
                Exit Sub
            End If
        End If
        If Game.inning = 9 Then
            If Game.GetResultPtr.resFirst = 4 Then
                runsScoredOnHit += 1
            End If
            If Game.GetResultPtr.resSecond = 4 Then
                runsScoredOnHit += 1
            End If
            If Game.GetResultPtr.resThird = 4 Then
                runsScoredOnHit += 1
            End If
            runLead = Game.BTeam.runs + runsScoredOnHit - Game.PTeam.runs
            If runLead < -1 Then
                'no incentive for batter to push for extra base
                Exit Sub
            End If
        End If
        If Game.outs = 2 And hitType = "2B" Then
            'no need to push for 3rd with 2 outs
            Exit Sub
        End If

        '1 in 130 chance of attempt to stretch hit and get thrown out
        Randomize()
        rndNum = Convert.ToInt32(130 * Rnd() + 1)
        If rndNum = 75 Then
            'batter is thrown out
            Game.GetResultPtr.resBatter = 0
            bolCountRun = True
            Select Case hitType
                Case "RL", "LR", "1B"
                    Game.GetResultPtr.description += " Runner is pushing for 2 and he is OUT at second."
                Case "2B"
                    Game.GetResultPtr.description += " Runner is pushing for third and he is TAGGED OUT."
                Case "3B"
                    Game.GetResultPtr.description += " Runner is trying to score and he is OUT AT THE PLATE."
            End Select

        End If
    End Sub
End Module