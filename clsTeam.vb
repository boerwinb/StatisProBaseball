Option Strict On
Option Explicit On
Friend Class clsTeam
	
    Private _Batter(38) As clsBatter
    Private _Pitcher(35) As clsPitcher
    Private _teamName As String
    Private _id As String
    Private _orderSpot As Integer
    Private _lineUp(9) As Integer
    Private _pitcherSel As Integer
    Private _starter As Integer
    Private _pr As Integer
    Private _runs As Integer
    Private _hits As Integer
    Private _errors As Integer
    Private _dps As Integer
    Private _pitcherUsed As Integer
    Private _pitchChange As Boolean
    Private _numInactiveHitters As Integer
    Private _numHitters As Integer
    Private _numPitchers As Integer
    Private _numInactivePitchers As Integer
    Private _consecutivePOE As Integer
    Private _adjK As Double
    Private _adjBB As Double
    Private _adjSingle As Double
    Private _singlePct As Double
    Private _shPct As Double

    Public Function GetBatterPtr(ByVal index As Integer) As clsBatter
        GetBatterPtr = _Batter(index)
    End Function

    Public Function GetPitcherPtr(ByVal index As Integer) As clsPitcher
        GetPitcherPtr = _Pitcher(index)
    End Function

    Public Property teamName() As String
        Get
            Return _teamName
        End Get
        Set(ByVal Value As String)
            _teamName = Value
        End Set
    End Property
    Public Property id() As String
        Get
            Return _id
        End Get
        Set(ByVal Value As String)
            _id = Value
        End Set
    End Property
    Public Property pitcherSel() As Integer
        Get
            Return _pitcherSel
        End Get
        Set(ByVal Value As Integer)
            _pitcherSel = Value
        End Set
    End Property
    Public Property starter() As Integer
        Get
            Return _starter
        End Get
        Set(ByVal Value As Integer)
            _starter = Value
        End Set
    End Property
    Public Property pr() As Integer
        Get
            Return _pr
        End Get
        Set(ByVal Value As Integer) 'Amount of remaining stamina for current pitcher
            _pr = Value
        End Set
    End Property
    Public Property runs() As Integer
        Get
            Return _runs
        End Get
        Set(ByVal Value As Integer)
            _runs = Value
        End Set
    End Property
    Public Property hits() As Integer
        Get
            Return _hits
        End Get
        Set(ByVal Value As Integer)
            _hits = Value
        End Set
    End Property
    Public Property errors() As Integer
        Get
            Return _errors
        End Get
        Set(ByVal Value As Integer)
            _errors = Value
        End Set
    End Property
    Public Property dps() As Integer
        Get
            Return _dps
        End Get
        Set(ByVal Value As Integer)
            _dps = Value
        End Set
    End Property
    Public Property order() As Integer
        Get
            Return _orderSpot
        End Get
        Set(ByVal Value As Integer)
            _orderSpot = Value
        End Set
    End Property
    Public Property nthPitcherUsed() As Integer
        Get
            Return _pitcherUsed
        End Get
        Set(ByVal Value As Integer)
            _pitcherUsed = Value
        End Set
    End Property
    Public Property pitchChange() As Boolean
        Get
            Return _pitchChange
        End Get
        Set(ByVal Value As Boolean)
            _pitchChange = Value
        End Set
    End Property
    Public Property inactiveHitters() As Integer
        Get
            Return _numInactiveHitters
        End Get
        Set(ByVal Value As Integer)
            _numInactiveHitters = Value
        End Set
    End Property
    Public Property hitters() As Integer
        Get
            Return _numHitters
        End Get
        Set(ByVal Value As Integer)
            _numHitters = Value
        End Set
    End Property
    Public Property pitchers() As Integer
        Get
            Return _numPitchers
        End Get
        Set(ByVal Value As Integer)
            _numPitchers = Value
        End Set
    End Property
    Public Property inactivePitchers() As Integer
        Get
            Return _numInactivePitchers
        End Get
        Set(ByVal Value As Integer)
            _numInactivePitchers = Value
        End Set
    End Property
    Public Property consecutivePOE() As Integer
        Get
            Return _consecutivePOE
        End Get
        Set(ByVal value As Integer)
            _consecutivePOE = value
        End Set
    End Property
    Public Property adjK As Double
        Get
            Return _adjK
        End Get
        Set(value As Double)
            _adjK = value
        End Set
    End Property
    Public Property adjBB As Double
        Get
            Return _adjBB
        End Get
        Set(value As Double)
            _adjBB = value
        End Set
    End Property
    Public Property adjSingle As Double
        Get
            Return _adjSingle
        End Get
        Set(value As Double)
            _adjSingle = value
        End Set
    End Property
    Public Property singlePct As Double
        Get
            Return _singlePct
        End Get
        Set(value As Double)
            _singlePct = value
        End Set
    End Property
    Public Property shPct As Double
        Get
            Return _shpct
        End Get
        Set(value As Double)
            _shpct = value
        End Set
    End Property

    Public Sub Clear()
        _TeamName = ""
        _ID = ""
        _Runs = 0
        _Hits = 0
        _Errors = 0
        _DPs = 0
        _OrderSpot = 1
        _PitcherUsed = 1
        _PitchChange = False
        _NumInactiveHitters = 0
        _NumHitters = 0
        _NumPitchers = 0
        _numInactivePitchers = 0
        _consecutivePOE = 0
        _adjK = 0
        _adjBB = 0
        _adjSingle = 0
        _singlePct = 0
        _shPct = 0
    End Sub

    Public Sub New()
        MyBase.New()

        For i As Integer = 1 To conMaxBatters + IIF(bolAmericanLeagueRules, 0, 10)
            _Batter(i) = New clsBatter
        Next i
        For i As Integer = 1 To conMaxPitchers
            _Pitcher(i) = New clsPitcher
        Next i
        Clear()
    End Sub
    ''' <summary>
    ''' add a player index to the lineup
    ''' </summary>
    ''' <param name="order"></param>
    ''' <param name="playerIndex"></param>
    ''' <remarks></remarks>
    Public Sub AddLineup(ByRef order As Integer, ByRef playerIndex As Integer)
        _lineUp(order) = playerIndex
    End Sub

    ''' <summary>
    ''' remove a player index from the lineup
    ''' </summary>
    ''' <param name="order"></param>
    ''' <remarks></remarks>
    Public Sub RemoveLineup(ByRef order As Integer)
        _lineUp(order) = 0 'Clear lineup
    End Sub

    ''' <summary>
    ''' Fetch the player index based on where they bat in the lineup
    ''' </summary>
    ''' <param name="order"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetPlayerNum(Optional ByRef order As Integer = 0) As Integer
        If IsNothing(order) Then
            Return _lineUp(_orderSpot)
        Else
            Return _lineUp(order)
        End If
    End Function

    ''' <summary>
    ''' returns the player index based on the submitted player name
    ''' </summary>
    ''' <param name="playerName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetIndexFromName(ByRef playerName As String) As Integer
        Dim index As Integer

        Try
            'For i As Integer = 1 To conMaxBatters
            'If Not isPitcher Then
            For i As Integer = 1 To _numHitters
                If playerName = _Batter(i).player Then
                    index = _Batter(i).playerIndex
                    Exit For
                End If
            Next i
            'End If
            If Not bolAmericanLeagueRules And index = 0 Then
                'Must be pitcher
                index = _numHitters + _pitcherUsed
            End If
        Catch ex As Exception
            Call MsgBox("Error in GetIndexFromName" & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return index
    End Function

    ''' <summary>
    ''' get pitcher index based on the player name
    ''' </summary>
    ''' <param name="playerName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetPIndexFromName(ByRef playerName As String) As Integer
        Dim index As Integer
        Try
            For i As Integer = 1 To Me._numPitchers
                If playerName = _Pitcher(i).player Then
                    index = _Pitcher(i).pitcherNum
                    Exit For
                End If
            Next i
        Catch ex As Exception
            Call MsgBox("Error in GetPIndexFromName" & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return index
    End Function

    ''' <summary>
    ''' Determines whether a player is in the lineup based on the submitted player index
    ''' </summary>
    ''' <param name="playerIndex"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function InLineup(ByRef playerIndex As Integer) As Boolean
        Dim isFound As Boolean
        Dim i As Integer
        Try
            i = 1
            While i < 10 And Not isFound
                If _lineUp(i) <> 0 Then
                    isFound = _lineUp(i) = playerIndex
                End If
                i += 1
            End While
        Catch ex As Exception
            Call MsgBox("Error in InLineup" & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return isFound
    End Function

    ''' <summary>
    ''' Updates the lineup index
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub UpdateLineUp()
        _orderSpot += 1
        If _orderSpot > 9 Then _orderSpot = 1
    End Sub

    ''' <summary>
    ''' Returns the number of available hitters
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetNumAvailablePlayers() As Integer
        Dim playerCount As Integer

        Try
            For i As Integer = 0 To _numHitters - 1
                With GetBatterPtr(i + 1)
                    If .player <> Nothing Then
                        If Not InLineup(.playerIndex) Then
                            If .available Then
                                playerCount += 1
                            End If
                        End If
                    End If
                End With
            Next i
        Catch ex As Exception
            Call MsgBox("Error in GetNumAvailablePlayers" & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return playerCount
    End Function

    ''' <summary>
    ''' returns the number of available pitchers
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetNumAvailablePitchers() As Integer
        Dim pitcherCount As Integer

        Try
            For i As Integer = 0 To _numPitchers - 1
                With GetPitcherPtr(i + 1)
                    If .available Then
                        'Only available relief pitchers
                        If Val(.rr) > 0 Then
                            pitcherCount += 1
                        End If
                    End If
                End With
            Next i
        Catch ex As Exception
            Call MsgBox("Error in GetNumAvailablePitchers" & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return pitcherCount
    End Function

    ''' <summary>
    ''' returns the stadium data based on the team and season
    ''' </summary>
    ''' <param name="attendance">out parameter representing today's attendance</param>
    ''' <returns>the ballpark based on the team and season</returns>
    ''' <remarks></remarks>
    Public Function GetStadiumStats(ByRef attendance As Int64) As String
        Dim dsStat As DataSet = Nothing
        Dim query As String
        Dim park As String = ""
        Dim games As Integer
        Dim annualAttendance As Int64
        Dim DataAccess As New clsDataAccess("lahman")

        Try
            query = "SELECT park, ghome, attendance FROM teams WHERE teamid='" & _id & "' AND yearid=" & gstrSeason
            dsStat = DataAccess.ExecuteDataSet(query)
            If dsStat.Tables(0).Rows.Count > 0 Then
                With dsStat.Tables(0).Rows(0)
                    park = .Item("park").ToString
                    games = CInt(.Item("ghome"))
                    annualAttendance = CInt(.Item("attendance"))
                End With
                Randomize()
                attendance = CType(Math.Ceiling(annualAttendance / games * 0.8 + annualAttendance / games * 0.4 * Rnd()), Int64)
            End If
        Catch ex As Exception
            Call MsgBox("Error in GetStadiumStats" & ex.ToString, MsgBoxStyle.OkOnly)
        Finally
            If Not dsStat Is Nothing Then dsStat.Dispose()
        End Try
        Return park
    End Function

    ''' <summary>
    ''' Determines the team logo based on the team and season.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetLogo() As String
        Dim logoFilename As String = ""
        Dim dsStat As DataSet = Nothing
        Dim query As String
        Dim DataAccess As New clsDataAccess("fac")

        Try
            query = "SELECT logofile FROM Logos WHERE teamid = '" & _teamName & "' AND year1 <= " & gstrSeason & _
                    " AND year2 >= " & gstrSeason
            dsStat = DataAccess.ExecuteDataSet(query)
            If dsStat.Tables(0).Rows.Count > 0 Then
                With dsStat.Tables(0).Rows(0)
                    logoFilename = .Item("logofile").ToString
                End With
            End If
        Catch ex As Exception
            Call MsgBox("Error in GetLogo" & ex.ToString, MsgBoxStyle.OkOnly)
        Finally
            If Not dsStat Is Nothing Then dsStat.Dispose()
        End Try
        Return logoFilename
    End Function

    Public Function GetPlayerAge(ByVal dob As String) As String
        Dim ds As DataSet = Nothing
        Dim query As String
        Dim teamGames As Integer
        Dim currentDate As Date
        Dim DataAccess As New clsDataAccess(gstrSeason & "data")

        Try
            query = "SELECT wins, losses FROM standings WHERE playerid = '" & _teamName & "'"
            ds = DataAccess.ExecuteDataSet(query)
            If ds.Tables(0).Rows.Count = 0 Then
                teamGames = 0
            Else
                With ds.Tables(0).Rows(0)
                    teamGames = CInt(.Item("wins")) + CInt(.Item("losses"))
                End With
            End If
            ds.Clear()

            currentDate = DateAdd(DateInterval.Day, teamGames, CDate("4/10/" & gstrSeason))
            Return CalcAge(CDate(dob), currentDate).ToString
        Catch ex As Exception
            Call MsgBox("Error in GetPlayerAge" & ex.ToString, MsgBoxStyle.OkOnly)
        Finally
            If Not ds Is Nothing Then ds.Dispose()
        End Try
        Return ""
    End Function

    ''' <summary>
    ''' Determines the manager based on the team, season, and the number of games that have player during the season
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetManager() As String
        Dim ds As DataSet = Nothing
        Dim query As String
        Dim managerName As String = ""
        Dim games As Integer
        Dim teamGames As Integer
        Dim DataAccess As New clsDataAccess(gstrSeason & "data")
        Dim LahmanDataAccess As New clsDataAccess("lahman")

        Try
            query = "SELECT wins, losses FROM standings WHERE playerid = '" & _teamName & "'"
            ds = DataAccess.ExecuteDataSet(query)
            If ds.Tables(0).Rows.Count = 0 Then
                teamGames = 0
            Else
                With ds.Tables(0).Rows(0)
                    teamGames = CInt(.Item("wins")) + CInt(.Item("losses"))
                End With
            End If
            ds.Clear()
            query = "SELECT managers.g, people.namefirst, people.namelast FROM MANAGERS INNER JOIN PEOPLE ON " & _
                    "managers.playerid = people.playerid " & _
                    "WHERE teamid='" & _id & "' AND yearid=" & gstrSeason & _
                    " ORDER BY managers.inseason ASC"
            ds = Lahmandataaccess.ExecuteDataSet(query)
            For Each dr As DataRow In ds.Tables(0).Rows
                games += CInt(dr.Item("g"))
                managerName = dr.Item("namefirst").ToString & " " & dr.Item("namelast").ToString
                If games >= teamGames Then
                    Return managerName
                End If
            Next dr
            'return last manager in recordset if none is returned above
            Return managerName
        Catch ex As Exception
            Call MsgBox("Error in GetManager" & ex.ToString, MsgBoxStyle.OkOnly)
        Finally
            If Not ds Is Nothing Then ds.Dispose()
        End Try
        Return ""
    End Function

    ''' <summary>
    ''' Determine real life league adjustments for Ks, BBs and Singles. These are used to determine high statistical variance
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub SetVarianceAdjustments()
        Dim ds As DataSet = Nothing
        Dim query As String
        'Dim managerName As String = ""
        'Dim games As Integer
        'Dim teamGames As Integer
        'Dim DataAccess As New clsDataAccess(gstrSeason & "data")
        Dim LahmanDataAccess As New clsDataAccess("lahman")
        Dim league As String = ""
        Dim year As Integer
        Dim outcomeRatio As Double
        Dim totalSingles As Integer
        Dim totalH As Integer
        Dim totalD As Integer
        Dim totalT As Integer
        Dim totalHR As Integer
        Dim totalK As Integer
        Dim totalBB As Integer
        Dim totalIBBS As Integer
        Dim totalBFP As Integer
        Dim totalSH As Integer

        Try
            league = GetDivision(_teamName).Substring(0, 2)
            year = CInt(gstrSeason)
            If year > 1915 Then
                'Batters Faced were tracked after 1915
                query = "SELECT sum(bfp) as totalbfp FROM pitching WHERE lgid = '" & _
                            league & "' AND yearid = " & year.ToString
                ds = LahmanDataAccess.ExecuteDataSet(query)
                With ds.Tables(0).Rows(0)
                    totalBFP = CInt(.Item("totalbfp"))
                End With
            Else
                'Otherwise, try to determine from other stats
                query = "SELECT sum(ipouts) as totalouts, sum(bb) as totalbb, sum(ibb) as totalibb, " & _
                                "sum(h) as totalhits FROM pitching WHERE lgid = '" & _
                            league & "' AND yearid = " & year.ToString
                ds = LahmanDataAccess.ExecuteDataSet(query)
                With ds.Tables(0).Rows(0)
                    totalBFP = CInt(CInt(.Item("totalouts")) * (26 / 27)) + CInt(.Item("totalbb")) + _
                                    CInt(.Item("totalhits")) ' - CheckField(.item("totalibb"), 0)
                End With
            End If

            query = "SELECT sum(h) as totalhits, sum(d) as totalds, sum(t) as totalts, " & _
                                  "sum(hr) as totalhrs, sum(so) as totalks, sum(bb) as totalbbs, " & _
                                  "sum(ibb) as totalibb, sum(sh) as totalsh " & _
                                  "FROM batting WHERE lgid = '" & league & "' AND " & _
                                  "yearid = " & year.ToString

            ds = LahmanDataAccess.ExecuteDataSet(query)
            With ds.Tables(0).Rows(0)
                totalH = CInt(.Item("totalhits"))
                totalD = CInt(.Item("totalds"))
                totalT = CInt(.Item("totalts"))
                totalHR = CInt(.Item("totalhrs"))
                totalK = CheckField(.Item("totalks"), 0)
                totalBB = CInt(.Item("totalbbs"))
                totalIBBS = CheckField(.Item("totalibb"), 0)
                totalSH = CheckField(.Item("totalsh"), 0)
                'rs.Close()
            End With

            totalSingles = totalH - totalD - totalT - totalHR
            outcomeRatio = (totalBFP - totalIBBS - totalSH) / 128
            _adjSingle = (totalSingles / outcomeRatio) / 2
            _adjK = (totalK / outcomeRatio) / 2
            _adjBB = ((totalBB - totalIBBS) / outcomeRatio) / 2
            _singlePct = totalSingles / totalH
            _shPct = totalSH / totalBFP
        Catch ex As Exception
            Call MsgBox("Error in GetManager" & ex.ToString, MsgBoxStyle.OkOnly)
        Finally
            If Not ds Is Nothing Then ds.Dispose()
        End Try

    End Sub

    ''' <summary>
    ''' determines adjustment factors for singles, strikeouts, walks based on league averages for that season
    ''' </summary>
    ''' <param name="league"></param>
    ''' <param name="year"></param>
    ''' <remarks></remarks>
    Private Sub DetermineLeagueAvgs(ByVal league As String, ByVal year As Integer, ByRef DataAccess As clsDataAccess)
        Dim outcomeRatio As Double
        Dim totalSingles As Integer
        Dim totalH As Integer
        Dim totalD As Integer
        Dim totalT As Integer
        Dim totalHR As Integer
        Dim totalK As Integer
        Dim totalBB As Integer
        Dim totalIBBS As Integer
        Dim totalBFP As Integer
        Dim totalSH As Integer
        Dim sqlQuery As String
        'Dim rs As ADODB.Recordset
        Dim ds As DataSet = Nothing
        'Dim cn As ADODB.Connection


        Try
            'rs = New ADODB.Recordset
            'cn = New ADODB.Connection
            'cn.Open(GetLahmanConnectString)
            If year > 1915 Then
                'Batters Faced were tracked after 1915
                sqlQuery = "SELECT sum(bfp) as totalbfp FROM pitching WHERE lgid = '" & _
                            league & "' AND yearid = " & year.ToString
                'rs.Open(sqlQuery, cn)
                ds = DataAccess.ExecuteDataSet(sqlQuery)
                With ds.Tables(0).Rows(0)
                    totalBFP = CInt(.Item("totalbfp"))
                End With
            Else
                'Otherwise, try to determine from other stats
                sqlQuery = "SELECT sum(ipouts) as totalouts, sum(bb) as totalbb, sum(ibb) as totalibb, " & _
                                "sum(h) as totalhits FROM pitching WHERE lgid = '" & _
                            league & "' AND yearid = " & year.ToString
                'rs.Open(sqlQuery, cn)
                ds = DataAccess.ExecuteDataSet(sqlQuery)
                With ds.Tables(0).Rows(0)
                    totalBFP = CInt(CInt(.Item("totalouts")) * (26 / 27)) + CInt(.Item("totalbb")) + _
                                    CInt(.Item("totalhits")) ' - CheckField(.item("totalibb"), 0)
                End With
            End If

            'rs.Close()

            sqlQuery = "SELECT sum(h) as totalhits, sum(d) as totalds, sum(t) as totalts, " & _
                                  "sum(hr) as totalhrs, sum(so) as totalks, sum(bb) as totalbbs, " & _
                                  "sum(ibb) as totalibb, sum(sh) as totalsh " & _
                                  "FROM batting WHERE lgid = '" & league & "' AND " & _
                                  "yearid = " & year.ToString

            'rs.Open(sqlQuery, cn)
            ds = DataAccess.ExecuteDataSet(sqlQuery)
            With ds.Tables(0).Rows(0)
                totalH = CInt(.Item("totalhits"))
                totalD = CInt(.Item("totalds"))
                totalT = CInt(.Item("totalts"))
                totalHR = CInt(.Item("totalhrs"))
                totalK = CheckField(.Item("totalks"), 0)
                totalBB = CInt(.Item("totalbbs"))
                totalIBBS = CheckField(.Item("totalibb"), 0)
                totalSH = CheckField(.Item("totalsh"), 0)
                'rs.Close()
            End With

            totalSingles = totalH - totalD - totalT - totalHR
            outcomeRatio = (totalBFP - totalIBBS - totalSH) / 128
            _adjSingle = (totalSingles / outcomeRatio) / 2
            _adjK = (totalK / outcomeRatio) / 2
            _adjBB = ((totalBB - totalIBBS) / outcomeRatio) / 2
        Catch ex As Exception
            Call MsgBox("DetermineLeagueAvgs " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub
End Class