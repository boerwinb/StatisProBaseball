Option Strict On
Option Explicit On
Friend Class clsResult
	
    Private _action As String
    Private _batter As Integer
    Private _first As Integer
    Private _second As Integer
    Private _third As Integer
    Private _description As String
    Private _ab As Integer
    Private _hit As Integer
    Private _double As Integer
    Private _triple As Integer
    Private _hr As Integer
    Private _rbi As Integer
    Private _strikeOut As Integer
    Private _walk As Integer
    Private _hpb As Integer
    Private _sf As Integer
    Private _sh As Integer
    Private _ip As Integer
    Private _run As Integer
    Private _wp As Integer
    Private _bk As Integer

    Private _po As String
    Private _a As String
    Private _baseThrow As Integer

    Public Property action() As String
        Get
            Action = _action
        End Get
        Set(ByVal Value As String)
            _action = Value
        End Set
    End Property
    Public Property resBatter() As Integer
        Get
            ResBatter = _batter
        End Get
        Set(ByVal Value As Integer)
            _batter = Value
        End Set
    End Property
    Public Property resFirst() As Integer
        Get
            ResFirst = _first
        End Get
        Set(ByVal Value As Integer)
            _first = Value
        End Set
    End Property
    Public Property resSecond() As Integer
        Get
            ResSecond = _second
        End Get
        Set(ByVal Value As Integer)
            _second = Value
        End Set
    End Property
    Public Property resThird() As Integer
        Get
            ResThird = _third
        End Get
        Set(ByVal Value As Integer)
            _third = Value
        End Set
    End Property
    Public Property description() As String
        Get
            Description = _description
        End Get
        Set(ByVal Value As String)
            _description = Value
        End Set
    End Property
    Public Property ab() As Integer
        Get
            AB = _ab
        End Get
        Set(ByVal Value As Integer)
            _ab = Value
        End Set
    End Property
    Public Property hit() As Integer
        Get
            Hit = _hit
        End Get
        Set(ByVal Value As Integer)
            _hit = Value
        End Set
    End Property
    Public Property doubl() As Integer
        Get
            Doubl = _double
        End Get
        Set(ByVal Value As Integer)
            _double = Value
        End Set
    End Property
    Public Property triple() As Integer
        Get
            Triple = _triple
        End Get
        Set(ByVal Value As Integer)
            _triple = Value
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
    Public Property rbi() As Integer
        Get
            RBI = _rbi
        End Get
        Set(ByVal Value As Integer)
            _rbi = Value
        End Set
    End Property
    Public Property strikeOut() As Integer
        Get
            StrikeOut = _strikeOut
        End Get
        Set(ByVal Value As Integer)
            _strikeOut = Value
        End Set
    End Property
    Public Property walk() As Integer
        Get
            Walk = _walk
        End Get
        Set(ByVal Value As Integer)
            _walk = Value
        End Set
    End Property
    Public Property hpb() As Integer
        Get
            HPB = _hpb
        End Get
        Set(ByVal Value As Integer)
            _hpb = Value
        End Set
    End Property
    Public Property sh() As Integer
        Get
            SH = _sh
        End Get
        Set(ByVal Value As Integer)
            _sh = Value
        End Set
    End Property
    Public Property sf() As Integer
        Get
            SF = _sf
        End Get
        Set(ByVal Value As Integer)
            _sf = Value
        End Set
    End Property
    Public Property ip() As Integer
        Get
            IP = _ip
        End Get
        Set(ByVal Value As Integer)
            _ip = Value
        End Set
    End Property
    Public Property run() As Integer
        Get
            RUN = _run
        End Get
        Set(ByVal Value As Integer)
            _run = Value
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
    Public Property bk() As Integer
        Get
            BK = _bk
        End Get
        Set(ByVal Value As Integer)
            _bk = Value
        End Set
    End Property
    Public Property po() As String
        Get
            po = _po
        End Get
        Set(ByVal value As String)
            _po = value
        End Set
    End Property
    Public Property a() As String
        Get
            a = _a
        End Get
        Set(ByVal value As String)
            _a = value
        End Set
    End Property
    Public Property baseThrow() As Integer
        Get
            baseThrow = _baseThrow
        End Get
        Set(value As Integer)
            _baseThrow = value
        End Set
    End Property

    ''' <summary>
    ''' Performs a lookup on the out charts to get details on the out that occurred.
    ''' </summary>
    ''' <param name="tableKey"></param>
    ''' <remarks></remarks>
    Public Sub ChartLookup(ByRef tableKey As String)
        Dim dsCharts As DataSet = Nothing
        Dim filDebug As Integer
        Dim rndNum As Integer

        Dim FACDataAccess As New clsDataAccess("fac")

        Try
            If Dir(gAppPath & "debug.bbb") <> Nothing Then
                filDebug = FreeFile()
                FileOpen(filDebug, gAppPath & "debug.out", OpenMode.Append)
                PrintLine(filDebug, "In Chart Lookup")
                FileClose(filDebug)
            End If

            System.Diagnostics.Debug.WriteLine(tableKey)
            If tableKey.IndexOf(conNoPlay) > -1 Then
                Exit Sub
            End If
            'Use In/Back indicator if it is ground ball/runner on third situation
            'Handle Offensive/Defensive Options
            If ThirdBase.occupied And tableKey.Substring(0, 1) = "G" Then
                If bolInfieldIn Then
                    tableKey += "IN"
                ElseIf bolCornersIn And (Left(tableKey, 3) = "G5A" Or Left(tableKey, 3) = "G3A") Then
                    tableKey += "CI"
                ElseIf bolGuardLine And Left(tableKey, 3) = "Gx5" Then
                    tableKey += "GL"
                Else
                    tableKey += "BA"
                End If
            ElseIf bolCornersIn And (Left(tableKey, 3) = "G5A" Or Left(tableKey, 3) = "G3A") Then
                tableKey += "CI"
            ElseIf bolGuardLine And Left(tableKey, 3) = "2B7" Then
                Randomize()
                rndNum = Convert.ToInt32(1000 * Rnd() + 1)
                If rndNum Mod 2 = 1 Then
                    tableKey = "Gx5" & tableKey.Substring(3)
                End If
            ElseIf bolGuardLine And Left(tableKey, 3) = "Gx5" Then
                tableKey = "2B7" & tableKey.Substring(3)
            End If

            dsCharts = FACDataAccess.ExecuteDataSet("Select * FROM CHARTS WHERE ACTION = '" & tableKey & "'")
            If dsCharts.Tables(0).Rows.Count = 0 Then
                'If .EOF Then
                MsgBox("ID not found! TableKey = " & tableKey)
            Else
                With dsCharts.Tables(0).Rows(0)
                    _action = .Item("action").ToString
                    _batter = CInt(.Item("batter").ToString)
                   _first = CInt(.Item("firstrunner").ToString)
                    _second = CInt(.Item("secondrunner").ToString)
                    _third = CInt(.Item("thirdrunner").ToString)
                    If Game.outs < 2 Then
                        If Not IsDBNull(.Item("description1")) Then
                            _description = .Item("description1").ToString
                        Else
                            _description = ""
                        End If
                        If Not IsDBNull(.Item("po1")) Then
                            _po = .Item("po1").ToString
                        Else
                            _po = ""
                        End If
                        If Not IsDBNull(.Item("a1")) Then
                            _a = .Item("a1").ToString
                        Else
                            _a = ""
                        End If
                    Else
                        If Not IsDBNull(.Item("description2")) Then
                            _description = .Item("description2").ToString
                        Else
                            _description = ""
                        End If
                        If Not IsDBNull(.Item("po2")) Then
                            _po = .Item("po2").ToString
                        Else
                            _po = ""
                        End If
                        If Not IsDBNull(.Item("a2")) Then
                            _a = .Item("a2").ToString
                        Else
                            _a = ""
                        End If
                    End If
                    If Not IsDBNull(.Item("ab")) Then
                        _ab = CInt(.Item("ab"))
                    Else
                        _ab = 0
                    End If
                    If Not IsDBNull(.Item("hit")) Then
                        _hit = CInt(.Item("hit"))
                    Else
                        _hit = 0
                    End If
                    If Not IsDBNull(.Item("double")) Then
                        _double = CInt(.Item("double"))
                    Else
                        _double = 0
                    End If
                    If Not IsDBNull(.Item("triple")) Then
                        _triple = CInt(.Item("triple"))
                    Else
                        _triple = 0
                    End If
                    If Not IsDBNull(.Item("hr")) Then
                        _hr = CInt(.Item("hr"))
                    Else
                        _hr = 0
                    End If
                    If Not IsDBNull(.Item("rbi")) Then
                        _rbi = CInt(.Item("rbi"))
                    Else
                        _rbi = 0
                    End If
                    If Not IsDBNull(.Item("strikeout")) Then
                        _strikeOut = CInt(.Item("strikeout"))
                    Else
                        _strikeOut = 0
                    End If
                    If Not IsDBNull(.Item("walk")) Then
                        _walk = CInt(.Item("walk"))
                    Else
                        _walk = 0
                    End If
                    If Not IsDBNull(.Item("hpb")) Then
                        _hpb = CInt(.Item("hpb"))
                    Else
                        _hpb = 0
                    End If
                    If Not IsDBNull(.Item("sh")) And Game.outs < 2 Then
                        _sh = CInt(.Item("sh"))
                    Else
                        _sh = 0
                    End If
                    If Not IsDBNull(.Item("sf")) Then
                        _sf = CInt(.Item("sf"))
                        If _sf = 1 And Game.outs >= 2 Then
                            'Remove the SF, give them credit for an AB
                            _sf = 0
                            _ab = 1
                        End If
                    Else
                        _sf = 0
                    End If
                    If Not IsDBNull(.Item("ip")) Then
                        _ip = CInt(.Item("ip"))
                    Else
                        _ip = 0
                    End If
                    If Not IsDBNull(.Item("run")) Then
                        _run = CInt(.Item("run"))
                    Else
                        _run = 0
                    End If
                    If Not IsDBNull(.Item("wp")) Then
                        _wp = CInt(.Item("wp"))
                    Else
                        _wp = 0
                    End If
                    If Not IsDBNull(.Item("bk")) Then
                        _bk = CInt(.Item("bk"))
                    Else
                        _bk = 0
                    End If
                End With
                _baseThrow = 0
            End If
        Catch ex As Exception
            Call MsgBox("Error in ChartLookup. " & ex.ToString & " " & tableKey, MsgBoxStyle.OkOnly)
        Finally
            'release objects
            If Not dsCharts Is Nothing Then dsCharts.Dispose()
        End Try
    End Sub

    Public Sub Clear()
        _action = ""
        _Batter = 0
        _First = 0
        _Second = 0
        _Third = 0
        _Description = ""
        _AB = 0
        _Hit = 0
        _Double = 0
        _Triple = 0
        _HR = 0
        _RBI = 0
        _StrikeOut = 0
        _Walk = 0
        _HPB = 0
        _SH = 0
        _SF = 0
        _IP = 0
        _RUN = 0
        _WP = 0
        _bk = 0
        _po = ""
        _a = ""
        _baseThrow = 0
    End Sub
End Class