Option Strict On
Option Explicit On

Imports System.Configuration.ConfigurationManager
Imports Microsoft.Win32

Public Class frmAdvance2

    Dim hitType As String
    Dim advHitType As Integer
    Dim description As String
    Dim userMessage As String
    Dim advanceResult As String
    Dim successRange As String
    Dim successRangeBatter As String
    Dim batterObr As String
    Dim runnerID As Integer
    Const RESULT_SAFE As String = "SAFE!"

    Private Sub cmdYes_Click(sender As Object, e As EventArgs) Handles cmdYes.Click
        Dim cutOffThrow As Boolean = False
        Dim structFAC As FACCard
        Dim doubleAdv As Boolean = False
        Dim frmBaseThrow As frmBaseThrow
        Dim batterAdvance As Boolean = False
        Dim dsChart As DataSet = Nothing
        Dim baseAdvanceSituation As String = ""

        'Dim sqlQuery As String = ""
        'Dim structFAC As FACCard
        'Dim facNumber As Integer
        Dim FACDataAccess As New clsDataAccess("fac")
        Dim successRangeTrailer As String = ""
        Dim obrRatingTrailer As String = ""
        Dim armRating As String = ""

        Try
            If hitType = "1B" Then
                Select Case BaseSituation()
                    Case "011", "111"
                        obrRatingTrailer = Game.BTeam.GetBatterPtr(FirstBase.runner).obr
                        Select Case Game.currentFieldPos
                            Case 7
                                baseAdvanceSituation = obrRatingTrailer & "1"
                            Case 8
                                baseAdvanceSituation = obrRatingTrailer & "2"
                            Case 9
                                baseAdvanceSituation = obrRatingTrailer & "3"
                        End Select
                        armRating = Game.PTeam.GetBatterPtr(Game.GetFielderNum).throwRating
                        If AppSettings("BaseAdvance_AdvanceRules").ToString.ToUpper = "ON" Then
                            baseAdvanceSituation = baseAdvanceSituation & (advHitType + 1).ToString
                            successRangeTrailer = GetSuccessRangeAdv(baseAdvanceSituation, armRating, advHitType)
                        Else
                            successRangeTrailer = GetSuccessRange(baseAdvanceSituation, obrRatingTrailer, armRating)
                        End If
                        doubleAdv = MsgBox("Will the runner on first advance to third? Success range is " & successRangeTrailer & ".", _
                                            MsgBoxStyle.YesNo, "Offensive option. Outs:" & Game.outs.ToString) = MsgBoxResult.Yes
                End Select

                If (BaseSituation() <> "111" And BaseSituation() <> "011") Or doubleAdv Then
                    'the defensive player must choose the base to which he is throwing the ball
                    frmBaseThrow = New frmBaseThrow
                    frmBaseThrow.ShowDialog(Me)
                End If

                If Game.GetResultPtr.baseThrow <> 2 Then
                    If (BaseSituation() <> "111" And BaseSituation() <> "011") Or doubleAdv Then
                        'The defensive player is NOT trying to keep the batter from taking second base, at least on the initial throw,
                        'and the base is open
                        dsChart = FACDataAccess.ExecuteDataSet("Select * FROM BATTERADVANCE WHERE OBR = '" & batterObr & "'")

                        If dsChart.Tables(0).Rows.Count = 0 Then
                            MsgBox("ID not found!")
                        Else
                            With dsChart.Tables(0).Rows(0)
                                successRangeBatter = .Item("Range").ToString
                            End With
                        End If
                        batterAdvance = MsgBox("Will the batter attempt to take an extra base on the throw? Success range is " & successRangeBatter & _
                                                " if the throw is cut off.", MsgBoxStyle.YesNo, "Offensive Option. Outs:" & Game.outs.ToString) = MsgBoxResult.Yes
                    End If
                End If
                If batterAdvance Then
                    cutOffThrow = MsgBox("Will the defense cut off throw?", MsgBoxStyle.YesNo, "Defensive Option. Outs:" & Game.outs.ToString) = MsgBoxResult.Yes
                End If
            End If
            If cutOffThrow Then
                gblUpdateMsg = gblUpdateMsg & userMessage & RESULT_SAFE & " The throw was cutoff."
                If runnerID = 1 Then
                    Game.GetResultPtr.resFirst = 3
                Else
                    Game.GetResultPtr.resSecond = 4
                    If doubleAdv Then
                        Game.GetResultPtr.resFirst = 3
                    End If
                End If
                structFAC = GetFAC("batterAdv", "Random")
                If InRange(successRangeBatter, CType(structFAC.Random, Integer)) Then
                    'Batter takes 2nd base even with the cut off throw
                    gblUpdateMsg = gblUpdateMsg & " Throw is to 2nd and the batter is...SAFE!"
                    Game.GetResultPtr.resBatter = 2
                Else
                    'Batter is out advancing for 2nd
                    gblUpdateMsg = gblUpdateMsg & " Throw is to 2nd and the batter is...TAGGED OUT!"
                    Game.GetResultPtr.resBatter = 0
                End If
            Else
                'could be a single or double, and there is no cutoff
                If Game.GetResultPtr.baseThrow = 2 And hitType = "1B" Then
                    'defense wants to keep the batter at first. All advancing runners are safe
                    gblUpdateMsg = gblUpdateMsg & description & RESULT_SAFE & "."
                    If runnerID = 1 Then
                        Game.GetResultPtr.resFirst = 3
                    Else
                        Game.GetResultPtr.resSecond = 4
                        If doubleAdv Then
                            Game.GetResultPtr.resFirst = 3
                        End If
                    End If
                ElseIf doubleAdv And Game.GetResultPtr.baseThrow = 3 Then
                    'defense is targeting the trailing runner
                    description = Game.BTeam.GetBatterPtr(FirstBase.runner).player
                    description = description.Substring(description.IndexOf(" ") + 1) & " is trying for third and he's "
                    advanceResult = "tagged OUT at third."
                    structFAC = GetFAC("baseAdv", "Random")
                    If InRange(successRangeTrailer, CType(structFAC.Random, Integer)) Then 'successful
                        gblUpdateMsg = gblUpdateMsg & description & RESULT_SAFE & "."
                        Game.GetResultPtr.resFirst = 3
                        Game.GetResultPtr.resSecond = 4
                        If batterAdvance Then 'Throw was not cutoff, so batter gets second base
                            gblUpdateMsg = gblUpdateMsg & " Batter to second on throw."
                            Game.GetResultPtr.resBatter = 2
                        End If
                    Else 'Thrown out
                        gblUpdateMsg = gblUpdateMsg & description & advanceResult & "."
                        Game.GetResultPtr.resFirst = 0
                        Game.GetResultPtr.resSecond = 4 'Assume runner scores before 3rd out
                        If batterAdvance And Game.outs < 2 Then 'Throw was not cutoff, so batter gets second base
                            gblUpdateMsg = gblUpdateMsg & " Runner scores from 2nd. Batter to second on throw."
                            Game.GetResultPtr.resBatter = 2
                        End If
                    End If
                Else
                    'Traditional path, where outfielder is targeting the lead advancing runner
                    structFAC = GetFAC("baseAdv", "Random")
                    If InRange(successRange, CType(structFAC.Random, Integer)) Then 'successful
                        gblUpdateMsg = gblUpdateMsg & description & RESULT_SAFE & "."
                        If batterAdvance Then 'Throw was not cutoff, so batter gets second base
                            gblUpdateMsg = gblUpdateMsg & " Batter to second on throw."
                            Game.GetResultPtr.resBatter = 2
                            If runnerID = 1 Then
                                Game.GetResultPtr.resFirst = 3
                            Else
                                Game.GetResultPtr.resSecond = 4
                            End If
                        Else
                            If hitType = "1B" Then
                                If runnerID = 1 Then
                                    Game.GetResultPtr.resFirst = 3
                                Else
                                    Game.GetResultPtr.resSecond = 4
                                End If
                            Else
                                Game.GetResultPtr.resFirst = 4
                            End If
                        End If
                    Else 'Thrown out
                        gblUpdateMsg = gblUpdateMsg & description & advanceResult & "."
                        If batterAdvance And Game.outs < 2 Then 'Throw was not cutoff, so batter gets second base
                            gblUpdateMsg = gblUpdateMsg & " Batter to second on throw."
                            Game.GetResultPtr.resBatter = 2
                            If runnerID = 1 Then
                                Game.GetResultPtr.resFirst = 0
                            Else
                                Game.GetResultPtr.resSecond = 0
                            End If
                        Else
                            If hitType = "1B" And runnerID > 1 Then
                                Game.GetResultPtr.resSecond = 0
                            Else
                                Game.GetResultPtr.resFirst = 0
                            End If
                        End If
                    End If
                    If doubleAdv And Game.GetResultPtr.baseThrow <> 3 Then
                        'The play was at the plate. Trailing runner takes 3rd.
                        Game.GetResultPtr.resFirst = 3
                    End If
                End If
            End If
            Me.Close()
        Catch ex As Exception
            Call MsgBox("Error in cmdYes_click. " & ex.ToString)
            Me.Close()
        End Try
    End Sub

    Private Sub cmdNo_Click(sender As Object, e As EventArgs) Handles cmdNo.Click
        Me.Close()
    End Sub

    
    Private Sub frmAdvance2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim baseAdvanceSituation As String = "" 'Base advance situation. There are 4 situations, one pertaining to each base advance table in the game.
        Dim obrRating As String = ""
        Dim armRating As String = ""
        'Chart Lookup
        Dim dsChart As DataSet = Nothing
        Dim sqlQuery As String = ""

        Try
            If AppSettings("BaseAdvance_AdvanceRules").ToString.ToUpper = "ON" Then
                UseAdvancedRules()
                Exit Sub
            End If

            hitType = Game.GetResultPtr.action.Substring(0, 2)
            'Determine AdvanceSituation
            Select Case BaseSituation()
                'Defined BaseAdvanceScenarios (BAS)
                '1 - First to third on single to left
                '2 - First to third on single to center
                '3 - First to third on single to right, or second to home on single to any field
                '4 - First to home on double to any field
                Case "000"
                    Call MsgBox("Base advance dialog launched incorrectly.", MsgBoxStyle.Exclamation)
                    Exit Sub
                Case "001"
                    obrRating = Game.BTeam.GetBatterPtr(FirstBase.runner).obr
                    If hitType = "1B" Then
                        Select Case Game.currentFieldPos
                            Case 7
                                baseAdvanceSituation = obrRating & "1"
                            Case 8
                                baseAdvanceSituation = obrRating & "2"
                            Case 9
                                baseAdvanceSituation = obrRating & "3"
                        End Select
                        'batterObr = Game.BTeam.GetBatterPtr(Game.currentBatter).obr
                        description = Game.BTeam.GetBatterPtr(FirstBase.runner).player
                        description = description.Substring(description.IndexOf(" ") + 1) & " is trying for third and he's "
                        advanceResult = "tagged OUT at third."
                        runnerID = 1
                    Else
                        baseAdvanceSituation = obrRating & "4"
                        description = Game.BTeam.GetBatterPtr(FirstBase.runner).player
                        description = description.Substring(description.IndexOf(" ") + 1) & " is trying for home and he's "
                        advanceResult = "OUT AT THE PLATE!"
                        runnerID = 1
                    End If
                Case "010"
                    obrRating = Game.BTeam.GetBatterPtr(SecondBase.runner).obr
                    baseAdvanceSituation = obrRating & "3"
                    'batterObr = Game.BTeam.GetBatterPtr(Game.currentBatter).obr
                    description = Game.BTeam.GetBatterPtr(SecondBase.runner).player
                    description = description.Substring(description.IndexOf(" ") + 1) & " is trying for home and he's "
                    advanceResult = "OUT AT THE PLATE!"
                    runnerID = 2
                Case "011"
                    If hitType = "1B" Then
                        obrRating = Game.BTeam.GetBatterPtr(SecondBase.runner).obr
                        baseAdvanceSituation = obrRating & "3"
                        description = Game.BTeam.GetBatterPtr(SecondBase.runner).player
                        description = description.Substring(description.IndexOf(" ") + 1) & " is trying for home and he's "
                        advanceResult = "OUT AT THE PLATE!"
                        runnerID = 2
                    Else
                        obrRating = Game.BTeam.GetBatterPtr(FirstBase.runner).obr
                        baseAdvanceSituation = obrRating & "4"
                        description = Game.BTeam.GetBatterPtr(FirstBase.runner).player
                        description = description.Substring(description.IndexOf(" ") + 1) & " is trying for home and he's "
                        advanceResult = "OUT AT THE PLATE!"
                        runnerID = 1
                    End If
                Case "100"
                    Call MsgBox("Base advance dialog launched incorrectly.", MsgBoxStyle.Exclamation)
                Case "101"
                    obrRating = Game.BTeam.GetBatterPtr(FirstBase.runner).obr
                    If hitType = "1B" Then
                        Select Case Game.currentFieldPos
                            Case 7
                                baseAdvanceSituation = obrRating & "1"
                            Case 8
                                baseAdvanceSituation = obrRating & "2"
                            Case 9
                                baseAdvanceSituation = obrRating & "3"
                        End Select
                        'batterObr = Game.BTeam.GetBatterPtr(Game.currentBatter).obr
                        description = Game.BTeam.GetBatterPtr(FirstBase.runner).player
                        description = description.Substring(description.IndexOf(" ") + 1) & " is trying for third and he's "
                        advanceResult = "tagged OUT at third."
                        runnerID = 1
                    Else
                        baseAdvanceSituation = obrRating & "4"
                        description = Game.BTeam.GetBatterPtr(FirstBase.runner).player
                        description = description.Substring(description.IndexOf(" ") + 1) & " is trying for home and he's "
                        advanceResult = "OUT AT THE PLATE!"
                        runnerID = 1
                    End If
                Case "110"
                    obrRating = Game.BTeam.GetBatterPtr(SecondBase.runner).obr
                    baseAdvanceSituation = obrRating & "3"
                    'batterObr = Game.BTeam.GetBatterPtr(Game.currentBatter).obr
                    description = Game.BTeam.GetBatterPtr(SecondBase.runner).player
                    description = description.Substring(description.IndexOf(" ") + 1) & " is trying for home and he's "
                    advanceResult = "OUT AT THE PLATE!"
                    runnerID = 2
                Case "111"
                    If hitType = "1B" Then
                        obrRating = Game.BTeam.GetBatterPtr(SecondBase.runner).obr
                        baseAdvanceSituation = obrRating & "3"
                        description = Game.BTeam.GetBatterPtr(SecondBase.runner).player
                        description = description.Substring(description.IndexOf(" ") + 1) & " is trying for home and he's "
                        advanceResult = "OUT AT THE PLATE!"
                        runnerID = 2
                    Else
                        obrRating = Game.BTeam.GetBatterPtr(FirstBase.runner).obr
                        baseAdvanceSituation = obrRating & "4"
                        description = Game.BTeam.GetBatterPtr(FirstBase.runner).player
                        description = description.Substring(description.IndexOf(" ") + 1) & " is trying for home and he's "
                        advanceResult = "OUT AT THE PLATE!"
                        runnerID = 1
                    End If
            End Select
            batterObr = Game.BTeam.GetBatterPtr(Game.currentBatter).obr
            armRating = Game.PTeam.GetBatterPtr(Game.GetFielderNum).throwRating
            successRange = GetSuccessRange(baseAdvanceSituation, obrRating, armRating)
            'isCutOffOpportunity = isCutOffOpportunity And (batterObr = "A" Or batterObr = "B")
            'tableKey = baseAdvanceSituation

            'dsChart = FACDataAccess.ExecuteDataSet("Select * FROM BASEADVANCE WHERE OBR = '" & tableKey & "'")

            'If dsChart.Tables(0).Rows.Count = 0 Then
            '    MsgBox("ID not found!")
            'Else
            '    With dsChart.Tables(0).Rows(0)
            '        armRating = Game.PTeam.GetBatterPtr(Game.GetFielderNum).throwRating
            '        Select Case armRating
            '            Case "2"
            '                successRange = .Item("T2").ToString
            '            Case "3"
            '                successRange = .Item("T3").ToString
            '            Case "4"
            '                successRange = .Item("T4").ToString
            '            Case "5"
            '                successRange = .Item("T5").ToString
            '        End Select
            '        If Game.outs = 2 Then
            '            'Increase chances for advance with 2 outs
            '            facNumber = CType(successRange.Substring(successRange.Length - 2), Integer) + 20
            '            If facNumber > 88 Then facNumber = 88
            '            successRange = successRange.Substring(0, successRange.Length - 2) & facNumber.ToString
            '        End If
            '        'Automatic scenarios for faster runners
            '        structFAC = GetFAC("hitType", "Random")
            '        If baseAdvanceSituation.Substring(baseAdvanceSituation.Length - 1) = "3" Then
            '            Select Case obrRating
            '                Case "A"
            '                    If Val(structFAC.Random) <= 58 Then
            '                        successRange = "11-82" 'Nearly automatic
            '                    End If
            '                Case "B"
            '                    If Val(structFAC.Random) <= 44 Then
            '                        successRange = "11-82" 'Nearly automatic
            '                    End If
            '                Case "C"
            '                    If Val(structFAC.Random) <= 31 Then
            '                        successRange = "11-82" 'Nearly automatic
            '                    End If
            '            End Select
            '        ElseIf baseAdvanceSituation.Substring(baseAdvanceSituation.Length - 1) = "4" Then
            '            Select Case obrRating
            '                Case "A"
            '                    If Val(structFAC.Random) <= 55 Then
            '                        successRange = "11-82" 'Nearly automatic
            '                    End If
            '                Case "B"
            '                    If Val(structFAC.Random) <= 41 Then
            '                        successRange = "11-82" 'Nearly automatic
            '                    End If
            '                Case "C"
            '                    If Val(structFAC.Random) <= 28 Then
            '                        successRange = "11-82" 'Nearly automatic
            '                    End If
            '            End Select
            '        End If
            '    End With
            'End If

            Select Case baseAdvanceSituation.Substring(baseAdvanceSituation.Length - 1)
                Case "1", "2"
                    userMessage = "Advance first to third? It is " & obrRating & " against " & armRating & ". Range is " & successRange & "."
                Case "3"
                    If SecondBase.occupied Then
                        userMessage = "Advance second to home? It is " & obrRating & " against " & armRating & ". Range is " & successRange & "."
                    Else
                        userMessage = "Advance first to third? It is " & obrRating & " against " & armRating & ". Range is " & successRange & "."
                    End If
                Case "4"
                    userMessage = "Advance first to home? It is " & obrRating & " against " & armRating & ". Range is " & successRange & "."
            End Select
            lblQuestion.Text = userMessage
        Catch ex As Exception
            Call MsgBox("Error in frmAdvanceLoad. " & sqlQuery & " " & ex.ToString, MsgBoxStyle.OkOnly)
            Err.Clear()
        Finally
            If Not dsChart Is Nothing Then dsChart.Dispose()
        End Try
    End Sub

    Private Function GetSuccessRange(baseAdvanceSituation As String, obrRating As String, armRating As String) As String
        Dim tableKey As String
        Dim dsChart As DataSet = Nothing
        'Dim sqlQuery As String = ""
        Dim structFAC As FACCard
        Dim facNumber As Integer
        Dim FACDataAccess As New clsDataAccess("fac")
        Dim lclSuccessRange As String = ""

        Try
            tableKey = baseAdvanceSituation

            dsChart = FACDataAccess.ExecuteDataSet("Select * FROM BASEADVANCE WHERE OBR = '" & tableKey & "'")

            If dsChart.Tables(0).Rows.Count = 0 Then
                MsgBox("ID not found!")
                Return ""
            Else
                With dsChart.Tables(0).Rows(0)
                    Select Case armRating
                        Case "2"
                            lclSuccessRange = .Item("T2").ToString
                        Case "3"
                            lclSuccessRange = .Item("T3").ToString
                        Case "4"
                            lclSuccessRange = .Item("T4").ToString
                        Case "5"
                            lclSuccessRange = .Item("T5").ToString
                    End Select
                    If Game.outs = 2 Then
                        'Increase chances for advance with 2 outs
                        facNumber = CType(lclSuccessRange.Substring(lclSuccessRange.Length - 2), Integer) + 20
                        If facNumber > 88 Then facNumber = 88
                        lclSuccessRange = lclSuccessRange.Substring(0, lclSuccessRange.Length - 2) & facNumber.ToString
                    End If
                    'Automatic scenarios for faster runners
                    structFAC = GetFAC("hitType", "Random")
                    If baseAdvanceSituation.Substring(baseAdvanceSituation.Length - 1) = "3" Then
                        Select Case obrRating
                            Case "A"
                                If Val(structFAC.Random) <= 58 Then
                                    lclSuccessRange = "11-82" 'Nearly automatic
                                End If
                            Case "B"
                                If Val(structFAC.Random) <= 44 Then
                                    lclSuccessRange = "11-82" 'Nearly automatic
                                End If
                            Case "C"
                                If Val(structFAC.Random) <= 31 Then
                                    lclSuccessRange = "11-82" 'Nearly automatic
                                End If
                        End Select
                    ElseIf baseAdvanceSituation.Substring(baseAdvanceSituation.Length - 1) = "4" Then
                        Select Case obrRating
                            Case "A"
                                If Val(structFAC.Random) <= 55 Then
                                    lclSuccessRange = "11-82" 'Nearly automatic
                                End If
                            Case "B"
                                If Val(structFAC.Random) <= 41 Then
                                    lclSuccessRange = "11-82" 'Nearly automatic
                                End If
                            Case "C"
                                If Val(structFAC.Random) <= 28 Then
                                    lclSuccessRange = "11-82" 'Nearly automatic
                                End If
                        End Select
                    End If
                End With
            End If
            Return lclSuccessRange
        Catch ex As Exception
            Call MsgBox("Error in GetSuccessRange. p1 = " & baseAdvanceSituation & " p2= " & obrRating & " p3= " & armRating & " " & ex.ToString, MsgBoxStyle.OkOnly)
            Err.Clear()
            Return lclSuccessRange
        End Try
    End Function

    Private Function GetSuccessRangeAdv(baseAdvanceSituation As String, armRating As String, advHitType As Integer) As String
        Dim tableKey As String
        Dim dsChart As DataSet = Nothing
        Dim facNumber As Integer
        Dim FACDataAccess As New clsDataAccess("fac")
        Dim lclSuccessRange As String = ""
        Dim modification As Integer

        Try
            tableKey = baseAdvanceSituation

            dsChart = FACDataAccess.ExecuteDataSet("Select * FROM BASEADVANCE WHERE OBR = '" & tableKey & "'")

            If dsChart.Tables(0).Rows.Count = 0 Then
                MsgBox("ID not found!")
            Else
                With dsChart.Tables(0).Rows(0)
                    Select Case armRating
                        Case "2"
                            lclSuccessRange = .Item("T2").ToString
                        Case "3"
                            lclSuccessRange = .Item("T3").ToString
                        Case "4"
                            lclSuccessRange = .Item("T4").ToString
                        Case "5"
                            lclSuccessRange = .Item("T5").ToString
                    End Select
                End With
                If Game.outs = 2 Then
                    'Increase chances for advance with 2 outs
                    Select Case advHitType
                        Case HITTYPES.TEXAS_LEAGUER
                            modification = 60
                        Case HITTYPES.BLOOP
                            modification = 40
                        Case HITTYPES.NORMAL
                            modification = 20
                        Case HITTYPES.SMASH
                            modification = 0
                    End Select
                    facNumber = CType(lclSuccessRange.Substring(lclSuccessRange.Length - 2), Integer) + modification
                    If facNumber > 88 Then facNumber = 88
                    lclSuccessRange = lclSuccessRange.Substring(0, lclSuccessRange.Length - 2) & facNumber.ToString
                End If
            End If
            Return lclSuccessRange
        Catch ex As Exception
            Call MsgBox("Error in GetSuccessRangeAdv. " & ex.ToString, MsgBoxStyle.OkOnly)
            Err.Clear()
            Return lclSuccessRange
        End Try
    End Function

    ''' <summary>
    ''' This method implements the base advance procedures listed in the Avalon Hill Advanced Rules. First, you determine
    ''' the hit type, and use hit type in conjunction with OBR and outfielder T rating to determine the base runner's 
    ''' chances
    ''' </summary>
    ''' <remarks>BB Created: 12-17-2008</remarks>
    Private Sub UseAdvancedRules()
        'Dim baseHit As String
        Dim obrRating As String = ""
        Dim baseAdvanceType As String = ""
        Dim armRating As String = ""
        Dim hitTypeText(4) As String
        Dim FACDataAccess As New clsDataAccess("fac")

        Try
            hitTypeText(HITTYPES.TEXAS_LEAGUER) = "Texas Leaguer"
            hitTypeText(HITTYPES.BLOOP) = "Bloop"
            hitTypeText(HITTYPES.NORMAL) = "Normal"
            hitTypeText(HITTYPES.SMASH) = "Smash"

            hitType = Game.GetResultPtr.action.Substring(0, 2)
            advHitType = DetermineHitType()
            'determine advance situation
            Select Case BaseSituation()
                'Defined BaseAdvanceScenarios (BAS)
                '1 - First to third on single to left
                '2 - First to third on single to center
                '3 - First to third on single to right
                '4 - Second to home on single to any field
                '5 - First to home on double to any field
                Case "000"
                    Call MsgBox("Base advance dialog launched incorrectly.", MsgBoxStyle.Exclamation)
                Case "001"
                    obrRating = Game.BTeam.GetBatterPtr(FirstBase.runner).obr
                    If hitType = "1B" Then
                        Select Case Game.currentFieldPos
                            Case 7
                                baseAdvanceType = obrRating & "1" & (advHitType + 1).ToString
                            Case 8
                                baseAdvanceType = obrRating & "2" & (advHitType + 1).ToString
                            Case 9
                                baseAdvanceType = obrRating & "3" & (advHitType + 1).ToString
                        End Select
                        description = Game.BTeam.GetBatterPtr(FirstBase.runner).player
                        description = description.Substring(description.IndexOf(" ") + 1) & " is trying for third and he's "
                        advanceResult = "tagged OUT at third."
                        runnerID = 1
                    Else
                        'double
                        baseAdvanceType = obrRating & "5" & (advHitType + 1).ToString
                        description = Game.BTeam.GetBatterPtr(FirstBase.runner).player
                        description = description.Substring(description.IndexOf(" ") + 1) & " is trying for home and he's "
                        advanceResult = "OUT AT THE PLATE!"
                        runnerID = 1
                    End If
                Case "010"
                    obrRating = Game.BTeam.GetBatterPtr(SecondBase.runner).obr
                    baseAdvanceType = obrRating & "4" & (advHitType + 1).ToString
                    description = Game.BTeam.GetBatterPtr(SecondBase.runner).player
                    description = description.Substring(description.IndexOf(" ") + 1) & " is trying for home and he's "
                    advanceResult = "OUT AT THE PLATE!"
                    runnerID = 2
                Case "011"
                    If hitType = "1B" Then
                        obrRating = Game.BTeam.GetBatterPtr(SecondBase.runner).obr
                        baseAdvanceType = obrRating & "4" & (advHitType + 1).ToString
                        description = Game.BTeam.GetBatterPtr(SecondBase.runner).player
                        description = description.Substring(description.IndexOf(" ") + 1) & " is trying for home and he's "
                        advanceResult = "OUT AT THE PLATE!"
                        runnerID = 2
                    Else
                        'double
                        obrRating = Game.BTeam.GetBatterPtr(FirstBase.runner).obr
                        baseAdvanceType = obrRating & "5" & (advHitType + 1).ToString
                        description = Game.BTeam.GetBatterPtr(FirstBase.runner).player
                        description = description.Substring(description.IndexOf(" ") + 1) & " is trying for home and he's "
                        advanceResult = "OUT AT THE PLATE!"
                        runnerID = 1
                    End If
                Case "100"
                    Call MsgBox("Base advance dialog launched incorrectly.", MsgBoxStyle.Exclamation)
                Case "101"
                    obrRating = Game.BTeam.GetBatterPtr(FirstBase.runner).obr
                    If hitType = "1B" Then
                        Select Case Game.currentFieldPos
                            Case 7
                                baseAdvanceType = obrRating & "1" & (advHitType + 1).ToString
                            Case 8
                                baseAdvanceType = obrRating & "2" & (advHitType + 1).ToString
                            Case 9
                                baseAdvanceType = obrRating & "3" & (advHitType + 1).ToString
                        End Select
                        description = Game.BTeam.GetBatterPtr(FirstBase.runner).player
                        description = description.Substring(description.IndexOf(" ") + 1) & " is trying for third and he's "
                        advanceResult = "tagged OUT at third."
                        runnerID = 1
                    Else
                        'double
                        baseAdvanceType = obrRating & "5" & (advHitType + 1).ToString
                        description = Game.BTeam.GetBatterPtr(FirstBase.runner).player
                        description = description.Substring(description.IndexOf(" ") + 1) & " is trying for home and he's "
                        advanceResult = "OUT AT THE PLATE!"
                        runnerID = 1
                    End If
                Case "110"
                    obrRating = Game.BTeam.GetBatterPtr(SecondBase.runner).obr
                    baseAdvanceType = obrRating & "4" & (advHitType + 1).ToString
                    description = Game.BTeam.GetBatterPtr(SecondBase.runner).player
                    description = description.Substring(description.IndexOf(" ") + 1) & " is trying for home and he's "
                    advanceResult = "OUT AT THE PLATE!"
                    runnerID = 2
                Case "111"
                    If hitType = "1B" Then
                        obrRating = Game.BTeam.GetBatterPtr(SecondBase.runner).obr
                        baseAdvanceType = obrRating & "4" & (advHitType + 1).ToString
                        description = Game.BTeam.GetBatterPtr(SecondBase.runner).player
                        description = description.Substring(description.IndexOf(" ") + 1) & " is trying for home and he's "
                        advanceResult = "OUT AT THE PLATE!"
                        runnerID = 2
                    Else
                        'double
                        obrRating = Game.BTeam.GetBatterPtr(FirstBase.runner).obr
                        baseAdvanceType = obrRating & "5" & (advHitType + 1).ToString
                        description = Game.BTeam.GetBatterPtr(FirstBase.runner).player
                        description = description.Substring(description.IndexOf(" ") + 1) & " is trying for home and he's "
                        advanceResult = "OUT AT THE PLATE!"
                        runnerID = 1
                    End If
            End Select
            batterObr = Game.BTeam.GetBatterPtr(Game.currentBatter).obr
            'isCutOffOpportunity = isCutOffOpportunity And (batterObr = "A" Or batterObr = "B")
            If advHitType = HITTYPES.BLOOP Then
                'Bloop is treated like texas leaguer in lookup table
                baseAdvanceType = baseAdvanceType.Substring(0, 2) & (HITTYPES.TEXAS_LEAGUER + 1).ToString
            End If

            armRating = Game.PTeam.GetBatterPtr(Game.GetFielderNum).throwRating
            successRange = GetSuccessRangeAdv(baseAdvanceType, armRating, advHitType)
            'tableKey = baseAdvanceType

            'dschart = FACDataAccess.ExecuteDataSet("Select * FROM BASEADVANCE WHERE OBR = '" & tableKey & "'")

            'If dschart.Tables(0).Rows.Count = 0 Then
            '    MsgBox("ID not found!")
            'Else
            '    With dschart.Tables(0).Rows(0)
            '        armRating = Game.PTeam.GetBatterPtr(Game.GetFielderNum).throwRating
            '        Select Case armRating
            '            Case "2"
            '                successRange = .Item("T2").ToString
            '            Case "3"
            '                successRange = .Item("T3").ToString
            '            Case "4"
            '                successRange = .Item("T4").ToString
            '            Case "5"
            '                successRange = .Item("T5").ToString
            '        End Select
            '    End With
            '    If Game.outs = 2 Then
            '        'Increase chances for advance with 2 outs
            '        Select Case advHitType
            '            Case HITTYPES.TEXAS_LEAGUER
            '                modification = 60
            '            Case HITTYPES.BLOOP
            '                modification = 40
            '            Case HITTYPES.NORMAL
            '                modification = 20
            '            Case HITTYPES.SMASH
            '                modification = 0
            '        End Select
            '        facNumber = CType(successRange.Substring(successRange.Length - 2), Integer) + modification
            '        If facNumber > 88 Then facNumber = 88
            '        successRange = successRange.Substring(0, successRange.Length - 2) & facNumber.ToString
            '    End If
            'End If
            Select Case baseAdvanceType.Substring(1, 1)
                Case "1", "2", "3"
                    userMessage = "Advance first to third? It is " & obrRating & " against " & armRating & _
                            ". Hit Type is " & hitTypeText(advHitType) & ". Range is " & successRange & "."
                Case "4"
                    userMessage = "Advance second to home? It is " & obrRating & " against " & armRating & _
                            ". Hit Type is " & hitTypeText(advHitType) & ". Range is " & successRange & "."
                Case "5"
                    userMessage = "Advance first to home? It is " & obrRating & " against " & armRating & _
                            ". Hit Type is " & hitTypeText(advHitType) & ". Range is " & successRange & "."
            End Select
            lblQuestion.Text = userMessage
        Catch ex As Exception
            Call MsgBox("Error in UseAdvancedRules. " & ex.ToString, MsgBoxStyle.OkOnly)
            Err.Clear()
        End Try
    End Sub

    ''' <summary>
    ''' 1-Texas Leaguer, 2-Bloop, 3-Normal, 4-Smash
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>BB Created: 12/17/2008</remarks>
    Private Function DetermineHitType() As HITTYPES
        Dim FACCard As FACCard

        Try
            FACCard = GetFAC("hitType", "Random")
            If CInt(FACCard.Random) Mod 12 = 0 Then
                Return HITTYPES.TEXAS_LEAGUER
            ElseIf CInt(FACCard.Random) Mod 4 = 0 Then
                Return HITTYPES.BLOOP
            ElseIf CInt(FACCard.Random) Mod 2 = 1 Then
                Return HITTYPES.NORMAL
            Else
                Return HITTYPES.SMASH
            End If
        Catch ex As Exception

        End Try
    End Function
End Class