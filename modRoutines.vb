Option Strict On
Option Explicit On 

Imports Microsoft.Win32
Imports System.Configuration.ConfigurationSettings

Module modRoutines
	
    Private Const SPB_REG_KEY As String = "Software\SPB"

    ''' <summary>
    ''' Returns index of matchValue within combo box
    ''' </summary>
    ''' <param name="cmbBox"></param>
    ''' <param name="matchValue"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SetListIndex(ByRef cmbBox As System.Windows.Forms.ComboBox, ByRef matchValue As String) As Integer
        Dim valueFound As Boolean
        Dim valueIsNumeric As Boolean
        Dim index As Integer
        Try

            valueIsNumeric = Val(matchValue) > 0
            index = 0
            valueFound = False
            While (index < cmbBox.Items.Count) And Not valueFound
                If valueIsNumeric Then
                    valueFound = cmbBox.Items.Item(index).ToString = matchValue
                Else
                    valueFound = (cmbBox.Items.Item(index).ToString.Trim = matchValue.Trim)
                End If
                If Not valueFound Then index += 1
            End While
            If Not valueFound Then
                index = 0
            End If
        Catch ex As Exception
            Call MsgBox("Error in SetListIndex. " & ex.ToString & " " & matchValue, MsgBoxStyle.OkOnly)
        End Try
        Return index
    End Function

    ''' <summary>
    ''' Need to turn single quotes into two single quotes.
    ''' so O'Malley becomes O''MALLEY and SQL inserts don't crash
    ''' </summary>
    ''' <param name="sqlStatement"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function HandleQuotes(ByRef sqlStatement As String) As String
        Dim tempValue As String = ""
        Dim positionIndex As Integer

        Try
            tempValue = sqlStatement
            positionIndex = 0
            While tempValue.IndexOf("'", positionIndex) > -1
                If tempValue.IndexOf("'", positionIndex) <> tempValue.IndexOf("''", positionIndex) Then
                    tempValue = tempValue.Substring(0, tempValue.IndexOf("'", positionIndex) + 1) & tempValue.Substring(tempValue.IndexOf("'", positionIndex))
                End If
                positionIndex = tempValue.IndexOf("''", positionIndex) + 2
            End While
            tempValue = tempValue.Trim
            If tempValue.Length > 1 Then
                If tempValue.Substring(tempValue.Length - 1) = "'" And tempValue.Substring(tempValue.Length - 2, 1) <> "'" Then
                    tempValue = tempValue.Substring(0, tempValue.Length - 1)
                End If
            End If
        Catch ex As Exception
            Call MsgBox("Error in HandleQuotes. " & ex.ToString & " " & sqlStatement, MsgBoxStyle.OkOnly)
        End Try
        Return tempValue
    End Function

    ''' <summary>
    ''' Verifies that there is a full lineup, a pitcher is selected, all positions are filled, and appropriate pitcher
    ''' (starter vs reliever) is on the mound.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ValidData() As Boolean
        Dim isValid As Boolean

        Try
            isValid = Home.GetPlayerNum(9) <> 0 And Visitor.GetPlayerNum(9) <> 0 And Home.PitcherSel > 0 And Visitor.PitcherSel > 0
            If isValid Then
                If Not bolStartGame Then
                    'Verify only before game starts
                    isValid = ValidPositions(Home) And ValidPositions(Visitor)
                End If
            End If
            If isValid Then
                If bolStartGame Then
                    If Game.HomeTeamBatting Then
                        isValid = CDbl(Visitor.GetPitcherPtr(Visitor.PitcherSel).RR) > 0
                    Else
                        isValid = CDbl(Home.GetPitcherPtr(Home.PitcherSel).RR) > 0
                    End If
                Else
                    isValid = CDbl(Home.GetPitcherPtr(Home.PitcherSel).SR) > 0
                    isValid = isValid And CDbl(Visitor.GetPitcherPtr(Visitor.PitcherSel).SR) > 0
                End If
            End If
        Catch ex As Exception
            Call MsgBox("Error in ValidData. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return isValid
    End Function

    ''' <summary>
    ''' Validates positions by verifying there are no duplicate positions in lineup
    ''' </summary>
    ''' <param name="team"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ValidPositions(ByRef team As clsTeam) As Boolean
        Dim tempValue As String
        Dim isValid As Boolean

        Try
            isValid = True
            For i As Integer = 1 To 9
                tempValue = team.GetBatterPtr(team.GetPlayerNum(i)).Position
                If Not team.GetBatterPtr(team.GetPlayerNum(i)).Available Then
                    'handle ejections and injuries
                    isValid = False
                    Call MsgBox("A player has been injured or ejected. Please replace them.")
                End If
                For j As Integer = (i + 1) To 9
                    If tempValue = team.GetBatterPtr(team.GetPlayerNum(j)).Position Then
                        isValid = False
                    End If
                Next j
            Next i
        Catch ex As Exception
            Call MsgBox("Error in ValidPositions. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return isValid
    End Function

    ''' <summary>
    ''' Determines whether a value is within a range shown on the player card. For example, a range may be listed as
    ''' 31-33, or simply 34. 
    ''' </summary>
    ''' <param name="valueRange"></param>
    ''' <param name="randomNumber"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function InRange(ByRef valueRange As String, ByRef randomNumber As Integer) As Boolean
        Dim index As Integer

        Try
            index = valueRange.IndexOf("-")
            If index > -1 Then
                Return randomNumber >= Val(valueRange.Substring(0, index)) And randomNumber <= Val(valueRange.Substring(index + 1))
            Else
                Return randomNumber = Val(valueRange)
            End If
        Catch ex As Exception
            Call MsgBox("Error in InRange. " & ex.ToString & " " & valueRange & " " & randomNumber.ToString, MsgBoxStyle.OkOnly)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' This method is used to determine whether we should refer to the other card to get the hit result. The purpose
    ''' is to disperse extra base hits correctly across all PB ratings. This prevents 2-9 pitchers from giving up too few
    ''' extra base hits and 2-6 pitcher from giving up too many.
    ''' </summary>
    ''' <param name="pbRating"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function UseOtherCard(ByVal pbRating As String) As Boolean
        Dim rndNum As Integer

        Randomize()
        rndNum = Convert.ToInt32(1000 * Rnd() + 1)
        Select Case pbRating
            Case "2-9"
                'switch to hitter card for 19.5% of hits given up on pitcher card
                Return rndNum <= 195
            Case "2-8"
                Return rndNum <= 194
            Case "2-7"
                Return rndNum <= 64
            Case "2-6"
                Return rndNum <= 83
            Case "2-5"
                Return rndNum <= 25
        End Select
    End Function

    ''' <summary>
    ''' This method is used to determine whether we should switch a K to an out.
    ''' The purpose is to reduce K's for 2-5 pitchers with no K range (almost half are this way).
    ''' There have been too many strikeouts for these pitchers
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SwitchFromK() As Boolean
        Dim rndNum As Integer

        Randomize()
        rndNum = Convert.ToInt32(1000 * Rnd() + 1)
        'switch to out 1/8 of the time
        Return rndNum <= 125
    End Function

    ''' <summary>
    ''' determines the number of incidents within a given range. For example, if Double range is 24 to 28,
    ''' this method will return 5.
    ''' </summary>
    ''' <param name="lowNumber"></param>
    ''' <param name="highNumber"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetRangeIncidents(ByVal lowNumber As Integer, ByVal highNumber As Integer) As Integer
        Dim incidents As Integer

        If lowNumber > highNumber Then
            Return 0
        End If
        For i As Integer = lowNumber To highNumber
            If i Mod 10 <> 9 And i Mod 10 <> 0 Then
                incidents += 1
            End If
        Next
        Return incidents
    End Function

    Public Function FetchRandomNumber(ByVal outcomes As Integer) As Integer
        Dim factor As Single
        Dim result As Integer

        Randomize()
        factor = Rnd()
        result = Convert.ToInt32(outcomes * factor + 1)
        If result > outcomes Then result = outcomes
        Return result
    End Function

    ''' <summary>
    ''' selects fac random number within a range. This is used when determining an outcome within a range of Singles.
    ''' </summary>
    ''' <param name="lowNumber"></param>
    ''' <param name="highNumber"></param>
    ''' <param name="incidentNumber"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SelectIncidentInRange(ByVal lowNumber As Integer, ByVal highNumber As Integer, ByVal incidentNumber As Integer, overrideType As String) As Integer
        Dim incidents As Integer
        Dim facViewRecord As clsFacView

        If lowNumber > highNumber Or incidentNumber = 0 Then
            Return 0
        End If

        For i As Integer = lowNumber To highNumber
            If i Mod 10 <> 9 And i Mod 10 <> 0 Then
                incidents += 1
                If incidents = incidentNumber Then
                    facViewRecord = New clsFacView(overrideType, "Number", i.ToString)
                    lstFacView.Add(facViewRecord)
                    Return i
                End If
            End If
        Next
        Return 0
    End Function

    ''' <summary>
    ''' GetFAC selects 1 of 384 random FACs (Fast Action Cards).
    ''' </summary>
    ''' <param name="action">intended action of FAC</param>
    ''' <param name="facField">field of interest</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetFAC(ByVal action As String, ByVal facField As String) As FACCard
        Dim dsFAC As DataSet = Nothing
        Dim sqlQuery As String = ""
        Dim cardID As Integer
        Dim tmpFACCard As FACCard = Nothing
        Dim filDebug As Integer
        Dim frmLaunch As frmDebugCard
        Dim facViewValue As String
        Dim facViewRecord As clsFacView
        Dim FACDataAccess As New clsDataAccess("fac")

        Try
            If Dir(gAppPath & "debug.bbb") <> Nothing Then
                filDebug = FreeFile()
                FileOpen(filDebug, gAppPath & "debug.out", OpenMode.Append)
                PrintLine(filDebug, "In GetFAC")
                FileClose(filDebug)
            End If

            If bolDebug Then
                frmLaunch = New frmDebugCard
                frmLaunch.ShowDialog()
                cardID = gblDebugCardID
            Else
                Randomize()
                cardID = Convert.ToInt32(384 * Rnd() + 1)
                If cardID > 384 Then
                    Randomize()
                    cardID = Convert.ToInt32(384 * Rnd() + 1)
                End If
            End If
            dsFAC = FACDataAccess.ExecuteDataSet("SELECT * FROM FAC WHERE cardid = " & cardID.ToString)
            If dsFAC.Tables(0).Rows.Count = 0 Then
                MsgBox("ID not found! CardID=" & cardID.ToString)
            Else
                With dsFAC.Tables(0).Rows(0)
                    System.Diagnostics.Debug.WriteLine(.Item("pb").ToString & Space(4) & .Item("random").ToString & Space(4) & .Item("cardid").ToString)
                    tmpFACCard.CD = .Item("CD").ToString
                    tmpFACCard.ErrorNum = .Item("Error").ToString
                    tmpFACCard.InfErr = .Item("InfErr").ToString
                    tmpFACCard.InfErr1B = .Item("InfErr1B").ToString
                    tmpFACCard.InfErrPC = .Item("InfErrPC").ToString
                    tmpFACCard.LN = .Item("LN").ToString
                    tmpFACCard.LP = .Item("LP").ToString
                    tmpFACCard.OFErr = .Item("OFErr").ToString
                    tmpFACCard.P = .Item("P").ToString
                    tmpFACCard.PB = .Item("PB").ToString
                    tmpFACCard.Random = .Item("Random").ToString
                    tmpFACCard.RN = .Item("RN").ToString
                    tmpFACCard.RP = .Item("RP").ToString
                    tmpFACCard.SN = .Item("SN").ToString
                    tmpFACCard.SP = .Item("SP").ToString
                    tmpFACCard.Pitch = Convert.ToBoolean(.Item("Pitch"))
                    If facField <> Nothing Then
                        facViewValue = .Item(facField).ToString
                        facViewRecord = New clsFacView(action, facField, facViewValue)
                        lstFacView.Add(facViewRecord)
                    End If

                End With
            End If
        Catch ex As Exception
            Call MsgBox("Error in GetFAC. " & ex.ToString, MsgBoxStyle.OkOnly)
        Finally
            'release objects
            If Not dsFAC Is Nothing Then dsFAC.Dispose()
        End Try
        Return tmpFACCard
    End Function

    '''' <summary>
    '''' returns full team name based on Lahman short name
    '''' </summary>
    '''' <param name="shortName"></param>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Function GetLongTeamTranslation(ByVal shortName As String) As String
    '    Dim dsTeam As DataSet = Nothing
    '    Dim sqlQuery As String = ""
    '    Dim longTeam As String = ""

    '    Try
    '        dsTeam = ExecuteDataSet("SELECT longteam FROM teamtranslation WHERE shortteam = '" & _
    '                                            shortName & "'", GetFACConnectString)

    '        If dsTeam.Tables(0).Rows.Count = 0 Then
    '            MsgBox("longname not found!")
    '        Else
    '            With dsTeam.Tables(0).Rows(0)
    '                longTeam = .Item("longteam").ToString
    '            End With
    '        End If

    '    Catch ex As Exception
    '        Call MsgBox("Error in GetLongTeamTranslation. " & ex.ToString, MsgBoxStyle.OkOnly)
    '    Finally
    '        If Not dsTeam Is Nothing Then dsTeam.Dispose()
    '    End Try
    '    Return longTeam
    'End Function

    ''' <summary>
    ''' returns Lahman short name based on full team name
    ''' </summary>
    ''' <param name="longName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetShortTeamTranslation(ByVal longName As String) As String
        Dim dsTeam As DataSet = Nothing
        Dim sqlQuery As String = ""
        Dim shortTeam As String = ""
        Dim FACDataAccess As New clsDataAccess("fac")

        Try
            dsTeam = FACDataAccess.ExecuteDataSet("SELECT shortteam FROM teamtranslation WHERE longteam = '" & _
                                    longName & "' AND year1 <= " & gstrSeason & " AND year2 >= " & _
                                    gstrSeason)

            If dsTeam.Tables(0).Rows.Count = 0 Then
                MsgBox("shortname not found!")
            Else
                With dsTeam.Tables(0).Rows(0)
                    shortTeam = .Item("shortteam").ToString
                End With
            End If

        Catch ex As Exception
            Call MsgBox("Error in GetShortTeamTranslation. " & ex.ToString, MsgBoxStyle.OkOnly)
        Finally
            If Not dsTeam Is Nothing Then dsTeam.Dispose()
        End Try
        Return shortTeam
    End Function


    ''' <summary>
    ''' BB - overrides the traditional IIF method so that any data type can be returned
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="Expression"></param>
    ''' <param name="TruePart"></param>
    ''' <param name="FalsePart"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IIF(Of T)(ByVal Expression As Boolean, ByVal TruePart As T, ByVal FalsePart As T) As T
        If Expression Then Return TruePart Else Return FalsePart
    End Function

    ''' <summary>
    ''' Determines the current basesituation
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function BaseSituation() As String
        Dim BS As String = ""

        Try
            BS = IIF(ThirdBase.occupied, "1", "0")
            BS += IIF(SecondBase.occupied, "1", "0")
            BS += IIF(FirstBase.occupied, "1", "0")
        Catch ex As Exception
            Call MsgBox("Error in BaseSituation. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return BS
    End Function

    ''' <summary>
    ''' Parses the last name from a full name
    ''' </summary>
    ''' <param name="fullName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetLastName(ByVal fullName As String) As String
        Dim index As Integer
        Dim lastName As String = ""

        Try
            index = fullName.IndexOf(" ")
            lastName = fullName.Substring(index + 1)
        Catch ex As Exception
            Call MsgBox("Error in GetLastName. " & ex.ToString & " " & fullName, MsgBoxStyle.OkOnly)
        End Try
        Return lastName
    End Function

    ''' <summary>
    ''' Gets numeric portion of string expression
    ''' </summary>
    ''' <param name="fullExpression"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetNum(ByVal fullExpression As String) As String
        Dim numberPortion As String = ""

        Try
            For i As Integer = 0 To fullExpression.Length - 1
                If IsNumeric(fullExpression.Substring(i, 1)) Then
                    numberPortion += fullExpression.Substring(i, 1)
                End If
            Next i
        Catch ex As Exception
            Call MsgBox("Error in GetNum. " & ex.ToString & " " & fullExpression, MsgBoxStyle.OkOnly)
        End Try
        Return numberPortion
    End Function

    ''' <summary>
    ''' Removes all instances of a character from a string expression
    ''' </summary>
    ''' <param name="expression"></param>
    ''' <param name="removeChar"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function StripChar(ByRef expression As String, ByRef removeChar As String) As String
        Dim tempValue As String = ""
        Try
            For i As Integer = 0 To expression.Length - 1
                If expression.Substring(i, 1) <> removeChar Then
                    tempValue = tempValue & expression.Substring(i, 1)
                End If
            Next i
        Catch ex As Exception
            Call MsgBox("Error in StripChar. " & ex.ToString & " " & expression, MsgBoxStyle.OkOnly)
        End Try
        Return tempValue
    End Function

    ''' <summary>
    ''' Gets throw rating for catchers and outfielders. Otherwise, error ratings.
    ''' </summary>
    ''' <param name="fieldLine"></param>
    ''' <param name="fieldPosition"></param>
    ''' <param name="category"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetRating(ByRef fieldLine As String, ByVal fieldPosition As String, ByRef category As String) As String
        Dim index As Integer

        Try
            If fieldPosition = "P" Then Return ""
            If fieldPosition.Substring(fieldPosition.Length - 1) = "F" Then fieldPosition = "OF"
            index = fieldLine.IndexOf(fieldPosition)
            If index = -1 Then
                'For Emergency players. Return defaults.
                If category = "E" Then
                    Return "10"
                Else
                    '"T" category
                    If fieldPosition = "C" Then
                        Return "C"
                    ElseIf fieldPosition = "OF" Then
                        Return "2"
                    Else
                        Return ""
                    End If
                End If
            End If
            index = fieldLine.IndexOf(category, index)
            If index = -1 Then
                Return ""
            Else
                Return fieldLine.Substring(index + 1, 1)
            End If
        Catch ex As Exception
            Call MsgBox("Error in GetRating. " & ex.ToString & " " & fieldLine, MsgBoxStyle.OkOnly)
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' This function returns the string portion before the specified character
    ''' </summary>
    ''' <param name="expression"></param>
    ''' <param name="cutoffChar"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetStrBefore(ByRef expression As String, ByRef cutoffChar As String) As String
        Dim index As Integer

        Try
            index = expression.IndexOf(cutoffChar)
            If index > -1 Then
                Return expression.Substring(0, expression.Length - (expression.Length - index))
            End If
        Catch ex As Exception
            Call MsgBox("Error in GetStrBefore. " & ex.ToString & " " & expression, MsgBoxStyle.OkOnly)
        End Try
        Return ""
    End Function

    ''' <summary>
    ''' This function counts the number of times a given char appears in a string
    ''' </summary>
    ''' <param name="expression"></param>
    ''' <param name="character"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetIncidents(ByRef expression As String, ByRef character As String) As Integer
        Dim incidents As Integer

        Try
            For i As Integer = 0 To expression.Length - 1
                If expression.Substring(i, 1) = character Then
                    incidents += 1
                End If
            Next i
        Catch ex As Exception
            Call MsgBox("Error in GetIncidents. " & ex.ToString & " " & expression, MsgBoxStyle.OkOnly)
        End Try
        Return incidents
    End Function

    ''' <summary>
    ''' Grabs current game in season from the registry
    ''' </summary>
    ''' <param name="Season"></param>
    ''' <param name="gamenumber"></param>
    ''' <param name="rownumber">row number in the .scd file</param>
    ''' <remarks></remarks>
    Public Sub GetGame(ByVal season As String, ByRef gameNumber As Integer, ByRef rowNumber As String)
        Dim regKey As String = REGISTRY_LOCAL_MACHINE & "\" & SPB_REG_KEY & "\" & season

        Try
            regKey = REGISTRY_LOCAL_MACHINE & "\" & SPB_REG_KEY & "\" & season
            gameNumber = Integer.Parse(Registry.GetValue(regKey, "Game", 1).ToString)
            rowNumber = Registry.GetValue(regKey, "SchedRow", 1).ToString
        Catch ex As Exception
            Call MsgBox("Error in GetGame. " & ex.ToString & " " & season & " " & gameNumber.ToString & _
                        " " & rowNumber.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' Updates the current game in season in the registry
    ''' </summary>
    ''' <param name="season"></param>
    ''' <param name="gameNumber"></param>
    ''' <param name="rowNumber">row number in the .scd file</param>
    ''' <remarks></remarks>
    Public Sub SetGame(ByVal season As String, ByVal gameNumber As String, ByVal rowNumber As String)
        Dim regKey As String = REGISTRY_LOCAL_MACHINE & "\" & SPB_REG_KEY & "\" & season

        Try
            Registry.SetValue(regKey, "SchedRow", rowNumber)
            Registry.SetValue(regKey, "Game", gameNumber)
        Catch ex As Exception
            Call MsgBox("Error in SetGame. " & ex.ToString & " " & season & " " & gameNumber.ToString & _
                        " " & rowNumber.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' Sets height and width for logo sizes. It would be better to have standard sized logos.
    ''' </summary>
    ''' <param name="teamName"></param>
    ''' <param name="height"></param>
    ''' <param name="width"></param>
    ''' <remarks></remarks>
    Public Sub GetLogoDimensions(ByRef teamName As String, ByRef height As Integer, ByRef width As Integer)
        height = 1100
        width = 1100
        Try
            'Select Case teamName.ToUpper
            '    Case "ATLANTA"
            '        height = 1800
            '        width = 1725
            '    Case "BALTIMORE"
            '        height = 1800
            '        width = 1845
            '    Case "BOSTON"
            '        height = 1500
            '        width = 1500
            '    Case "CALIFORNIA"
            '        height = 1800
            '        width = 1800
            '    Case "CHICAGO (A)"
            '        height = 1800
            '        width = 1395
            '    Case "CHICAGO (N)"
            '        height = 1800
            '        width = 1800
            '    Case "CINCINNATI"
            '        height = 1800
            '        width = 2055
            '    Case "CLEVELAND"
            '        height = 1500
            '        width = 1500
            '    Case "DETROIT"
            '        height = 1800
            '        width = 1800
            '    Case "HOUSTON"
            '        height = 1800
            '        width = 1815
            '    Case "KANSAS CITY"
            '        height = 1500
            '        width = 1500
            '    Case "LOS ANGELES"
            '        height = 1500
            '        width = 1500
            '    Case "MILWAUKEE"
            '        height = 1800
            '        width = 1515
            '    Case "MINNESOTA"
            '        height = 2295
            '        width = 1965
            '    Case "MONTREAL"
            '        height = 1800
            '        width = 2055
            '    Case "NEW YORK (A)"
            '        height = 1500
            '        width = 1500
            '    Case "NEW YORK (N)"
            '        height = 1500
            '        width = 1500
            '    Case "OAKLAND"
            '        height = 1500
            '        width = 1500
            '    Case "PHILADELPHIA"
            '        height = 1800
            '        width = 1440
            '    Case "PITTSBURGH"
            '        height = 2115
            '        width = 1575
            '    Case "SAN DIEGO"
            '        height = 1800
            '        width = 1935
            '    Case "SAN FRANCISCO"
            '        height = 795
            '        width = 1500
            '    Case "SEATTLE"
            '        height = 1800
            '        width = 1695
            '    Case "ST LOUIS"
            '        height = 1800
            '        width = 1800
            '    Case "TEXAS"
            '        height = 1800
            '        width = 1815
            '    Case "TORONTO"
            '        height = 1800
            '        width = 1800
            '    Case Else
            '        height = 1800
            '        width = 1725
            'End Select
        Catch ex As Exception
            Call MsgBox("Error in GetLogoDimensions. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' copies the contents of one Standing array to another
    ''' </summary>
    ''' <param name="arr1"></param>
    ''' <param name="arr2"></param>
    ''' <param name="arraySize"></param>
    ''' <remarks></remarks>
    Public Sub CopyStandingArray(ByRef arr1() As Standing, ByRef arr2() As Standing, ByRef arraySize As Integer)
        Try
            For i As Integer = 1 To arraySize
                arr1(i) = arr2(i)
            Next i
        Catch ex As Exception
            Call MsgBox("Error in CopyStandingArray. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    '''' <summary>
    '''' Builds Access connection string for season tables
    '''' </summary>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Function GetConnectString() As String
    '    Return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gAppPath & _
    '                        "\data\" & gstrSeason & "data.mdb;User Id=admin;Password=;"
    'End Function

    '''' <summary>
    '''' Builds Access connection string for FAC tables
    '''' </summary>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Function GetFACConnectString() As String
    '    Return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gAppPath & _
    '                        "\data\fac.mdb;User Id=admin;Password=;"
    'End Function

    '''' <summary>
    '''' Builds Access connection string for Lahman tables
    '''' </summary>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Function GetLahmanConnectString() As String
    '    Return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gAppPath & _
    '                        "\data\lahman56.mdb;User Id=admin;Password=;"
    'End Function

    ''' <summary>
    ''' returns the average value of the contents of a collection
    ''' </summary>
    ''' <param name="colValues"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetAverageCollection(ByVal colValues As Collection) As Double
        Dim sumValues As Integer
        If colValues.Count = 0 Then Return 0
        For Each oValue As Object In colValues
            sumValues += CInt(oValue)
        Next
        Return sumValues / colValues.Count
    End Function

    Public Function GetAppLocation() As String
        Dim appPath As String
        Dim stringPosition As Integer

        appPath = Reflection.Assembly.GetExecutingAssembly.Location
        stringPosition = InStrRev(appPath, "\")
        If stringPosition > 0 Then
            appPath = appPath.Substring(0, stringPosition - 1)
        End If
        Return appPath
    End Function

    'Public Function ExecuteMultipleScalars(ByVal colQuery As Collection, ByVal connectString As String) As Object
    '    For Each query As String In colQuery
    '        ExecuteScalar()
    '    Next
    'End Function

    'Public Function ExecuteDataSetMany(ByVal sqlQuery As String, ByRef accessConn As OleDb.OleDbConnection) As DataSet
    '    Try
    '        Using accessCmd As OleDb.OleDbCommand = accessConn.CreateCommand
    '            Using accessAdapter As New OleDb.OleDbDataAdapter(accessCmd)
    '                Using ds As New DataSet
    '                    accessCmd.CommandText = sqlQuery
    '                    accessAdapter.Fill(ds)
    '                    Return ds
    '                End Using
    '            End Using
    '        End Using
    '    Catch ex As Exception
    '        Call MsgBox("Error in ExecuteDataSetMany. " & ex.ToString, MsgBoxStyle.OkOnly)
    '    End Try
    '    Return Nothing
    'End Function

    'Public Function ExecuteDataSet(ByVal sqlQuery As String, ByVal connectString As String) As DataSet
    '    Try
    '        Using accessConn As OleDb.OleDbConnection = New OleDb.OleDbConnection
    '            accessConn.ConnectionString = connectString
    '            accessConn.Open()
    '            Using accessCmd As OleDb.OleDbCommand = accessConn.CreateCommand
    '                Using accessAdapter As New OleDb.OleDbDataAdapter(accessCmd)
    '                    Using ds As New DataSet
    '                        accessCmd.CommandText = sqlQuery
    '                        accessAdapter.Fill(ds)
    '                        Return ds
    '                    End Using
    '                End Using
    '            End Using
    '        End Using

    '    Catch ex As Exception
    '        Call MsgBox("Error in ExecuteDataSet. " & ex.ToString, MsgBoxStyle.OkOnly)
    '    End Try
    '    Return Nothing
    'End Function

    'Public Function ExecuteScalarMany(ByVal sqlStatement As String, ByRef accessConn As OleDb.OleDbConnection) As Object
    '    Try
    '        Using accessCmd As OleDb.OleDbCommand = accessConn.CreateCommand()
    '            accessCmd.CommandText = sqlStatement
    '            Return accessCmd.ExecuteScalar()
    '        End Using
    '    Catch ex As Exception
    '        Call MsgBox("Error in ExecuteScalar. " & ex.ToString & " SQL: " & sqlStatement, MsgBoxStyle.OkOnly)
    '    End Try
    '    Return Nothing
    'End Function
    'Public Function ExecuteScalar(ByVal sqlStatement As String, ByVal connectString As String) As Object
    '    Try
    '        Using accessConn As OleDb.OleDbConnection = New OleDb.OleDbConnection
    '            accessConn.ConnectionString = connectString
    '            accessConn.Open()
    '            Using accessCmd As OleDb.OleDbCommand = accessConn.CreateCommand()
    '                accessCmd.CommandText = sqlStatement
    '                Return accessCmd.ExecuteScalar()
    '            End Using
    '        End Using
    '    Catch ex As Exception
    '        Call MsgBox("Error in ExecuteScalar. " & ex.ToString, MsgBoxStyle.OkOnly)
    '    End Try
    '    Return Nothing
    'End Function
    ''' <summary>
    ''' Increases a given range by x numbers
    ''' </summary>
    ''' <param name="currentRange"></param>
    ''' <param name="increment"></param>
    ''' <returns></returns>
    ''' <remarks>BB: Created 12/26/2008</remarks>
    Public Function IncreaseRange(ByVal currentRange As String, ByVal increment As Integer, _
                        ByVal previousNumber As Integer) As String
        Dim firstNumber As Integer
        Dim currentNumber As Integer
        Dim lastNumber As Integer

        If currentRange = Nothing Then
            'there are no numbers 
            If previousNumber = 0 Then
                currentNumber = 8
            Else
                currentNumber = previousNumber
            End If
            For i As Integer = 1 To increment
                If currentNumber < 88 Then
                    If currentNumber Mod 10 = 8 Then
                        currentNumber += 3
                    Else
                        currentNumber += 1
                    End If
                    If i = 1 Then
                        firstNumber = currentNumber
                    End If
                End If
            Next i
            lastNumber = currentNumber
        Else
            'there is a current range
            currentNumber = CInt(Right(currentRange, 2))
            firstNumber = CInt(Left(currentRange, 2))
            For i As Integer = 1 To increment
                If currentNumber < 88 Then
                    If currentNumber Mod 10 = 8 Then
                        currentNumber += 3
                    Else
                        currentNumber += 1
                    End If
                End If
            Next i
            lastNumber = currentNumber
        End If
        If firstNumber = 0 Then
            Return ""
        ElseIf firstNumber = lastNumber Then
            Return firstNumber.ToString
        Else
            Return firstNumber.ToString & "-" & lastNumber.ToString
        End If
    End Function

    Public Function CalcAge(ByVal StartDate As Date, ByVal EndDate As Date) As Integer
        ' Returns the number of years between the passed dates
        If EndDate.Month < StartDate.Month OrElse (EndDate.Month = StartDate.Month AndAlso EndDate.Day < StartDate.Day) Then
            Return EndDate.Year - StartDate.Year - 1
        Else
            Return EndDate.Year - StartDate.Year
        End If
    End Function

    Public Function CheckField(ByVal value As Object, ByVal defaultValue As Integer) As Integer
        If IsDBNull(value) Then
            Return defaultValue
        Else
            Return CInt(value)
        End If
    End Function

    Public Function DetermineCardNumbers(ByVal cardLabel As String, ByRef startRange As Integer, ByRef endRange As Integer) As Integer
        Dim cardNumbers As Integer
        Dim startPos As Integer
        Dim endPos As Integer

        If cardLabel = "" Then
            Return 0
        ElseIf cardLabel.Contains("-") Then
            startPos = CInt(cardLabel.Substring(0, 2))
            endPos = CInt(cardLabel.Substring(3, 2))
            For x As Integer = startPos To endPos
                If (x Mod 10) >= 1 And (x Mod 10) <= 8 Then
                    'this excludes card numbers ending in 9 or 0
                    cardNumbers += 1
                End If
            Next
            startRange = startPos
            endRange = endPos
            Return cardNumbers
        Else
            'should be a single number
            startRange = CInt(cardLabel)
            endRange = startRange
            Return 1
        End If
    End Function
End Module