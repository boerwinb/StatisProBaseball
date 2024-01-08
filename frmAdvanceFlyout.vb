Public Class frmAdvanceFlyout

    Dim advanceResult As String
    Dim description As String
    Dim runnerID As Integer
    Dim successRange As String
    Dim userMessage As String
    Const RESULT_SAFE As String = "SAFE!"

    Private Sub cmdNo_Click(sender As Object, e As EventArgs) Handles cmdNo.Click
        Me.Close()
    End Sub


    Private Sub frmAdvanceFlyout_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim armRating As String = ""
        Dim baseAdvanceSituation As String = "" 'Base advance situation. There are 4 situations, one pertaining to each base advance table in the game.
        Dim tableKey As String
        Dim obrRating As String = ""
        'Chart Lookup
        Dim dsChart As DataSet = Nothing
        Dim sqlQuery As String = ""
        ' Dim structFAC As FACCard
        Dim facNumber As Integer
        Dim FACDataAccess As New clsDataAccess("fac")

        Try
            Select Case BaseSituation()
                Case "000"
                    Call MsgBox("Base advance dialog launched incorrectly.", MsgBoxStyle.Exclamation)
                    Me.Close()
                    Exit Sub
                Case "001"
                    runnerID = 1
                Case "010"
                    runnerID = 2
                Case "011"
                    If SecondBase.occupied And Game.GetResultPtr.resSecond = 2 Then
                        'runner on 2nd holds. This is the runner to offer advance.
                        runnerID = 2
                    Else
                        'Runner on 2nd advances already. This is for the runner on first.
                        runnerID = 1
                    End If
                Case "100"
                    runnerID = 3
                Case "101"
                    If ThirdBase.occupied And Game.GetResultPtr.resThird = 3 Then
                        'runner on 3rd holds. This is the runner to offer advance.
                        runnerID = 3
                    Else
                        'Runner on 3rd advances already. This is for the runner on first.
                        runnerID = 1
                    End If
                Case "110"
                    If ThirdBase.occupied And Game.GetResultPtr.resThird = 3 Then
                        'runner on 3rd holds. This is the runner to offer advance.
                        runnerID = 3
                    Else
                        'Runner on 3rd advances already. This is for the runner on second.
                        runnerID = 2
                    End If
                Case "111"
                    If ThirdBase.occupied And Game.GetResultPtr.resThird = 3 Then
                        'runner on 3rd holds. This is the runner to offer advance.
                        runnerID = 3
                    ElseIf SecondBase.occupied And Game.GetResultPtr.resSecond = 2 Then
                        'runner on 2nd holds. This is the runner to offer advance.
                        runnerID = 2
                    Else
                        'Runners on 2nd and 3rd advances already. This is for the runner on first.
                        runnerID = 1
                    End If
            End Select

            Select Case runnerID
                Case 1
                    obrRating = Game.BTeam.GetBatterPtr(FirstBase.runner).obr
                    description = Game.BTeam.GetBatterPtr(FirstBase.runner).player
                    description = description.Substring(description.IndexOf(" ") + 1) & " is trying for second and he's "
                    advanceResult = "tagged OUT at second."
                Case 2
                    obrRating = Game.BTeam.GetBatterPtr(SecondBase.runner).obr
                    description = Game.BTeam.GetBatterPtr(SecondBase.runner).player
                    description = description.Substring(description.IndexOf(" ") + 1) & " is trying for third and he's "
                    advanceResult = "tagged OUT at third."
                Case 3
                    obrRating = Game.BTeam.GetBatterPtr(ThirdBase.runner).obr
                    description = Game.BTeam.GetBatterPtr(ThirdBase.runner).player
                    description = description.Substring(description.IndexOf(" ") + 1) & " is trying for home and he's "
                    advanceResult = "OUT AT THE PLATE!"
            End Select

            If obrRating = "E" Then
                'no chance
                Me.Close()
                Exit Sub
            End If
            tableKey = obrRating
            sqlQuery = "Select * FROM BASEADVANCEFO WHERE OBR = '" & tableKey & "'"

            dsChart = FACDataAccess.ExecuteDataSet(sqlQuery)

            If dsChart.Tables(0).Rows.Count = 0 Then
                MsgBox("ID not found!")
            Else
                With dsChart.Tables(0).Rows(0)
                    armRating = Game.PTeam.GetBatterPtr(Game.GetFielderNum).throwRating
                    Select Case armRating
                        Case "2"
                            successRange = .Item("T2").ToString
                        Case "3"
                            successRange = .Item("T3").ToString
                        Case "4"
                            successRange = .Item("T4").ToString
                        Case "5"
                            successRange = .Item("T5").ToString
                    End Select
                    If Game.GetResultPtr.action.Substring(0, 2) = "FD" Then
                        'Increase chances on a deep fly
                        facNumber = CType(successRange.Substring(successRange.Length - 2), Integer) + 20
                        If facNumber > 88 Then facNumber = 88
                        successRange = successRange.Substring(0, successRange.Length - 2) & facNumber.ToString
                    End If
               End With
            End If

            Select Case runnerID
                Case 1
                    userMessage = "Will runner on 1st tag and advance? It is " & obrRating & " against " & armRating & ". Range is " & successRange & "."
                Case 2
                    userMessage = "Will runner on 2nd tag and advance? It is " & obrRating & " against " & armRating & ". Range is " & successRange & "."
                Case 3
                    userMessage = "Will runner on 3rd tag and score? It is " & obrRating & " against " & armRating & ". Range is " & successRange & "."
            End Select
            lblQuestion.Text = userMessage
        Catch ex As Exception
            Call MsgBox("Error in frmAdvanceFlyout_Load. " & sqlQuery & " " & ex.ToString, MsgBoxStyle.OkOnly)
            Err.Clear()
        Finally
            If Not dsChart Is Nothing Then dsChart.Dispose()
        End Try
    End Sub

    
    Private Sub cmdYes_Click(sender As Object, e As EventArgs) Handles cmdYes.Click
        Dim structFAC As FACCard

        Try
            structFAC = GetFAC("TagAdv", "Random")
            If InRange(successRange, CType(structFAC.Random, Integer)) Then
                'successful
                gblUpdateMsg = gblUpdateMsg & description & RESULT_SAFE & "."
                Select Case runnerID
                    Case 1
                        Game.GetResultPtr.resFirst = 2
                    Case 2
                        Game.GetResultPtr.resSecond = 3
                        If Game.GetResultPtr.resFirst = 1 Then
                            'trailing runners advance
                            Game.GetResultPtr.resFirst = 2
                        End If
                    Case 3
                        Game.GetResultPtr.resThird = 4
                        If Game.GetResultPtr.resSecond = 2 Then
                            'trailing runners advance
                            Game.GetResultPtr.resSecond = 3
                        End If
                        If Game.GetResultPtr.resFirst = 1 Then
                            'trailing runners advance
                            Game.GetResultPtr.resFirst = 2
                        End If
                End Select
            Else
                'Thrown out
                gblUpdateMsg = gblUpdateMsg & description & advanceResult & "."
                Select Case runnerID
                    Case 1
                        Game.GetResultPtr.resFirst = 0
                    Case 2
                        Game.GetResultPtr.resSecond = 0
                        If Game.GetResultPtr.resFirst = 1 Then
                            'trailing runners advance
                            Game.GetResultPtr.resFirst = 2
                        End If
                    Case 3
                        Game.GetResultPtr.resThird = 0
                        If Game.GetResultPtr.resSecond = 2 Then
                            'trailing runners advance
                            Game.GetResultPtr.resSecond = 3
                        End If
                        If Game.GetResultPtr.resFirst = 1 Then
                            'trailing runners advance
                            Game.GetResultPtr.resFirst = 2
                        End If
                End Select
            End If
            Me.Close()
        Catch ex As Exception
            Call MsgBox("Error in cmdYes_click. " & ex.ToString)
            Me.Close()
        End Try
    End Sub

    
End Class