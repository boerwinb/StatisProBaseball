Option Strict On
Option Explicit On

Imports System.Configuration.configurationManager
Imports Microsoft.win32

Friend Class frmAdvance
	Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
	Public Sub New()
		MyBase.New()
		If m_vb6FormDefInstance Is Nothing Then
			If m_InitializingDefInstance Then
				m_vb6FormDefInstance = Me
			Else
				Try 
					'For the start-up form, the first instance created is the default instance.
					If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
						m_vb6FormDefInstance = Me
					End If
				Catch
				End Try
			End If
		End If
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
    Public WithEvents cmdNo As System.Windows.Forms.Button
    Public WithEvents cmdYes As System.Windows.Forms.Button
	Public WithEvents lblQuestion As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cmdNo = New System.Windows.Forms.Button()
        Me.cmdYes = New System.Windows.Forms.Button()
        Me.lblQuestion = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'cmdNo
        '
        Me.cmdNo.BackColor = System.Drawing.SystemColors.Control
        Me.cmdNo.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdNo.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdNo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdNo.Location = New System.Drawing.Point(168, 112)
        Me.cmdNo.Name = "cmdNo"
        Me.cmdNo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdNo.Size = New System.Drawing.Size(65, 25)
        Me.cmdNo.TabIndex = 2
        Me.cmdNo.Text = "No"
        Me.cmdNo.UseVisualStyleBackColor = False
        '
        'cmdYes
        '
        Me.cmdYes.BackColor = System.Drawing.SystemColors.Control
        Me.cmdYes.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdYes.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdYes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdYes.Location = New System.Drawing.Point(64, 112)
        Me.cmdYes.Name = "cmdYes"
        Me.cmdYes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdYes.Size = New System.Drawing.Size(65, 25)
        Me.cmdYes.TabIndex = 1
        Me.cmdYes.Text = "Yes"
        Me.cmdYes.UseVisualStyleBackColor = False
        '
        'lblQuestion
        '
        Me.lblQuestion.BackColor = System.Drawing.SystemColors.Control
        Me.lblQuestion.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblQuestion.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQuestion.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblQuestion.Location = New System.Drawing.Point(48, 32)
        Me.lblQuestion.Name = "lblQuestion"
        Me.lblQuestion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblQuestion.Size = New System.Drawing.Size(217, 33)
        Me.lblQuestion.TabIndex = 0
        '
        'frmAdvance
        '
        Me.AcceptButton = Me.cmdYes
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(313, 153)
        Me.Controls.Add(Me.cmdNo)
        Me.Controls.Add(Me.cmdYes)
        Me.Controls.Add(Me.lblQuestion)
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(4, 25)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmAdvance"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Advance Extra Base"
        Me.ResumeLayout(False)

    End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As frmAdvance
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmAdvance
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmAdvance()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	
    Dim isCutOffOpportunity As Boolean
    Dim hitType As String
    Dim description As String
    Dim userMessage As String
    Dim advanceResult As String
    Dim successRange As String
    Dim batterObr As String
    Dim runnerID As Integer
    Const RESULT_SAFE As String = "SAFE!"

    Private Sub cmdNo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdNo.Click
        Me.Close()
    End Sub

    Private Sub cmdYes_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdYes.Click
        Dim cutOffThrow As Boolean
        Dim structFAC As FACCard

        Try
            If isCutOffOpportunity Then
                If batterObr = "A" Or batterObr = "B" Then
                    cutOffThrow = MsgBox("Will the defense cut off throw?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes
                End If
            End If
            If cutOffThrow Then
                gblUpdateMsg = gblUpdateMsg & userMessage & RESULT_SAFE & " The throw was cutoff."
                If runnerID = 1 Then
                    Game.GetResultPtr.resFirst = 3
                Else
                    Game.GetResultPtr.resSecond = 4
                End If
            Else
                structFAC = GetFAC("baseAdv", "Random")
                If InRange(successRange, CType(structFAC.Random, Integer)) Then 'successful
                    gblUpdateMsg = gblUpdateMsg & description & RESULT_SAFE & "."
                    If isCutOffOpportunity Then 'Throw was not cutoff, so batter gets second base
                        If batterObr = "A" Or batterObr = "B" Then
                            gblUpdateMsg = gblUpdateMsg & " Batter to second on throw."
                            Game.GetResultPtr.resBatter = 2
                        End If
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
                    If isCutOffOpportunity And Game.outs < 2 Then 'Throw was not cutoff, so batter gets second base
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
            End If
            Me.Close()
        Catch ex As Exception
            Call MsgBox("Error in cmdYes_click. " & ex.ToString)
            Me.Close()
        End Try
    End Sub

    Private Sub frmAdvance_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dim armRating As String = ""
        Dim baseAdvanceSituation As String = "" 'Base advance situation. There are 4 situations, one pertaining to each base advance table in the game.
        Dim tableKey As String
        Dim obrRating As String = ""
        'Chart Lookup
        Dim dsChart As DataSet = Nothing
        Dim sqlQuery As String = ""
        Dim structFAC As FACCard
        Dim facNumber As Integer
        Dim FACDataAccess As New clsDataAccess("fac")

        Try
            If AppSettings("BaseAdvance_AdvanceRules").ToString.ToUpper = "ON" Then
                UseAdvancedRules()
                Exit Sub
            End If

            hitType = Game.GetResultPtr.action.Substring(0, 2)
            'Determine AdvanceSituation
            isCutOffOpportunity = False
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
                        batterObr = Game.BTeam.GetBatterPtr(Game.currentBatter).obr
                        isCutOffOpportunity = True
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
                    batterObr = Game.BTeam.GetBatterPtr(Game.currentBatter).obr
                    isCutOffOpportunity = True
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
                        batterObr = Game.BTeam.GetBatterPtr(Game.currentBatter).obr
                        isCutOffOpportunity = True
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
                    batterObr = Game.BTeam.GetBatterPtr(Game.currentBatter).obr
                    isCutOffOpportunity = True
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
            isCutOffOpportunity = isCutOffOpportunity And (batterObr = "A" Or batterObr = "B")
            tableKey = baseAdvanceSituation

            dsChart = FACDataAccess.ExecuteDataSet("Select * FROM BASEADVANCE WHERE OBR = '" & tableKey & "'")

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
                    If Game.outs = 2 Then
                        'Increase chances for advance with 2 outs
                        facNumber = CType(successRange.Substring(successRange.Length - 2), Integer) + 20
                        If facNumber > 88 Then facNumber = 88
                        successRange = successRange.Substring(0, successRange.Length - 2) & facNumber.ToString
                    End If
                    'Automatic scenarios for faster runners
                    structFAC = GetFAC("hitType", "Random")
                    If baseAdvanceSituation.Substring(baseAdvanceSituation.Length - 1) = "3" Then
                        Select Case obrRating
                            Case "A"
                                If Val(structFAC.Random) <= 58 Then
                                    successRange = "11-82" 'Nearly automatic
                                End If
                            Case "B"
                                If Val(structFAC.Random) <= 44 Then
                                    successRange = "11-82" 'Nearly automatic
                                End If
                            Case "C"
                                If Val(structFAC.Random) <= 31 Then
                                    successRange = "11-82" 'Nearly automatic
                                End If
                        End Select
                    ElseIf baseAdvanceSituation.Substring(baseAdvanceSituation.Length - 1) = "4" Then
                        Select Case obrRating
                            Case "A"
                                If Val(structFAC.Random) <= 55 Then
                                    successRange = "11-82" 'Nearly automatic
                                End If
                            Case "B"
                                If Val(structFAC.Random) <= 41 Then
                                    successRange = "11-82" 'Nearly automatic
                                End If
                            Case "C"
                                If Val(structFAC.Random) <= 28 Then
                                    successRange = "11-82" 'Nearly automatic
                                End If
                        End Select
                    End If
                End With
            End If

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

    ''' <summary>
    ''' This method implements the base advance procedures listed in the Avalon Hill Advanced Rules. First, you determine
    ''' the hit type, and use hit type in conjunction with OBR and outfielder T rating to determine the base runner's 
    ''' chances
    ''' </summary>
    ''' <remarks>BB Created: 12-17-2008</remarks>
    Private Sub UseAdvancedRules()
        'Dim baseHit As String
        Dim advHitType As Integer
        Dim obrRating As String = ""
        Dim baseAdvanceType As String = ""
        Dim tableKey As String
        Dim dschart As System.Data.DataSet
        Dim armRating As String = ""
        Dim modification As Integer
        Dim facNumber As Integer
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
            isCutOffOpportunity = False
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
                        isCutOffOpportunity = True
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
                    isCutOffOpportunity = True
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
                        isCutOffOpportunity = True
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
                    isCutOffOpportunity = True
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
            isCutOffOpportunity = isCutOffOpportunity And (batterObr = "A" Or batterObr = "B")
            If advHitType = HITTYPES.BLOOP Then
                'Bloop is treated like texas leaguer in lookup table
                baseAdvanceType = baseAdvanceType.Substring(0, 2) & (HITTYPES.TEXAS_LEAGUER + 1).ToString
            End If
            tableKey = baseAdvanceType

            dschart = FACDataAccess.ExecuteDataSet("Select * FROM BASEADVANCE WHERE OBR = '" & tableKey & "'")

            If dschart.Tables(0).Rows.Count = 0 Then
                MsgBox("ID not found!")
            Else
                With dschart.Tables(0).Rows(0)
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
                    facNumber = CType(successRange.Substring(successRange.Length - 2), Integer) + modification
                    If facNumber > 88 Then facNumber = 88
                    successRange = successRange.Substring(0, successRange.Length - 2) & facNumber.ToString
                End If
            End If
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