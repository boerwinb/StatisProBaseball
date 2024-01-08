Option Strict On
Option Explicit On

Imports System.Text


Friend Class frmStats
    Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        ' Create an instance of a ListView column sorter and assign it 
        ' to the ListView control.
        lvwColumnSorter = New ListViewColumnSorter
        Me.lvwStats.ListViewItemSorter = lvwColumnSorter

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
    Public WithEvents cmdNL As System.Windows.Forms.Button
    Public WithEvents frmAL As System.Windows.Forms.GroupBox
    Public WithEvents cmdAL As System.Windows.Forms.Button
    Public WithEvents lstTeams As System.Windows.Forms.ListBox
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Public WithEvents optStandings As System.Windows.Forms.RadioButton
    Public WithEvents optPitchingLeadersTeam As System.Windows.Forms.RadioButton
    Public WithEvents optHittingLeadersTeam As System.Windows.Forms.RadioButton
    Public WithEvents optPitchingLeaders As System.Windows.Forms.RadioButton
    Public WithEvents optHittingLeaders As System.Windows.Forms.RadioButton
    Public WithEvents optTeamPitching As System.Windows.Forms.RadioButton
    Public WithEvents optTeamHitting As System.Windows.Forms.RadioButton
    Friend WithEvents grpRegPostSeason As System.Windows.Forms.GroupBox
    Friend WithEvents rbRegSeason As System.Windows.Forms.RadioButton
    Friend WithEvents rbPostSeason As System.Windows.Forms.RadioButton
    Friend WithEvents chkBatTitle As System.Windows.Forms.CheckBox
    Friend WithEvents chkERATitle As System.Windows.Forms.CheckBox
    Friend WithEvents optTeamResults As System.Windows.Forms.RadioButton
    Friend WithEvents dgResults As System.Windows.Forms.DataGridView
    Friend WithEvents lvwStats As System.Windows.Forms.ListView
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cmdNL = New System.Windows.Forms.Button
        Me.frmAL = New System.Windows.Forms.GroupBox
        Me.optTeamResults = New System.Windows.Forms.RadioButton
        Me.optStandings = New System.Windows.Forms.RadioButton
        Me.optPitchingLeadersTeam = New System.Windows.Forms.RadioButton
        Me.optHittingLeadersTeam = New System.Windows.Forms.RadioButton
        Me.optPitchingLeaders = New System.Windows.Forms.RadioButton
        Me.optHittingLeaders = New System.Windows.Forms.RadioButton
        Me.optTeamPitching = New System.Windows.Forms.RadioButton
        Me.optTeamHitting = New System.Windows.Forms.RadioButton
        Me.cmdAL = New System.Windows.Forms.Button
        Me.lstTeams = New System.Windows.Forms.ListBox
        Me.lvwStats = New System.Windows.Forms.ListView
        Me.grpRegPostSeason = New System.Windows.Forms.GroupBox
        Me.rbPostSeason = New System.Windows.Forms.RadioButton
        Me.rbRegSeason = New System.Windows.Forms.RadioButton
        Me.chkBatTitle = New System.Windows.Forms.CheckBox
        Me.chkERATitle = New System.Windows.Forms.CheckBox
        Me.dgResults = New System.Windows.Forms.DataGridView
        Me.frmAL.SuspendLayout()
        Me.grpRegPostSeason.SuspendLayout()
        CType(Me.dgResults, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdNL
        '
        Me.cmdNL.BackColor = System.Drawing.SystemColors.Control
        Me.cmdNL.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdNL.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdNL.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdNL.Location = New System.Drawing.Point(840, 24)
        Me.cmdNL.Name = "cmdNL"
        Me.cmdNL.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdNL.Size = New System.Drawing.Size(33, 33)
        Me.cmdNL.TabIndex = 11
        Me.cmdNL.Text = "NL"
        Me.cmdNL.UseVisualStyleBackColor = False
        '
        'frmAL
        '
        Me.frmAL.BackColor = System.Drawing.SystemColors.Control
        Me.frmAL.Controls.Add(Me.optTeamResults)
        Me.frmAL.Controls.Add(Me.optStandings)
        Me.frmAL.Controls.Add(Me.optPitchingLeadersTeam)
        Me.frmAL.Controls.Add(Me.optHittingLeadersTeam)
        Me.frmAL.Controls.Add(Me.optPitchingLeaders)
        Me.frmAL.Controls.Add(Me.optHittingLeaders)
        Me.frmAL.Controls.Add(Me.optTeamPitching)
        Me.frmAL.Controls.Add(Me.optTeamHitting)
        Me.frmAL.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.frmAL.ForeColor = System.Drawing.SystemColors.ControlText
        Me.frmAL.Location = New System.Drawing.Point(224, 8)
        Me.frmAL.Name = "frmAL"
        Me.frmAL.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.frmAL.Size = New System.Drawing.Size(201, 178)
        Me.frmAL.TabIndex = 3
        Me.frmAL.TabStop = False
        Me.frmAL.Text = "Reports"
        '
        'optTeamResults
        '
        Me.optTeamResults.AutoSize = True
        Me.optTeamResults.Location = New System.Drawing.Point(8, 56)
        Me.optTeamResults.Name = "optTeamResults"
        Me.optTeamResults.Size = New System.Drawing.Size(90, 18)
        Me.optTeamResults.TabIndex = 11
        Me.optTeamResults.TabStop = True
        Me.optTeamResults.Text = "Team Results"
        Me.optTeamResults.UseVisualStyleBackColor = True
        '
        'optStandings
        '
        Me.optStandings.BackColor = System.Drawing.SystemColors.Control
        Me.optStandings.Cursor = System.Windows.Forms.Cursors.Default
        Me.optStandings.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optStandings.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optStandings.Location = New System.Drawing.Point(8, 156)
        Me.optStandings.Name = "optStandings"
        Me.optStandings.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optStandings.Size = New System.Drawing.Size(153, 17)
        Me.optStandings.TabIndex = 10
        Me.optStandings.TabStop = True
        Me.optStandings.Text = "Standings"
        Me.optStandings.UseVisualStyleBackColor = False
        '
        'optPitchingLeadersTeam
        '
        Me.optPitchingLeadersTeam.BackColor = System.Drawing.SystemColors.Control
        Me.optPitchingLeadersTeam.Cursor = System.Windows.Forms.Cursors.Default
        Me.optPitchingLeadersTeam.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optPitchingLeadersTeam.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optPitchingLeadersTeam.Location = New System.Drawing.Point(8, 136)
        Me.optPitchingLeadersTeam.Name = "optPitchingLeadersTeam"
        Me.optPitchingLeadersTeam.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optPitchingLeadersTeam.Size = New System.Drawing.Size(153, 17)
        Me.optPitchingLeadersTeam.TabIndex = 9
        Me.optPitchingLeadersTeam.TabStop = True
        Me.optPitchingLeadersTeam.Text = "Pitching Leaders by Team"
        Me.optPitchingLeadersTeam.UseVisualStyleBackColor = False
        '
        'optHittingLeadersTeam
        '
        Me.optHittingLeadersTeam.BackColor = System.Drawing.SystemColors.Control
        Me.optHittingLeadersTeam.Cursor = System.Windows.Forms.Cursors.Default
        Me.optHittingLeadersTeam.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optHittingLeadersTeam.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optHittingLeadersTeam.Location = New System.Drawing.Point(8, 116)
        Me.optHittingLeadersTeam.Name = "optHittingLeadersTeam"
        Me.optHittingLeadersTeam.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optHittingLeadersTeam.Size = New System.Drawing.Size(137, 17)
        Me.optHittingLeadersTeam.TabIndex = 8
        Me.optHittingLeadersTeam.TabStop = True
        Me.optHittingLeadersTeam.Text = "Hitting Leaders by Team"
        Me.optHittingLeadersTeam.UseVisualStyleBackColor = False
        '
        'optPitchingLeaders
        '
        Me.optPitchingLeaders.BackColor = System.Drawing.SystemColors.Control
        Me.optPitchingLeaders.Cursor = System.Windows.Forms.Cursors.Default
        Me.optPitchingLeaders.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optPitchingLeaders.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optPitchingLeaders.Location = New System.Drawing.Point(8, 96)
        Me.optPitchingLeaders.Name = "optPitchingLeaders"
        Me.optPitchingLeaders.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optPitchingLeaders.Size = New System.Drawing.Size(137, 25)
        Me.optPitchingLeaders.TabIndex = 7
        Me.optPitchingLeaders.TabStop = True
        Me.optPitchingLeaders.Text = "Pitching Leaders"
        Me.optPitchingLeaders.UseVisualStyleBackColor = False
        '
        'optHittingLeaders
        '
        Me.optHittingLeaders.BackColor = System.Drawing.SystemColors.Control
        Me.optHittingLeaders.Cursor = System.Windows.Forms.Cursors.Default
        Me.optHittingLeaders.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optHittingLeaders.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optHittingLeaders.Location = New System.Drawing.Point(8, 76)
        Me.optHittingLeaders.Name = "optHittingLeaders"
        Me.optHittingLeaders.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optHittingLeaders.Size = New System.Drawing.Size(105, 17)
        Me.optHittingLeaders.TabIndex = 6
        Me.optHittingLeaders.TabStop = True
        Me.optHittingLeaders.Text = "Hitting Leaders"
        Me.optHittingLeaders.UseVisualStyleBackColor = False
        '
        'optTeamPitching
        '
        Me.optTeamPitching.BackColor = System.Drawing.SystemColors.Control
        Me.optTeamPitching.Cursor = System.Windows.Forms.Cursors.Default
        Me.optTeamPitching.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optTeamPitching.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optTeamPitching.Location = New System.Drawing.Point(8, 36)
        Me.optTeamPitching.Name = "optTeamPitching"
        Me.optTeamPitching.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optTeamPitching.Size = New System.Drawing.Size(105, 17)
        Me.optTeamPitching.TabIndex = 5
        Me.optTeamPitching.TabStop = True
        Me.optTeamPitching.Text = "Team Pitching"
        Me.optTeamPitching.UseVisualStyleBackColor = False
        '
        'optTeamHitting
        '
        Me.optTeamHitting.BackColor = System.Drawing.SystemColors.Control
        Me.optTeamHitting.Cursor = System.Windows.Forms.Cursors.Default
        Me.optTeamHitting.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optTeamHitting.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optTeamHitting.Location = New System.Drawing.Point(8, 16)
        Me.optTeamHitting.Name = "optTeamHitting"
        Me.optTeamHitting.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optTeamHitting.Size = New System.Drawing.Size(89, 17)
        Me.optTeamHitting.TabIndex = 4
        Me.optTeamHitting.TabStop = True
        Me.optTeamHitting.Text = "Team Hitting"
        Me.optTeamHitting.UseVisualStyleBackColor = False
        '
        'cmdAL
        '
        Me.cmdAL.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAL.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAL.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAL.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAL.Location = New System.Drawing.Point(792, 24)
        Me.cmdAL.Name = "cmdAL"
        Me.cmdAL.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAL.Size = New System.Drawing.Size(33, 33)
        Me.cmdAL.TabIndex = 1
        Me.cmdAL.Text = "AL"
        Me.cmdAL.UseVisualStyleBackColor = False
        '
        'lstTeams
        '
        Me.lstTeams.BackColor = System.Drawing.SystemColors.Window
        Me.lstTeams.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstTeams.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstTeams.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstTeams.ItemHeight = 14
        Me.lstTeams.Location = New System.Drawing.Point(64, 8)
        Me.lstTeams.Name = "lstTeams"
        Me.lstTeams.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstTeams.Size = New System.Drawing.Size(137, 88)
        Me.lstTeams.TabIndex = 0
        '
        'lvwStats
        '
        Me.lvwStats.Location = New System.Drawing.Point(8, 192)
        Me.lvwStats.Name = "lvwStats"
        Me.lvwStats.Size = New System.Drawing.Size(1000, 443)
        Me.lvwStats.TabIndex = 15
        Me.lvwStats.UseCompatibleStateImageBehavior = False
        '
        'grpRegPostSeason
        '
        Me.grpRegPostSeason.Controls.Add(Me.rbPostSeason)
        Me.grpRegPostSeason.Controls.Add(Me.rbRegSeason)
        Me.grpRegPostSeason.Location = New System.Drawing.Point(13, 104)
        Me.grpRegPostSeason.Name = "grpRegPostSeason"
        Me.grpRegPostSeason.Size = New System.Drawing.Size(205, 65)
        Me.grpRegPostSeason.TabIndex = 16
        Me.grpRegPostSeason.TabStop = False
        Me.grpRegPostSeason.Text = "Season Type"
        '
        'rbPostSeason
        '
        Me.rbPostSeason.AutoSize = True
        Me.rbPostSeason.Location = New System.Drawing.Point(7, 41)
        Me.rbPostSeason.Name = "rbPostSeason"
        Me.rbPostSeason.Size = New System.Drawing.Size(86, 18)
        Me.rbPostSeason.TabIndex = 1
        Me.rbPostSeason.TabStop = True
        Me.rbPostSeason.Text = "Post Season"
        Me.rbPostSeason.UseVisualStyleBackColor = True
        '
        'rbRegSeason
        '
        Me.rbRegSeason.AutoSize = True
        Me.rbRegSeason.Checked = True
        Me.rbRegSeason.Location = New System.Drawing.Point(7, 20)
        Me.rbRegSeason.Name = "rbRegSeason"
        Me.rbRegSeason.Size = New System.Drawing.Size(102, 18)
        Me.rbRegSeason.TabIndex = 0
        Me.rbRegSeason.TabStop = True
        Me.rbRegSeason.Text = "Regular Season"
        Me.rbRegSeason.UseVisualStyleBackColor = True
        '
        'chkBatTitle
        '
        Me.chkBatTitle.AutoSize = True
        Me.chkBatTitle.Location = New System.Drawing.Point(467, 14)
        Me.chkBatTitle.Name = "chkBatTitle"
        Me.chkBatTitle.Size = New System.Drawing.Size(239, 18)
        Me.chkBatTitle.TabIndex = 17
        Me.chkBatTitle.Text = "Only show hitters that qualify for batting title"
        Me.chkBatTitle.UseVisualStyleBackColor = True
        '
        'chkERATitle
        '
        Me.chkERATitle.AutoSize = True
        Me.chkERATitle.Location = New System.Drawing.Point(467, 38)
        Me.chkERATitle.Name = "chkERATitle"
        Me.chkERATitle.Size = New System.Drawing.Size(237, 18)
        Me.chkERATitle.TabIndex = 18
        Me.chkERATitle.Text = "Only show pitchers that qualify for ERA title"
        Me.chkERATitle.UseVisualStyleBackColor = True
        '
        'dgResults
        '
        Me.dgResults.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgResults.Location = New System.Drawing.Point(8, 192)
        Me.dgResults.Name = "dgResults"
        Me.dgResults.Size = New System.Drawing.Size(889, 431)
        Me.dgResults.TabIndex = 19
        Me.dgResults.Visible = False
        '
        'frmStats
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1028, 635)
        Me.Controls.Add(Me.dgResults)
        Me.Controls.Add(Me.chkERATitle)
        Me.Controls.Add(Me.chkBatTitle)
        Me.Controls.Add(Me.grpRegPostSeason)
        Me.Controls.Add(Me.lvwStats)
        Me.Controls.Add(Me.cmdNL)
        Me.Controls.Add(Me.frmAL)
        Me.Controls.Add(Me.cmdAL)
        Me.Controls.Add(Me.lstTeams)
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(4, 30)
        Me.Name = "frmStats"
        Me.Text = "Statistics"
        Me.frmAL.ResumeLayout(False)
        Me.frmAL.PerformLayout()
        Me.grpRegPostSeason.ResumeLayout(False)
        Me.grpRegPostSeason.PerformLayout()
        CType(Me.dgResults, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
#Region "Upgrade Support "
    Private Shared m_vb6FormDefInstance As frmStats
    Private Shared m_InitializingDefInstance As Boolean
    Public Shared Property DefInstance() As frmStats
        Get
            If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
                m_InitializingDefInstance = True
                m_vb6FormDefInstance = New frmStats()
                m_InitializingDefInstance = False
            End If
            DefInstance = m_vb6FormDefInstance
        End Get
        Set(ByVal value As frmStats)
            m_vb6FormDefInstance = value
        End Set
    End Property
#End Region

    Private lvwColumnSorter As ListViewColumnSorter

    'Const TEAM_NAMES_AL As String = "'Toronto','Baltimore','Detroit'," & "'Milwaukee','Cleveland','Boston','New York (A)'," & "'Oakland','Minnesota','Seattle','California'," & "'Kansas City','Chicago (A)','Texas'"
    'Const TEAM_NAMES_NL As String = "'Chicago (N)','Pittsburgh','Montreal','Philadelphia','St Louis'," & "'New York (N)','Atlanta','San Francisco','Houston','San Diego'," & "'Cincinnati','Los Angeles','Colorado','Florida'"
    Const TEAM_DISPLAY_ORDER As String = "152364"

    ''' <summary>
    ''' displays the current season standings
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ShowStandings()
        Dim arrALEast(10) As Standing
        Dim arrALWest(10) As Standing
        Dim arrALCentral(10) As Standing
        Dim arrNLEast(10) As Standing
        Dim arrNLWest(10) As Standing
        Dim arrNLCentral(10) As Standing
        Dim arrStanding(10) As Standing
        Dim totalWins As Double
        Dim totalLosses As Double
        Dim teamsUsed As String = ""
        Dim dsStat As DataSet = Nothing
        Dim rowIndex As Integer
        Dim arrQuery(6) As String
        Dim highestIndex As Integer
        Dim firstIndex As Integer
        Dim filDebug As Integer
        Dim index As Integer
        Dim DataAccess As New clsDataAccess(gstrSeason & "data")

        Try
            If Dir(gAppPath & "debug.bbb") <> Nothing Then
                filDebug = FreeFile()
                FileOpen(filDebug, gAppPath & "debug.out", OpenMode.Append)
                PrintLine(filDebug, "In ShowStandings")
                FileClose(filDebug)
            End If

            'If gAccessConnectStr = Nothing Then
            '    gAccessConnectStr = GetConnectString()
            'End If
            'Using accessConn As OleDb.OleDbConnection = New OleDb.OleDbConnection
            '    accessConn.ConnectionString = GetConnectString()
            '    accessConn.Open()
            ResetStandingHeaders()
            Select Case Val(gstrSeason)
                Case 1982
                    arrQuery(1) = "SELECT * from standings where division = '" & conALBill & "'"
                    arrQuery(2) = "SELECT * from standings where division = '" & conALJohn & "'"
                    arrQuery(3) = "SELECT * from standings where division = '" & conNLBill & "'"
                    arrQuery(4) = "SELECT * from standings where division = '" & conNLJohn & "'"
                Case 1969 To 1981, 1983 To 1993
                    arrQuery(1) = "SELECT * from standings where division = '" & conALEast & "'"
                    arrQuery(2) = "SELECT * from standings where division = '" & conALWest & "'"
                    arrQuery(3) = "SELECT * from standings where division = '" & conNLEast & "'"
                    arrQuery(4) = "SELECT * from standings where division = '" & conNLWest & "'"
                Case Is >= 1994
                    arrQuery(1) = "SELECT * from standings where division = '" & conALEast & "'"
                    arrQuery(2) = "SELECT * from standings where division = '" & conALWest & "'"
                    arrQuery(3) = "SELECT * from standings where division = '" & conNLEast & "'"
                    arrQuery(4) = "SELECT * from standings where division = '" & conNLWest & "'"
                    arrQuery(5) = "SELECT * from standings where division = '" & conALCentral & "'"
                    arrQuery(6) = "SELECT * from standings where division = '" & conNLCentral & "'"
                Case Is < 1969
                    arrQuery(1) = "SELECT * from standings where division = 'AL'"
                    arrQuery(2) = "SELECT * from standings where division = 'NL'"
            End Select

            For divisionIndex As Integer = 1 To 6
                If arrQuery(divisionIndex) <> Nothing Then
                    dsStat = DataAccess.ExecuteDataSet(arrQuery(divisionIndex))
                    rowIndex = 0
                    highestIndex = 0
                    firstIndex = 0

                    For Each dr As DataRow In dsStat.Tables(0).Rows
                        rowIndex += 1
                        With arrStanding(rowIndex)
                            .Wins = CInt(dr.Item("wins"))
                            .Losses = CInt(dr.Item("losses"))
                            .team = dr.Item("playerid").ToString
                            .Games = .Wins - .Losses
                            .Streak = dr.Item("streak").ToString
                            .Last10Wins = 0
                            .Last10Losses = 0
                            For i As Integer = 1 To 10
                                If dr.Item("last" & i.ToString).ToString = "W" Then
                                    .Last10Wins += 1
                                ElseIf dr.Item("last" & i.ToString).ToString = "L" Then
                                    .Last10Losses += 1
                                End If
                            Next i
                        End With
                    Next dr
                    For divIndex As Integer = 1 To 10
                        highestIndex = 0
                        For i As Integer = 1 To 10
                            If arrStanding(i).team <> Nothing Then
                                If teamsUsed.IndexOf(arrStanding(i).team & "|") = -1 Then
                                    If highestIndex = 0 Then
                                        highestIndex = i
                                    Else
                                        If arrStanding(i).Games > arrStanding(highestIndex).Games Then
                                            highestIndex = i
                                        End If
                                    End If
                                End If
                            End If
                        Next i
                        If highestIndex <> 0 Then
                            Select Case divisionIndex
                                Case 1
                                    arrALEast(divIndex) = arrStanding(highestIndex)
                                    teamsUsed += arrALEast(divIndex).team & "|"
                                Case 2
                                    arrALWest(divIndex) = arrStanding(highestIndex)
                                    teamsUsed += arrALWest(divIndex).team & "|"
                                Case 3
                                    arrNLEast(divIndex) = arrStanding(highestIndex)
                                    teamsUsed += arrNLEast(divIndex).team & "|"
                                Case 4
                                    arrNLWest(divIndex) = arrStanding(highestIndex)
                                    teamsUsed += arrNLWest(divIndex).team & "|"
                                Case 5
                                    arrALCentral(divIndex) = arrStanding(highestIndex)
                                    teamsUsed += arrALCentral(divIndex).team & "|"
                                Case 6
                                    arrNLCentral(divIndex) = arrStanding(highestIndex)
                                    teamsUsed += arrNLCentral(divIndex).team & "|"
                            End Select
                        End If
                    Next divIndex
                    dsStat.Clear()
                End If
            Next divisionIndex

            For j As Integer = 0 To 5
                index = CType(TEAM_DISPLAY_ORDER.Substring(j, 1), Integer)
                If arrQuery(index) <> Nothing Then
                    Select Case index
                        Case 1
                            lvwStats.Items.Add(IIF(CInt(gstrSeason) >= 1969, conALEast, "AL"))
                            Call CopyStandingArray(arrStanding, arrALEast, 10)
                        Case 2
                            lvwStats.Items.Add(IIF(CInt(gstrSeason) >= 1969, conALWest, "NL"))
                            Call CopyStandingArray(arrStanding, arrALWest, 10)
                        Case 3
                            lvwStats.Items.Add(conNLEast)
                            Call CopyStandingArray(arrStanding, arrNLEast, 10)
                        Case 4
                            lvwStats.Items.Add(conNLWest)
                            Call CopyStandingArray(arrStanding, arrNLWest, 10)
                        Case 5
                            lvwStats.Items.Add(conALCentral)
                            Call CopyStandingArray(arrStanding, arrALCentral, 10)
                        Case 6
                            lvwStats.Items.Add(conNLCentral)
                            Call CopyStandingArray(arrStanding, arrNLCentral, 10)
                    End Select
                    For i As Integer = 1 To 10
                        If arrStanding(i).team <> Nothing Then
                            lvwStats.Items.Add(arrStanding(i).team)
                            With lvwStats.Items(lvwStats.Items.Count - 1)
                                totalWins = arrStanding(i).Wins
                                .SubItems.Add(totalWins.ToString)
                                totalLosses = arrStanding(i).Losses
                                .SubItems.Add(totalLosses.ToString)
                                If totalWins + totalLosses > 0 Then
                                    .SubItems.Add(Format(totalWins / (totalWins + totalLosses), "#.000"))
                                End If
                                If i = 1 Then
                                    firstIndex = arrStanding(i).Games
                                    .SubItems.Add("-")
                                Else
                                    .SubItems.Add(Format((firstIndex - arrStanding(i).Games) / 2, "##.#"))
                                End If
                                .SubItems.Add(arrStanding(i).Streak)
                                .SubItems.Add(arrStanding(i).Last10Wins.ToString & "-" & arrStanding(i).Last10Losses.ToString)
                            End With
                        End If
                    Next i
                End If
            Next j
            'End Using
        Catch ex As Exception
            Call MsgBox("Error in ShowStandings. " & ex.ToString, MsgBoxStyle.OkOnly)
        Finally
            If Not dsStat Is Nothing Then dsStat.Dispose()
        End Try
    End Sub

    ''' <summary>
    ''' display stats for a selected team
    ''' </summary>
    ''' <param name="isAmericanLeague"></param>
    ''' <remarks></remarks>
    Private Sub GetTeamStats(ByRef isAmericanLeague As Boolean)
        Dim checkPitching As Boolean
        Dim checkHitting As Boolean
        Dim checkStandings As Boolean
        Dim teamName As String = ""
        Dim hitTable As String
        Dim pitchTable As String
        Dim sqlQuery As String = ""
        'Dim leagueTeams As String
        Dim league As String
        'Dim sqlSort As String = ""
        Dim sqlOrder As String
        Dim colGames As Collection

        Dim dsStat As New DataSet
        Dim filDebug As Integer
        Dim bs As New BindingSource
        Dim DataAccess As New clsDataAccess(gstrSeason & "data")

        Try
            If Dir(gAppPath & "debug.bbb") <> Nothing Then
                filDebug = FreeFile()
                FileOpen(filDebug, gAppPath & "debug.out", OpenMode.Append)
                PrintLine(filDebug, "In Chart Lookup")
                FileClose(filDebug)
            End If
            lvwStats.Items.Clear()
            lvwColumnSorter.Order = 0
            lvwColumnSorter.SortColumn = 0
            lvwStats.ListViewItemSorter = lvwColumnSorter
            If optStandings.Checked Then
                ShowStandings()
                Exit Sub
            End If
            lvwStats.Visible = Not optTeamResults.Checked
            dgResults.Visible = optTeamResults.Checked
            If isAmericanLeague Then
                'leagueTeams = TEAM_NAMES_AL
                league = "'AL'"
            Else
                'leagueTeams = TEAM_NAMES_NL
                league = "'NL'"
            End If

            hitTable = IIF(rbRegSeason.Checked, "HITTINGSTATS", "PSHITTINGSTATS")
            pitchTable = IIF(rbRegSeason.Checked, "PITCHINGSTATS", "PSPITCHINGSTATS")

            colGames = New Collection
            'If gAccessConnectStr = Nothing Then
            '    gAccessConnectStr = GetConnectString()
            'End If

            If optTeamHitting.Checked Or optTeamPitching.Checked Or optTeamResults.Checked Then
                If lstTeams.SelectedItems.Count < 1 Then
                    Call MsgBox("Must select a team to display.", MsgBoxStyle.OkOnly)
                    Exit Sub
                End If
                teamName = lstTeams.Items.Item(lstTeams.SelectedIndex).ToString
            End If
            'Using accessConn As OleDb.OleDbConnection = New OleDb.OleDbConnection
            '    accessConn.ConnectionString = GetConnectString()
            '    accessConn.Open()
            If optTeamHitting.Checked Then
                checkHitting = True
                ResetHittingHeaders(False)
                'sqlSort = GetSort(checkHitting, checkPitching)
                'If sqlSort = Nothing Then
                sqlOrder = " ORDER BY AB DESC"
                'Else
                '    sqlOrder = " ORDER BY " & sqlSort & " DESC"
                'End If
                sqlQuery = "SELECT * FROM " & hitTable & " WHERE " & "team = '" & teamName & "'" & sqlOrder
            ElseIf optTeamPitching.Checked Then
                checkPitching = True
                ResetPitchingHeaders(False)
                'sqlSort = GetSort(checkHitting, checkPitching)
                'If sqlSort = Nothing Then
                sqlOrder = " ORDER BY IP DESC"
                'Else
                '    sqlOrder = " ORDER BY " & sqlSort & " DESC"
                'End If
                sqlQuery = "SELECT * FROM " & pitchTable & " WHERE " & "team = '" & teamName & "'" & sqlOrder
            ElseIf optTeamResults.Checked Then
                sqlOrder = " ORDER BY GAME"
                sqlQuery = "SELECT * FROM results where team = '" & teamName & "'" & sqlOrder
            ElseIf optHittingLeaders.Checked Then
                checkHitting = True
                ResetHittingHeaders(True)
                'sqlSort = GetSort(checkHitting, checkPitching)
                'If sqlSort = Nothing Then
                '    sqlOrder = " ORDER BY AB DESC"
                'Else
                '    sqlOrder = " ORDER BY " & sqlSort & " DESC"
                'End If
                sqlOrder = ""
                sqlQuery = "SELECT wins, losses, playerid FROM STANDINGS"
                dsStat = DataAccess.ExecuteDataSet(sqlQuery)
                For Each dr As DataRow In dsStat.Tables(0).Rows
                    colGames.Add(Val(dr.Item("wins")) + Val(dr.Item("losses")), dr.Item("playerid").ToString)
                Next
                dsStat.Clear()
                sqlQuery = "SELECT distinct name FROM " & hitTable & " WHERE league = " & _
                                            league & " AND " & "games >= 20" & sqlOrder
                'select distinct name
            ElseIf optPitchingLeaders.Checked Then
                checkPitching = True
                ResetPitchingHeaders(True)
                'sqlSort = GetSort(checkHitting, checkPitching)
                sqlOrder = ""
                sqlQuery = "SELECT wins, losses, playerid FROM STANDINGS"
                dsStat = DataAccess.ExecuteDataSet(sqlQuery)
                For Each dr As DataRow In dsStat.Tables(0).Rows
                    colGames.Add(Val(dr.Item("wins")) + Val(dr.Item("losses")), dr.Item("playerid").ToString)
                Next
                dsStat.Clear()
                'If sqlSort.ToUpper = "SAVES" Then
                sqlQuery = "SELECT distinct name FROM " & pitchTable & " WHERE league = " & league & " " & sqlOrder
                'Else
                'sqlQuery = "SELECT distinct name FROM " & pitchTable & " WHERE " & _
                '                        "league = " & league & " AND " & "ip >= 75 " & sqlOrder
                'End If
            ElseIf optHittingLeadersTeam.Checked Then
                With lvwStats
                    .View = View.Details
                    .Columns.Clear()
                End With
                checkHitting = True
                ResetHittingHeaders(False)
                sqlQuery = "SELECT DISTINCT team FROM " & hitTable & " WHERE league = " & league
            ElseIf optPitchingLeadersTeam.Checked Then
                checkPitching = True
                ResetPitchingHeaders(False)
                sqlQuery = "SELECT DISTINCT team FROM " & pitchTable & " WHERE league = " & league
            ElseIf optStandings.Checked Then
                checkStandings = True
                ResetStandingHeaders()

            End If

            lvwStats.Visible = False
            lvwStats.View = View.Details

            dsStat = DataAccess.ExecuteDataSet(sqlQuery)
            For Each dr As DataRow In dsStat.Tables(0).Rows
                If checkHitting Then
                    If optHittingLeadersTeam.Checked Then
                        Call GetSumHittingStats((dr.Item("team").ToString), hitTable)
                    ElseIf optHittingLeaders.Checked Then
                        GetPlayerHittingStats(dr.Item("name").ToString, hitTable, _
                                        GetAverageCollection(colGames), teamName, True)
                    Else
                        GetPlayerHittingStats(dr.Item("name").ToString, hitTable, _
                                        GetAverageCollection(colGames), teamName, False)
                    End If
                ElseIf checkPitching Then
                    If optPitchingLeadersTeam.Checked Then
                        'teamwise pitching
                        Call GetSumPitchingStats(dr.Item("team").ToString, pitchTable)
                    ElseIf optPitchingLeaders.Checked Then
                        GetPlayerPitchingStats(dr.Item("name").ToString, pitchTable, _
                                        GetAverageCollection(colGames), teamName, True)
                    Else
                        GetPlayerPitchingStats(dr.Item("name").ToString, pitchTable, _
                                        GetAverageCollection(colGames), teamName, False)
                    End If
                End If
            Next

            If checkHitting And optTeamHitting.Checked Then
                Call GetSumHittingStats(teamName, hitTable)
            End If
            If checkPitching And optTeamPitching.Checked Then
                Call GetSumPitchingStats(teamName, pitchTable)
            End If
            If optTeamResults.Checked Then
                bs.DataSource = dsStat
                bs.DataMember = dsStat.Tables(0).TableName
                dgResults.DataSource = bs
                dgResults.AutoGenerateColumns = True

                dgResults.Columns.RemoveAt(7)
                dgResults.Columns.RemoveAt(7)
                Dim dgvlc As New DataGridViewLinkColumn
                dgvlc.DataPropertyName = "boxscore"
                dgvlc.HeaderText = "BoxScore"
                dgvlc.Width = 260
                dgResults.Columns.Add(dgvlc)
                'dgResults.DataSource = bs
                'dgResults.AutoGenerateColumns = True
                dgResults.AutoResizeColumns()
            Else
                lvwStats.Visible = True
            End If

            'End Using
        Catch ex As Exception
            Call MsgBox("Error in GetTeamStats. " & sqlQuery & " " & ex.ToString, MsgBoxStyle.OkOnly)
        Finally
            If Not dsStat Is Nothing Then dsStat.Dispose()
        End Try
    End Sub

    '''' <summary>
    '''' builds sorting sql based on checked items
    '''' </summary>
    '''' <param name="sortHitting"></param>
    '''' <param name="sortPitching"></param>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Private Function GetSort(ByRef sortHitting As Boolean, ByRef sortPitching As Boolean) As String
    '    Dim sqlSort As String = ""

    '    Try
    '        If sortPitching Then
    '            If optPSWins.Checked Then
    '                sqlSort = "wins"
    '            ElseIf optPSLosses.Checked Then
    '                sqlSort = "losses"
    '            ElseIf optPSStrikeouts.Checked Then
    '                sqlSort = "strikeouts"
    '            ElseIf optPSWalks.Checked Then
    '                sqlSort = "walks"
    '            ElseIf optPSSaves.Checked Then
    '                sqlSort = "saves"
    '            ElseIf optPSCG.Checked Then
    '                sqlSort = "cg"
    '            ElseIf optPSShutouts.Checked Then
    '                sqlSort = "shutouts"
    '            End If
    '        ElseIf sortHitting Then
    '            If optHSHits.Checked Then
    '                sqlSort = "hits"
    '            ElseIf optHSDoubles.Checked Then
    '                sqlSort = "doubles"
    '            ElseIf optHSHomeRuns.Checked Then
    '                sqlSort = "homeruns"
    '            ElseIf optHSRBI.Checked Then
    '                sqlSort = "rbi"
    '            ElseIf optHSRuns.Checked Then
    '                sqlSort = "runs"
    '            ElseIf optHSSB.Checked Then
    '                sqlSort = "sb"
    '            ElseIf optHSWalks.Checked Then
    '                sqlSort = "walks"
    '            End If
    '        End If
    '    Catch ex As Exception
    '        Call MsgBox("Error in GetSort. " & ex.ToString, MsgBoxStyle.OkOnly)
    '    End Try
    '    Return sqlSort
    'End Function

    ''' <summary>
    ''' resets column headers
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ResetHittingHeaders(ByVal leagueLeaders As Boolean)
        Try
            With lvwStats
                .View = View.Details
                .Columns.Clear()
                .Columns.Add("Player", 100, HorizontalAlignment.Center)
                If leagueLeaders Then
                    .Columns.Add("Team", 60, HorizontalAlignment.Center)
                End If
                .Columns.Add("AB", 60, HorizontalAlignment.Center)
                .Columns.Add("H", 50, HorizontalAlignment.Center)
                .Columns.Add("D", 30, HorizontalAlignment.Center)
                .Columns.Add("T", 30, HorizontalAlignment.Center)
                .Columns.Add("HR", 30, HorizontalAlignment.Center)
                .Columns.Add("RBI", 30, HorizontalAlignment.Center)
                .Columns.Add("AVG", 50, HorizontalAlignment.Center)
                .Columns.Add("R", 30, HorizontalAlignment.Center)
                .Columns.Add("SB", 30, HorizontalAlignment.Center)
                .Columns.Add("SBA", 30, HorizontalAlignment.Center)
                .Columns.Add("E", 30, HorizontalAlignment.Center)
                .Columns.Add("K", 50, HorizontalAlignment.Center)
                .Columns.Add("W", 30, HorizontalAlignment.Center)
                .Columns.Add("OBP", 50, HorizontalAlignment.Center)
                .Columns.Add("SLUG", 50, HorizontalAlignment.Center)
                .Columns.Add("HB", 30, HorizontalAlignment.Center)
                .Columns.Add("PB", 30, HorizontalAlignment.Center)
                .Columns.Add("HS", 30, HorizontalAlignment.Center)
                .Columns.Add("G", 30, HorizontalAlignment.Center)
                .Columns.Add("P-G", 175, HorizontalAlignment.Center)
                .Columns.Add("INJ", 30, HorizontalAlignment.Center)
                .Columns.Add("HSH", 50, HorizontalAlignment.Center)
                .Columns.Add("RISP", 50, HorizontalAlignment.Center)
                .Columns.Add("SF", 30, HorizontalAlignment.Center)
                .Columns.Add("SH", 30, HorizontalAlignment.Center)
                .Columns.Add("IBB", 30, HorizontalAlignment.Center)
                .Columns.Add("GIDP", 60, HorizontalAlignment.Center)
                .Columns.Add("PO", 45, HorizontalAlignment.Center)
                .Columns.Add("A", 30, HorizontalAlignment.Center)
                .Columns.Add("DP", 30, HorizontalAlignment.Center)
            End With
        Catch ex As Exception
            Call MsgBox("Error in ResetHittingHeaders. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' resets column headers
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ResetPitchingHeaders(ByVal leagueLeaders As Boolean)
        Try
            With lvwStats
                .View = View.Details
                .Columns.Clear()
                .Columns.Add("Player", 100, HorizontalAlignment.Center)
                If leagueLeaders Then
                    .Columns.Add("Team", 60, HorizontalAlignment.Center)
                End If
                .Columns.Add("IP", 60, HorizontalAlignment.Center)
                .Columns.Add("H", 60, HorizontalAlignment.Center)
                .Columns.Add("R", 30, HorizontalAlignment.Center)
                .Columns.Add("ER", 30, HorizontalAlignment.Center)
                .Columns.Add("W", 30, HorizontalAlignment.Center)
                .Columns.Add("L", 30, HorizontalAlignment.Center)
                .Columns.Add("K", 60, HorizontalAlignment.Center)
                .Columns.Add("W", 30, HorizontalAlignment.Center)
                .Columns.Add("S", 30, HorizontalAlignment.Center)
                .Columns.Add("ERA", 50, HorizontalAlignment.Center)
                .Columns.Add("HR", 30, HorizontalAlignment.Center)
                .Columns.Add("WP", 30, HorizontalAlignment.Center)
                .Columns.Add("HB", 30, HorizontalAlignment.Center)
                .Columns.Add("BK", 30, HorizontalAlignment.Center)
                .Columns.Add("E", 30, HorizontalAlignment.Center)
                .Columns.Add("S", 30, HorizontalAlignment.Center)
                .Columns.Add("R", 30, HorizontalAlignment.Center)
                .Columns.Add("WHIP", 50, HorizontalAlignment.Center)
                .Columns.Add("K/9", 50, HorizontalAlignment.Center)
                .Columns.Add("CG", 30, HorizontalAlignment.Center)
                .Columns.Add("SHO", 60, HorizontalAlignment.Center)
                .Columns.Add("ONEH", 60, HorizontalAlignment.Center)
                .Columns.Add("NOH", 60, HorizontalAlignment.Center)
                .Columns.Add("PER", 60, HorizontalAlignment.Center)
            End With
        Catch ex As Exception
            Call MsgBox("Error in ResetPitchingHeaders. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' resets column headers
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ResetStandingHeaders()
        Try
            With lvwStats
                .View = View.Details
                .Columns.Clear()
                .Columns.Add("Team", 100, HorizontalAlignment.Center)
                .Columns.Add("Wins", 50, HorizontalAlignment.Center)
                .Columns.Add("Losses", 50, HorizontalAlignment.Center)
                .Columns.Add("Pct", 50, HorizontalAlignment.Center)
                .Columns.Add("GB", 50, HorizontalAlignment.Center)
                .Columns.Add("Streak", 50, HorizontalAlignment.Center)
                .Columns.Add("Last10", 50, HorizontalAlignment.Center)
            End With
        Catch ex As Exception
            Call MsgBox("Error in ResetStandingHeaders. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' displays individual hitting stats
    ''' </summary>
    ''' <param name="playerName"></param>
    ''' <param name="hittable"></param>
    ''' <param name="avgGames"></param>
    ''' <remarks></remarks>
    Private Sub GetPlayerHittingStats(ByVal playerName As String, ByVal hittable As String, _
                            ByRef avgGames As Double, ByVal teamName As String, ByVal leagueLeaders As Boolean)
        Dim sqlQuery As New StringBuilder
        Dim dsStat As DataSet = Nothing
        Dim drStat As DataRow = Nothing
        Dim sacrificeHits As Integer
        Dim sacrificeFlies As Integer
        Dim numHB As Integer
        Dim playerQualifies As Boolean
        Dim tempDecimalValue As Double
        Dim tempValue As String = ""
        Dim sqlSingle As String = ""
        Dim team As String = ""
        Dim DataAccess As New clsDataAccess(gstrSeason & "data")

        Try
            sqlQuery.Append("SELECT SUM(AB) AS sumAB, SUM(HITS) AS sumH, SUM(DOUBLES) AS sumD," & " SUM(TRIPLES) AS sumT, ")
            sqlQuery.Append("SUM(HOMERUNS) AS sumHR, SUM(RBI) AS sumRBI," & " SUM(RUNS) AS sumR, SUM(SB) AS sumSB, ")
            sqlQuery.Append("SUM(SBA) AS sumSBA," & " SUM(ERRORS) AS sumE, SUM(STRIKEOUTS) AS sumK, SUM(WALKS) AS sumW, ")
            sqlQuery.Append("SUM(HITBYPITCH) AS sumHB, SUM(PASSEDBALLS) AS sumPB, SUM(SH) AS sumSH, SUM(SF) AS sumSF, ")
            sqlQuery.Append("SUM(GAMES) AS sumG, SUM(HS) as sumHS, SUM(GAMESC) as sumGamesC, SUM(GAMESP) as sumGamesP, ")
            sqlQuery.Append("SUM(GAMES1B) as sumGames1B, SUM(GAMES2B) as sumGames2B, SUM(GAMESSS) as sumGamesSS, ")
            sqlQuery.Append("SUM(GAMES3B) as sumGames3B, SUM(GAMESOF) as sumGamesOF, SUM(GAMESDH) as sumGamesDH, ")
            sqlQuery.Append("MAX(HSH) as maxhsh, SUM(RISPAB) as sumRispAb, SUM(RISPH) as sumRispH, SUM(GIDP) as sumGIDP, ")
            sqlQuery.Append("SUM(IBB) as sumIBB, SUM(PO) as sumPO, SUM(A) as sumA, SUM(DP) as sumDP ")
            sqlQuery.Append("FROM " & hittable & " " & "WHERE name = '" & HandleQuotes(playerName) & "'")
            If teamName <> Nothing Then
                'for hitting leaders, you want to include stats for all league teams (mid season trades). If not,
                'let's just show the stats obtained while with the requested team
                sqlQuery.Append(" AND team = '" & teamName & "'")
            Else
                sqlSingle = "select team from " & hittable & " where name = '" & HandleQuotes(playerName) & "'"
                dsStat = DataAccess.ExecuteDataSet(sqlSingle)
                For Each dr As DataRow In dsStat.Tables(0).Rows
                    team += GetShortTeamTranslation(dr.Item("team").ToString)
                Next
                dsStat.Clear()
            End If

            dsStat = DataAccess.ExecuteDataSet(sqlQuery.ToString)
            drStat = dsStat.Tables(0).Rows(0)

            If Not IsDBNull(drStat.Item("sumSH")) Then
                sacrificeHits = CInt(drStat.Item("sumSH"))
            End If
            If Not IsDBNull(drStat.Item("sumSF")) Then
                sacrificeFlies = CInt(drStat.Item("sumSF"))
            End If

            If optHittingLeaders.Checked And chkBatTitle.CheckState = CheckState.Checked Then
                playerQualifies = avgGames * 3.1 <= _
                                    CInt(drStat.Item("sumAB")) + _
                                    CInt(drStat.Item("sumHB")) + _
                                    CInt(drStat.Item("sumW")) + _
                                    sacrificeHits + _
                                    sacrificeFlies
            Else
                playerQualifies = True
            End If

            If playerQualifies Then
                lvwStats.Items.Add(playerName)
                With lvwStats.Items(lvwStats.Items.Count - 1)
                    If leagueLeaders Then
                        .SubItems.Add(team)
                    End If
                    If Not IsDBNull(drStat.Item("sumAB")) Then
                        .SubItems.Add(drStat.Item("sumAB").ToString)
                    End If
                    If Not IsDBNull(drStat.Item("sumH")) Then
                        .SubItems.Add(drStat.Item("sumH").ToString)
                    End If
                    If Not IsDBNull(drStat.Item("sumD")) Then
                        .SubItems.Add(drStat.Item("sumD").ToString)
                    End If
                    If Not IsDBNull(drStat.Item("sumT")) Then
                        .SubItems.Add(drStat.Item("sumT").ToString)
                    End If
                    If Not IsDBNull(drStat.Item("sumHR")) Then
                        .SubItems.Add(drStat.Item("sumHR").ToString)
                    End If
                    If Not IsDBNull(drStat.Item("sumRBI")) Then
                        .SubItems.Add(drStat.Item("sumRBI").ToString)
                    End If
                    If CDbl(drStat.Item("sumAB")) > 0 Then
                        tempDecimalValue = Val(drStat.Item("sumH")) / Val(drStat.Item("sumAB"))
                        .SubItems.Add(Format(tempDecimalValue, ".###"))
                    Else
                        .SubItems.Add(".000")
                    End If
                    If Not IsDBNull(drStat.Item("sumR")) Then
                        .SubItems.Add(drStat.Item("sumR").ToString)
                    End If
                    If Not IsDBNull(drStat.Item("sumSB")) Then
                        .SubItems.Add(drStat.Item("sumSB").ToString)
                    End If
                    If Not IsDBNull(drStat.Item("sumSBA")) Then
                        .SubItems.Add(drStat.Item("sumSBA").ToString)
                    End If
                    If Not IsDBNull(drStat.Item("sumE")) Then
                        .SubItems.Add(drStat.Item("sumE").ToString)
                    End If
                    If Not IsDBNull(drStat.Item("sumK")) Then
                        .SubItems.Add(drStat.Item("sumK").ToString)
                    End If
                    If Not IsDBNull(drStat.Item("sumW")) Then
                        .SubItems.Add(drStat.Item("sumW").ToString)
                    End If

                    If Not IsDBNull(drStat.Item("sumHB")) Then
                        numHB = CInt(drStat.Item("sumHB"))
                    Else
                        numHB = 0
                    End If
                    If numHB + Val(drStat.Item("sumW")) + Val(drStat.Item("sumAB")) + sacrificeFlies > 0 Then
                        tempDecimalValue = (numHB + Val(drStat.Item("sumW")) + Val(drStat.Item("sumH"))) / _
                                            (sacrificeFlies + numHB + Val(drStat.Item("sumW")) + Val(drStat.Item("sumAB")))
                        .SubItems.Add(Format(tempDecimalValue, "0.###"))
                    Else
                        .SubItems.Add(".000")
                    End If

                    If CDbl(drStat.Item("sumAB")) > 0 Then
                        tempDecimalValue = (Val(drStat.Item("sumH")) + Val(drStat.Item("sumD")) + _
                                        Val(drStat.Item("sumT")) * 2 + _
                                        Val(drStat.Item("sumHR")) * 3) / Val(drStat.Item("sumAB"))
                        .SubItems.Add(Format(tempDecimalValue, "0.###"))
                    Else
                        .SubItems.Add(".000")
                    End If
                    .SubItems.Add(numHB.ToString)
                    If Not IsDBNull(drStat.Item("sumPB")) Then
                        .SubItems.Add(drStat.Item("sumPB").ToString)
                    End If
                    If Not IsDBNull(drStat.Item("sumHS")) Then
                        .SubItems.Add(drStat.Item("sumHS").ToString)
                    End If
                    If Not IsDBNull(drStat.Item("sumG")) Then
                        .SubItems.Add(drStat.Item("sumG").ToString)
                    End If

                    If CInt(drStat.Item("sumGamesC")) > 0 Then
                        tempValue += "2-" & drStat.Item("sumGamesC").ToString & Space(1)
                    End If
                    If CInt(drStat.Item("sumGamesP")) > 0 Then
                        tempValue += "1-" & drStat.Item("sumGamesP").ToString & Space(1)
                    End If
                    If CInt(drStat.Item("sumGames1B")) > 0 Then
                        tempValue += "3-" & drStat.Item("sumGames1B").ToString & Space(1)
                    End If
                    If CInt(drStat.Item("sumGames2B")) > 0 Then
                        tempValue += "4-" & drStat.Item("sumGames2B").ToString & Space(1)
                    End If
                    If CInt(drStat.Item("sumGamesSS")) > 0 Then
                        tempValue += "6-" & drStat.Item("sumGamesSS").ToString & Space(1)
                    End If
                    If CInt(drStat.Item("sumGames3B")) > 0 Then
                        tempValue += "5-" & drStat.Item("sumGames3B").ToString & Space(1)
                    End If
                    If CInt(drStat.Item("sumGamesOF")) > 0 Then
                        tempValue += "OF-" & drStat.Item("sumGamesOF").ToString & Space(1)
                    End If
                    If CInt(drStat.Item("sumGamesDH")) > 0 Then
                        tempValue += "DH-" & drStat.Item("sumGamesDH").ToString & Space(1)
                    End If
                    .SubItems.Add(tempValue)
                    .SubItems.Add("")
                    .SubItems.Add(drStat.Item("maxHSH").ToString)

                    If Not IsDBNull(drStat.Item("sumRispAb")) And Not IsDBNull(drStat.Item("sumRispH")) Then
                        If Val(drStat.Item("sumRispAb")) > 0 Then
                            tempDecimalValue = Val(drStat.Item("sumRispH")) / Val(drStat.Item("sumRispAb"))
                            .SubItems.Add(Format(tempDecimalValue, "0.###"))
                        Else
                            .SubItems.Add(".000")
                        End If
                    Else
                        .SubItems.Add(".000")
                    End If

                    .SubItems.Add(sacrificeFlies.ToString)
                    .SubItems.Add(sacrificeHits.ToString)
                    If Not IsDBNull(drStat.Item("sumIBB")) Then
                        .SubItems.Add(drStat.Item("sumIBB").ToString)
                    End If
                    If Not IsDBNull(drStat.Item("sumGIDP")) Then
                        .SubItems.Add(drStat.Item("sumGIDP").ToString)
                    End If
                    If Not IsDBNull(drStat.Item("sumPO")) Then
                        .SubItems.Add(drStat.Item("sumPO").ToString)
                    End If
                    If Not IsDBNull(drStat.Item("sumA")) Then
                        .SubItems.Add(drStat.Item("sumA").ToString)
                    End If
                    If Not IsDBNull(drStat.Item("sumDP")) Then
                        .SubItems.Add(drStat.Item("sumDP").ToString)
                    End If

                End With
            End If
        Catch ex As Exception
            Call MsgBox("Error in GetPlayerHittingStats. " & ex.ToString, MsgBoxStyle.OkOnly)
        Finally
            If Not dsStat Is Nothing Then dsStat.Dispose()
        End Try
    End Sub

    ''' <summary>
    ''' Determines team hitting stats
    ''' </summary>
    ''' <param name="teamName"></param>
    ''' <param name="hitTable"></param>
    ''' <remarks></remarks>
    Private Sub GetSumHittingStats(ByRef teamName As String, ByRef hitTable As String)
        Dim sqlQuery As String = ""
        Dim dsTeam As DataSet = Nothing
        Dim drTeam As DataRow = Nothing
        Dim tempDoubleValue As Double
        Dim filDebug As Integer
        Dim numHB As Integer
        Dim DataAccess As New clsDataAccess(gstrSeason & "data")

        Try
            If Dir(gAppPath & "debug.bbb") <> Nothing Then
                filDebug = FreeFile()
                FileOpen(filDebug, gAppPath & "debug.out", OpenMode.Append)
                PrintLine(filDebug, "In GetSumHittingsStats")
                FileClose(filDebug)
            End If

            sqlQuery = "SELECT SUM(AB) AS sumAB, SUM(HITS) AS sumH, SUM(DOUBLES) AS sumD, SUM(TRIPLES) AS sumT, " & _
                        "SUM(HOMERUNS) AS sumHR, SUM(RBI) AS sumRBI," & " SUM(RUNS) AS sumR, SUM(SB) AS sumSB, " & _
                        "SUM(SBA) AS sumSBA," & " SUM(ERRORS) AS sumE, SUM(STRIKEOUTS) AS sumK, SUM(WALKS) AS sumW, " & _
                        "SUM(HITBYPITCH) AS sumHB, SUM(PASSEDBALLS) AS sumPB, SUM(RISPAB) AS sumRispAb, " & _
                        "SUM(RISPH) AS sumRispH, SUM(SF) AS sumSF, SUM(SH) AS sumSH, SUM(IBB) AS sumIBB, " & _
                        "SUM(GIDP) AS sumGIDP FROM " & hitTable & " WHERE team = '" & teamName & "'"
            dsTeam = DataAccess.ExecuteDataSet(sqlQuery)

            If dsTeam.Tables(0).Rows.Count > 0 Then
                drTeam = dsTeam.Tables(0).Rows(0)
                lvwStats.Items.Add(teamName)
                With lvwStats.Items(lvwStats.Items.Count - 1)
                    If Not IsDBNull(drTeam.Item("sumAB")) Then
                        .SubItems.Add(drTeam.Item("sumAB").ToString)
                    End If
                    If Not IsDBNull(drTeam.Item("sumH")) Then
                        .SubItems.Add(drTeam.Item("sumH").ToString)
                    End If
                    If Not IsDBNull(drTeam.Item("sumD")) Then
                        .SubItems.Add(drTeam.Item("sumD").ToString)
                    End If
                    If Not IsDBNull(drTeam.Item("sumT")) Then
                        .SubItems.Add(drTeam.Item("sumT").ToString)
                    End If
                    If Not IsDBNull(drTeam.Item("sumHR")) Then
                        .SubItems.Add(drTeam.Item("sumHR").ToString)
                    End If
                    If Not IsDBNull(drTeam.Item("sumRBI")) Then
                        .SubItems.Add(drTeam.Item("sumRBI").ToString)
                    End If
                    If CDbl(.SubItems(1).Text) > 0 Then
                        tempDoubleValue = CInt(.SubItems(2).Text) / CInt(.SubItems(1).Text)
                        .SubItems.Add(Format(tempDoubleValue, ".###"))
                    Else
                        .SubItems.Add(".000")
                    End If
                    If Not IsDBNull(drTeam.Item("sumR")) Then
                        .SubItems.Add(drTeam.Item("sumR").ToString)
                    End If
                    If Not IsDBNull(drTeam.Item("sumSB")) Then
                        .SubItems.Add(drTeam.Item("sumSB").ToString)
                    End If
                    If Not IsDBNull(drTeam.Item("sumSBA")) Then
                        .SubItems.Add(drTeam.Item("sumSBA").ToString)
                    End If
                    If Not IsDBNull(drTeam.Item("sumE")) Then
                        .SubItems.Add(drTeam.Item("sumE").ToString)
                    End If
                    If Not IsDBNull(drTeam.Item("sumK")) Then
                        .SubItems.Add(drTeam.Item("sumK").ToString)
                    End If
                    If Not IsDBNull(drTeam.Item("sumW")) Then
                        .SubItems.Add(drTeam.Item("sumW").ToString)
                    End If
                    If Not IsDBNull(drTeam.Item("sumHB")) Then
                        numHB = CInt(drTeam.Item("sumHB"))
                    Else
                        numHB = 0
                    End If
                    If numHB + CInt(.SubItems(13).Text) + CInt(.SubItems(1).Text) > 0 Then
                        tempDoubleValue = (numHB + CInt(.SubItems(13).Text) + _
                                        CInt(.SubItems(2).Text)) / (numHB + CInt(.SubItems(13).Text) + _
                                        CInt(.SubItems(1).Text))
                        .SubItems.Add(Format(tempDoubleValue, "0.###"))
                    Else
                        .SubItems.Add(".000")
                    End If
                    If CDbl(.SubItems(1).Text) > 0 Then
                        tempDoubleValue = (CInt(.SubItems(2).Text) + CInt(.SubItems(3).Text) + _
                                        CInt(.SubItems(4).Text) * 2 + _
                                        CInt(.SubItems(5).Text) * 3) / CInt(.SubItems(1).Text)
                        .SubItems.Add(Format(tempDoubleValue, "0.###"))
                    Else
                        .SubItems.Add(".000")
                    End If
                    .SubItems.Add(numHB.ToString)
                    If Not IsDBNull(drTeam.Item("sumPB")) Then
                        .SubItems.Add(drTeam.Item("sumPB").ToString)
                    End If
                    .SubItems.Add("")
                    .SubItems.Add("")
                    .SubItems.Add("")
                    .SubItems.Add("")
                    .SubItems.Add("")
                    If Not IsDBNull(drTeam.Item("sumRispAb")) And Not IsDBNull(drTeam.Item("sumRispH")) Then
                        If CDbl(drTeam.Item("sumRispAb")) > 0 Then
                            tempDoubleValue = CDbl(drTeam.Item("sumRispH")) / CDbl(drTeam.Item("sumRispAb"))
                            .SubItems.Add(Format(tempDoubleValue, "0.###"))
                        Else
                            .SubItems.Add(".000")
                        End If
                    Else
                        .SubItems.Add("0.000")
                    End If
                    If Not IsDBNull(drTeam.Item("sumSF")) Then
                        .SubItems.Add(drTeam.Item("sumSF").ToString)
                    End If
                    If Not IsDBNull(drTeam.Item("sumSH")) Then
                        .SubItems.Add(drTeam.Item("sumSH").ToString)
                    End If
                    If Not IsDBNull(drTeam.Item("sumIBB")) Then
                        .SubItems.Add(drTeam.Item("sumIBB").ToString)
                    End If
                    If Not IsDBNull(drTeam.Item("sumGIDP")) Then
                        .SubItems.Add(drTeam.Item("sumGIDP").ToString)
                    End If
                End With
            End If
        Catch ex As Exception
            Call MsgBox("Error in GetSumHittingStats. " & sqlQuery & " " & ex.ToString, MsgBoxStyle.OkOnly)
        Finally
            If Not dsTeam Is Nothing Then dsTeam.Dispose()
        End Try
    End Sub

    ''' <summary>
    ''' displays individual pitching stats
    ''' </summary>
    ''' <param name="playerName"></param>
    ''' <param name="pitchTable"></param>
    ''' <param name="avgGames"></param>
    ''' <remarks></remarks>
    Private Sub GetPlayerPitchingStats(ByVal playerName As String, ByVal pitchTable As String, ByRef avgGames As Double, _
                            ByVal teamName As String, ByVal leagueLeaders As Boolean)
        Dim playerQualifies As Boolean
        Dim sqlQuery As New StringBuilder
        Dim inningsPitched As Double
        Dim tempDecimalValue As Double
        Dim dsStat As DataSet = Nothing
        Dim drStat As DataRow = Nothing
        Dim sqlSingle As String = ""
        Dim team As String = ""
        Dim er As Integer
        Dim walks As Integer
        Dim hits As Integer
        Dim strikeOuts As Integer
        Dim DataAccess As New clsDataAccess(gstrSeason & "data")

        Try
            sqlQuery.Append("SELECT SUM(IP) AS sumIP, SUM(HITS) AS sumH, SUM(RUNS) AS sumRUNS," & "SUM(EARNEDRUNS) AS sumER, ")
            sqlQuery.Append("SUM(WINS) AS sumWINS, SUM(LOSSES) AS sumL," & "SUM(STRIKEOUTS) AS sumK, SUM(WALKS) AS sumW, ")
            sqlQuery.Append("SUM(SAVES) AS sumSV," & "SUM(HOMERUNS) AS sumHR, SUM(WP) AS sumWP, SUM(HB) AS sumHB,")
            sqlQuery.Append("SUM(BK) AS sumBK, SUM(ERRORS) AS sumE, SUM(STARTS) AS sumS," & "SUM(RELIEFS) AS sumR, ")
            sqlQuery.Append("SUM(CG) AS sumCG, SUM(SHUTOUTS) AS sumSH," & "SUM(ONEHITTERS) AS sumOH, ")
            sqlQuery.Append("SUM(NOHITTERS) AS sumNH, SUM(PERFECT) AS sumP " & "FROM " & pitchTable & " ")
            sqlQuery.Append("WHERE name = '" & HandleQuotes(playerName) & "'")
            If teamName <> Nothing Then
                'for pitching leaders, you want to include stats for all league teams (mid season trades). If not,
                'let's just show the stats obtained while with the requested team
                sqlQuery.Append(" AND team = '" & teamName & "'")
            Else
                sqlSingle = "select team from " & pitchTable & " where name = '" & HandleQuotes(playerName) & "'"
                dsStat = DataAccess.ExecuteDataSet(sqlSingle)
                For Each dr As DataRow In dsStat.Tables(0).Rows
                    team += GetShortTeamTranslation(dr.Item("team").ToString)
                Next
                dsStat.Clear()
            End If

            dsStat = DataAccess.ExecuteDataSet(sqlQuery.ToString)
            drStat = dsStat.Tables(0).Rows(0)

            If optPitchingLeaders.Checked And chkERATitle.CheckState = CheckState.Checked Then
                playerQualifies = avgGames <= _
                                        CInt(drStat.Item("sumIP")) * 0.333333
            Else
                playerQualifies = True
            End If
            'other pitching options
            'If playerQualifies Or sqlSort.ToUpper = "SAVES" Then
            If playerQualifies Then
                lvwStats.Items.Add(playerName)
                With lvwStats.Items(lvwStats.Items.Count - 1)
                    If leagueLeaders Then
                        .SubItems.Add(team)
                    End If
                    If Not IsDBNull(drStat.Item("sumIP")) Then
                        inningsPitched = CInt(drStat.Item("sumIP")) * 0.333333
                        .SubItems.Add(Format(inningsPitched, "####.#"))
                    End If
                    If Not IsDBNull(drStat.Item("sumH")) Then
                        .SubItems.Add(drStat.Item("sumH").ToString)
                        hits = CInt(drStat.Item("sumH"))
                    End If
                    If Not IsDBNull(drStat.Item("sumRUNS")) Then
                        .SubItems.Add(drStat.Item("sumRUNS").ToString)
                    End If
                    If Not IsDBNull(drStat.Item("sumER")) Then
                        .SubItems.Add(drStat.Item("sumER").ToString)
                        er = CInt(drStat.Item("sumER"))
                    End If
                    If Not IsDBNull(drStat.Item("sumWINS")) Then
                        .SubItems.Add(drStat.Item("sumWINS").ToString)
                    End If
                    If Not IsDBNull(drStat.Item("sumL")) Then
                        .SubItems.Add(drStat.Item("sumL").ToString)
                    End If
                    If Not IsDBNull(drStat.Item("sumK")) Then
                        .SubItems.Add(drStat.Item("sumK").ToString)
                        strikeOuts = CInt(drStat.Item("sumK"))
                    End If
                    If Not IsDBNull(drStat.Item("sumW")) Then
                        .SubItems.Add(drStat.Item("sumW").ToString)
                        walks = CInt(drStat.Item("sumW"))
                    End If
                    If Not IsDBNull(drStat.Item("sumSV")) Then
                        .SubItems.Add(drStat.Item("sumSV").ToString)
                    End If
                    If inningsPitched > 0 Then
                        tempDecimalValue = er * 9 / inningsPitched
                        .SubItems.Add(Format(tempDecimalValue, "##.##"))
                    End If
                    If Not IsDBNull(drStat.Item("sumHR")) Then
                        .SubItems.Add(drStat.Item("sumHR").ToString)
                    End If
                    If Not IsDBNull(drStat.Item("sumWP")) Then
                        .SubItems.Add(drStat.Item("sumWP").ToString)
                    End If
                    If Not IsDBNull(drStat.Item("sumHB")) Then
                        .SubItems.Add(drStat.Item("sumHB").ToString)
                    End If
                    If Not IsDBNull(drStat.Item("sumBK")) Then
                        .SubItems.Add(drStat.Item("sumBK").ToString)
                    End If
                    If Not IsDBNull(drStat.Item("sumE")) Then
                        .SubItems.Add(drStat.Item("sumE").ToString)
                    End If
                    If Not IsDBNull(drStat.Item("sumS")) Then
                        .SubItems.Add(drStat.Item("sumS").ToString)
                    End If
                    If Not IsDBNull(drStat.Item("sumR")) Then
                        .SubItems.Add(drStat.Item("sumR").ToString)
                    End If
                    If inningsPitched > 0 Then
                        tempDecimalValue = (hits + walks) / inningsPitched
                        .SubItems.Add(Format(tempDecimalValue, "#.##"))
                    End If
                    If inningsPitched > 0 Then
                        tempDecimalValue = strikeOuts * 9 / inningsPitched
                        .SubItems.Add(Format(tempDecimalValue, "#.##"))
                    End If
                    If Not IsDBNull(drStat.Item("sumCG")) Then
                        .SubItems.Add(drStat.Item("sumCG").ToString)
                    End If
                    If Not IsDBNull(drStat.Item("sumSH")) Then
                        .SubItems.Add(drStat.Item("sumSH").ToString)
                    End If
                    If Not IsDBNull(drStat.Item("sumOH")) Then
                        .SubItems.Add(drStat.Item("sumOH").ToString)
                    End If
                    If Not IsDBNull(drStat.Item("sumNH")) Then
                        .SubItems.Add(drStat.Item("sumNH").ToString)
                    End If
                    If Not IsDBNull(drStat.Item("sumP")) Then
                        .SubItems.Add(drStat.Item("sumP").ToString)
                    End If
                End With
            End If
        Catch ex As Exception
            Call MsgBox("Error in GetPlayerPitchingStats. " & sqlQuery.ToString & " " & ex.ToString, MsgBoxStyle.OkOnly)
        Finally
            If Not dsStat Is Nothing Then dsStat.Dispose()
        End Try
    End Sub


    ''' <summary>
    ''' determines team pitching stats
    ''' </summary>
    ''' <param name="teamName"></param>
    ''' <param name="pitchTable"></param>
    ''' <remarks></remarks>
    Private Sub GetSumPitchingStats(ByRef teamName As String, ByRef pitchTable As String)
        Dim sqlQuery As String = ""
        Dim dsTeam As DataSet = Nothing
        Dim drTeam As DataRow = Nothing
        Dim tempDoubleValue As Double
        Dim inningsPitched As Double
        Dim er As Integer
        Dim walks As Integer
        Dim hits As Integer
        Dim strikeOuts As Integer
        Dim filDebug As Integer
        Dim DataAccess As New clsDataAccess(gstrSeason & "data")

        Try
            If Dir(gAppPath & "debug.bbb") <> Nothing Then
                filDebug = FreeFile()
                FileOpen(filDebug, gAppPath & "debug.out", OpenMode.Append)
                PrintLine(filDebug, "In GetSumPitchingStats")
                FileClose(filDebug)
            End If

            sqlQuery = "SELECT SUM(IP) AS sumIP, SUM(HITS) AS sumH, SUM(RUNS) AS sumRUNS," & "SUM(EARNEDRUNS) AS sumER, " & _
                            "SUM(WINS) AS sumWINS, SUM(LOSSES) AS sumL," & "SUM(STRIKEOUTS) AS sumK, SUM(WALKS) AS sumW, " & _
                            "SUM(SAVES) AS sumSV," & "SUM(HOMERUNS) AS sumHR, SUM(WP) AS sumWP, SUM(HB) AS sumHB," & _
                            "SUM(BK) AS sumBK, SUM(ERRORS) AS sumE, SUM(STARTS) AS sumS," & "SUM(RELIEFS) AS sumR, " & _
                            "SUM(CG) AS sumCG, SUM(SHUTOUTS) AS sumSH," & "SUM(ONEHITTERS) AS sumOH, " & _
                            "SUM(NOHITTERS) AS sumNH, SUM(PERFECT) AS sumP " & "FROM " & pitchTable & " " & _
                            "WHERE team = '" & teamName & "'"
            dsTeam = DataAccess.ExecuteDataSet(sqlQuery)
            If dsTeam.Tables(0).Rows.Count > 0 Then
                drTeam = dsTeam.Tables(0).Rows(0)
                lvwStats.Items.Add(teamName)
                With lvwStats.Items(lvwStats.Items.Count - 1)
                    If Not IsDBNull(drTeam.Item("sumIP")) Then
                        inningsPitched = CInt(drTeam.Item("sumIP")) * 0.333333
                        .SubItems.Add(Format(inningsPitched, "####.#"))
                    End If
                    If Not IsDBNull(drTeam.Item("sumH")) Then
                        .SubItems.Add(drTeam.Item("sumH").ToString)
                        hits = CInt(drTeam.Item("sumH"))
                    End If
                    If Not IsDBNull(drTeam.Item("sumRUNS")) Then
                        .SubItems.Add(drTeam.Item("sumRUNS").ToString)
                    End If
                    If Not IsDBNull(drTeam.Item("sumER")) Then
                        .SubItems.Add(drTeam.Item("sumER").ToString)
                        er = CInt(drTeam.Item("sumER"))
                    End If
                    If Not IsDBNull(drTeam.Item("sumWINS")) Then
                        .SubItems.Add(drTeam.Item("sumWINS").ToString)
                    End If
                    If Not IsDBNull(drTeam.Item("sumL")) Then
                        .SubItems.Add(drTeam.Item("sumL").ToString)
                    End If
                    If Not IsDBNull(drTeam.Item("sumK")) Then
                        .SubItems.Add(drTeam.Item("sumK").ToString)
                        strikeOuts = CInt(drTeam.Item("sumK"))
                    End If
                    If Not IsDBNull(drTeam.Item("sumW")) Then
                        .SubItems.Add(drTeam.Item("sumW").ToString)
                        walks = CInt(drTeam.Item("sumW"))
                    End If
                    If Not IsDBNull(drTeam.Item("sumSV")) Then
                        .SubItems.Add(drTeam.Item("sumSV").ToString)
                    End If
                    If inningsPitched > 0 Then
                        tempDoubleValue = er * 9 / inningsPitched
                        .SubItems.Add(Format(tempDoubleValue, "##.##"))
                    End If
                    If Not IsDBNull(drTeam.Item("sumHR")) Then
                        .SubItems.Add(drTeam.Item("sumHR").ToString)
                    End If
                    If Not IsDBNull(drTeam.Item("sumWP")) Then
                        .SubItems.Add(drTeam.Item("sumWP").ToString)
                    End If
                    If Not IsDBNull(drTeam.Item("sumHB")) Then
                        .SubItems.Add(drTeam.Item("sumHB").ToString)
                    End If
                    If Not IsDBNull(drTeam.Item("sumBK")) Then
                        .SubItems.Add(drTeam.Item("sumBK").ToString)
                    End If
                    If Not IsDBNull(drTeam.Item("sumE")) Then
                        .SubItems.Add(drTeam.Item("sumE").ToString)
                    End If
                    If Not IsDBNull(drTeam.Item("sumS")) Then
                        .SubItems.Add(drTeam.Item("sumS").ToString)
                    End If
                    If Not IsDBNull(drTeam.Item("sumR")) Then
                        .SubItems.Add(drTeam.Item("sumR").ToString)
                    End If
                    If inningsPitched > 0 Then
                        tempDoubleValue = (hits + walks) / inningsPitched
                        .SubItems.Add(Format(tempDoubleValue, "#.##"))
                    End If
                    If inningsPitched > 0 Then
                        tempDoubleValue = strikeOuts * 9 / inningsPitched
                        .SubItems.Add(Format(tempDoubleValue, "#.##"))
                    End If
                    If Not IsDBNull(drTeam.Item("sumCG")) Then
                        .SubItems.Add(drTeam.Item("sumCG").ToString)
                    End If
                    If Not IsDBNull(drTeam.Item("sumSH")) Then
                        .SubItems.Add(drTeam.Item("sumSH").ToString)
                    End If
                    If Not IsDBNull(drTeam.Item("sumOH")) Then
                        .SubItems.Add(drTeam.Item("sumOH").ToString)
                    End If
                    If Not IsDBNull(drTeam.Item("sumNH")) Then
                        .SubItems.Add(drTeam.Item("sumNH").ToString)
                    End If
                    If Not IsDBNull(drTeam.Item("sumP")) Then
                        .SubItems.Add(drTeam.Item("sumP").ToString)
                    End If
                End With
            End If

        Catch ex As Exception
            Call MsgBox("Error in GetSumPitchingStats. " & sqlQuery & " " & ex.ToString, MsgBoxStyle.OkOnly)
        Finally
            If Not dsTeam Is Nothing Then dsTeam.Dispose()
        End Try
    End Sub

    Private Sub cmdAL_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAL.Click
        GetTeamStats(True)
    End Sub

    Private Sub cmdNL_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdNL.Click
        GetTeamStats(False)
    End Sub

    ''' <summary>
    ''' load available teams based on the current season
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    Private Sub frmStats_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dim teamName As String

        Try
            lvwStats.View = View.Details
            teamName = Dir(gAppPath & "teams\" & gstrSeason & "\*.*", FileAttribute.Directory)
            While teamName <> Nothing
                If teamName.IndexOf(".") = -1 Then
                    lstTeams.Items.Add(teamName)
                End If
                teamName = Dir()
            End While
        Catch ex As Exception
            Call MsgBox("Error in frmStats_Load. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub lvwStats_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lvwStats.ColumnClick
        If (e.Column = lvwColumnSorter.SortColumn) Then
            ' Reverse the current sort direction for this column.
            If (lvwColumnSorter.Order = SortOrder.Ascending) Then
                lvwColumnSorter.Order = SortOrder.Descending
            Else
                lvwColumnSorter.Order = SortOrder.Ascending
            End If
        Else
            ' Set the column number that is to be sorted; default to ascending.
            lvwColumnSorter.SortColumn = e.Column
            lvwColumnSorter.Order = SortOrder.Ascending
        End If

        ' Perform the sort with these new sort options.
        Me.lvwStats.Sort()
    End Sub


    Private Sub dgResults_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgResults.CellContentClick
        If Me.dgResults.Columns(e.ColumnIndex).GetType.ToString = "System.Windows.Forms.DataGridViewLinkColumn" Then
            Dim link As String
            link = Me.dgResults.Item(e.ColumnIndex, e.RowIndex).Value.ToString
            If Not link.Contains("\") Then
                link = gAppPath & "archive\" & gstrSeason & "\" & link & ".txt"
            End If
            Dim sysPath As String
            sysPath = System.Environment.GetFolderPath _
               (Environment.SpecialFolder.System)
            Shell(sysPath & "\notepad.exe " & link, _
               AppWinStyle.NormalFocus, True)
            Me.TopLevel = True

        End If
    End Sub
End Class


' Implements the manual sorting of items by columns.
'Class ListViewItemComparer
'    Implements IComparer

'    Private col As Integer

'    Public Sub New()
'        col = 0
'    End Sub

'    Public Sub New(ByVal column As Integer)
'        col = column
'    End Sub

'    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer _
'       Implements IComparer.Compare
'        Return [String].Compare(CType(x, ListViewItem).SubItems(col).Text, CType(y, ListViewItem).SubItems(col).Text)
'    End Function
'End Class