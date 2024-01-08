Option Strict On
Option Explicit On

Imports System.Configuration.configurationManager

Friend Class frmPitching
    Inherits System.Windows.Forms.Form

    Dim colPitCardGroupBoxes As clsFrameArray
    Friend WithEvents pnlPitching As System.Windows.Forms.Panel
    Friend WithEvents dgUsage As System.Windows.Forms.DataGridView
    Dim arrPitCardLabelsCols(22) As clsLabelArray

#Region "Windows Form Designer generated code "
    Public Sub New()
        MyBase.New()
        'If m_vb6FormDefInstance Is Nothing Then
        '	If m_InitializingDefInstance Then
        '		m_vb6FormDefInstance = Me
        '	Else
        '		Try 
        '			'For the start-up form, the first instance created is the default instance.
        '			If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
        '				m_vb6FormDefInstance = Me
        '			End If
        '		Catch
        '		End Try
        '	End If
        'End If
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
    Public WithEvents lstPitchers As System.Windows.Forms.ListBox
    Public WithEvents cmdHitting As System.Windows.Forms.Button
    Public WithEvents cmdOtherTeam As System.Windows.Forms.Button
    Public WithEvents cmdPlayBall As System.Windows.Forms.Button
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lstPitchers = New System.Windows.Forms.ListBox
        Me.cmdHitting = New System.Windows.Forms.Button
        Me.cmdOtherTeam = New System.Windows.Forms.Button
        Me.cmdPlayBall = New System.Windows.Forms.Button
        Me.pnlPitching = New System.Windows.Forms.Panel
        Me.dgUsage = New System.Windows.Forms.DataGridView
        CType(Me.dgUsage, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lstPitchers
        '
        Me.lstPitchers.BackColor = System.Drawing.SystemColors.Window
        Me.lstPitchers.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstPitchers.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstPitchers.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstPitchers.ItemHeight = 14
        Me.lstPitchers.Location = New System.Drawing.Point(152, 522)
        Me.lstPitchers.Name = "lstPitchers"
        Me.lstPitchers.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstPitchers.Size = New System.Drawing.Size(193, 158)
        Me.lstPitchers.TabIndex = 497
        '
        'cmdHitting
        '
        Me.cmdHitting.BackColor = System.Drawing.SystemColors.Control
        Me.cmdHitting.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdHitting.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdHitting.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdHitting.Location = New System.Drawing.Point(12, 566)
        Me.cmdHitting.Name = "cmdHitting"
        Me.cmdHitting.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdHitting.Size = New System.Drawing.Size(113, 25)
        Me.cmdHitting.TabIndex = 496
        Me.cmdHitting.Text = "Hitting"
        Me.cmdHitting.UseVisualStyleBackColor = False
        '
        'cmdOtherTeam
        '
        Me.cmdOtherTeam.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOtherTeam.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOtherTeam.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOtherTeam.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOtherTeam.Location = New System.Drawing.Point(12, 533)
        Me.cmdOtherTeam.Name = "cmdOtherTeam"
        Me.cmdOtherTeam.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOtherTeam.Size = New System.Drawing.Size(134, 25)
        Me.cmdOtherTeam.TabIndex = 495
        Me.cmdOtherTeam.Text = "Opponent Pitching"
        Me.cmdOtherTeam.UseVisualStyleBackColor = False
        '
        'cmdPlayBall
        '
        Me.cmdPlayBall.BackColor = System.Drawing.SystemColors.Control
        Me.cmdPlayBall.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdPlayBall.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPlayBall.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPlayBall.Location = New System.Drawing.Point(12, 597)
        Me.cmdPlayBall.Name = "cmdPlayBall"
        Me.cmdPlayBall.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdPlayBall.Size = New System.Drawing.Size(73, 25)
        Me.cmdPlayBall.TabIndex = 464
        Me.cmdPlayBall.Text = "Play Ball"
        Me.cmdPlayBall.UseVisualStyleBackColor = False
        '
        'pnlPitching
        '
        Me.pnlPitching.AutoScroll = True
        Me.pnlPitching.Location = New System.Drawing.Point(0, 2)
        Me.pnlPitching.Name = "pnlPitching"
        Me.pnlPitching.Size = New System.Drawing.Size(725, 514)
        Me.pnlPitching.TabIndex = 498
        '
        'dgUsage
        '
        Me.dgUsage.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgUsage.Location = New System.Drawing.Point(391, 526)
        Me.dgUsage.Name = "dgUsage"
        Me.dgUsage.Size = New System.Drawing.Size(194, 150)
        Me.dgUsage.TabIndex = 499
        '
        'frmPitching
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(726, 688)
        Me.Controls.Add(Me.dgUsage)
        Me.Controls.Add(Me.pnlPitching)
        Me.Controls.Add(Me.lstPitchers)
        Me.Controls.Add(Me.cmdHitting)
        Me.Controls.Add(Me.cmdOtherTeam)
        Me.Controls.Add(Me.cmdPlayBall)
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(4, 23)
        Me.Name = "frmPitching"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Pitchers"
        CType(Me.dgUsage, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
#End Region
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As frmPitching
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmPitching
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmPitching()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region
    ''' <summary>
    ''' show hitters
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
	Private Sub cmdHitting_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHitting.Click
        Dim frmHitting As frmSettings

        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        SaveSelection(IIF(bolHomeActive, Home, Visitor))
        frmHitting = New frmSettings()
        frmHitting.Show()
        Me.Cursor = System.Windows.Forms.Cursors.Arrow
        Me.Close()
    End Sub

    ''' <summary>
    ''' switch team screens
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
	Private Sub cmdOtherTeam_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOtherTeam.Click
        SaveSelection(IIF(bolHomeActive, Home, Visitor))
		bolHomeActive = Not bolHomeActive
		If bolHomeActive Then
            Me.Text = Home.TeamName & " Bullpen"
			cmdOtherTeam.Text = Visitor.TeamName & " Bullpen"
			Call LoadPitchers(Home)
		Else
            Me.Text = Visitor.TeamName & " Bullpen"
			cmdOtherTeam.Text = Home.TeamName & " Bullpen"
			Call LoadPitchers(Visitor)
		End If
	End Sub

    ''' <summary>
    ''' resume play
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    Private Sub cmdPlayBall_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPlayBall.Click
        Dim frmMainForm As frmMain

        SaveSelection(IIF(bolHomeActive, Home, Visitor))
        If Not bolStartGame Then
            If ValidData() Then
                Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
                frmMainForm = New frmMain
                frmMainForm.Show()
                Me.Cursor = System.Windows.Forms.Cursors.Arrow
            Else
                Call MsgBox("Data invalid. Recheck lineups or select pitcher.", MsgBoxStyle.OkOnly)
            End If
        End If
        Me.Close()
    End Sub

    ''' <summary>
    ''' load pitcher selection screen
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    Private Sub frmPitching_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'colPitCardGroupBoxes = New clsFrameArray(Me)
        colPitCardGroupBoxes = New clsFrameArray(pnlPitching)
        If bolStartGame Then
            bolHomeActive = Not Game.HomeTeamBatting
        Else
            bolHomeActive = True
        End If
        Me.cmdOtherTeam.Visible = Not bolStartGame
        Me.cmdHitting.Visible = Not bolStartGame
        If bolHomeActive Then
            Me.Text = Home.TeamName & " Bullpen"
            cmdOtherTeam.Text = Visitor.TeamName & " Bullpen"
            Call LoadPitchers(Home)
        Else
            Me.Text = Visitor.TeamName & " Bullpen"
            cmdOtherTeam.Text = Home.TeamName & " Bullpen"
            Call LoadPitchers(Visitor)
        End If
    End Sub

    ''' <summary>
    ''' fill pitching cards for display
    ''' </summary>
    ''' <param name="Team"></param>
    ''' <remarks></remarks>
    Private Sub LoadPitchers(ByRef Team As clsTeam)
        Dim objPCard As New PCard
        Dim sqlQuery As String = ""
        Dim pitchTable As String = ""
        Dim ds As New DataSet
        Dim bs As New BindingSource
        Dim DataAccess As New clsDataAccess(gstrSeason & "data")

        Try
            pitchTable = IIF(gbolPostSeason, "PSPITCHINGSTATS", "PITCHINGSTATS")
            lstPitchers.Items.Clear()
            For i As Integer = 0 To Team.pitchers - 1
                Call FillCards(Team, i)
            Next

            If bolStartGame Then
                sqlQuery = "SELECT Name,Reliefs FROM " & pitchTable & " WHERE team = '" & Team.teamName & _
                            "' and reliefs > 0"
            Else
                sqlQuery = "SELECT Name,Starts FROM " & pitchTable & " WHERE team = '" & Team.teamName & _
                            "' and starts > 0"
            End If
            'If gAccessConnectStr = Nothing Then
            '    gAccessConnectStr = GetConnectString()
            'End If
            'Using accessConn As OleDb.OleDbConnection = New OleDb.OleDbConnection
            '    accessConn.ConnectionString = GetConnectString()
            '    accessConn.Open()
            ds = DataAccess.ExecuteDataSet(sqlQuery)
            'End Using
            bs.DataSource = ds
            bs.DataMember = ds.Tables(0).TableName
            dgUsage.DataSource = bs
            dgUsage.AutoGenerateColumns = True
            dgUsage.AutoResizeColumns()

        Catch ex As Exception
            Call MsgBox("LoadPitchers " & ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' fills the pitching cards
    ''' </summary>
    ''' <param name="Team"></param>
    ''' <param name="index"></param>
    ''' <remarks></remarks>
    Private Sub FillCards(ByRef Team As clsTeam, ByRef index As Integer)
        Dim objPCard As New PCard
        Dim playerIndex As Integer

        Try
            playerIndex = index
            If playerIndex = 0 Then
                'Clear group boxes
                colPitCardGroupBoxes.RemoveAll()
            End If

            With Team.GetPitcherPtr(playerIndex + 1)
                If .player = Nothing Then
                    objPCard.Visible = False
                Else

                    objPCard.Visible = True
                    objPCard.grpPCard.Text = .player
                    objPCard.lblThrows.Text = "Field: " & .field & Space(2) & "Throws: " & .throwField
                    objPCard.lblPB.Text = .pbRange
                    objPCard.lblSR.Text = .sr
                    objPCard.lblRR.Text = .rr
                    objPCard.lbl1bf.Text = .hit1Bf
                    objPCard.lbl1b7.Text = .hit1B7
                    objPCard.lbl1b8.Text = .hit1B8
                    objPCard.lbl1b9.Text = .hit1B9
                    objPCard.lblBK.Text = .bk
                    objPCard.lblK.Text = .k
                    objPCard.lblW.Text = .w
                    objPCard.lblPBall.Text = .pb
                    objPCard.lblWP.Text = .wp
                    objPCard.lblOut.Text = .out
                    objPCard.lblStartRel.Text = .StoR
                    If .available Then
                        If bolStartGame Then
                            'Only available relief pitchers
                            If CDbl(.rr) > 0 Then
                                lstPitchers.Items.Add(.player)
                            End If
                        Else
                            'Only available starting pitchers
                            If CDbl(.sr) > 0 Then
                                lstPitchers.Items.Add(.player)
                            End If
                        End If
                    Else
                        'disable pitcher card
                        If .PitStatPtr.gamesInj > 0 Then
                            objPCard.BackColor = System.Drawing.Color.Red
                        Else
                            objPCard.Enabled = False
                        End If
                    End If
                End If
            End With
            colPitCardGroupBoxes.AddNewPCard(objPCard)
        Catch ex As Exception
            Call MsgBox("FillCards " & ex.ToString)
        End Try
    End Sub


    'Private Sub LoadPitcherSel(ByRef pitcherIndex As Integer, ByRef Team As clsTeam)
    '    Dim objPCard As PCard

    '    Try
    '        objPCard = CType(colPitCardGroupBoxes.Item(conMaxPitchers), PCard)
    '        With Team.GetPitcherPtr(pitcherIndex)
    '            objPCard.Visible = True
    '            objPCard.grpPCard.Text = .player
    '            objPCard.lblThrows.Text = "Field: " & .field & Space(2) & "Throws: " & .throwField
    '            objPCard.lblPB.Text = .pbRange
    '            objPCard.lblSR.Text = .sr
    '            objPCard.lblRR.Text = .rr
    '            objPCard.lbl1bf.Text = .hit1Bf
    '            objPCard.lbl1b7.Text = .hit1B7
    '            objPCard.lbl1b8.Text = .hit1B8
    '            objPCard.lbl1b9.Text = .hit1B9
    '            objPCard.lblBK.Text = .bk
    '            objPCard.lblK.Text = .k
    '            objPCard.lblW.Text = .w
    '            objPCard.lblPBall.Text = .pb
    '            objPCard.lblWP.Text = .wp
    '            objPCard.lblOut.Text = .out
    '            objPCard.lblStartRel.Text = .StoR
    '        End With
    '    Catch ex As Exception
    '        Call MsgBox("LoadPitcherSel " & ex.ToString)
    '    End Try
    'End Sub

    'Private Sub lstPitchers_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstPitchers.SelectedIndexChanged
    '    Dim playerName As String

    '    playerName = lstPitchers.SelectedItem.ToString
    '    If bolHomeActive Then
    '        Call LoadPitcherSel(Home.GetPIndexFromName(playerName), Home)
    '    Else
    '        Call LoadPitcherSel(Visitor.GetPIndexFromName(playerName), Visitor)
    '    End If
    'End Sub

    ''' <summary>
    ''' Save the selected pitcher. Also determine if there is a Save situation for a reliever.
    ''' </summary>
    ''' <param name="Team"></param>
    ''' <remarks></remarks>
    Private Sub SaveSelection(ByRef Team As clsTeam)
        Dim runnersOnBase As Integer
        Dim runsLeadingBy As Integer
        Dim isNewPitcher As Boolean
        
        Try
            If lstPitchers.SelectedItems.Count > 0 Then
                isNewPitcher = Team.pitcherSel <> Team.GetPIndexFromName(lstPitchers.SelectedItem.ToString)
                Team.pitcherSel = Team.GetPIndexFromName(lstPitchers.SelectedItem.ToString)
                Team.GetPitcherPtr(Team.pitcherSel).available = False 'Set so that they cannot reenter
                If Not bolAmericanLeagueRules Then
                    If bolStartGame Then
                        Team.nthPitcherUsed += 1
                    End If
                    Call LoadBattingCard(Team)
                End If
                If bolStartGame Then
                    'It is a new reliever, check for save situation
                    If bolHomeActive Then
                        runsLeadingBy = Team.runs - Visitor.runs
                    Else
                        runsLeadingBy = Team.runs - Home.runs
                    End If
                    If runsLeadingBy > 0 And isNewPitcher Then
                        runnersOnBase = GetIncidents(BaseSituation, "1")
                        If ((Game.inning < 9 Or (Game.inning >= 9 And Game.outs = 0)) And runsLeadingBy <= 3) Or runnersOnBase + 2 >= runsLeadingBy Then
                            Game.save = Team.pitcherSel
                        Else
                            Game.save = 0
                        End If
                    End If
                    With Team.GetPitcherPtr(Team.pitcherSel)
                        DetermineStuff(True, .pbRange, .player)
                    End With
                Else
                    'record starting pitcher
                    Team.starter = Team.pitcherSel
                    With Team.GetPitcherPtr(Team.pitcherSel)
                        DetermineStuff(False, .pbRange, .player)
                    End With
                End If
            End If
        Catch ex As Exception
            Call MsgBox("SaveSelection " & ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' determines pitchers stuff before entering game. The resulting factor will be applied to the pitcher's PB rating
    ''' </summary>
    ''' <param name="isReliever"></param>
    ''' <remarks>BB Created: 12/20/2008</remarks>
    Private Sub DetermineStuff(ByVal isReliever As Boolean, ByRef pbRange As String, ByVal name As String)
        Dim rndNum As Integer
        Dim factor As Integer = Stuff.NORMAL
        Dim pbHigh As Integer
        Dim stuffType As String

        If AppSettings("GoodStuff").ToString.ToUpper = "ON" Then
            Randomize()
            rndNum = Convert.ToInt32(64 * Rnd() + 1)
            Select Case rndNum
                Case 1 To 4
                    factor = IIF(isReliever, Stuff.GOOD, Stuff.GREAT)
                    stuffType = IIF(isReliever, "Good", "Great")
                Case 5 To 8
                    factor = Stuff.GOOD
                    stuffType = "Good"
                Case 57 To 60
                    factor = Stuff.BAD
                    stuffType = "Bad"
                Case 61 To 64
                    factor = Stuff.TERRIBLE
                    stuffType = "Terrible"
                Case Else
                    factor = Stuff.NORMAL
                    stuffType = "Normal"
            End Select
            Call MsgBox(name & " has " & stuffType & " stuff.", MsgBoxStyle.OkOnly)
        End If
        pbHigh = CInt(pbRange.Substring(2))
        pbHigh += factor
        If pbHigh < 2 Then
            pbHigh = 2
        End If
        pbRange = "2-" & pbHigh.ToString
    End Sub
End Class