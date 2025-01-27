Option Strict On
Option Explicit On

Friend Class frmStartup
	Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
	Public Sub New()
		MyBase.New()
        'If m_vb6FormDefInstance Is Nothing Then
        '    If m_InitializingDefInstance Then
        '        m_vb6FormDefInstance = Me
        '    Else
        '        Try
        '            'For the start-up form, the first instance created is the default instance.
        '            If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
        '                m_vb6FormDefInstance = Me
        '            End If
        '        Catch
        '        End Try
        '    End If
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
    Public WithEvents cmdStats As System.Windows.Forms.Button
    Public WithEvents cmdRosters As System.Windows.Forms.Button
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents cmbSeason As System.Windows.Forms.ComboBox
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents cmbVisitor As System.Windows.Forms.ComboBox
	Public WithEvents cmbHome As System.Windows.Forms.ComboBox
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
    'Public WithEvents chkRules As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
    'Public WithEvents optMode As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
    Public WithEvents rbPartSeason As System.Windows.Forms.RadioButton
    Public WithEvents rbSeason As System.Windows.Forms.RadioButton
    Public WithEvents rbSingleGame As System.Windows.Forms.RadioButton
    Public WithEvents rbNational As System.Windows.Forms.RadioButton
    Public WithEvents rbAmerican As System.Windows.Forms.RadioButton
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmStartup))
        Me.cmdStats = New System.Windows.Forms.Button
        Me.cmdRosters = New System.Windows.Forms.Button
        Me.Frame2 = New System.Windows.Forms.GroupBox
        Me.rbPartSeason = New System.Windows.Forms.RadioButton
        Me.rbSeason = New System.Windows.Forms.RadioButton
        Me.rbSingleGame = New System.Windows.Forms.RadioButton
        Me.cmbSeason = New System.Windows.Forms.ComboBox
        Me.Frame1 = New System.Windows.Forms.GroupBox
        Me.rbNational = New System.Windows.Forms.RadioButton
        Me.rbAmerican = New System.Windows.Forms.RadioButton
        Me.cmbVisitor = New System.Windows.Forms.ComboBox
        Me.cmbHome = New System.Windows.Forms.ComboBox
        Me.cmdOK = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Frame2.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdStats
        '
        Me.cmdStats.BackColor = System.Drawing.SystemColors.Control
        Me.cmdStats.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdStats.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdStats.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdStats.Location = New System.Drawing.Point(223, 312)
        Me.cmdStats.Name = "cmdStats"
        Me.cmdStats.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdStats.Size = New System.Drawing.Size(57, 25)
        Me.cmdStats.TabIndex = 14
        Me.cmdStats.Text = "Stats"
        Me.cmdStats.UseVisualStyleBackColor = False
        '
        'cmdRosters
        '
        Me.cmdRosters.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.cmdRosters.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdRosters.Enabled = False
        Me.cmdRosters.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRosters.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdRosters.Location = New System.Drawing.Point(160, 136)
        Me.cmdRosters.Name = "cmdRosters"
        Me.cmdRosters.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdRosters.Size = New System.Drawing.Size(105, 20)
        Me.cmdRosters.TabIndex = 13
        Me.cmdRosters.Text = "Set Active Rosters"
        Me.cmdRosters.UseVisualStyleBackColor = False
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.rbPartSeason)
        Me.Frame2.Controls.Add(Me.rbSeason)
        Me.Frame2.Controls.Add(Me.rbSingleGame)
        Me.Frame2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(48, 16)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(185, 89)
        Me.Frame2.TabIndex = 10
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Game Mode"
        '
        'rbPartSeason
        '
        Me.rbPartSeason.BackColor = System.Drawing.SystemColors.Control
        Me.rbPartSeason.Cursor = System.Windows.Forms.Cursors.Default
        Me.rbPartSeason.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbPartSeason.ForeColor = System.Drawing.SystemColors.ControlText
        Me.rbPartSeason.Location = New System.Drawing.Point(8, 56)
        Me.rbPartSeason.Name = "rbPartSeason"
        Me.rbPartSeason.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.rbPartSeason.Size = New System.Drawing.Size(161, 17)
        Me.rbPartSeason.TabIndex = 15
        Me.rbPartSeason.TabStop = True
        Me.rbPartSeason.Text = "Partial Season (about half)"
        Me.rbPartSeason.UseVisualStyleBackColor = False
        '
        'rbSeason
        '
        Me.rbSeason.BackColor = System.Drawing.SystemColors.Control
        Me.rbSeason.Cursor = System.Windows.Forms.Cursors.Default
        Me.rbSeason.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbSeason.ForeColor = System.Drawing.SystemColors.ControlText
        Me.rbSeason.Location = New System.Drawing.Point(8, 36)
        Me.rbSeason.Name = "rbSeason"
        Me.rbSeason.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.rbSeason.Size = New System.Drawing.Size(97, 17)
        Me.rbSeason.TabIndex = 12
        Me.rbSeason.TabStop = True
        Me.rbSeason.Text = "Season Play"
        Me.rbSeason.UseVisualStyleBackColor = False
        '
        'rbSingleGame
        '
        Me.rbSingleGame.BackColor = System.Drawing.SystemColors.Control
        Me.rbSingleGame.Checked = True
        Me.rbSingleGame.Cursor = System.Windows.Forms.Cursors.Default
        Me.rbSingleGame.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbSingleGame.ForeColor = System.Drawing.SystemColors.ControlText
        Me.rbSingleGame.Location = New System.Drawing.Point(8, 16)
        Me.rbSingleGame.Name = "rbSingleGame"
        Me.rbSingleGame.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.rbSingleGame.Size = New System.Drawing.Size(105, 25)
        Me.rbSingleGame.TabIndex = 11
        Me.rbSingleGame.TabStop = True
        Me.rbSingleGame.Text = "Single Game"
        Me.rbSingleGame.UseVisualStyleBackColor = False
        '
        'cmbSeason
        '
        Me.cmbSeason.BackColor = System.Drawing.SystemColors.Window
        Me.cmbSeason.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbSeason.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbSeason.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbSeason.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbSeason.Location = New System.Drawing.Point(80, 136)
        Me.cmbSeason.Name = "cmbSeason"
        Me.cmbSeason.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbSeason.Size = New System.Drawing.Size(57, 22)
        Me.cmbSeason.TabIndex = 9
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.rbNational)
        Me.Frame1.Controls.Add(Me.rbAmerican)
        Me.Frame1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(80, 232)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(128, 65)
        Me.Frame1.TabIndex = 5
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Rules"
        '
        'rbNational
        '
        Me.rbNational.BackColor = System.Drawing.SystemColors.Control
        Me.rbNational.Cursor = System.Windows.Forms.Cursors.Default
        Me.rbNational.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbNational.ForeColor = System.Drawing.SystemColors.ControlText
        Me.rbNational.Location = New System.Drawing.Point(8, 40)
        Me.rbNational.Name = "rbNational"
        Me.rbNational.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.rbNational.Size = New System.Drawing.Size(107, 17)
        Me.rbNational.TabIndex = 7
        Me.rbNational.TabStop = True
        Me.rbNational.Text = "National League"
        Me.rbNational.UseVisualStyleBackColor = False
        '
        'rbAmerican
        '
        Me.rbAmerican.BackColor = System.Drawing.SystemColors.Control
        Me.rbAmerican.Checked = True
        Me.rbAmerican.Cursor = System.Windows.Forms.Cursors.Default
        Me.rbAmerican.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbAmerican.ForeColor = System.Drawing.SystemColors.ControlText
        Me.rbAmerican.Location = New System.Drawing.Point(8, 16)
        Me.rbAmerican.Name = "rbAmerican"
        Me.rbAmerican.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.rbAmerican.Size = New System.Drawing.Size(112, 17)
        Me.rbAmerican.TabIndex = 6
        Me.rbAmerican.TabStop = True
        Me.rbAmerican.Text = "American League"
        Me.rbAmerican.UseVisualStyleBackColor = False
        '
        'cmbVisitor
        '
        Me.cmbVisitor.BackColor = System.Drawing.SystemColors.Window
        Me.cmbVisitor.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbVisitor.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbVisitor.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbVisitor.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbVisitor.Location = New System.Drawing.Point(160, 200)
        Me.cmbVisitor.MaxDropDownItems = 30
        Me.cmbVisitor.Name = "cmbVisitor"
        Me.cmbVisitor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbVisitor.Size = New System.Drawing.Size(97, 22)
        Me.cmbVisitor.TabIndex = 4
        '
        'cmbHome
        '
        Me.cmbHome.BackColor = System.Drawing.SystemColors.Window
        Me.cmbHome.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbHome.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbHome.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbHome.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbHome.Location = New System.Drawing.Point(16, 200)
        Me.cmbHome.MaxDropDownItems = 30
        Me.cmbHome.Name = "cmbHome"
        Me.cmbHome.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbHome.Size = New System.Drawing.Size(105, 22)
        Me.cmbHome.TabIndex = 3
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOK.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOK.Location = New System.Drawing.Point(40, 312)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOK.Size = New System.Drawing.Size(65, 25)
        Me.cmdOK.TabIndex = 0
        Me.cmdOK.Text = "OK"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(24, 136)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(49, 17)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Season:"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(160, 176)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(73, 17)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Visitor"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(24, 176)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(49, 17)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Home"
        '
        'frmStartup
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(311, 345)
        Me.Controls.Add(Me.cmdStats)
        Me.Controls.Add(Me.cmdRosters)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.cmbSeason)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.cmbVisitor)
        Me.Controls.Add(Me.cmbHome)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(4, 23)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmStartup"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Choose Teams"
        Me.Frame2.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
#End Region
#Region "Upgrade Support "
    Private Shared m_vb6FormDefInstance As frmStartup
    Private Shared m_InitializingDefInstance As Boolean
    Public Shared Property DefInstance() As frmStartup
        Get
            If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
                m_InitializingDefInstance = True
                m_vb6FormDefInstance = New frmStartup()
                m_InitializingDefInstance = False
            End If
            DefInstance = m_vb6FormDefInstance
        End Get
        Set(ByVal value As frmStartup)
            m_vb6FormDefInstance = Value
        End Set
    End Property
#End Region

    ''' <summary>
    ''' loads the current season
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
        Dim visitingTeam As String = ""
        Dim homeTeam As String = ""
        Dim frmSettings As frmSettings

        Try
            If rbSeason.Checked Then
                gbolSeason = True
                gstrSeason = cmbSeason.Text
                'gAccessConnectStr = GetConnectString()
                Season = New clsSeason()
                Call Season.LoadTeamsFromSched(gstrSeason, visitingTeam, homeTeam)
                Home.teamName = homeTeam
                Visitor.teamName = visitingTeam
            ElseIf rbPartSeason.Checked Then
                gbolSeason = True
                gbolHalfSeason = True
                gstrSeason = cmbSeason.Text
                Season = New clsSeason()
                Call Season.LoadTeamsFromSched(gstrSeason, visitingTeam, homeTeam)
                Home.teamName = homeTeam
                Visitor.teamName = visitingTeam
            Else
                gbolSeason = False
                gstrSeason = cmbSeason.Text
                bolAmericanLeagueRules = (rbAmerican.Checked And CInt(gstrSeason) >= 1973) Or CInt(gstrSeason) >= 2022
                'Home.Clear
                Home.teamName = cmbHome.Text
                'Visitor.Clear
                Visitor.teamName = cmbVisitor.Text
            End If
            If gbolSeasonOver Then
                'season is over
                Me.Close()
                End
                Exit Sub
            End If
            Call LoadData(Home)
            Call LoadData(Visitor)
            frmSettings = New frmSettings()
            frmSettings.Show()
            Me.Visible = False
        Catch ex As Exception
            Call MsgBox("frmStartup.cmdOK_click " & ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' launch rosters
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    Private Sub cmdRosters_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdRosters.Click
        Dim homeTeam As String = ""
        Dim visitingTeam As String = ""
        Dim oSeason As clsSeason
        Dim frmLaunch As frmRoster

        oSeason = New clsSeason()
        gstrSeason = cmbSeason.Text
        maxRosterSize = IIF(CInt(gstrSeason) >= 2021, 26, 25)
        'Determine if season if post season or not
        Call oSeason.LoadTeamsFromSched(gstrSeason, homeTeam, visitingTeam)
        frmLaunch = New frmRoster()
        frmLaunch.Show()
    End Sub

    ''' <summary>
    ''' launch stats form
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    Private Sub cmdStats_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdStats.Click
        Dim frmLaunch As frmStats
        gstrSeason = cmbSeason.Text
        frmLaunch = New frmStats()
        frmLaunch.Show()
    End Sub

    ''' <summary>
    ''' load event for application
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    Private Sub frmStartup_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dim teamName As String
        Dim currentSeason As String
        Dim stringPosition As Integer
        Dim firstYear As String

        Try
            gAppPath = System.Reflection.Assembly.GetExecutingAssembly.Location
            stringPosition = InStrRev(gAppPath, "\")
            If stringPosition > 0 Then
                gAppPath = gAppPath.Substring(0, stringPosition)
            End If

            currentSeason = Dir(gAppPath & "schedule\*.scd")
            firstYear = currentSeason.Substring(0, currentSeason.IndexOf("."))
            gstrSeason = firstYear
            While currentSeason <> Nothing
                stringPosition = currentSeason.IndexOf(".")
                If stringPosition > -1 Then
                    cmbSeason.Items.Add(currentSeason.Substring(0, stringPosition))
                End If
                currentSeason = Dir()
            End While

            teamName = Dir(gAppPath & "teams\" & firstYear & "\*.*", FileAttribute.Directory)
            While teamName <> Nothing
                If teamName.IndexOf(".") = -1 Then
                    cmbHome.Items.Add(teamName)
                    stringPosition = cmbHome.Items.Count()
                    cmbVisitor.Items.Add(teamName)
                End If
                teamName = Dir()
            End While

            'cmbSeason2.Items.Add("Test1")
            'cmbSeason2.Items.Add("Test2")
            Home = New clsTeam
            Visitor = New clsTeam
            Game = New clsGame
            FirstBase = New clsBase
            SecondBase = New clsBase
            ThirdBase = New clsBase
            gblPositions(1) = "P"
            gblPositions(2) = "C"
            gblPositions(3) = "1B"
            gblPositions(4) = "2B"
            gblPositions(5) = "3B"
            gblPositions(6) = "SS"
            gblPositions(7) = "LF"
            gblPositions(8) = "CF"
            gblPositions(9) = "RF"
            cmbHome.SelectedIndex = 0
            cmbVisitor.SelectedIndex = 0
            cmbSeason.SelectedIndex = 0
            'cmbSeason2.SelectedIndex = 0

            bolDebug = Dir(gAppPath & "debug") <> Nothing
            bolJohn = Dir(gAppPath & "john") <> Nothing
        Catch ex As Exception
            Call MsgBox("frmStartup_Load " & ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' enables/disables settings based on SingleGame selection
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub rbSingleGame_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbSingleGame.CheckedChanged
        cmbSeason.Enabled = Not rbSingleGame.Checked
        cmdRosters.Enabled = Not rbSingleGame.Checked
        cmbHome.Enabled = rbSingleGame.Checked
        cmbVisitor.Enabled = rbSingleGame.Checked
        rbAmerican.Enabled = rbSingleGame.Checked
        rbNational.Enabled = rbSingleGame.Checked
    End Sub

End Class