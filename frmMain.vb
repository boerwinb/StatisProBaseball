Option Strict On
Option Explicit On

Imports System.Configuration.configurationManager

Friend Class frmMain
    Inherits System.Windows.Forms.Form

    Dim colLSVisitorTextBoxes As clsTextBoxArray
    Dim colLSHomeTextBoxes As clsTextBoxArray
    Dim arrPitCardLabelsCols(2) As clsLabelArray
    Dim arrBatCardLabelsCols(2) As clsLabelArray
    Dim colHomeKLabels As clsLabelArray
    Dim colVisitorKLabels As clsLabelArray
    Dim colBaseTextBoxes As clsTextBoxArray
    Dim colVisitorBSLabels As clsLabelArray
    Dim colHomeBSLabels As clsLabelArray
    Dim colVisitorPosLabels As clsLabelArray
    Dim colHomePosLabels As clsLabelArray
    Dim colVisitorBatterLabels As clsLabelArray
    Friend WithEvents fraDefOptions As System.Windows.Forms.GroupBox
    Friend WithEvents chkPitchAround1 As System.Windows.Forms.CheckBox
    Friend WithEvents chkGuardLine As System.Windows.Forms.CheckBox
    Friend WithEvents chkCornersIn As System.Windows.Forms.CheckBox
    Friend WithEvents optSqueeze As System.Windows.Forms.RadioButton
    Friend WithEvents chkPitchAround2 As System.Windows.Forms.CheckBox
    Friend WithEvents dgFAC As System.Windows.Forms.DataGridView
    Friend WithEvents lblHBatBorn As System.Windows.Forms.Label
    Friend WithEvents lblHBatAge As System.Windows.Forms.Label
    Friend WithEvents lblVBatBorn As System.Windows.Forms.Label
    Friend WithEvents lblVBatAge As System.Windows.Forms.Label
    Friend WithEvents lblHPitBorn As System.Windows.Forms.Label
    Friend WithEvents lblVPitBorn As System.Windows.Forms.Label
    Friend WithEvents lblHPitAge As System.Windows.Forms.Label
    Friend WithEvents lblVPitAge As System.Windows.Forms.Label
    Dim colHomeBatterLabels As clsLabelArray
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
    Public WithEvents cmdStats As System.Windows.Forms.Button
    Public WithEvents cmdLineupChange As System.Windows.Forms.Button
    Public WithEvents prgVisitorPR As System.Windows.Forms.ProgressBar
    Public WithEvents prgHomePR As System.Windows.Forms.ProgressBar 'AxMSComctlLib.AxProgressBar
    Public WithEvents optIntentionalWalk As System.Windows.Forms.RadioButton
    Public WithEvents chkDNS As System.Windows.Forms.CheckBox
    Public WithEvents cmdPitch As System.Windows.Forms.Button
    Public WithEvents chkInfieldIn As System.Windows.Forms.CheckBox
    Public WithEvents optBunt As System.Windows.Forms.RadioButton
    Public WithEvents optSac As System.Windows.Forms.RadioButton
    Public WithEvents optSteal As System.Windows.Forms.RadioButton
    Public WithEvents optHitandRun As System.Windows.Forms.RadioButton
    Public WithEvents optNormal As System.Windows.Forms.RadioButton
    Public WithEvents Label42 As System.Windows.Forms.Label
    Public WithEvents Label35 As System.Windows.Forms.Label
    Public WithEvents fraOffOptions As System.Windows.Forms.GroupBox
    Public WithEvents picVisitor As System.Windows.Forms.PictureBox
    Public WithEvents picHome As System.Windows.Forms.PictureBox
    Public WithEvents cmdEndGame As System.Windows.Forms.Button
    Public WithEvents cmdTime As System.Windows.Forms.Button
    Public WithEvents Text4 As System.Windows.Forms.TextBox
    Public WithEvents lblStats As System.Windows.Forms.Label
    Public WithEvents lblHome As System.Windows.Forms.Label
    Public WithEvents lblVisitor As System.Windows.Forms.Label
    Public WithEvents lblVisitorPR As System.Windows.Forms.Label
    Public WithEvents Label41 As System.Windows.Forms.Label
    Public WithEvents lblHomePR As System.Windows.Forms.Label
    Public WithEvents Label40 As System.Windows.Forms.Label
    Public WithEvents lblSecond As System.Windows.Forms.Label
    Public WithEvents lblFirst As System.Windows.Forms.Label
    Public WithEvents lblThird As System.Windows.Forms.Label
    Public WithEvents lblOuts As System.Windows.Forms.Label
    Public WithEvents Label36 As System.Windows.Forms.Label
    Public WithEvents lblMessage As System.Windows.Forms.Label
    Public WithEvents Label34 As System.Windows.Forms.Label
    Public WithEvents Label33 As System.Windows.Forms.Label
    Public WithEvents Label32 As System.Windows.Forms.Label
    Public WithEvents Label31 As System.Windows.Forms.Label
    Public WithEvents Label30 As System.Windows.Forms.Label
    Public WithEvents Label29 As System.Windows.Forms.Label
    Public WithEvents Label28 As System.Windows.Forms.Label
    Public WithEvents Label27 As System.Windows.Forms.Label
    Public WithEvents Label26 As System.Windows.Forms.Label
    Public WithEvents Label25 As System.Windows.Forms.Label
    Public WithEvents Label24 As System.Windows.Forms.Label
    Public WithEvents Label23 As System.Windows.Forms.Label
    Public WithEvents Label22 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Public WithEvents grpPitcherVisitor As PCard
    Public WithEvents grpPitcherHome As PCard
    Public WithEvents grpBatterHome As BCard
    Public WithEvents grpBatterVisitor As BCard
    Friend WithEvents lblHPStats As System.Windows.Forms.Label
    Friend WithEvents lblVPStats As System.Windows.Forms.Label
    Friend WithEvents lblHomeMgr As System.Windows.Forms.Label
    Friend WithEvents lblVisitorMgr As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.cmdStats = New System.Windows.Forms.Button
        Me.cmdLineupChange = New System.Windows.Forms.Button
        Me.prgVisitorPR = New System.Windows.Forms.ProgressBar
        Me.prgHomePR = New System.Windows.Forms.ProgressBar
        Me.fraOffOptions = New System.Windows.Forms.GroupBox
        Me.optSqueeze = New System.Windows.Forms.RadioButton
        Me.optIntentionalWalk = New System.Windows.Forms.RadioButton
        Me.chkDNS = New System.Windows.Forms.CheckBox
        Me.optBunt = New System.Windows.Forms.RadioButton
        Me.optSac = New System.Windows.Forms.RadioButton
        Me.optSteal = New System.Windows.Forms.RadioButton
        Me.optHitandRun = New System.Windows.Forms.RadioButton
        Me.optNormal = New System.Windows.Forms.RadioButton
        Me.Label42 = New System.Windows.Forms.Label
        Me.cmdPitch = New System.Windows.Forms.Button
        Me.chkInfieldIn = New System.Windows.Forms.CheckBox
        Me.Label35 = New System.Windows.Forms.Label
        Me.picVisitor = New System.Windows.Forms.PictureBox
        Me.picHome = New System.Windows.Forms.PictureBox
        Me.cmdEndGame = New System.Windows.Forms.Button
        Me.cmdTime = New System.Windows.Forms.Button
        Me.Text4 = New System.Windows.Forms.TextBox
        Me.lblStats = New System.Windows.Forms.Label
        Me.lblHome = New System.Windows.Forms.Label
        Me.lblVisitor = New System.Windows.Forms.Label
        Me.lblVisitorPR = New System.Windows.Forms.Label
        Me.Label41 = New System.Windows.Forms.Label
        Me.lblHomePR = New System.Windows.Forms.Label
        Me.Label40 = New System.Windows.Forms.Label
        Me.lblSecond = New System.Windows.Forms.Label
        Me.lblFirst = New System.Windows.Forms.Label
        Me.lblThird = New System.Windows.Forms.Label
        Me.lblOuts = New System.Windows.Forms.Label
        Me.Label36 = New System.Windows.Forms.Label
        Me.lblMessage = New System.Windows.Forms.Label
        Me.Label34 = New System.Windows.Forms.Label
        Me.Label33 = New System.Windows.Forms.Label
        Me.Label32 = New System.Windows.Forms.Label
        Me.Label31 = New System.Windows.Forms.Label
        Me.Label30 = New System.Windows.Forms.Label
        Me.Label29 = New System.Windows.Forms.Label
        Me.Label28 = New System.Windows.Forms.Label
        Me.Label27 = New System.Windows.Forms.Label
        Me.Label26 = New System.Windows.Forms.Label
        Me.Label25 = New System.Windows.Forms.Label
        Me.Label24 = New System.Windows.Forms.Label
        Me.Label23 = New System.Windows.Forms.Label
        Me.Label22 = New System.Windows.Forms.Label
        Me.lblHPStats = New System.Windows.Forms.Label
        Me.lblVPStats = New System.Windows.Forms.Label
        Me.lblHomeMgr = New System.Windows.Forms.Label
        Me.lblVisitorMgr = New System.Windows.Forms.Label
        Me.grpPitcherVisitor = New StatisProBaseball.PCard
        Me.grpPitcherHome = New StatisProBaseball.PCard
        Me.grpBatterVisitor = New StatisProBaseball.BCard
        Me.grpBatterHome = New StatisProBaseball.BCard
        Me.fraDefOptions = New System.Windows.Forms.GroupBox
        Me.chkPitchAround2 = New System.Windows.Forms.CheckBox
        Me.chkPitchAround1 = New System.Windows.Forms.CheckBox
        Me.chkGuardLine = New System.Windows.Forms.CheckBox
        Me.chkCornersIn = New System.Windows.Forms.CheckBox
        Me.dgFAC = New System.Windows.Forms.DataGridView
        Me.lblHBatBorn = New System.Windows.Forms.Label
        Me.lblHBatAge = New System.Windows.Forms.Label
        Me.lblVBatBorn = New System.Windows.Forms.Label
        Me.lblVBatAge = New System.Windows.Forms.Label
        Me.lblHPitBorn = New System.Windows.Forms.Label
        Me.lblVPitBorn = New System.Windows.Forms.Label
        Me.lblHPitAge = New System.Windows.Forms.Label
        Me.lblVPitAge = New System.Windows.Forms.Label
        Me.fraOffOptions.SuspendLayout()
        CType(Me.picVisitor, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picHome, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraDefOptions.SuspendLayout()
        CType(Me.dgFAC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdStats
        '
        Me.cmdStats.BackColor = System.Drawing.Color.Red
        Me.cmdStats.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdStats.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdStats.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdStats.Location = New System.Drawing.Point(416, 656)
        Me.cmdStats.Name = "cmdStats"
        Me.cmdStats.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdStats.Size = New System.Drawing.Size(81, 25)
        Me.cmdStats.TabIndex = 322
        Me.cmdStats.Text = "Stats"
        Me.cmdStats.UseVisualStyleBackColor = False
        '
        'cmdLineupChange
        '
        Me.cmdLineupChange.BackColor = System.Drawing.Color.Red
        Me.cmdLineupChange.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdLineupChange.Font = New System.Drawing.Font("Century Gothic", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdLineupChange.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdLineupChange.Location = New System.Drawing.Point(344, 656)
        Me.cmdLineupChange.Name = "cmdLineupChange"
        Me.cmdLineupChange.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdLineupChange.Size = New System.Drawing.Size(73, 25)
        Me.cmdLineupChange.TabIndex = 258
        Me.cmdLineupChange.Text = "Lineup"
        Me.cmdLineupChange.UseVisualStyleBackColor = False
        '
        'prgVisitorPR
        '
        Me.prgVisitorPR.Location = New System.Drawing.Point(712, 320)
        Me.prgVisitorPR.Name = "prgVisitorPR"
        Me.prgVisitorPR.Size = New System.Drawing.Size(100, 18)
        Me.prgVisitorPR.TabIndex = 253
        '
        'prgHomePR
        '
        Me.prgHomePR.Location = New System.Drawing.Point(240, 320)
        Me.prgHomePR.Name = "prgHomePR"
        Me.prgHomePR.Size = New System.Drawing.Size(100, 18)
        Me.prgHomePR.TabIndex = 252
        '
        'fraOffOptions
        '
        Me.fraOffOptions.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.fraOffOptions.Controls.Add(Me.optSqueeze)
        Me.fraOffOptions.Controls.Add(Me.optIntentionalWalk)
        Me.fraOffOptions.Controls.Add(Me.chkDNS)
        Me.fraOffOptions.Controls.Add(Me.optBunt)
        Me.fraOffOptions.Controls.Add(Me.optSac)
        Me.fraOffOptions.Controls.Add(Me.optSteal)
        Me.fraOffOptions.Controls.Add(Me.optHitandRun)
        Me.fraOffOptions.Controls.Add(Me.optNormal)
        Me.fraOffOptions.Controls.Add(Me.Label42)
        Me.fraOffOptions.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraOffOptions.ForeColor = System.Drawing.Color.Yellow
        Me.fraOffOptions.Location = New System.Drawing.Point(592, 8)
        Me.fraOffOptions.Name = "fraOffOptions"
        Me.fraOffOptions.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraOffOptions.Size = New System.Drawing.Size(160, 152)
        Me.fraOffOptions.TabIndex = 231
        Me.fraOffOptions.TabStop = False
        Me.fraOffOptions.Text = "Offensive Options"
        '
        'optSqueeze
        '
        Me.optSqueeze.AutoSize = True
        Me.optSqueeze.Location = New System.Drawing.Point(16, 112)
        Me.optSqueeze.Name = "optSqueeze"
        Me.optSqueeze.Size = New System.Drawing.Size(68, 18)
        Me.optSqueeze.TabIndex = 324
        Me.optSqueeze.TabStop = True
        Me.optSqueeze.Text = "Squeeze"
        Me.optSqueeze.UseVisualStyleBackColor = True
        '
        'optIntentionalWalk
        '
        Me.optIntentionalWalk.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.optIntentionalWalk.Cursor = System.Windows.Forms.Cursors.Default
        Me.optIntentionalWalk.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optIntentionalWalk.ForeColor = System.Drawing.Color.Yellow
        Me.optIntentionalWalk.Location = New System.Drawing.Point(16, 96)
        Me.optIntentionalWalk.Name = "optIntentionalWalk"
        Me.optIntentionalWalk.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optIntentionalWalk.Size = New System.Drawing.Size(105, 17)
        Me.optIntentionalWalk.TabIndex = 323
        Me.optIntentionalWalk.TabStop = True
        Me.optIntentionalWalk.Text = "Intentional Walk"
        Me.optIntentionalWalk.UseVisualStyleBackColor = False
        '
        'chkDNS
        '
        Me.chkDNS.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.chkDNS.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkDNS.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkDNS.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkDNS.Location = New System.Drawing.Point(16, 128)
        Me.chkDNS.Name = "chkDNS"
        Me.chkDNS.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkDNS.Size = New System.Drawing.Size(17, 17)
        Me.chkDNS.TabIndex = 279
        Me.chkDNS.Text = "Check1"
        Me.chkDNS.UseVisualStyleBackColor = False
        '
        'optBunt
        '
        Me.optBunt.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.optBunt.Cursor = System.Windows.Forms.Cursors.Default
        Me.optBunt.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optBunt.ForeColor = System.Drawing.Color.Yellow
        Me.optBunt.Location = New System.Drawing.Point(16, 80)
        Me.optBunt.Name = "optBunt"
        Me.optBunt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optBunt.Size = New System.Drawing.Size(137, 17)
        Me.optBunt.TabIndex = 236
        Me.optBunt.TabStop = True
        Me.optBunt.Text = "Bunt for a hit"
        Me.optBunt.UseVisualStyleBackColor = False
        '
        'optSac
        '
        Me.optSac.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.optSac.Cursor = System.Windows.Forms.Cursors.Default
        Me.optSac.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optSac.ForeColor = System.Drawing.Color.Yellow
        Me.optSac.Location = New System.Drawing.Point(16, 64)
        Me.optSac.Name = "optSac"
        Me.optSac.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optSac.Size = New System.Drawing.Size(137, 17)
        Me.optSac.TabIndex = 235
        Me.optSac.TabStop = True
        Me.optSac.Text = "Sacrifice"
        Me.optSac.UseVisualStyleBackColor = False
        '
        'optSteal
        '
        Me.optSteal.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.optSteal.Cursor = System.Windows.Forms.Cursors.Default
        Me.optSteal.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optSteal.ForeColor = System.Drawing.Color.Yellow
        Me.optSteal.Location = New System.Drawing.Point(16, 48)
        Me.optSteal.Name = "optSteal"
        Me.optSteal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optSteal.Size = New System.Drawing.Size(137, 17)
        Me.optSteal.TabIndex = 234
        Me.optSteal.TabStop = True
        Me.optSteal.Text = "Steal a base"
        Me.optSteal.UseVisualStyleBackColor = False
        '
        'optHitandRun
        '
        Me.optHitandRun.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.optHitandRun.Cursor = System.Windows.Forms.Cursors.Default
        Me.optHitandRun.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optHitandRun.ForeColor = System.Drawing.Color.Yellow
        Me.optHitandRun.Location = New System.Drawing.Point(16, 32)
        Me.optHitandRun.Name = "optHitandRun"
        Me.optHitandRun.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optHitandRun.Size = New System.Drawing.Size(137, 17)
        Me.optHitandRun.TabIndex = 233
        Me.optHitandRun.TabStop = True
        Me.optHitandRun.Text = "Hit and Run"
        Me.optHitandRun.UseVisualStyleBackColor = False
        '
        'optNormal
        '
        Me.optNormal.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.optNormal.Checked = True
        Me.optNormal.Cursor = System.Windows.Forms.Cursors.Default
        Me.optNormal.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optNormal.ForeColor = System.Drawing.Color.Yellow
        Me.optNormal.Location = New System.Drawing.Point(16, 16)
        Me.optNormal.Name = "optNormal"
        Me.optNormal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optNormal.Size = New System.Drawing.Size(137, 17)
        Me.optNormal.TabIndex = 232
        Me.optNormal.TabStop = True
        Me.optNormal.Text = "Normal"
        Me.optNormal.UseVisualStyleBackColor = False
        '
        'Label42
        '
        Me.Label42.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label42.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label42.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label42.ForeColor = System.Drawing.Color.Yellow
        Me.Label42.Location = New System.Drawing.Point(32, 128)
        Me.Label42.Name = "Label42"
        Me.Label42.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label42.Size = New System.Drawing.Size(65, 17)
        Me.Label42.TabIndex = 280
        Me.Label42.Text = "Do not steal"
        '
        'cmdPitch
        '
        Me.cmdPitch.BackColor = System.Drawing.Color.Blue
        Me.cmdPitch.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdPitch.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPitch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPitch.Location = New System.Drawing.Point(198, 19)
        Me.cmdPitch.Name = "cmdPitch"
        Me.cmdPitch.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdPitch.Size = New System.Drawing.Size(65, 25)
        Me.cmdPitch.TabIndex = 239
        Me.cmdPitch.Text = "Pitch"
        Me.cmdPitch.UseVisualStyleBackColor = False
        '
        'chkInfieldIn
        '
        Me.chkInfieldIn.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.chkInfieldIn.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkInfieldIn.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkInfieldIn.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkInfieldIn.Location = New System.Drawing.Point(113, 24)
        Me.chkInfieldIn.Name = "chkInfieldIn"
        Me.chkInfieldIn.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkInfieldIn.Size = New System.Drawing.Size(17, 17)
        Me.chkInfieldIn.TabIndex = 237
        Me.chkInfieldIn.Text = "Check1"
        Me.chkInfieldIn.UseVisualStyleBackColor = False
        '
        'Label35
        '
        Me.Label35.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label35.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label35.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label35.ForeColor = System.Drawing.Color.Yellow
        Me.Label35.Location = New System.Drawing.Point(129, 24)
        Me.Label35.Name = "Label35"
        Me.Label35.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label35.Size = New System.Drawing.Size(55, 17)
        Me.Label35.TabIndex = 238
        Me.Label35.Text = "Infield in"
        '
        'picVisitor
        '
        Me.picVisitor.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.picVisitor.BackColor = System.Drawing.Color.White
        Me.picVisitor.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.picVisitor.Cursor = System.Windows.Forms.Cursors.Default
        Me.picVisitor.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.picVisitor.ForeColor = System.Drawing.SystemColors.ControlText
        Me.picVisitor.Location = New System.Drawing.Point(824, 8)
        Me.picVisitor.Name = "picVisitor"
        Me.picVisitor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.picVisitor.Size = New System.Drawing.Size(150, 105)
        Me.picVisitor.TabIndex = 217
        Me.picVisitor.TabStop = False
        '
        'picHome
        '
        Me.picHome.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.picHome.BackColor = System.Drawing.Color.White
        Me.picHome.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.picHome.Cursor = System.Windows.Forms.Cursors.Default
        Me.picHome.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.picHome.ForeColor = System.Drawing.SystemColors.ControlText
        Me.picHome.Location = New System.Drawing.Point(8, 4)
        Me.picHome.Name = "picHome"
        Me.picHome.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.picHome.Size = New System.Drawing.Size(150, 105)
        Me.picHome.TabIndex = 216
        Me.picHome.TabStop = False
        '
        'cmdEndGame
        '
        Me.cmdEndGame.BackColor = System.Drawing.Color.Red
        Me.cmdEndGame.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdEndGame.Font = New System.Drawing.Font("Century Gothic", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdEndGame.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdEndGame.Location = New System.Drawing.Point(496, 656)
        Me.cmdEndGame.Name = "cmdEndGame"
        Me.cmdEndGame.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdEndGame.Size = New System.Drawing.Size(97, 25)
        Me.cmdEndGame.TabIndex = 31
        Me.cmdEndGame.Text = "End Game"
        Me.cmdEndGame.UseVisualStyleBackColor = False
        '
        'cmdTime
        '
        Me.cmdTime.BackColor = System.Drawing.Color.Red
        Me.cmdTime.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdTime.Font = New System.Drawing.Font("Century Gothic", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdTime.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdTime.Location = New System.Drawing.Point(272, 656)
        Me.cmdTime.Name = "cmdTime"
        Me.cmdTime.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdTime.Size = New System.Drawing.Size(73, 25)
        Me.cmdTime.TabIndex = 30
        Me.cmdTime.Text = "Pitching"
        Me.cmdTime.UseVisualStyleBackColor = False
        '
        'Text4
        '
        Me.Text4.AcceptsReturn = True
        Me.Text4.BackColor = System.Drawing.SystemColors.Window
        Me.Text4.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text4.Location = New System.Drawing.Point(472, 440)
        Me.Text4.MaxLength = 0
        Me.Text4.Name = "Text4"
        Me.Text4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text4.Size = New System.Drawing.Size(17, 20)
        Me.Text4.TabIndex = 3
        '
        'lblStats
        '
        Me.lblStats.BackColor = System.Drawing.Color.Black
        Me.lblStats.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblStats.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStats.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblStats.Location = New System.Drawing.Point(192, 80)
        Me.lblStats.Name = "lblStats"
        Me.lblStats.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblStats.Size = New System.Drawing.Size(345, 49)
        Me.lblStats.TabIndex = 321
        Me.lblStats.Text = "StatLine"
        '
        'lblHome
        '
        Me.lblHome.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblHome.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblHome.Font = New System.Drawing.Font("Century Gothic", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHome.ForeColor = System.Drawing.Color.White
        Me.lblHome.Location = New System.Drawing.Point(164, 50)
        Me.lblHome.Name = "lblHome"
        Me.lblHome.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblHome.Size = New System.Drawing.Size(125, 17)
        Me.lblHome.TabIndex = 260
        Me.lblHome.Text = "Home"
        '
        'lblVisitor
        '
        Me.lblVisitor.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblVisitor.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblVisitor.Font = New System.Drawing.Font("Century Gothic", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVisitor.ForeColor = System.Drawing.Color.White
        Me.lblVisitor.Location = New System.Drawing.Point(164, 26)
        Me.lblVisitor.Name = "lblVisitor"
        Me.lblVisitor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblVisitor.Size = New System.Drawing.Size(125, 17)
        Me.lblVisitor.TabIndex = 259
        Me.lblVisitor.Text = "Visitor"
        '
        'lblVisitorPR
        '
        Me.lblVisitorPR.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblVisitorPR.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblVisitorPR.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVisitorPR.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblVisitorPR.Location = New System.Drawing.Point(672, 320)
        Me.lblVisitorPR.Name = "lblVisitorPR"
        Me.lblVisitorPR.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblVisitorPR.Size = New System.Drawing.Size(33, 17)
        Me.lblVisitorPR.TabIndex = 257
        '
        'Label41
        '
        Me.Label41.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label41.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label41.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label41.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label41.Location = New System.Drawing.Point(624, 320)
        Me.Label41.Name = "Label41"
        Me.Label41.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label41.Size = New System.Drawing.Size(33, 17)
        Me.Label41.TabIndex = 256
        Me.Label41.Text = "PR:"
        '
        'lblHomePR
        '
        Me.lblHomePR.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblHomePR.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblHomePR.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHomePR.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblHomePR.Location = New System.Drawing.Point(192, 320)
        Me.lblHomePR.Name = "lblHomePR"
        Me.lblHomePR.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblHomePR.Size = New System.Drawing.Size(33, 17)
        Me.lblHomePR.TabIndex = 255
        '
        'Label40
        '
        Me.Label40.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label40.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label40.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label40.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label40.Location = New System.Drawing.Point(152, 320)
        Me.Label40.Name = "Label40"
        Me.Label40.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label40.Size = New System.Drawing.Size(33, 17)
        Me.Label40.TabIndex = 254
        Me.Label40.Text = "PR:"
        '
        'lblSecond
        '
        Me.lblSecond.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblSecond.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSecond.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSecond.ForeColor = System.Drawing.Color.Red
        Me.lblSecond.Location = New System.Drawing.Point(424, 320)
        Me.lblSecond.Name = "lblSecond"
        Me.lblSecond.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSecond.Size = New System.Drawing.Size(121, 17)
        Me.lblSecond.TabIndex = 245
        Me.lblSecond.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblFirst
        '
        Me.lblFirst.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblFirst.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFirst.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFirst.ForeColor = System.Drawing.Color.Red
        Me.lblFirst.Location = New System.Drawing.Point(560, 392)
        Me.lblFirst.Name = "lblFirst"
        Me.lblFirst.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFirst.Size = New System.Drawing.Size(113, 17)
        Me.lblFirst.TabIndex = 244
        '
        'lblThird
        '
        Me.lblThird.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblThird.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblThird.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblThird.ForeColor = System.Drawing.Color.Red
        Me.lblThird.Location = New System.Drawing.Point(288, 392)
        Me.lblThird.Name = "lblThird"
        Me.lblThird.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblThird.Size = New System.Drawing.Size(96, 17)
        Me.lblThird.TabIndex = 243
        Me.lblThird.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblOuts
        '
        Me.lblOuts.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblOuts.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblOuts.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOuts.ForeColor = System.Drawing.Color.Yellow
        Me.lblOuts.Location = New System.Drawing.Point(488, 392)
        Me.lblOuts.Name = "lblOuts"
        Me.lblOuts.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblOuts.Size = New System.Drawing.Size(17, 17)
        Me.lblOuts.TabIndex = 242
        Me.lblOuts.Text = "0"
        '
        'Label36
        '
        Me.Label36.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label36.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label36.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label36.ForeColor = System.Drawing.Color.Yellow
        Me.Label36.Location = New System.Drawing.Point(448, 392)
        Me.Label36.Name = "Label36"
        Me.Label36.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label36.Size = New System.Drawing.Size(33, 17)
        Me.Label36.TabIndex = 241
        Me.Label36.Text = "Outs:"
        '
        'lblMessage
        '
        Me.lblMessage.BackColor = System.Drawing.Color.Black
        Me.lblMessage.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMessage.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMessage.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblMessage.Location = New System.Drawing.Point(152, 224)
        Me.lblMessage.Name = "lblMessage"
        Me.lblMessage.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMessage.Size = New System.Drawing.Size(657, 81)
        Me.lblMessage.TabIndex = 240
        Me.lblMessage.Text = "Message board"
        '
        'Label34
        '
        Me.Label34.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label34.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label34.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label34.ForeColor = System.Drawing.Color.White
        Me.Label34.Location = New System.Drawing.Point(536, 8)
        Me.Label34.Name = "Label34"
        Me.Label34.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label34.Size = New System.Drawing.Size(9, 17)
        Me.Label34.TabIndex = 230
        Me.Label34.Text = "E"
        Me.Label34.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label33
        '
        Me.Label33.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label33.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label33.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label33.ForeColor = System.Drawing.Color.White
        Me.Label33.Location = New System.Drawing.Point(512, 8)
        Me.Label33.Name = "Label33"
        Me.Label33.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label33.Size = New System.Drawing.Size(9, 17)
        Me.Label33.TabIndex = 229
        Me.Label33.Text = "H"
        Me.Label33.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label32
        '
        Me.Label32.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label32.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label32.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label32.ForeColor = System.Drawing.Color.White
        Me.Label32.Location = New System.Drawing.Point(488, 8)
        Me.Label32.Name = "Label32"
        Me.Label32.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label32.Size = New System.Drawing.Size(9, 17)
        Me.Label32.TabIndex = 228
        Me.Label32.Text = "R"
        Me.Label32.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label31
        '
        Me.Label31.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label31.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label31.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label31.ForeColor = System.Drawing.Color.White
        Me.Label31.Location = New System.Drawing.Point(444, 8)
        Me.Label31.Name = "Label31"
        Me.Label31.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label31.Size = New System.Drawing.Size(20, 17)
        Me.Label31.TabIndex = 227
        Me.Label31.Text = "10"
        Me.Label31.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label30
        '
        Me.Label30.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label30.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label30.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label30.ForeColor = System.Drawing.Color.White
        Me.Label30.Location = New System.Drawing.Point(432, 8)
        Me.Label30.Name = "Label30"
        Me.Label30.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label30.Size = New System.Drawing.Size(9, 17)
        Me.Label30.TabIndex = 226
        Me.Label30.Text = "9"
        Me.Label30.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label29
        '
        Me.Label29.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label29.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label29.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label29.ForeColor = System.Drawing.Color.White
        Me.Label29.Location = New System.Drawing.Point(416, 8)
        Me.Label29.Name = "Label29"
        Me.Label29.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label29.Size = New System.Drawing.Size(9, 17)
        Me.Label29.TabIndex = 225
        Me.Label29.Text = "8"
        Me.Label29.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label28
        '
        Me.Label28.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label28.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label28.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label28.ForeColor = System.Drawing.Color.White
        Me.Label28.Location = New System.Drawing.Point(400, 8)
        Me.Label28.Name = "Label28"
        Me.Label28.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label28.Size = New System.Drawing.Size(9, 17)
        Me.Label28.TabIndex = 224
        Me.Label28.Text = "7"
        Me.Label28.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label27
        '
        Me.Label27.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label27.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label27.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label27.ForeColor = System.Drawing.Color.White
        Me.Label27.Location = New System.Drawing.Point(384, 8)
        Me.Label27.Name = "Label27"
        Me.Label27.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label27.Size = New System.Drawing.Size(9, 17)
        Me.Label27.TabIndex = 223
        Me.Label27.Text = "6"
        Me.Label27.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label26
        '
        Me.Label26.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label26.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label26.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label26.ForeColor = System.Drawing.Color.White
        Me.Label26.Location = New System.Drawing.Point(368, 8)
        Me.Label26.Name = "Label26"
        Me.Label26.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label26.Size = New System.Drawing.Size(9, 17)
        Me.Label26.TabIndex = 222
        Me.Label26.Text = "5"
        Me.Label26.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label25
        '
        Me.Label25.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label25.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label25.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label25.ForeColor = System.Drawing.Color.White
        Me.Label25.Location = New System.Drawing.Point(352, 8)
        Me.Label25.Name = "Label25"
        Me.Label25.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label25.Size = New System.Drawing.Size(9, 17)
        Me.Label25.TabIndex = 221
        Me.Label25.Text = "4"
        Me.Label25.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label24
        '
        Me.Label24.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label24.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label24.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label24.ForeColor = System.Drawing.Color.White
        Me.Label24.Location = New System.Drawing.Point(336, 8)
        Me.Label24.Name = "Label24"
        Me.Label24.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label24.Size = New System.Drawing.Size(9, 17)
        Me.Label24.TabIndex = 220
        Me.Label24.Text = "3"
        Me.Label24.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label23
        '
        Me.Label23.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label23.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label23.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label23.ForeColor = System.Drawing.Color.White
        Me.Label23.Location = New System.Drawing.Point(320, 8)
        Me.Label23.Name = "Label23"
        Me.Label23.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label23.Size = New System.Drawing.Size(9, 17)
        Me.Label23.TabIndex = 219
        Me.Label23.Text = "2"
        Me.Label23.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label22
        '
        Me.Label22.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label22.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label22.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label22.ForeColor = System.Drawing.Color.White
        Me.Label22.Location = New System.Drawing.Point(304, 8)
        Me.Label22.Name = "Label22"
        Me.Label22.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label22.Size = New System.Drawing.Size(9, 17)
        Me.Label22.TabIndex = 218
        Me.Label22.Text = "1"
        Me.Label22.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblHPStats
        '
        Me.lblHPStats.BackColor = System.Drawing.Color.Black
        Me.lblHPStats.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHPStats.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblHPStats.Location = New System.Drawing.Point(8, 168)
        Me.lblHPStats.Name = "lblHPStats"
        Me.lblHPStats.Size = New System.Drawing.Size(240, 23)
        Me.lblHPStats.TabIndex = 323
        Me.lblHPStats.Text = "HPStatLine"
        '
        'lblVPStats
        '
        Me.lblVPStats.BackColor = System.Drawing.Color.Black
        Me.lblVPStats.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVPStats.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblVPStats.Location = New System.Drawing.Point(760, 168)
        Me.lblVPStats.Name = "lblVPStats"
        Me.lblVPStats.Size = New System.Drawing.Size(216, 23)
        Me.lblVPStats.TabIndex = 324
        Me.lblVPStats.Text = "VPStatLine"
        '
        'lblHomeMgr
        '
        Me.lblHomeMgr.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHomeMgr.ForeColor = System.Drawing.Color.White
        Me.lblHomeMgr.Location = New System.Drawing.Point(8, 144)
        Me.lblHomeMgr.Name = "lblHomeMgr"
        Me.lblHomeMgr.Size = New System.Drawing.Size(168, 16)
        Me.lblHomeMgr.TabIndex = 325
        Me.lblHomeMgr.Text = "Manager"
        '
        'lblVisitorMgr
        '
        Me.lblVisitorMgr.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVisitorMgr.ForeColor = System.Drawing.Color.White
        Me.lblVisitorMgr.Location = New System.Drawing.Point(840, 144)
        Me.lblVisitorMgr.Name = "lblVisitorMgr"
        Me.lblVisitorMgr.Size = New System.Drawing.Size(136, 16)
        Me.lblVisitorMgr.TabIndex = 326
        Me.lblVisitorMgr.Text = "Manager"
        '
        'grpPitcherVisitor
        '
        Me.grpPitcherVisitor.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.grpPitcherVisitor.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpPitcherVisitor.ForeColor = System.Drawing.Color.Blue
        Me.grpPitcherVisitor.Location = New System.Drawing.Point(840, 200)
        Me.grpPitcherVisitor.Name = "grpPitcherVisitor"
        Me.grpPitcherVisitor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.grpPitcherVisitor.Size = New System.Drawing.Size(134, 145)
        Me.grpPitcherVisitor.TabIndex = 186
        Me.grpPitcherVisitor.TabStop = False
        '
        'grpPitcherHome
        '
        Me.grpPitcherHome.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.grpPitcherHome.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpPitcherHome.ForeColor = System.Drawing.Color.Blue
        Me.grpPitcherHome.Location = New System.Drawing.Point(8, 200)
        Me.grpPitcherHome.Name = "grpPitcherHome"
        Me.grpPitcherHome.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.grpPitcherHome.Size = New System.Drawing.Size(134, 145)
        Me.grpPitcherHome.TabIndex = 156
        Me.grpPitcherHome.TabStop = False
        '
        'grpBatterVisitor
        '
        Me.grpBatterVisitor.BackColor = System.Drawing.SystemColors.Control
        Me.grpBatterVisitor.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpBatterVisitor.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grpBatterVisitor.Location = New System.Drawing.Point(624, 424)
        Me.grpBatterVisitor.Name = "grpBatterVisitor"
        Me.grpBatterVisitor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.grpBatterVisitor.Size = New System.Drawing.Size(144, 172)
        Me.grpBatterVisitor.TabIndex = 112
        Me.grpBatterVisitor.TabStop = False
        '
        'grpBatterHome
        '
        Me.grpBatterHome.BackColor = System.Drawing.SystemColors.Control
        Me.grpBatterHome.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpBatterHome.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grpBatterHome.Location = New System.Drawing.Point(213, 424)
        Me.grpBatterHome.Name = "grpBatterHome"
        Me.grpBatterHome.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.grpBatterHome.Size = New System.Drawing.Size(144, 172)
        Me.grpBatterHome.TabIndex = 68
        Me.grpBatterHome.TabStop = False
        '
        'fraDefOptions
        '
        Me.fraDefOptions.Controls.Add(Me.chkPitchAround2)
        Me.fraDefOptions.Controls.Add(Me.chkPitchAround1)
        Me.fraDefOptions.Controls.Add(Me.cmdPitch)
        Me.fraDefOptions.Controls.Add(Me.chkInfieldIn)
        Me.fraDefOptions.Controls.Add(Me.chkGuardLine)
        Me.fraDefOptions.Controls.Add(Me.chkCornersIn)
        Me.fraDefOptions.Controls.Add(Me.Label35)
        Me.fraDefOptions.ForeColor = System.Drawing.Color.Yellow
        Me.fraDefOptions.Location = New System.Drawing.Point(475, 160)
        Me.fraDefOptions.Name = "fraDefOptions"
        Me.fraDefOptions.Size = New System.Drawing.Size(279, 61)
        Me.fraDefOptions.TabIndex = 327
        Me.fraDefOptions.TabStop = False
        Me.fraDefOptions.Text = "Defensive Options"
        '
        'chkPitchAround2
        '
        Me.chkPitchAround2.AutoSize = True
        Me.chkPitchAround2.Location = New System.Drawing.Point(9, 34)
        Me.chkPitchAround2.Name = "chkPitchAround2"
        Me.chkPitchAround2.Size = New System.Drawing.Size(97, 18)
        Me.chkPitchAround2.TabIndex = 240
        Me.chkPitchAround2.Text = "Pitch Around 2"
        Me.chkPitchAround2.UseVisualStyleBackColor = True
        '
        'chkPitchAround1
        '
        Me.chkPitchAround1.AutoSize = True
        Me.chkPitchAround1.Location = New System.Drawing.Point(9, 18)
        Me.chkPitchAround1.Name = "chkPitchAround1"
        Me.chkPitchAround1.Size = New System.Drawing.Size(97, 18)
        Me.chkPitchAround1.TabIndex = 2
        Me.chkPitchAround1.Text = "Pitch Around 1"
        Me.chkPitchAround1.UseVisualStyleBackColor = True
        '
        'chkGuardLine
        '
        Me.chkGuardLine.AutoSize = True
        Me.chkGuardLine.Location = New System.Drawing.Point(113, 41)
        Me.chkGuardLine.Name = "chkGuardLine"
        Me.chkGuardLine.Size = New System.Drawing.Size(79, 18)
        Me.chkGuardLine.TabIndex = 1
        Me.chkGuardLine.Text = "Guard Line"
        Me.chkGuardLine.UseVisualStyleBackColor = True
        '
        'chkCornersIn
        '
        Me.chkCornersIn.AutoSize = True
        Me.chkCornersIn.Location = New System.Drawing.Point(113, 7)
        Me.chkCornersIn.Name = "chkCornersIn"
        Me.chkCornersIn.Size = New System.Drawing.Size(76, 18)
        Me.chkCornersIn.TabIndex = 0
        Me.chkCornersIn.Text = "Corners In"
        Me.chkCornersIn.UseVisualStyleBackColor = True
        '
        'dgFAC
        '
        Me.dgFAC.AllowUserToAddRows = False
        Me.dgFAC.AllowUserToDeleteRows = False
        Me.dgFAC.AllowUserToResizeColumns = False
        Me.dgFAC.AllowUserToResizeRows = False
        Me.dgFAC.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgFAC.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgFAC.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.dgFAC.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgFAC.GridColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgFAC.Location = New System.Drawing.Point(380, 524)
        Me.dgFAC.Name = "dgFAC"
        Me.dgFAC.ReadOnly = True
        Me.dgFAC.Size = New System.Drawing.Size(222, 113)
        Me.dgFAC.TabIndex = 328
        '
        'lblHBatBorn
        '
        Me.lblHBatBorn.AutoSize = True
        Me.lblHBatBorn.ForeColor = System.Drawing.Color.Yellow
        Me.lblHBatBorn.Location = New System.Drawing.Point(182, 606)
        Me.lblHBatBorn.Name = "lblHBatBorn"
        Me.lblHBatBorn.Size = New System.Drawing.Size(30, 14)
        Me.lblHBatBorn.TabIndex = 329
        Me.lblHBatBorn.Text = "Born"
        '
        'lblHBatAge
        '
        Me.lblHBatAge.AutoSize = True
        Me.lblHBatAge.ForeColor = System.Drawing.Color.Yellow
        Me.lblHBatAge.Location = New System.Drawing.Point(182, 620)
        Me.lblHBatAge.Name = "lblHBatAge"
        Me.lblHBatAge.Size = New System.Drawing.Size(27, 14)
        Me.lblHBatAge.TabIndex = 330
        Me.lblHBatAge.Text = "Age"
        '
        'lblVBatBorn
        '
        Me.lblVBatBorn.AutoSize = True
        Me.lblVBatBorn.ForeColor = System.Drawing.Color.Yellow
        Me.lblVBatBorn.Location = New System.Drawing.Point(624, 606)
        Me.lblVBatBorn.Name = "lblVBatBorn"
        Me.lblVBatBorn.Size = New System.Drawing.Size(30, 14)
        Me.lblVBatBorn.TabIndex = 331
        Me.lblVBatBorn.Text = "Born"
        '
        'lblVBatAge
        '
        Me.lblVBatAge.AutoSize = True
        Me.lblVBatAge.ForeColor = System.Drawing.Color.Yellow
        Me.lblVBatAge.Location = New System.Drawing.Point(624, 620)
        Me.lblVBatAge.Name = "lblVBatAge"
        Me.lblVBatAge.Size = New System.Drawing.Size(27, 14)
        Me.lblVBatAge.TabIndex = 332
        Me.lblVBatAge.Text = "Age"
        '
        'lblHPitBorn
        '
        Me.lblHPitBorn.AutoSize = True
        Me.lblHPitBorn.ForeColor = System.Drawing.Color.Yellow
        Me.lblHPitBorn.Location = New System.Drawing.Point(8, 373)
        Me.lblHPitBorn.Name = "lblHPitBorn"
        Me.lblHPitBorn.Size = New System.Drawing.Size(30, 14)
        Me.lblHPitBorn.TabIndex = 333
        Me.lblHPitBorn.Text = "Born"
        '
        'lblVPitBorn
        '
        Me.lblVPitBorn.AutoSize = True
        Me.lblVPitBorn.ForeColor = System.Drawing.Color.Yellow
        Me.lblVPitBorn.Location = New System.Drawing.Point(812, 373)
        Me.lblVPitBorn.Name = "lblVPitBorn"
        Me.lblVPitBorn.Size = New System.Drawing.Size(30, 14)
        Me.lblVPitBorn.TabIndex = 334
        Me.lblVPitBorn.Text = "Born"
        '
        'lblHPitAge
        '
        Me.lblHPitAge.AutoSize = True
        Me.lblHPitAge.ForeColor = System.Drawing.Color.Yellow
        Me.lblHPitAge.Location = New System.Drawing.Point(8, 387)
        Me.lblHPitAge.Name = "lblHPitAge"
        Me.lblHPitAge.Size = New System.Drawing.Size(27, 14)
        Me.lblHPitAge.TabIndex = 335
        Me.lblHPitAge.Text = "Age"
        '
        'lblVPitAge
        '
        Me.lblVPitAge.AutoSize = True
        Me.lblVPitAge.ForeColor = System.Drawing.Color.Yellow
        Me.lblVPitAge.Location = New System.Drawing.Point(812, 387)
        Me.lblVPitAge.Name = "lblVPitAge"
        Me.lblVPitAge.Size = New System.Drawing.Size(27, 14)
        Me.lblVPitAge.TabIndex = 336
        Me.lblVPitAge.Text = "Age"
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(987, 713)
        Me.ControlBox = False
        Me.Controls.Add(Me.lblVPitAge)
        Me.Controls.Add(Me.lblHPitAge)
        Me.Controls.Add(Me.lblVPitBorn)
        Me.Controls.Add(Me.lblHPitBorn)
        Me.Controls.Add(Me.lblVBatAge)
        Me.Controls.Add(Me.lblVBatBorn)
        Me.Controls.Add(Me.lblHBatAge)
        Me.Controls.Add(Me.lblHBatBorn)
        Me.Controls.Add(Me.dgFAC)
        Me.Controls.Add(Me.fraDefOptions)
        Me.Controls.Add(Me.lblVisitorMgr)
        Me.Controls.Add(Me.lblHomeMgr)
        Me.Controls.Add(Me.lblVPStats)
        Me.Controls.Add(Me.lblHPStats)
        Me.Controls.Add(Me.cmdStats)
        Me.Controls.Add(Me.cmdLineupChange)
        Me.Controls.Add(Me.prgVisitorPR)
        Me.Controls.Add(Me.prgHomePR)
        Me.Controls.Add(Me.fraOffOptions)
        Me.Controls.Add(Me.picVisitor)
        Me.Controls.Add(Me.picHome)
        Me.Controls.Add(Me.grpPitcherVisitor)
        Me.Controls.Add(Me.grpPitcherHome)
        Me.Controls.Add(Me.grpBatterVisitor)
        Me.Controls.Add(Me.grpBatterHome)
        Me.Controls.Add(Me.cmdEndGame)
        Me.Controls.Add(Me.cmdTime)
        Me.Controls.Add(Me.Text4)
        Me.Controls.Add(Me.lblStats)
        Me.Controls.Add(Me.lblHome)
        Me.Controls.Add(Me.lblVisitor)
        Me.Controls.Add(Me.lblVisitorPR)
        Me.Controls.Add(Me.Label41)
        Me.Controls.Add(Me.lblHomePR)
        Me.Controls.Add(Me.Label40)
        Me.Controls.Add(Me.lblSecond)
        Me.Controls.Add(Me.lblFirst)
        Me.Controls.Add(Me.lblThird)
        Me.Controls.Add(Me.lblOuts)
        Me.Controls.Add(Me.Label36)
        Me.Controls.Add(Me.lblMessage)
        Me.Controls.Add(Me.Label34)
        Me.Controls.Add(Me.Label33)
        Me.Controls.Add(Me.Label32)
        Me.Controls.Add(Me.Label31)
        Me.Controls.Add(Me.Label30)
        Me.Controls.Add(Me.Label29)
        Me.Controls.Add(Me.Label28)
        Me.Controls.Add(Me.Label27)
        Me.Controls.Add(Me.Label26)
        Me.Controls.Add(Me.Label25)
        Me.Controls.Add(Me.Label24)
        Me.Controls.Add(Me.Label23)
        Me.Controls.Add(Me.Label22)
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmMain"
        Me.Text = "Play Ball!"
        Me.fraOffOptions.ResumeLayout(False)
        Me.fraOffOptions.PerformLayout()
        CType(Me.picVisitor, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picHome, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraDefOptions.ResumeLayout(False)
        Me.fraDefOptions.PerformLayout()
        CType(Me.dgFAC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
#Region "Upgrade Support "
    Private Shared m_vb6FormDefInstance As frmMain
    Private Shared m_InitializingDefInstance As Boolean
    Public Shared Property DefInstance() As frmMain
        Get
            If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
                m_InitializingDefInstance = True
                m_vb6FormDefInstance = New frmMain()
                m_InitializingDefInstance = False
            End If
            DefInstance = m_vb6FormDefInstance
        End Get
        Set(ByVal value As frmMain)
            m_vb6FormDefInstance = value
        End Set
    End Property
#End Region

    Private Sub cmdEndGame_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdEndGame.Click
        If MsgBox("Are you sure?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            EndGame()
        End If
     End Sub

    Private Sub cmdLineupChange_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdLineupChange.Click
        Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        HandleLineupChange()
        UpdateStatLine()
        Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub cmdPitch_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPitch.Click
        Dim playResult As String
        Dim isBaseAdvanceOpportunity As Boolean
        Dim isBaseAdvanceOpportunityFO As Boolean
        Dim switchingSides As Boolean
        'Dim ElapsedTime As TimeSpan

        Try
            'ElapsedTime = DateTime.Now.Subtract(lastInvoked)
            'If ElapsedTime.Milliseconds < 50 Then
            '    Exit Sub
            'End If
            'lastInvoked = Now
            If bolInPlay Then
                bolInPlay = False
                Exit Sub
            Else
                bolInPlay = True
            End If
            gblDescBatter = Game.BTeam.GetBatterPtr(Game.currentBatter).player
            If Game.PTeam.pitchChange Then
                Game.PTeam.pitchChange = False
                Call MsgBox("Must replace pitcher.", MsgBoxStyle.OkOnly)
                Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
                Call PitchingChange()
                Me.Cursor = System.Windows.Forms.Cursors.Default
                bolInPlay = False
                Exit Sub
            End If
            'Verify positions at the beginning of each half inning
            If Not ValidPositions(Game.PTeam) Then
                'Force lineup change
                Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
                HandleLineupChange()
                Me.Cursor = System.Windows.Forms.Cursors.Default
                bolInPlay = False
                Exit Sub
            ElseIf Not bolAmericanLeagueRules Then
                If Not Game.PTeam.InLineup(Game.PTeam.hitters + Game.PTeam.nthPitcherUsed) Then
                    'Force lineup change
                    Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
                    HandleLineupChange()
                    Me.Cursor = System.Windows.Forms.Cursors.Default
                    bolInPlay = False
                    Exit Sub
                End If
            End If
            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
            System.Diagnostics.Debug.WriteLine("*****NEW PLAY********")
            bolNormal = optNormal.Checked
            bolHitRun = optHitandRun.Checked
            bolSteal = optSteal.Checked
            gbolAutoSteal = True
            bolSac = optSac.Checked
            bolBunt = optBunt.Checked
            bolSqueeze = optSqueeze.Checked
            bolIntentionalWalk = optIntentionalWalk.Checked
            bolInfieldIn = chkInfieldIn.CheckState = 1
            bolCornersIn = chkCornersIn.CheckState = 1
            bolGuardLine = chkGuardLine.CheckState = 1
            bolPitchAround1 = chkPitchAround1.CheckState = 1
            bolPitchAround2 = chkPitchAround2.CheckState = 1

            bolDNS = chkDNS.CheckState = 1
            Me.lblMessage.Text = ""
            lstFacView.Clear()
            playResult = GetResult(isBaseAdvanceOpportunity, isBaseAdvanceOpportunityFO)
            If isBaseAdvanceOpportunity And playResult.IndexOf(conNoPlay) = -1 Then
                Me.AdvanceRunner()
            End If
            If isBaseAdvanceOpportunityFO And playResult.IndexOf(conNoPlay) = -1 Then
                Me.AdvanceRunnerOnFlyout()
            End If

            'Handle ERR situation
            If playResult.IndexOf(conNoPlay) > -1 Then
                'Normal action
                bolNormal = True
                bolHitRun = False
                If bolSteal Then
                    'Runner cannot get jump when stealing. Turn off autosteal
                    gbolAutoSteal = False
                End If
                bolSteal = False
                bolSac = False
                bolBunt = False
                bolSqueeze = False
                While playResult.IndexOf(conNoPlay) > -1
                    playResult = GetResult(isBaseAdvanceOpportunity, isBaseAdvanceOpportunityFO)
                    If isBaseAdvanceOpportunity And playResult.IndexOf(conNoPlay) = -1 Then
                        Me.AdvanceRunner()
                    End If
                    If isBaseAdvanceOpportunityFO And playResult.IndexOf(conNoPlay) = -1 Then
                        Me.AdvanceRunnerOnFlyout()
                    End If
                End While
            End If
            gblUpdateMsg = ""
            UpdateGame(switchingSides)
            UpdateBoard(switchingSides)
            Game.GetResultPtr.description += gblUpdateMsg
            If Game.homeTeamBatting Then
                Call DispResult(Game.GetResultPtr.description, Game.BTeam, Game.PTeam)
            Else
                Call DispResult(Game.GetResultPtr.description, Game.PTeam, Game.BTeam)
            End If
            'reset options
            optNormal.Checked = True
            optHitandRun.Checked = False
            optHitandRun.Enabled = BaseSituation() = "001" Or BaseSituation() = "101"
            optSteal.Checked = False
            optSteal.Enabled = BaseSituation() <> "000"
            optSac.Checked = False
            optSac.Enabled = Not (BaseSituation() = "000" Or ThirdBase.occupied)
            optBunt.Checked = False
            optBunt.Enabled = BaseSituation() = "000"
            optSqueeze.Checked = False
            optSqueeze.Enabled = ThirdBase.occupied
            chkInfieldIn.CheckState = CheckState.Unchecked
            chkInfieldIn.Enabled = Game.outs < 2
            chkCornersIn.CheckState = CheckState.Unchecked
            chkCornersIn.Enabled = Game.outs < 2
            chkGuardLine.CheckState = CheckState.Unchecked
            chkPitchAround1.CheckState = CheckState.Unchecked
            chkPitchAround2.CheckState = CheckState.Unchecked
            chkDNS.CheckState = System.Windows.Forms.CheckState.Unchecked
            Me.Cursor = System.Windows.Forms.Cursors.Default
            If bolFinal Then
                bolStartGame = False
                'Handle Win/Loss/Save
                Call Game.CreditWinLossSave()
                Call MsgBox("Game Over.", MsgBoxStyle.OkOnly)
                DisplayDecisions()
            End If
            bolInPlay = False
            'Write Code for DP stats and IP correction
        Catch ex As Exception
            Call MsgBox("Error in cmdPitch_Click " & ex.ToString, MsgBoxStyle.OkOnly)
            'Finally
            '    lastInvoked = Now
        End Try
    End Sub

    Private Sub cmdStats_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdStats.Click
        Dim frmLaunch As frmStats

        frmLaunch = New frmStats
        frmLaunch.Show()
    End Sub

    Private Sub cmdTime_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdTime.Click
        Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Call PitchingChange()
        Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub frmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dim height As Integer
        Dim width As Integer
        Dim attendance As Int64 = 0
        Dim logofile As String
        Dim isPOE As Boolean

        Try
            isPOE = AppSettings("PointsOfEffectiveness").ToString.ToUpper = "ON"
            colLSVisitorTextBoxes = New clsTextBoxArray(Me)
            colLSHomeTextBoxes = New clsTextBoxArray(Me)
            For i As Integer = 1 To 13
                colLSVisitorTextBoxes.AddNewTextBoxLS(False)
                colLSHomeTextBoxes.AddNewTextBoxLS(True)
            Next i
            colHomeKLabels = New clsLabelArray(Me)
            colVisitorKLabels = New clsLabelArray(Me)
            For i As Integer = 1 To 20
                'Create K labels
                colHomeKLabels.AddNewKLabel(conHome)
                colVisitorKLabels.AddNewKLabel(conVisitor)
            Next i
            colBaseTextBoxes = New clsTextBoxArray(Me)
            For i As Integer = 1 To 3
                'Create bases
                colBaseTextBoxes.AddNewTextBoxBase()
            Next i
            colVisitorBSLabels = New clsLabelArray(Me)
            colHomeBSLabels = New clsLabelArray(Me)
            colVisitorPosLabels = New clsLabelArray(Me)
            colHomePosLabels = New clsLabelArray(Me)
            colVisitorBatterLabels = New clsLabelArray(Me)
            colHomeBatterLabels = New clsLabelArray(Me)
            For i As Integer = 1 To 9
                colVisitorBSLabels.AddNewBatSideLabel(conVisitor)
                colHomeBSLabels.AddNewBatSideLabel(conHome)
                colVisitorPosLabels.AddNewPosLabel(conVisitor)
                colHomePosLabels.AddNewPosLabel(conHome)
                colVisitorBatterLabels.AddNewBatterLabel(conVisitor)
                colHomeBatterLabels.AddNewBatterLabel(conHome)
            Next i
            Call LoadLineups(Home.GetPlayerNum(1), conHome, Home)
            Call LoadLineups(Visitor.GetPlayerNum(1), conVisitor, Visitor)
            Call LoadPitcher(Home, conHome)
            Call LoadPitcher(Visitor, conVisitor)
            NewGame() 'New game
            lblVisitor.Text = Visitor.teamName
            lblHome.Text = Home.teamName
            Me.Text = Home.GetStadiumStats(attendance) & "  A-" & attendance.ToString
            Me.lblHomeMgr.Text = "Mgr " & Home.GetManager()
            Me.lblVisitorMgr.Text = "Mgr " & Visitor.GetManager()
            Home.SetVarianceAdjustments()
            Visitor.SetVarianceAdjustments()
            Call GetLogoDimensions((Visitor.teamName), height, width)
            'picVisitor.Height = VB6.TwipsToPixelsY(height)
            'picVisitor.Width = VB6.TwipsToPixelsX(width)
            TwipsToPixels(width, height, CDbl(picVisitor.Width), CDbl(picVisitor.Height))
            logofile = Visitor.GetLogo
            If Dir(gAppPath & "graphics\" & logofile & ".jpg") <> Nothing Then
                picVisitor.Image = System.Drawing.Image.FromFile(gAppPath & "graphics\" & logofile & ".jpg")
            Else
                picVisitor.Image = System.Drawing.Image.FromFile(gAppPath & "graphics\" & logofile & ".gif")
            End If
            Call GetLogoDimensions((Home.teamName), height, width)

            TwipsToPixels(width, height, CDbl(picHome.Width), CDbl(picHome.Height))
            logofile = Home.GetLogo
            If Dir(gAppPath & "graphics\" & logofile & ".jpg") <> Nothing Then
                picHome.Image = System.Drawing.Image.FromFile(gAppPath & "graphics\" & logofile & ".jpg")
            Else
                picHome.Image = System.Drawing.Image.FromFile(gAppPath & "graphics\" & logofile & ".gif")
            End If
            optHitandRun.Enabled = False
            prgHomePR.Minimum = 0
            prgVisitorPR.Minimum = 0
            If isPOE Then
                prgHomePR.Maximum = Home.GetPitcherPtr(Home.pitcherSel).DeterminePOE(False, True)
                prgVisitorPR.Maximum = Visitor.GetPitcherPtr(Visitor.pitcherSel).DeterminePOE(False, True)
            Else
                prgHomePR.Maximum = CInt(Home.GetPitcherPtr(Home.pitcherSel).sr)
                prgVisitorPR.Maximum = CInt(Visitor.GetPitcherPtr(Visitor.pitcherSel).sr)
            End If
            prgHomePR.Value = prgHomePR.Maximum
            prgVisitorPR.Value = prgVisitorPR.Maximum
            lblHomePR.Text = prgHomePR.Maximum.ToString
            lblVisitorPR.Text = prgVisitorPR.Maximum.ToString
            Home.pr = prgHomePR.Maximum
            Visitor.pr = prgVisitorPR.Maximum
            grpBatterVisitor.BackColor = System.Drawing.Color.Cyan
            grpPitcherHome.BackColor = System.Drawing.Color.Cyan
            UpdateStatLine()
            'Call PlaySound(App.Path & "\playball.wav")
        Catch ex As Exception
            Call MsgBox("frmMain_Load " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' load batting cards
    ''' </summary>
    ''' <param name="currentHitterIndex"></param>
    ''' <param name="battingTeamIndicator">0=home 1=visitor</param>
    ''' <param name="team"></param>
    ''' <remarks></remarks>
    Private Sub LoadLineups(ByRef currentHitterIndex As Integer, ByRef battingTeamIndicator As Integer, _
                                ByRef team As clsTeam)
        Dim grpBatter As BCard

        Try
            With team.GetBatterPtr(currentHitterIndex)
                If battingTeamIndicator = conHome Then
                    grpBatter = grpBatterHome
                    lblHBatBorn.Text = "Born: " & .city & ", " & .state
                    lblHBatAge.Text = "Age: " & .age
                Else
                    grpBatter = grpBatterVisitor
                    lblVBatBorn.Text = "Born: " & .city & ", " & .state
                    lblVBatAge.Text = "Age: " & .age
                End If

                If .player = Nothing Then
                    grpBatter.Visible = False
                Else
                    grpBatter.Visible = True
                    grpBatter.grpBCard.Text = .player
                    grpBatter.lblField.Text = .field
                    grpBatter.lblOBR.Text = .obr
                    grpBatter.lblSP.Text = .sp
                    grpBatter.lblHitandRun.Text = .hitRun
                    grpBatter.lblCD.Text = .cd
                    grpBatter.lblSAC.Text = .sac
                    grpBatter.lblInj.Text = .inj
                    grpBatter.lbl1bf.Text = .hit1Bf
                    grpBatter.lbl1b7.Text = .hit1B7
                    grpBatter.lbl1b8.Text = .hit1B8
                    grpBatter.lbl1b9.Text = .hit1B9
                    grpBatter.lbl2b7.Text = .hit2B7
                    grpBatter.lbl2b8.Text = .hit2B8
                    grpBatter.lbl2b9.Text = .hit2B9
                    grpBatter.lbl3b8.Text = .hit3B8
                    grpBatter.lblHR.Text = .hr
                    grpBatter.lblK.Text = .k
                    grpBatter.lblW.Text = .w
                    grpBatter.lblHPB.Text = .hpb
                    grpBatter.lblOut.Text = .out
                    grpBatter.lblCht.Text = .cht
                    grpBatter.lblBD.Text = .bd
                End If
            End With
            For i As Integer = 1 To 9
                If battingTeamIndicator = conHome Then
                    If Home.GetPlayerNum(i) > 0 Then
                        colHomeBatterLabels.Item(i - 1).Text = Home.GetBatterPtr(Home.GetPlayerNum(i)).player
                        colHomeBSLabels.Item(i - 1).Text = Home.GetBatterPtr(Home.GetPlayerNum(i)).cht.Substring(0, 1).ToUpper
                        colHomePosLabels.Item(i - 1).Text = Home.GetBatterPtr(Home.GetPlayerNum(i)).position
                    Else
                        colHomeBatterLabels.Item(i - 1).Text = ""
                        colHomeBSLabels.Item(i - 1).Text = ""
                    End If
                Else
                    If Visitor.GetPlayerNum(i) > 0 Then
                        colVisitorBatterLabels.Item(i - 1).Text = Visitor.GetBatterPtr(Visitor.GetPlayerNum(i)).player
                        colVisitorBSLabels.Item(i - 1).Text = Visitor.GetBatterPtr(Visitor.GetPlayerNum(i)).cht.Substring(0, 1).ToUpper
                        colVisitorPosLabels.Item(i - 1).Text = Visitor.GetBatterPtr(Visitor.GetPlayerNum(i)).position
                    Else
                        colVisitorBatterLabels.Item(i - 1).Text = ""
                        colVisitorBSLabels.Item(i - 1).Text = ""
                    End If
                End If
            Next i
        Catch ex As Exception
            Call MsgBox("LoadLineups " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' loads the pitching cards
    ''' </summary>
    ''' <param name="team"></param>
    ''' <param name="homeTeamIndicator">0=home 1=visitor</param>
    ''' <remarks></remarks>
    Private Sub LoadPitcher(ByRef team As clsTeam, ByVal homeTeamIndicator As Integer)
        Dim grpPitcher As PCard

        Try
            With team.GetPitcherPtr(team.pitcherSel)
                If homeTeamIndicator = conHome Then
                    grpPitcher = grpPitcherHome
                    lblHPitBorn.Text = "Born: " & .city & ", " & .state
                    lblHPitAge.Text = "Age: " & .age
                Else
                    grpPitcher = grpPitcherVisitor
                    lblVPitBorn.Text = "Born: " & .city & ", " & .state
                    lblVPitAge.Text = "Age: " & .age
                End If

                If .player = Nothing Then
                    grpPitcher.Visible = False
                Else
                    grpPitcher.Visible = True
                    grpPitcher.grpPCard.Text = .player
                    grpPitcher.lblThrows.Text = "Field: " & .field & Space(2) & .throwField
                    grpPitcher.lblPB.Text = .pbRange
                    grpPitcher.lblSR.Text = .sr
                    grpPitcher.lblRR.Text = .rr
                    grpPitcher.lbl1bf.Text = .hit1Bf
                    grpPitcher.lbl1b7.Text = .hit1B7
                    grpPitcher.lbl1b8.Text = .hit1B8
                    grpPitcher.lbl1b9.Text = .hit1B9
                    grpPitcher.lblBK.Text = .bk
                    grpPitcher.lblK.Text = .k
                    grpPitcher.lblW.Text = .w
                    grpPitcher.lblPBall.Text = .pb
                    grpPitcher.lblWP.Text = .wp
                    grpPitcher.lblOut.Text = .out
                    grpPitcher.lblStartRel.Text = .StoR

                End If
            End With
        Catch ex As Exception
            Call MsgBox("LoadPitcher " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' displays result of play on main board
    ''' </summary>
    ''' <param name="playResult"></param>
    ''' <param name="HTeam"></param>
    ''' <param name="VTeam"></param>
    ''' <remarks></remarks>
    Public Sub DispResult(ByRef playResult As String, ByRef HTeam As clsTeam, ByRef VTeam As clsTeam)
        Dim lineScoreIndex As Integer 'The current inning to show active
        Dim filLog As Integer
        Dim bs As New BindingSource

        Try
            filLog = FreeFile()
            FileOpen(filLog, gAppPath & "desclog.txt", OpenMode.Append)
            For i As Integer = 1 To Game.inning
                If Game.inning <= 10 Then
                    colLSVisitorTextBoxes.Item(i - 1).Text = Game.GetLineScore(i, 1).ToString
                    colLSHomeTextBoxes.Item(i - 1).Text = Game.GetLineScore(i, 2).ToString
                    lineScoreIndex = i - 1
                    colLSVisitorTextBoxes.Item(i - 1).BackColor = System.Drawing.Color.White
                    colLSHomeTextBoxes.Item(i - 1).BackColor = System.Drawing.Color.White
                ElseIf i = Game.inning Then
                    'Display current extra inning only
                    colLSVisitorTextBoxes.Item(9).Text = Game.GetLineScore(i, 1).ToString
                    colLSHomeTextBoxes.Item(9).Text = Game.GetLineScore(i, 2).ToString
                    lineScoreIndex = 9
                    colLSVisitorTextBoxes.Item(9).BackColor = System.Drawing.Color.White
                    colLSHomeTextBoxes.Item(9).BackColor = System.Drawing.Color.White
                End If
            Next i
            If Game.homeTeamBatting Then
                colLSHomeTextBoxes.Item(lineScoreIndex).BackColor = System.Drawing.Color.Yellow
            Else
                colLSVisitorTextBoxes.Item(lineScoreIndex).BackColor = System.Drawing.Color.Yellow
            End If
            colLSVisitorTextBoxes.Item(10).Text = VTeam.runs.ToString
            colLSVisitorTextBoxes.Item(11).Text = VTeam.hits.ToString
            colLSVisitorTextBoxes.Item(12).Text = VTeam.errors.ToString
            colLSHomeTextBoxes.Item(10).Text = HTeam.runs.ToString
            colLSHomeTextBoxes.Item(11).Text = HTeam.hits.ToString
            colLSHomeTextBoxes.Item(12).Text = HTeam.errors.ToString
            If Game.lastBatterHome Then
                Call UpdateLineups(HTeam)
            Else
                Call UpdateLineups(VTeam)
            End If
            colBaseTextBoxes.Item(0).BackColor = IIF(FirstBase.occupied, System.Drawing.Color.Red, _
                                                System.Drawing.Color.White)
            colBaseTextBoxes.Item(1).BackColor = IIF(SecondBase.occupied, System.Drawing.Color.Red, _
                                            System.Drawing.Color.White)
            colBaseTextBoxes.Item(2).BackColor = IIF(ThirdBase.occupied, System.Drawing.Color.Red, _
                                            System.Drawing.Color.White)
            lblOuts.Text = Game.outs.ToString
            PrintLine(filLog, gblDescBatter & ": " & playResult)
            lblMessage.Text = playResult
            'Select Case Game.BTeam.GetBatterPtr(Game.currentBatter).cht.Substring(0, 1)
            '    Case "L"
            '        If Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).throwField = "Right" Then
            '            lblAdvantageLR.Text = "Batter"
            '        Else
            '            lblAdvantageLR.Text = "Pitcher"
            '        End If
            '    Case "R"
            '        If Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).throwField = "Right" Then
            '            lblAdvantageLR.Text = "Pitcher"
            '        Else
            '            lblAdvantageLR.Text = "Batter"
            '        End If
            '    Case "S"
            '        lblAdvantageLR.Text = "None"
            'End Select

            'Update PR display
            lblVisitorPR.Text = Visitor.pr.ToString
            lblHomePR.Text = Home.pr.ToString
            prgHomePR.Value = Home.pr
            prgVisitorPR.Value = Visitor.pr
            If Visitor.GetPitcherPtr(Visitor.pitcherSel).PitStatPtr.k < 21 Then
                For i As Integer = 0 To Visitor.GetPitcherPtr(Visitor.pitcherSel).PitStatPtr.k - 1
                    colVisitorKLabels.Item(i).Visible = True
                Next i
            End If
            If Home.GetPitcherPtr(Home.pitcherSel).PitStatPtr.k < 21 Then
                For i As Integer = 0 To Home.GetPitcherPtr(Home.pitcherSel).PitStatPtr.k - 1
                    colHomeKLabels.Item(i).Visible = True
                Next i
            End If
            'Update the stat line
            If gbolSeason Then
                UpdateStatLine()
            End If
            FileClose(filLog)
            bs.DataSource = lstFacView
            dgFAC.AutoGenerateColumns = True
            dgFAC.DataSource = bs
        Catch ex As Exception
            Call MsgBox("DispResult " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' updates current hitter card after the last play
    ''' </summary>
    ''' <param name="team"></param>
    ''' <remarks></remarks>
    Private Sub UpdateLineups(ByRef team As clsTeam)
        Dim homeIndicator As Integer
        Dim currentHitter As Integer
        Dim grpBatter As BCard

        Try
            If Game.lastBatterHome Then
                homeIndicator = conHome
            Else
                homeIndicator = conVisitor
            End If

            currentHitter = team.GetPlayerNum(team.order)
            With team.GetBatterPtr(currentHitter)
                If homeIndicator = conHome Then
                    grpBatter = grpBatterHome
                    lblHBatBorn.Text = "Born: " & .city & ", " & .state
                    lblHBatAge.Text = "Age: " & .age
                Else
                    grpBatter = grpBatterVisitor
                    lblVBatBorn.Text = "Born: " & .city & ", " & .state
                    lblVBatAge.Text = "Age: " & .age
                End If

                If .player = Nothing Then
                    grpBatter.Visible = False
                Else
                    grpBatter.Visible = True
                    grpBatter.grpBCard.Text = .player
                    grpBatter.lblField.Text = .field
                    grpBatter.lblOBR.Text = .obr
                    grpBatter.lblSP.Text = .sp
                    grpBatter.lblHitandRun.Text = .hitRun
                    grpBatter.lblCD.Text = .cd
                    grpBatter.lblSAC.Text = .sac
                    grpBatter.lblInj.Text = .inj
                    grpBatter.lbl1bf.Text = .hit1Bf
                    grpBatter.lbl1b7.Text = .hit1B7
                    grpBatter.lbl1b8.Text = .hit1B8
                    grpBatter.lbl1b9.Text = .hit1B9
                    grpBatter.lbl2b7.Text = .hit2B7
                    grpBatter.lbl2b8.Text = .hit2B8
                    grpBatter.lbl2b9.Text = .hit2B9
                    grpBatter.lbl3b8.Text = .hit3B8
                    grpBatter.lblHR.Text = .hr
                    grpBatter.lblK.Text = .k
                    grpBatter.lblW.Text = .w
                    grpBatter.lblHPB.Text = .hpb
                    grpBatter.lblOut.Text = .out
                    grpBatter.lblCht.Text = .cht
                    grpBatter.lblBD.Text = .bd
                End If
            End With
        Catch ex As Exception
            Call MsgBox("UpdateLineups " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' Updates display after a lineup change
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub HandleLineupChange()
        Dim frmLaunch As frmLineupChange

        Try
            frmLaunch = New frmLineupChange
            frmLaunch.ShowDialog()
            Call LoadLineups(Home.GetPlayerNum(Home.order), conHome, Home)
            Call LoadLineups(Visitor.GetPlayerNum(Visitor.order), conVisitor, Visitor)
            Game.currentBatter = Game.BTeam.GetPlayerNum(Game.BTeam.order)
            'Select Case Game.BTeam.GetBatterPtr(Game.currentBatter).cht.Substring(0, 1)
            '    Case "L"
            '        If Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).throwField = "Right" Then
            '            lblAdvantageLR.Text = "Batter"
            '        Else
            '            lblAdvantageLR.Text = "Pitcher"
            '        End If
            '    Case "R"
            '        If Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).throwField = "Right" Then
            '            lblAdvantageLR.Text = "Pitcher"
            '        Else
            '            lblAdvantageLR.Text = "Batter"
            '        End If
            '    Case "S"
            '        lblAdvantageLR.Text = "None"
            'End Select

            'Update runners
            If FirstBase.occupied Then
                With Game.BTeam.GetBatterPtr(FirstBase.runner)
                    Me.lblFirst.Text = GetLastName(.player) & Space(1) & .obr & Space(1) & .sp
                End With
            End If
            If SecondBase.occupied Then
                With Game.BTeam.GetBatterPtr(SecondBase.runner)
                    Me.lblSecond.Text = GetLastName(.player) & Space(1) & .obr & Space(1) & .sp
                End With
            End If
            If ThirdBase.occupied Then
                With Game.BTeam.GetBatterPtr(ThirdBase.runner)
                    Me.lblThird.Text = GetLastName(.player) & Space(1) & .obr & Space(1) & .sp
                End With
            End If
        Catch ex As Exception
            Call MsgBox("HandleLineupChange " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' Updates current pitcher display after a pitching change
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub PitchingChange()
        Dim pitcherIndex As Integer
        Dim frmLaunch As frmPitching
        Dim isPOE As Boolean

        Try
            isPOE = AppSettings("PointsOfEffectiveness").ToString.ToUpper = "ON"
            pitcherIndex = IIF(Game.homeTeamBatting, Visitor.pitcherSel, Home.pitcherSel)
            frmLaunch = New frmPitching
            frmLaunch.ShowDialog()
            If pitcherIndex <> IIF(Game.homeTeamBatting, Visitor.pitcherSel, Home.pitcherSel) Then
                'It's a new pitcher
                If Game.homeTeamBatting Then
                    Visitor.pitchChange = False
                Else
                    Home.pitchChange = False
                End If
                If Game.homeTeamBatting Then
                    If isPOE Then
                        prgVisitorPR.Maximum = Visitor.GetPitcherPtr(Visitor.pitcherSel).DeterminePOE(True, True)
                    Else
                        prgVisitorPR.Maximum = CInt(Visitor.GetPitcherPtr(Visitor.pitcherSel).rr)
                    End If
                    prgVisitorPR.Value = prgVisitorPR.Maximum
                    lblVisitorPR.Text = prgVisitorPR.Maximum.ToString
                    Visitor.pr = CType(prgVisitorPR.Maximum, Integer)
                    Call LoadPitcher(Visitor, conVisitor)
                    'Reset K counter
                    For i As Integer = 0 To 19
                        colVisitorKLabels.Item(i).Visible = False
                    Next i
                Else
                    If isPOE Then
                        prgHomePR.Maximum = Home.GetPitcherPtr(Home.pitcherSel).DeterminePOE(True, True)
                    Else
                        prgHomePR.Maximum = CInt(Home.GetPitcherPtr(Home.pitcherSel).rr)
                    End If
                    prgHomePR.Value = prgHomePR.Maximum
                    lblHomePR.Text = prgHomePR.Maximum.ToString
                    Home.pr = CType(prgHomePR.Maximum, Integer)
                    Call LoadPitcher(Home, conHome)
                    'Reset K counter
                    For i As Integer = 0 To 19
                        colHomeKLabels.Item(i).Visible = False
                    Next i
                End If
                'Select Case Game.BTeam.GetBatterPtr(Game.currentBatter).cht.Substring(0, 1)
                '    Case "L"
                '        If Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).throwField = "Right" Then
                '            lblAdvantageLR.Text = "Batter"
                '        Else
                '            lblAdvantageLR.Text = "Pitcher"
                '        End If
                '    Case "R"
                '        If Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).throwField = "Right" Then
                '            lblAdvantageLR.Text = "Pitcher"
                '        Else
                '            lblAdvantageLR.Text = "Batter"
                '        End If
                '    Case "S"
                '        lblAdvantageLR.Text = "None"
                'End Select
                UpdateStatLine()
            End If
        Catch ex As Exception
            Call MsgBox("PitchingChange " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' calls all routines necessary at the end of the game, including printing the box score, saving the stats
    ''' and updating the season
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub EndGame()
        Dim filBox As Integer
        Dim fileName As String
        Dim frmLaunch As frmStartup
        Dim fullPath As String

        Try
            If gbolSeason And bolFinal Then
                filBox = FreeFile()
                If gbolPostSeason Then
                    fileName = gstrPostSeason & "_Game_" & Season.game
                Else
                    fileName = Season.seriesRow & "X" & Season.game
                End If
                If Dir(GetAppLocation() & "\archive", vbDirectory) = Nothing Then
                    My.Computer.FileSystem.CreateDirectory(GetAppLocation() & "\archive")
                End If
                If Dir(GetAppLocation() & "\archive\" & gstrSeason, vbDirectory) = Nothing Then
                    My.Computer.FileSystem.CreateDirectory(GetAppLocation() & "\archive\" & gstrSeason)
                End If

                fullPath = gAppPath & "archive\" & gstrSeason & "\" & fileName & ".txt"
                FileOpen(filBox, fullPath, OpenMode.Output)
                Print(filBox, Visitor.teamName & Space(16 - Visitor.teamName.Length))
                For i As Integer = 1 To Game.inning
                    Print(filBox, Game.GetLineScore(i, 1) & Space(2))
                Next i
                PrintLine(filBox, Visitor.runs & Space(2) & Visitor.hits & Space(2) & Visitor.errors)
                Print(filBox, Home.teamName & Space(16 - Home.teamName.Length))
                For i As Integer = 1 To Game.inning
                    Print(filBox, Game.GetLineScore(i, 2) & Space(2))
                Next i
                PrintLine(filBox, Home.runs & Space(2) & Home.hits & Space(2) & Home.errors)
                PrintLine(filBox)
                PrintLine(filBox)
                If MsgBox("Would you like to save this game?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    'Call Season.SaveGame(filBox, fullPath)
                    Call Season.SaveGame(filBox, fileName)
                Else
                    If MsgBox("Are you sure?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                        'Call Season.SaveGame(filBox, fullPath)
                        Call Season.SaveGame(filBox, fileName)
                    End If
                End If
                'If Home.runs > Visitor.runs Then
                '    PrintLine(filBox, "WP - " & Home.GetPitcherPtr(Game.winner).player)
                '    PrintLine(filBox, "LP - " & Visitor.GetPitcherPtr(Game.loser).player)
                '    If Game.save > 0 Then
                '        PrintLine(filBox, "S  - " & Home.GetPitcherPtr(Game.save).player)
                '    End If
                'Else
                PrintLine(filBox, Me.lblMessage.Text)
                'End If
                FileClose(filBox)
            End If
            bolFinal = False
            Me.Close()
            If MsgBox("Would you like to play again?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                frmLaunch = New frmStartup
                frmLaunch.Show()
            Else
                End
            End If
        Catch ex As Exception
            Call MsgBox("EndGame " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' Displays the pitchers of record at the conclusion of a game
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub DisplayDecisions()
        Dim dsStat As DataSet = Nothing
        Dim sqlQuery As String = ""
        Dim WinTeam As clsTeam
        Dim LoseTeam As clsTeam
        Dim winsWP As Integer
        Dim lossesWP As Integer
        Dim winsLP As Integer
        Dim lossesLP As Integer
        Dim saves As Integer
        Dim tableKey As String
        Dim filDebug As Integer
        Dim DataAccess As New clsDataAccess(gstrSeason & "data")

        Try
            If Dir(gAppPath & "debug.bbb") <> Nothing Then
                filDebug = FreeFile()
                FileOpen(filDebug, gAppPath & "debug.out", OpenMode.Append)
                PrintLine(filDebug, "In DisplayDecisions")
                FileClose(filDebug)
            End If

            If Home.runs > Visitor.runs Then
                WinTeam = Home
                LoseTeam = Visitor
            Else
                WinTeam = Visitor
                LoseTeam = Home
            End If

            If Not gbolSeason Then
                Me.lblMessage.Text = "WP - " & WinTeam.GetPitcherPtr(Game.winner).player & _
                            " (1-0)" & vbCrLf & _
                            "LP - " & LoseTeam.GetPitcherPtr(Game.loser).player & _
                            " (0-1)" & vbCrLf
                Exit Sub
            End If

            'If gAccessConnectStr = Nothing Then
            '    gAccessConnectStr = GetConnectString()
            'End If

            tableKey = HandleQuotes(StripChar(WinTeam.GetPitcherPtr(Game.winner).player & _
                                        WinTeam.teamName, " "))
            tableKey = HandleQuotes(tableKey)
            'Using accessConn As OleDb.OleDbConnection = New OleDb.OleDbConnection
            '    accessConn.ConnectionString = GetConnectString()
            '    accessConn.Open()
            dsStat = DataAccess.ExecuteDataSet("SELECT * FROM " & gstrPitchingTable & " WHERE playerid = '" & tableKey & "'")

            If dsStat.Tables(0).Rows.Count = 0 Then
                winsWP = 1
                lossesWP = 0
            Else
                With dsStat.Tables(0).Rows(0)
                    winsWP = CInt(.Item("wins")) + 1
                    lossesWP = CInt(.Item("losses"))
                End With
            End If

            dsStat.Clear()
            tableKey = HandleQuotes(StripChar(LoseTeam.GetPitcherPtr(Game.loser).player & _
                                    LoseTeam.teamName, " "))
            tableKey = HandleQuotes(tableKey)
            dsStat = DataAccess.ExecuteDataSet("SELECT * FROM " & gstrPitchingTable & " WHERE playerid = '" & tableKey & "'")

            If dsStat.Tables(0).Rows.Count = 0 Then
                winsLP = 0
                lossesLP = 1
            Else
                With dsStat.Tables(0).Rows(0)
                    winsLP = CInt(.Item("wins"))
                    lossesLP = CInt(.Item("losses")) + 1
                End With
            End If

            If Game.save > 0 Then
                dsStat.Clear()
                tableKey = StripChar(WinTeam.GetPitcherPtr(Game.save).player & WinTeam.teamName, " ")
                tableKey = HandleQuotes(tableKey)

                sqlQuery = "SELECT * FROM " & gstrPitchingTable & " WHERE playerid = '" & tableKey & "'"
                dsStat = DataAccess.ExecuteDataSet(sqlQuery)

                If dsStat.Tables(0).Rows.Count = 0 Then
                    saves = 1
                Else
                    With dsStat.Tables(0).Rows(0)
                        saves = CInt(.Item("saves")) + 1
                    End With
                End If
            End If
            Me.lblMessage.Text = "WP - " & WinTeam.GetPitcherPtr(Game.winner).player & _
                        " (" & winsWP & "-" & lossesWP & ")" & vbCrLf & _
                        "LP - " & LoseTeam.GetPitcherPtr(Game.loser).player & _
                        " (" & winsLP & "-" & lossesLP & ")" & vbCrLf
            If Game.save > 0 Then
                Me.lblMessage.Text = Me.lblMessage.Text & _
                            "S  - " & WinTeam.GetPitcherPtr(Game.save).player & " (" & saves & ")"
            End If
            'End Using
        Catch ex As Exception
            Call MsgBox("DisplayDecisions " & sqlQuery & " " & ex.ToString, MsgBoxStyle.OkOnly)
        Finally
            If Not dsStat Is Nothing Then dsStat.Dispose()
        End Try
    End Sub

    Private Sub frmMain_Closing(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Dim isCancel As Boolean = eventArgs.Cancel
        If bolFinal Then
            EndGame()
        End If
        eventArgs.Cancel = isCancel
    End Sub

    ''' <summary>
    ''' updates the stat line on the main board
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub UpdateStatLine()
        Dim dsStat As DataSet = Nothing
        Dim drStat As DataRow = Nothing
        Dim sqlQuery As String = ""
        Dim baseHits As Integer
        Dim homeRuns As Integer
        Dim rbis As Integer
        Dim battingAverage As Double
        Dim runs As Integer
        Dim atBats As Integer
        Dim tableKey As String
        Dim tempvalue As String = ""
        Dim filDebug As Integer

        'Pitch stats
        Dim ipOuts As Integer
        Dim earnedRuns As Integer
        Dim pitchRecord As String
        Dim era As Double

        Dim DataAccess As New clsDataAccess(gstrSeason & "data")

        Try
            If Dir(gAppPath & "debug.bbb") <> Nothing Then
                filDebug = FreeFile()
                FileOpen(filDebug, gAppPath & "debug.out", OpenMode.Append)
                PrintLine(filDebug, "In UpdateStatLine")
                FileClose(filDebug)
            End If

            'If gAccessConnectStr = Nothing Then
            '    gAccessConnectStr = GetConnectString()
            'End If
            If Not gbolSeason Then
                If Game.BTeam.GetBatterPtr(Game.currentBatter).BatStatPtr.AB > 0 Then
                    tempvalue = "Today: " & Game.BTeam.GetBatterPtr(Game.currentBatter).BatStatPtr.H & " for " & _
                                Game.BTeam.GetBatterPtr(Game.currentBatter).BatStatPtr.AB
                    era = IIF(ipOuts = 0, CDbl(0), earnedRuns * 27 / ipOuts)
                    tempvalue = "IP: " & Format(Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).PitStatPtr.ip / 3, "0.0") & "  W-L-S:" & _
                                "0-0-0" & "   ERA:" & Format(era, "#.00")
                    If Game.homeTeamBatting Then
                        Me.lblHPStats.Text = tempvalue
                    Else
                        Me.lblVPStats.Text = tempvalue
                    End If
                End If
                Me.lblStats.Text = tempvalue

                Exit Sub
            End If

            'Using accessConn As OleDb.OleDbConnection = New OleDb.OleDbConnection
            '    accessConn.ConnectionString = GetConnectString()
            '    accessConn.Open()
            With Game.BTeam.GetBatterPtr(Game.currentBatter)
                tableKey = StripChar(.player & Game.BTeam.teamName, " ")
                tableKey = HandleQuotes(tableKey)
                sqlQuery = "SELECT * FROM " & gstrHittingTable & " WHERE playerid = '" & tableKey & "'"
                dsStat = DataAccess.ExecuteDataSet(sqlQuery)

                If dsStat.Tables(0).Rows.Count = 0 Then
                    baseHits = .BatStatPtr.H
                    homeRuns = .BatStatPtr.HR
                    rbis = .BatStatPtr.RBI
                    runs = .BatStatPtr.R
                    atBats = .BatStatPtr.AB
                Else
                    drStat = dsStat.Tables(0).Rows(0)
                    baseHits = .BatStatPtr.H + CInt(drStat.Item("Hits"))
                    homeRuns = .BatStatPtr.HR + CInt(drStat.Item("HomeRuns"))
                    rbis = .BatStatPtr.RBI + CInt(drStat.Item("RBI"))
                    runs = .BatStatPtr.R + CInt(drStat.Item("Runs"))
                    atBats = .BatStatPtr.AB + CInt(drStat.Item("AB"))
                End If
                If atBats > 0 Then
                    battingAverage = CDbl(Format(CDbl(baseHits) / CDbl(atBats), ".000").ToString)
                Else
                    battingAverage = CDbl(".000")
                End If
                tempvalue = Space(20) & "AB  Hits  HR  RBI   Avg  Runs" & vbCrLf & "Season: " & _
                            atBats.ToString.PadLeft(9) & baseHits.ToString.PadLeft(6) & _
                            homeRuns.ToString.PadLeft(6) & rbis.ToString.PadLeft(6) & _
                            battingAverage.ToString.PadLeft(9) & runs.ToString.PadLeft(7)
                If .BatStatPtr.AB > 0 Then
                    tempvalue = tempvalue & vbCrLf & "Today: " & .BatStatPtr.H & " for " & _
                                .BatStatPtr.AB
                End If
                Me.lblStats.Text = tempvalue
            End With

            'Update defensive team pitch stats
            With Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel)
                tableKey = StripChar(.player & Game.PTeam.teamName, " ")
                tableKey = HandleQuotes(tableKey)
                sqlQuery = "SELECT * FROM " & gstrPitchingTable & " WHERE playerid = '" & tableKey & "'"
                dsStat.Clear()
                dsStat = DataAccess.ExecuteDataSet(sqlQuery)

                If dsStat.Tables(0).Rows.Count = 0 Then
                    ipOuts = .PitStatPtr.ip
                    earnedRuns = .PitStatPtr.er
                    pitchRecord = "0-0-0"
                Else
                    drStat = dsStat.Tables(0).Rows(0)
                    ipOuts = .PitStatPtr.ip + CInt(drStat.Item("ip"))
                    earnedRuns = .PitStatPtr.er + CInt(drStat.Item("earnedruns"))
                    pitchRecord = drStat.Item("wins").ToString & "-" & _
                                drStat.Item("losses").ToString & "-" & _
                                drStat.Item("saves").ToString
                End If
                era = IIF(ipOuts = 0, CDbl(0), earnedRuns * 27 / ipOuts)
                tempvalue = "IP: " & Format(.PitStatPtr.ip / 3, "0.0") & "  W-L-S:" & _
                            pitchRecord & "   ERA:" & Format(era, "#.00")
                If Game.homeTeamBatting Then
                    Me.lblVPStats.Text = tempvalue
                Else
                    Me.lblHPStats.Text = tempvalue
                End If
            End With

            'Update offensive team pitch stats
            With Game.BTeam.GetPitcherPtr(Game.BTeam.pitcherSel)
                tableKey = StripChar(.player & Game.BTeam.teamName, " ")
                tableKey = HandleQuotes(tableKey)
                sqlQuery = "SELECT * FROM " & gstrPitchingTable & " WHERE playerid = '" & tableKey & "'"
                dsStat.Clear()
                dsStat = DataAccess.ExecuteDataSet(sqlQuery)

                If dsStat.Tables(0).Rows.Count = 0 Then
                    ipOuts = .PitStatPtr.ip
                    earnedRuns = .PitStatPtr.er
                    pitchRecord = "0-0-0"
                Else
                    drStat = dsStat.Tables(0).Rows(0)
                    ipOuts = .PitStatPtr.ip + CInt(drStat.Item("ip"))
                    earnedRuns = .PitStatPtr.er + CInt(drStat.Item("earnedruns"))
                    pitchRecord = drStat.Item("wins").ToString & "-" & _
                                drStat.Item("losses").ToString & "-" & _
                                drStat.Item("saves").ToString
                End If
                era = IIF(ipOuts = 0, CDbl(0), earnedRuns * 27 / ipOuts)
                tempvalue = "IP: " & Format(.PitStatPtr.ip / 3, "0.0") & "  W-L-S:" & _
                            pitchRecord & "   ERA:" & Format(era, "#.00")
                If Game.homeTeamBatting Then
                    Me.lblHPStats.Text = tempvalue
                Else
                    Me.lblVPStats.Text = tempvalue
                End If
            End With
            'End Using
        Catch ex As Exception
            Call MsgBox("UpdateStatLine " & sqlQuery & " " & ex.ToString, MsgBoxStyle.OkOnly)
        Finally
            If Not dsStat Is Nothing Then dsStat.Dispose()
        End Try
    End Sub

    ''' <summary>
    ''' update the main display
    ''' </summary>
    ''' <param name="isSwitch"></param>
    ''' <remarks></remarks>
    Private Sub UpdateBoard(ByVal isSwitch As Boolean)
        Try
            If Not bolAnyBS Then
                If isSwitch Then
                    Me.lblFirst.Text = ""
                    Me.lblSecond.Text = ""
                    Me.lblThird.Text = ""
                    If bolFinal Then
                        Me.grpBatterHome.BackColor = System.Drawing.Color.White
                        Me.grpBatterVisitor.BackColor = System.Drawing.Color.White
                        Me.grpPitcherHome.BackColor = System.Drawing.Color.White
                        Me.grpPitcherVisitor.BackColor = System.Drawing.Color.White
                    Else
                        If Game.homeTeamBatting Then
                            Me.grpBatterHome.BackColor = System.Drawing.Color.Cyan
                            Me.grpBatterVisitor.BackColor = System.Drawing.Color.White
                            Me.grpPitcherHome.BackColor = System.Drawing.Color.White
                            Me.grpPitcherVisitor.BackColor = System.Drawing.Color.Cyan
                        Else
                            Me.grpBatterHome.BackColor = System.Drawing.Color.White
                            Me.grpBatterVisitor.BackColor = System.Drawing.Color.Cyan
                            Me.grpPitcherHome.BackColor = System.Drawing.Color.Cyan
                            Me.grpPitcherVisitor.BackColor = System.Drawing.Color.White
                        End If
                    End If
                Else
                    If FirstBase.occupied Then
                        With Game.BTeam.GetBatterPtr(FirstBase.runner)
                            Me.lblFirst.Text = GetLastName(.player) & Space(1) & .obr & Space(1) & .sp
                        End With
                    Else
                        Me.lblFirst.Text = ""
                    End If
                    If SecondBase.occupied Then
                        With Game.BTeam.GetBatterPtr(SecondBase.runner)
                            Me.lblSecond.Text = GetLastName(.player) & Space(1) & .obr & Space(1) & .sp
                        End With
                    Else
                        Me.lblSecond.Text = ""
                    End If
                    If ThirdBase.occupied Then
                        With Game.BTeam.GetBatterPtr(ThirdBase.runner)
                            Me.lblThird.Text = GetLastName(.player) & Space(1) & .obr & Space(1) & .sp
                        End With
                    Else
                        Me.lblThird.Text = ""
                    End If
                End If
            End If
            If Home.pr = 0 Then
                LoadPitcher(Home, conHome)
            End If
            If Visitor.pr = 0 Then
                LoadPitcher(Visitor, conVisitor)
            End If

        Catch ex As Exception
            Call MsgBox("Update Board - " & ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' This routine is called in the event a fly out has a runner hold. The lead hold runner can advance. When a play is made on the advance
    ''' attempt, all trailing runners advance.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub AdvanceRunnerOnFlyout()
        Dim frmLaunch As frmAdvanceFlyout
        Dim tempValue As Integer

        Try
            With Game.GetResultPtr
                If Game.homeTeamBatting Then
                    Call Me.DispResult(Game.GetResultPtr.description, Game.BTeam, Game.PTeam)
                Else
                    Call Me.DispResult(Game.GetResultPtr.description, Game.PTeam, Game.BTeam)
                End If
                System.Windows.Forms.Application.DoEvents()
                System.Threading.Thread.Sleep(1000)
                gblUpdateMsg = ""
                frmLaunch = New frmAdvanceFlyout
                frmLaunch.ShowDialog(Me)
                .description = .description & gblUpdateMsg
                tempValue = 0
                'Recalc RBIs and Runs
                If .resFirst = 4 Then tempValue = tempValue + 1
                If .resSecond = 4 Then tempValue = tempValue + 1
                If .resThird = 4 Then tempValue = tempValue + 1
                .rbi = tempValue
                .run = tempValue
                bolCountRun = True
                bolGiveRBI = True
            End With
        Catch ex As Exception
            Call MsgBox("frmMain.AdvanceRunnerOnFlyout - " & ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' This routine is called in the event a runner can advance an extra base.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub AdvanceRunner()
        Dim tempValue As Integer
        Dim frmLaunch As frmAdvance
        Dim frmLaunch2 As frmAdvance2

        Try
            With Game.GetResultPtr
                If Game.homeTeamBatting Then
                    Call Me.DispResult(Game.GetResultPtr.description, Game.BTeam, Game.PTeam)
                Else
                    Call Me.DispResult(Game.GetResultPtr.description, Game.PTeam, Game.BTeam)
                End If
                System.Windows.Forms.Application.DoEvents()
                System.Threading.Thread.Sleep(1000)
                gblUpdateMsg = ""
                If AppSettings("CutOff").ToString.ToUpper = "ON" Then
                    frmLaunch2 = New frmAdvance2
                    frmLaunch2.ShowDialog(Me)
                Else
                    frmLaunch = New frmAdvance
                    frmLaunch.ShowDialog(Me)
                End If
                .description = .description & gblUpdateMsg
                tempValue = 0
                'Recalc RBIs and Runs
                If .resFirst = 4 Then tempValue = tempValue + 1
                If .resSecond = 4 Then tempValue = tempValue + 1
                If .resThird = 4 Then tempValue = tempValue + 1
                .rbi = tempValue
                .run = tempValue
                bolCountRun = True
                bolGiveRBI = True
            End With
        Catch ex As Exception
            Call MsgBox("frmMain.AdvanceRunner - " & ex.ToString)
        End Try
    End Sub

    Private Sub TwipsToPixels(ByVal twipsX As Double, ByVal twipsY As Double, ByRef pixelsX As Double, _
                                ByRef pixelsY As Double)
        Dim g As Graphics = Me.CreateGraphics

        pixelsX = twipsX / 1440.0F * g.DpiX
        pixelsY = twipsY / 1440.0F * g.DpiY

    End Sub

    Private Sub chkInfieldIn_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkInfieldIn.CheckedChanged
        If chkInfieldIn.CheckState = 1 Then
            chkCornersIn.CheckState = CheckState.Unchecked
            chkGuardLine.CheckState = CheckState.Unchecked
        End If
    End Sub

    Private Sub chkCornersIn_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCornersIn.CheckedChanged
        If chkCornersIn.CheckState = 1 Then
            chkInfieldIn.CheckState = CheckState.Unchecked
            chkGuardLine.CheckState = CheckState.Unchecked
        End If
    End Sub

    Private Sub chkGuardLine_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkGuardLine.CheckedChanged
        If chkGuardLine.CheckState = 1 Then
            chkCornersIn.CheckState = CheckState.Unchecked
            chkInfieldIn.CheckState = CheckState.Unchecked
        End If
    End Sub

    Private Sub chkPitchAround1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPitchAround1.CheckedChanged
        If chkPitchAround1.CheckState = 1 Then
            chkPitchAround2.CheckState = CheckState.Unchecked
        End If
    End Sub

    Private Sub chkPitchAround2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPitchAround2.CheckedChanged
        If chkPitchAround2.CheckState = 1 Then
            chkPitchAround1.CheckState = CheckState.Unchecked
        End If
    End Sub

    
End Class