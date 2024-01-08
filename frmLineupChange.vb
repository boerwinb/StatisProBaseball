Option Strict On
Option Explicit On

Friend Class frmLineupChange
	Inherits System.Windows.Forms.Form
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
    Public WithEvents chkAddPositions As System.Windows.Forms.CheckBox
    Public WithEvents cmbAddPos As System.Windows.Forms.ComboBox
	Public WithEvents cmdAddPos As System.Windows.Forms.Button
	Public WithEvents cmdOther As System.Windows.Forms.Button
	Public WithEvents lstUsed As System.Windows.Forms.ListBox
	Public WithEvents lstAvailable As System.Windows.Forms.ListBox
    Public WithEvents cmdExit As System.Windows.Forms.Button
    Public WithEvents cmdReset As System.Windows.Forms.Button
	Public WithEvents cmdSwitch As System.Windows.Forms.Button
    Public WithEvents _Label2_0 As System.Windows.Forms.Label
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
    Public WithEvents grpLineup As System.Windows.Forms.GroupBox
    Friend WithEvents pnlBatters As System.Windows.Forms.Panel
    Friend WithEvents BCard1 As StatisProBaseball.BCard
    Friend WithEvents BCard4 As StatisProBaseball.BCard
    Friend WithEvents BCard3 As StatisProBaseball.BCard
    Friend WithEvents BCard2 As StatisProBaseball.BCard
    Friend WithEvents BCard5 As StatisProBaseball.BCard
    Friend WithEvents BCard7 As StatisProBaseball.BCard
    Friend WithEvents BCard6 As StatisProBaseball.BCard
    Friend WithEvents BCard8 As StatisProBaseball.BCard
    Friend WithEvents BCard11 As StatisProBaseball.BCard
    Friend WithEvents BCard10 As StatisProBaseball.BCard
    Friend WithEvents BCard9 As StatisProBaseball.BCard
    Friend WithEvents BCard14 As StatisProBaseball.BCard
    Friend WithEvents BCard13 As StatisProBaseball.BCard
    Friend WithEvents BCard12 As StatisProBaseball.BCard
    Friend WithEvents BCard21 As StatisProBaseball.BCard
    Friend WithEvents BCard20 As StatisProBaseball.BCard
    Friend WithEvents BCard19 As StatisProBaseball.BCard
    Friend WithEvents BCard18 As StatisProBaseball.BCard
    Friend WithEvents BCard17 As StatisProBaseball.BCard
    Friend WithEvents BCard16 As StatisProBaseball.BCard
    Friend WithEvents BCard15 As StatisProBaseball.BCard
    Friend WithEvents BCard27 As StatisProBaseball.BCard
    Friend WithEvents BCard26 As StatisProBaseball.BCard
    Friend WithEvents BCard25 As StatisProBaseball.BCard
    Friend WithEvents BCard24 As StatisProBaseball.BCard
    Friend WithEvents BCard23 As StatisProBaseball.BCard
    Friend WithEvents BCard28 As StatisProBaseball.BCard
    Friend WithEvents BCard22 As StatisProBaseball.BCard
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmLineupChange))
        Me.chkAddPositions = New System.Windows.Forms.CheckBox
        Me.cmbAddPos = New System.Windows.Forms.ComboBox
        Me.cmdAddPos = New System.Windows.Forms.Button
        Me.cmdOther = New System.Windows.Forms.Button
        Me.lstUsed = New System.Windows.Forms.ListBox
        Me.lstAvailable = New System.Windows.Forms.ListBox
        Me.grpLineup = New System.Windows.Forms.GroupBox
        Me.cmdExit = New System.Windows.Forms.Button
        Me.cmdReset = New System.Windows.Forms.Button
        Me.cmdSwitch = New System.Windows.Forms.Button
        Me._Label2_0 = New System.Windows.Forms.Label
        Me._Label1_0 = New System.Windows.Forms.Label
        Me.pnlBatters = New System.Windows.Forms.Panel
        Me.BCard28 = New StatisProBaseball.BCard
        Me.BCard27 = New StatisProBaseball.BCard
        Me.BCard26 = New StatisProBaseball.BCard
        Me.BCard25 = New StatisProBaseball.BCard
        Me.BCard24 = New StatisProBaseball.BCard
        Me.BCard23 = New StatisProBaseball.BCard
        Me.BCard22 = New StatisProBaseball.BCard
        Me.BCard21 = New StatisProBaseball.BCard
        Me.BCard20 = New StatisProBaseball.BCard
        Me.BCard19 = New StatisProBaseball.BCard
        Me.BCard18 = New StatisProBaseball.BCard
        Me.BCard17 = New StatisProBaseball.BCard
        Me.BCard16 = New StatisProBaseball.BCard
        Me.BCard15 = New StatisProBaseball.BCard
        Me.BCard14 = New StatisProBaseball.BCard
        Me.BCard13 = New StatisProBaseball.BCard
        Me.BCard12 = New StatisProBaseball.BCard
        Me.BCard11 = New StatisProBaseball.BCard
        Me.BCard10 = New StatisProBaseball.BCard
        Me.BCard9 = New StatisProBaseball.BCard
        Me.BCard8 = New StatisProBaseball.BCard
        Me.BCard7 = New StatisProBaseball.BCard
        Me.BCard6 = New StatisProBaseball.BCard
        Me.BCard5 = New StatisProBaseball.BCard
        Me.BCard4 = New StatisProBaseball.BCard
        Me.BCard3 = New StatisProBaseball.BCard
        Me.BCard2 = New StatisProBaseball.BCard
        Me.BCard1 = New StatisProBaseball.BCard
        Me.pnlBatters.SuspendLayout()
        Me.SuspendLayout()
        '
        'chkAddPositions
        '
        Me.chkAddPositions.BackColor = System.Drawing.SystemColors.Control
        Me.chkAddPositions.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkAddPositions.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkAddPositions.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkAddPositions.Location = New System.Drawing.Point(400, 488)
        Me.chkAddPositions.Name = "chkAddPositions"
        Me.chkAddPositions.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkAddPositions.Size = New System.Drawing.Size(97, 17)
        Me.chkAddPositions.TabIndex = 997
        Me.chkAddPositions.Text = "Add Positions"
        Me.chkAddPositions.UseVisualStyleBackColor = False
        '
        'cmbAddPos
        '
        Me.cmbAddPos.BackColor = System.Drawing.SystemColors.Window
        Me.cmbAddPos.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbAddPos.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbAddPos.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbAddPos.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbAddPos.Location = New System.Drawing.Point(472, 520)
        Me.cmbAddPos.Name = "cmbAddPos"
        Me.cmbAddPos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbAddPos.Size = New System.Drawing.Size(41, 22)
        Me.cmbAddPos.TabIndex = 996
        '
        'cmdAddPos
        '
        Me.cmdAddPos.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAddPos.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAddPos.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAddPos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAddPos.Location = New System.Drawing.Point(352, 520)
        Me.cmdAddPos.Name = "cmdAddPos"
        Me.cmdAddPos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAddPos.Size = New System.Drawing.Size(105, 25)
        Me.cmdAddPos.TabIndex = 995
        Me.cmdAddPos.Text = "Add Position"
        Me.cmdAddPos.UseVisualStyleBackColor = False
        '
        'cmdOther
        '
        Me.cmdOther.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOther.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOther.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOther.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOther.Image = CType(resources.GetObject("cmdOther.Image"), System.Drawing.Image)
        Me.cmdOther.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdOther.Location = New System.Drawing.Point(864, 464)
        Me.cmdOther.Name = "cmdOther"
        Me.cmdOther.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOther.Size = New System.Drawing.Size(105, 41)
        Me.cmdOther.TabIndex = 994
        Me.cmdOther.Text = "Other Team"
        Me.cmdOther.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdOther.UseVisualStyleBackColor = False
        '
        'lstUsed
        '
        Me.lstUsed.BackColor = System.Drawing.SystemColors.Window
        Me.lstUsed.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstUsed.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstUsed.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstUsed.ItemHeight = 14
        Me.lstUsed.Location = New System.Drawing.Point(520, 608)
        Me.lstUsed.Name = "lstUsed"
        Me.lstUsed.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstUsed.Size = New System.Drawing.Size(153, 88)
        Me.lstUsed.TabIndex = 993
        '
        'lstAvailable
        '
        Me.lstAvailable.BackColor = System.Drawing.SystemColors.Window
        Me.lstAvailable.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstAvailable.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstAvailable.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstAvailable.ItemHeight = 14
        Me.lstAvailable.Location = New System.Drawing.Point(520, 480)
        Me.lstAvailable.Name = "lstAvailable"
        Me.lstAvailable.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstAvailable.Size = New System.Drawing.Size(153, 88)
        Me.lstAvailable.TabIndex = 992
        '
        'grpLineup
        '
        Me.grpLineup.BackColor = System.Drawing.SystemColors.Control
        Me.grpLineup.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpLineup.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grpLineup.Location = New System.Drawing.Point(768, 512)
        Me.grpLineup.Name = "grpLineup"
        Me.grpLineup.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.grpLineup.Size = New System.Drawing.Size(193, 193)
        Me.grpLineup.TabIndex = 973
        Me.grpLineup.TabStop = False
        Me.grpLineup.Text = "Current Lineup"
        '
        'cmdExit
        '
        Me.cmdExit.BackColor = System.Drawing.SystemColors.Control
        Me.cmdExit.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdExit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdExit.Image = CType(resources.GetObject("cmdExit.Image"), System.Drawing.Image)
        Me.cmdExit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdExit.Location = New System.Drawing.Point(696, 648)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdExit.Size = New System.Drawing.Size(57, 49)
        Me.cmdExit.TabIndex = 972
        Me.cmdExit.Text = "Exit"
        Me.cmdExit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdExit.UseVisualStyleBackColor = False
        '
        'cmdReset
        '
        Me.cmdReset.BackColor = System.Drawing.SystemColors.Control
        Me.cmdReset.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdReset.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdReset.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdReset.Image = CType(resources.GetObject("cmdReset.Image"), System.Drawing.Image)
        Me.cmdReset.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdReset.Location = New System.Drawing.Point(696, 584)
        Me.cmdReset.Name = "cmdReset"
        Me.cmdReset.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdReset.Size = New System.Drawing.Size(57, 49)
        Me.cmdReset.TabIndex = 971
        Me.cmdReset.Text = "Reset"
        Me.cmdReset.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdReset.UseVisualStyleBackColor = False
        '
        'cmdSwitch
        '
        Me.cmdSwitch.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSwitch.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSwitch.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSwitch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSwitch.Image = CType(resources.GetObject("cmdSwitch.Image"), System.Drawing.Image)
        Me.cmdSwitch.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdSwitch.Location = New System.Drawing.Point(696, 520)
        Me.cmdSwitch.Name = "cmdSwitch"
        Me.cmdSwitch.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSwitch.Size = New System.Drawing.Size(57, 48)
        Me.cmdSwitch.TabIndex = 970
        Me.cmdSwitch.Text = "Switch"
        Me.cmdSwitch.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdSwitch.UseVisualStyleBackColor = False
        '
        '_Label2_0
        '
        Me._Label2_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label2_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label2_0.Location = New System.Drawing.Point(528, 592)
        Me._Label2_0.Name = "_Label2_0"
        Me._Label2_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_0.Size = New System.Drawing.Size(89, 17)
        Me._Label2_0.TabIndex = 1
        Me._Label2_0.Text = "Not Available:"
        '
        '_Label1_0
        '
        Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_0.Location = New System.Drawing.Point(528, 464)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_0.Size = New System.Drawing.Size(145, 17)
        Me._Label1_0.TabIndex = 0
        Me._Label1_0.Text = "Available Bench Players:"
        '
        'pnlBatters
        '
        Me.pnlBatters.AutoScroll = True
        Me.pnlBatters.Controls.Add(Me.BCard28)
        Me.pnlBatters.Controls.Add(Me.BCard27)
        Me.pnlBatters.Controls.Add(Me.BCard26)
        Me.pnlBatters.Controls.Add(Me.BCard25)
        Me.pnlBatters.Controls.Add(Me.BCard24)
        Me.pnlBatters.Controls.Add(Me.BCard23)
        Me.pnlBatters.Controls.Add(Me.BCard22)
        Me.pnlBatters.Controls.Add(Me.BCard21)
        Me.pnlBatters.Controls.Add(Me.BCard20)
        Me.pnlBatters.Controls.Add(Me.BCard19)
        Me.pnlBatters.Controls.Add(Me.BCard18)
        Me.pnlBatters.Controls.Add(Me.BCard17)
        Me.pnlBatters.Controls.Add(Me.BCard16)
        Me.pnlBatters.Controls.Add(Me.BCard15)
        Me.pnlBatters.Controls.Add(Me.BCard14)
        Me.pnlBatters.Controls.Add(Me.BCard13)
        Me.pnlBatters.Controls.Add(Me.BCard12)
        Me.pnlBatters.Controls.Add(Me.BCard11)
        Me.pnlBatters.Controls.Add(Me.BCard10)
        Me.pnlBatters.Controls.Add(Me.BCard9)
        Me.pnlBatters.Controls.Add(Me.BCard8)
        Me.pnlBatters.Controls.Add(Me.BCard7)
        Me.pnlBatters.Controls.Add(Me.BCard6)
        Me.pnlBatters.Controls.Add(Me.BCard5)
        Me.pnlBatters.Controls.Add(Me.BCard4)
        Me.pnlBatters.Controls.Add(Me.BCard3)
        Me.pnlBatters.Controls.Add(Me.BCard2)
        Me.pnlBatters.Controls.Add(Me.BCard1)
        Me.pnlBatters.Location = New System.Drawing.Point(0, 1)
        Me.pnlBatters.Name = "pnlBatters"
        Me.pnlBatters.Size = New System.Drawing.Size(1026, 460)
        Me.pnlBatters.TabIndex = 1020
        '
        'BCard28
        '
        Me.BCard28.Location = New System.Drawing.Point(864, 516)
        Me.BCard28.Name = "BCard28"
        Me.BCard28.Size = New System.Drawing.Size(144, 172)
        Me.BCard28.TabIndex = 27
        Me.BCard28.Tag = "28"
        '
        'BCard27
        '
        Me.BCard27.Location = New System.Drawing.Point(720, 516)
        Me.BCard27.Name = "BCard27"
        Me.BCard27.Size = New System.Drawing.Size(144, 172)
        Me.BCard27.TabIndex = 26
        Me.BCard27.Tag = "27"
        '
        'BCard26
        '
        Me.BCard26.Location = New System.Drawing.Point(576, 516)
        Me.BCard26.Name = "BCard26"
        Me.BCard26.Size = New System.Drawing.Size(144, 172)
        Me.BCard26.TabIndex = 25
        Me.BCard26.Tag = "26"
        '
        'BCard25
        '
        Me.BCard25.Location = New System.Drawing.Point(432, 516)
        Me.BCard25.Name = "BCard25"
        Me.BCard25.Size = New System.Drawing.Size(144, 172)
        Me.BCard25.TabIndex = 24
        Me.BCard25.Tag = "25"
        '
        'BCard24
        '
        Me.BCard24.Location = New System.Drawing.Point(288, 516)
        Me.BCard24.Name = "BCard24"
        Me.BCard24.Size = New System.Drawing.Size(144, 172)
        Me.BCard24.TabIndex = 23
        Me.BCard24.Tag = "24"
        '
        'BCard23
        '
        Me.BCard23.Location = New System.Drawing.Point(144, 546)
        Me.BCard23.Name = "BCard23"
        Me.BCard23.Size = New System.Drawing.Size(144, 172)
        Me.BCard23.TabIndex = 22
        Me.BCard23.Tag = "23"
        '
        'BCard22
        '
        Me.BCard22.Location = New System.Drawing.Point(0, 516)
        Me.BCard22.Name = "BCard22"
        Me.BCard22.Size = New System.Drawing.Size(144, 172)
        Me.BCard22.TabIndex = 21
        Me.BCard22.Tag = "22"
        '
        'BCard21
        '
        Me.BCard21.Location = New System.Drawing.Point(864, 344)
        Me.BCard21.Name = "BCard21"
        Me.BCard21.Size = New System.Drawing.Size(144, 172)
        Me.BCard21.TabIndex = 20
        Me.BCard21.Tag = "21"
        '
        'BCard20
        '
        Me.BCard20.Location = New System.Drawing.Point(720, 344)
        Me.BCard20.Name = "BCard20"
        Me.BCard20.Size = New System.Drawing.Size(144, 172)
        Me.BCard20.TabIndex = 19
        Me.BCard20.Tag = "20"
        '
        'BCard19
        '
        Me.BCard19.Location = New System.Drawing.Point(576, 344)
        Me.BCard19.Name = "BCard19"
        Me.BCard19.Size = New System.Drawing.Size(144, 172)
        Me.BCard19.TabIndex = 18
        Me.BCard19.Tag = "19"
        '
        'BCard18
        '
        Me.BCard18.Location = New System.Drawing.Point(432, 344)
        Me.BCard18.Name = "BCard18"
        Me.BCard18.Size = New System.Drawing.Size(144, 172)
        Me.BCard18.TabIndex = 17
        Me.BCard18.Tag = "18"
        '
        'BCard17
        '
        Me.BCard17.Location = New System.Drawing.Point(288, 344)
        Me.BCard17.Name = "BCard17"
        Me.BCard17.Size = New System.Drawing.Size(144, 172)
        Me.BCard17.TabIndex = 16
        Me.BCard17.Tag = "17"
        '
        'BCard16
        '
        Me.BCard16.Location = New System.Drawing.Point(144, 344)
        Me.BCard16.Name = "BCard16"
        Me.BCard16.Size = New System.Drawing.Size(144, 172)
        Me.BCard16.TabIndex = 15
        Me.BCard16.Tag = "16"
        '
        'BCard15
        '
        Me.BCard15.Location = New System.Drawing.Point(0, 344)
        Me.BCard15.Name = "BCard15"
        Me.BCard15.Size = New System.Drawing.Size(144, 172)
        Me.BCard15.TabIndex = 14
        Me.BCard15.Tag = "15"
        '
        'BCard14
        '
        Me.BCard14.Location = New System.Drawing.Point(864, 172)
        Me.BCard14.Name = "BCard14"
        Me.BCard14.Size = New System.Drawing.Size(144, 172)
        Me.BCard14.TabIndex = 13
        Me.BCard14.Tag = "14"
        '
        'BCard13
        '
        Me.BCard13.Location = New System.Drawing.Point(720, 172)
        Me.BCard13.Name = "BCard13"
        Me.BCard13.Size = New System.Drawing.Size(144, 172)
        Me.BCard13.TabIndex = 12
        Me.BCard13.Tag = "13"
        '
        'BCard12
        '
        Me.BCard12.Location = New System.Drawing.Point(576, 172)
        Me.BCard12.Name = "BCard12"
        Me.BCard12.Size = New System.Drawing.Size(144, 172)
        Me.BCard12.TabIndex = 11
        Me.BCard12.Tag = "12"
        '
        'BCard11
        '
        Me.BCard11.Location = New System.Drawing.Point(432, 172)
        Me.BCard11.Name = "BCard11"
        Me.BCard11.Size = New System.Drawing.Size(144, 172)
        Me.BCard11.TabIndex = 10
        Me.BCard11.Tag = "11"
        '
        'BCard10
        '
        Me.BCard10.Location = New System.Drawing.Point(288, 172)
        Me.BCard10.Name = "BCard10"
        Me.BCard10.Size = New System.Drawing.Size(144, 172)
        Me.BCard10.TabIndex = 9
        Me.BCard10.Tag = "10"
        '
        'BCard9
        '
        Me.BCard9.Location = New System.Drawing.Point(144, 172)
        Me.BCard9.Name = "BCard9"
        Me.BCard9.Size = New System.Drawing.Size(144, 172)
        Me.BCard9.TabIndex = 8
        Me.BCard9.Tag = "9"
        '
        'BCard8
        '
        Me.BCard8.Location = New System.Drawing.Point(0, 172)
        Me.BCard8.Name = "BCard8"
        Me.BCard8.Size = New System.Drawing.Size(144, 172)
        Me.BCard8.TabIndex = 7
        Me.BCard8.Tag = "8"
        '
        'BCard7
        '
        Me.BCard7.Location = New System.Drawing.Point(864, 0)
        Me.BCard7.Name = "BCard7"
        Me.BCard7.Size = New System.Drawing.Size(144, 172)
        Me.BCard7.TabIndex = 6
        Me.BCard7.Tag = "7"
        '
        'BCard6
        '
        Me.BCard6.Location = New System.Drawing.Point(720, 0)
        Me.BCard6.Name = "BCard6"
        Me.BCard6.Size = New System.Drawing.Size(144, 172)
        Me.BCard6.TabIndex = 5
        Me.BCard6.Tag = "6"
        '
        'BCard5
        '
        Me.BCard5.Location = New System.Drawing.Point(576, 0)
        Me.BCard5.Name = "BCard5"
        Me.BCard5.Size = New System.Drawing.Size(144, 172)
        Me.BCard5.TabIndex = 4
        Me.BCard5.Tag = "5"
        '
        'BCard4
        '
        Me.BCard4.Location = New System.Drawing.Point(432, 0)
        Me.BCard4.Name = "BCard4"
        Me.BCard4.Size = New System.Drawing.Size(144, 172)
        Me.BCard4.TabIndex = 3
        Me.BCard4.Tag = "4"
        '
        'BCard3
        '
        Me.BCard3.Location = New System.Drawing.Point(288, 0)
        Me.BCard3.Name = "BCard3"
        Me.BCard3.Size = New System.Drawing.Size(144, 172)
        Me.BCard3.TabIndex = 2
        Me.BCard3.Tag = "3"
        '
        'BCard2
        '
        Me.BCard2.Location = New System.Drawing.Point(144, 0)
        Me.BCard2.Name = "BCard2"
        Me.BCard2.Size = New System.Drawing.Size(144, 172)
        Me.BCard2.TabIndex = 1
        Me.BCard2.Tag = "2"
        '
        'BCard1
        '
        Me.BCard1.Location = New System.Drawing.Point(0, 0)
        Me.BCard1.Name = "BCard1"
        Me.BCard1.Size = New System.Drawing.Size(144, 172)
        Me.BCard1.TabIndex = 0
        Me.BCard1.Tag = "1"
        '
        'frmLineupChange
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1028, 713)
        Me.Controls.Add(Me.pnlBatters)
        Me.Controls.Add(Me.chkAddPositions)
        Me.Controls.Add(Me.cmbAddPos)
        Me.Controls.Add(Me.cmdAddPos)
        Me.Controls.Add(Me.cmdOther)
        Me.Controls.Add(Me.lstUsed)
        Me.Controls.Add(Me.lstAvailable)
        Me.Controls.Add(Me.grpLineup)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.cmdReset)
        Me.Controls.Add(Me.cmdSwitch)
        Me.Controls.Add(Me._Label2_0)
        Me.Controls.Add(Me._Label1_0)
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(4, 23)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmLineupChange"
        Me.Text = "Lineup Change"
        Me.pnlBatters.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
#End Region
#Region "Upgrade Support "
    Private Shared m_vb6FormDefInstance As frmLineupChange
    Private Shared m_InitializingDefInstance As Boolean
    Public Shared Property DefInstance() As frmLineupChange
        Get
            If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
                m_InitializingDefInstance = True
                m_vb6FormDefInstance = New frmLineupChange()
                m_InitializingDefInstance = False
            End If
            DefInstance = m_vb6FormDefInstance
        End Get
        Set(ByVal value As frmLineupChange)
            m_vb6FormDefInstance = value
        End Set
    End Property
#End Region
    Dim bolDefensive As Boolean
    Dim bolPromptSave As Boolean
    Dim bolPHPitchTemp As Boolean
    Dim arrBatCardLabelsCols(22) As clsLabelArray
    Dim colLineupButtons As clsRadioButtonArray
    Dim colPosCombos As clsComboArray

    'UPGRADE_WARNING: Event chkAddPositions.CheckStateChanged may fire when form is intialized. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup2075"'
    Private Sub chkAddPositions_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkAddPositions.CheckStateChanged
        Try
            If chkAddPositions.CheckState = 1 Then
                cmdAddPos.Visible = True
                With cmbAddPos
                    .Visible = True
                    .Items.Add("C")
                    .Items.Add("1B")
                    .Items.Add("2B")
                    .Items.Add("3B")
                    .Items.Add("SS")
                    .Items.Add("OF")
                    If bolAmericanLeagueRules Then
                        .Items.Add("DH")
                    End If
                End With
            Else
                cmdAddPos.Visible = False
                cmbAddPos.Visible = False
            End If
        Catch ex As Exception
            Call MsgBox("Error in chkAddPositions_CheckStateChanged " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub cmdAddPos_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAddPos.Click
        Dim fieldPosition As String
        Dim team As clsTeam
        Dim isFound As Boolean
        Dim index As Integer

        fieldPosition = cmbAddPos.Text
        Try
            'Load current lineup
            If Game.homeTeamBatting Then
                If bolDefensive Then
                    team = Visitor
                Else
                    'Offensive changes
                    team = Home
                End If
            Else
                If bolDefensive Then
                    team = Home
                Else
                    'Offensive changes
                    team = Visitor
                End If
            End If

            For index = 0 To 8
                If colLineupButtons.Item(index).Checked Then
                    isFound = True
                    Exit For
                End If
            Next index
            If Not isFound Then
                Call MsgBox("Must select a player to replace.", MsgBoxStyle.OkOnly)
                Exit Sub
            End If
            With team.GetBatterPtr(team.GetIndexFromName(colLineupButtons.Item(index).Text))
                Select Case fieldPosition
                    Case "C"
                        .field = .field & " C-2 E10 TC"
                    Case "OF"
                        .field = .field & " OF-2 E10 T2"
                    Case "DH"
                        .field = .field & " DH-2"
                    Case Else
                        .field = .field & Space(1) & fieldPosition & "-2 E10"
                End Select
            End With
            'Changes are persistent here
            If Game.homeTeamBatting Then
                If bolDefensive Then
                    Visitor = team
                Else
                    'Offensive changes
                    Home = team
                End If
            Else
                If bolDefensive Then
                    Home = team
                Else
                    'Offensive changes
                    Visitor = team
                End If
            End If

            bolPromptSave = False
            FillPlayers(False)
        Catch ex As Exception
            Call MsgBox("cmdAddPos_Click " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub cmdExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdExit.Click
        'First save the positions. This functionality was removed with 
        'the cmbPos click event
        Dim tmpHome As clsTeam = Home
        Dim tmpVisitor As clsTeam = Visitor

        Me.SavePositions()

        If SaveChanges() Then
            Me.Close()
        Else
            Home = tmpHome
            Visitor = tmpVisitor
            BoardReset()
        End If
    End Sub

    Private Sub cmdReset_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdReset.Click
        BoardReset()
    End Sub

    Private Sub cmdSwitch_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSwitch.Click
        SwitchPlayers()
    End Sub

    Private Sub cmdOther_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOther.Click
        'First save the positions. This functionality was removed with 
        'the cmbPos click event
        Me.SavePositions()

        If bolPromptSave Or PositionsChanged() Then
            If MsgBox("Save current changes?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                SaveChanges()
            End If
        End If
        bolDefensive = Not bolDefensive
        FillPlayers(False)
    End Sub

    Private Sub frmLineupChange_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        bolDefensive = False
        bolPromptSave = False
        colLineupButtons = New clsRadioButtonArray(Me.grpLineup)
        colPosCombos = New clsComboArray(Me.grpLineup)
        FillPlayers(False)
    End Sub

    ''' <summary>
    ''' Fills the batting cards
    ''' </summary>
    ''' <param name="isReset"></param>
    ''' <remarks></remarks>
    Private Sub FillPlayers(ByRef isReset As Boolean)
        Dim team As clsTeam
        Dim objBCard As BCard = Nothing
        'Dim cardIndex As Integer
        'Dim inactiveCount As Integer
        'Dim maxInactive As Integer

        Try
            colLineupButtons.RemoveAll()
            colPosCombos.RemoveAll()
            For j As Integer = 1 To 9
                colLineupButtons.AddNewRadioButton()
                colPosCombos.AddNewCombo()
            Next j
            team = New clsTeam
            'Load current lineup
            If Game.homeTeamBatting Then
                If bolDefensive Then
                    team = Visitor
                Else
                    'Offensive changes
                    team = Home
                End If
            Else
                If bolDefensive Then
                    team = Home
                Else
                    'Offensive changes
                    team = Visitor
                End If
            End If
            For i As Integer = 0 To 8
                colLineupButtons.Item(i).Text = team.GetBatterPtr(team.GetPlayerNum(i + 1)).player
                Call FillPositions(i + 1, team.GetPlayerNum(i + 1), team, False)
                colPosCombos.Item(i).SelectedIndex = SetListIndex(colPosCombos.Item(i), (team.GetBatterPtr(team.GetPlayerNum(i + 1)).position))
            Next i

            lstAvailable.Items.Clear()
            lstUsed.Items.Clear()

            'If team.hitters > conMaxBatters Then
            '    If team.hitters - team.inactiveHitters > conMaxBatters Then
            '        'Just list the first conmaxmatters
            '        For i As Integer = 0 To conMaxBatters - 1
            '            Call FillCards(team, i, i + 1, isReset)
            '        Next i
            '    Else
            '        maxInactive = conMaxBatters - (team.hitters - team.inactiveHitters)
            '        For i As Integer = 0 To team.hitters - 1
            '            If team.GetBatterPtr(i + 1).BatStatPtr.Active Then
            '                cardIndex += 1
            '                Call FillCards(team, i, cardIndex, isReset)
            '            ElseIf inactiveCount < maxInactive Then
            '                cardIndex += 1
            '                inactiveCount += 1
            '                Call FillCards(team, i, cardIndex, isReset)
            '            End If
            '        Next i
            '    End If
            'Else
            '    For i As Integer = 0 To conMaxBatters - 1
            '        Call FillCards(team, i, i + 1, isReset)
            '    Next i
            'End If
            For i As Integer = 0 To conMaxBatters - 1
                Call FillCards(team, i, i + 1, isReset)
            Next
            
            'Handle NL Pitcher
            If Not bolAmericanLeagueRules Then
                With team.GetBatterPtr(team.hitters + team.nthPitcherUsed)
                    If .player <> Nothing Then
                        If .available Then
                            lstAvailable.Items.Add(.player)
                        End If
                    End If
                End With
            End If

        Catch ex As Exception
            Call MsgBox("FillPlayers " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub FillCards(ByRef Team As clsTeam, ByRef playerIndex As Integer, ByVal cardIndex As Integer, ByVal isReset As Boolean)
        Dim objBCard As BCard = Nothing

        'For k As Integer = 0 To Me.Controls.Count - 1
        For k As Integer = 0 To Me.pnlBatters.Controls.Count - 1
            If CType(Me.pnlBatters.Controls(k).Tag, Integer) = cardIndex Then
                objBCard = CType(Me.pnlBatters.Controls(k), BCard)
                k = Me.pnlBatters.Controls.Count
            End If
        Next k
        With Team.GetBatterPtr(playerIndex + 1)
            If playerIndex <= Team.hitters - 1 Then
                'If .player <> Nothing Then
                If Not Team.InLineup(.playerIndex) Then
                    If .available Then
                        lstAvailable.Items.Add(.player)
                    Else
                        lstUsed.Items.Add(.player)
                    End If
                End If

                If Not isReset Then
                    objBCard.Visible = True
                    objBCard.grpBCard.Text = .player
                    objBCard.lblField.Text = .field
                    objBCard.lblOBR.Text = .obr
                    objBCard.lblSP.Text = .sp
                    objBCard.lblHitandRun.Text = .hitRun
                    objBCard.lblCD.Text = .cd
                    objBCard.lblSAC.Text = .sac
                    objBCard.lblInj.Text = .inj
                    objBCard.lbl1bf.Text = .hit1Bf
                    objBCard.lbl1b7.Text = .hit1B7
                    objBCard.lbl1b8.Text = .hit1B8
                    objBCard.lbl1b9.Text = .hit1B9
                    objBCard.lbl2b7.Text = .hit2B7
                    objBCard.lbl2b8.Text = .hit2B8
                    objBCard.lbl2b9.Text = .hit2B9
                    objBCard.lbl3b8.Text = .hit3B8
                    objBCard.lblHR.Text = .hr
                    objBCard.lblK.Text = .k
                    objBCard.lblW.Text = .w
                    objBCard.lblHPB.Text = .hpb
                    objBCard.lblOut.Text = .out
                    objBCard.lblCht.Text = .cht
                    objBCard.lblBD.Text = .bd

                    objBCard.BackColor = System.Drawing.Color.White
                    If Not .available Then
                        If .BatStatPtr.GamesInj > 0 Then
                            objBCard.BackColor = System.Drawing.Color.Red
                        Else
                            objBCard.Enabled = False
                        End If
                    End If
                    objBCard.Enabled = False
                End If
            ElseIf Not isReset Then
                objBCard.Visible = False
            End If
        End With
   End Sub

    ''' <summary>
    ''' fills the position options for each player
    ''' </summary>
    ''' <param name="lineupIndex"></param>
    ''' <param name="playerIndex"></param>
    ''' <param name="team"></param>
    ''' <remarks></remarks>
    Private Sub FillPositions(ByRef lineupIndex As Integer, ByRef playerIndex As Integer, ByRef team As clsTeam, _
                        ByVal possiblePitcher As Boolean)
        Dim iPos As Integer
        Dim fieldLine As String
        Dim fieldPosition As String
        Dim hasDH As Boolean

        Try
            If playerIndex >= team.hitters + 1 Then
                'Pitcher
                colPosCombos.Item(lineupIndex - 1).Items.Add("P")
            Else
                fieldLine = team.GetBatterPtr(playerIndex).field
                fieldLine = "*" & fieldLine 'Add a dummy text buffer
                iPos = fieldLine.IndexOf("-")
                colPosCombos.Item(lineupIndex - 1).Items.Clear()

                While iPos <> -1
                    fieldPosition = fieldLine.Substring(iPos - 2, 2)
                    Select Case fieldPosition.Substring(fieldPosition.Length - 1)
                        Case "C"
                            fieldPosition = "C"
                        Case "P"
                            fieldPosition = "P"
                    End Select
                    If fieldPosition = "OF" Then
                        colPosCombos.Item(lineupIndex - 1).Items.Add("LF")
                        colPosCombos.Item(lineupIndex - 1).Items.Add("CF")
                        colPosCombos.Item(lineupIndex - 1).Items.Add("RF")
                    Else
                        colPosCombos.Item(lineupIndex - 1).Items.Add(fieldPosition)
                        If fieldPosition = "DH" Then
                            hasDH = True
                        End If
                    End If
                    fieldLine = fieldLine.Substring(iPos + 1)
                    iPos = fieldLine.IndexOf("-")
                End While
                If fieldLine.ToUpper.IndexOf("PINCH HITTER ONLY") > -1 Then
                    colPosCombos.Item(lineupIndex - 1).Items.Add("PH")
                End If
                If gstrPostSeason = conWS And bolAmericanLeagueRules And Not hasDH Then
                    'Allow DH for all players
                    colPosCombos.Item(lineupIndex - 1).Items.Add("DH")
                End If
                If possiblePitcher Then
                    colPosCombos.Item(lineupIndex - 1).Items.Add("P")
                End If
            End If
        Catch ex As Exception
            Call MsgBox("FillPositions " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' check to make sure there are enough position players to field a team after a batter is pinch hit for
    ''' </summary>
    ''' <param name="team"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ValidPinchHit(ByRef team As clsTeam) As Boolean
        Dim position As Integer
        Dim fieldLine As String
        Dim fieldPosition As String
        Dim tempValue As String = ""
        Dim playerUsed As Boolean
        Dim canPinchHit As Boolean

        Try
            For i As Integer = 1 To team.hitters
                If team.GetBatterPtr(i).available Then
                    playerUsed = False
                    For intIndex As Integer = 0 To lstUsed.Items.Count - 1
                        If lstUsed.Items.Item(intIndex).ToString = team.GetBatterPtr(i).player Then
                            playerUsed = True
                        End If
                    Next intIndex
                    If Not playerUsed Then
                        fieldLine = team.GetBatterPtr(i).field
                        fieldLine = "*" & fieldLine 'Add a dummy text buffer
                        position = fieldLine.IndexOf("-")
                        While position <> -1
                            fieldPosition = fieldLine.Substring(position - 2, 2)
                            Select Case fieldPosition.Substring(fieldPosition.Length - 1)
                                Case "C"
                                    fieldPosition = "C"
                                Case "P"
                                    fieldPosition = "P"
                            End Select
                            tempValue = tempValue & fieldPosition
                            fieldLine = fieldLine.Substring(position + 1)
                            position = fieldLine.IndexOf("-")
                        End While
                    End If
                End If
            Next i
            canPinchHit = True
            canPinchHit = canPinchHit And tempValue.IndexOf("C") > -1
            canPinchHit = canPinchHit And tempValue.IndexOf("1B") > -1
            canPinchHit = canPinchHit And tempValue.IndexOf("2B") > -1
            canPinchHit = canPinchHit And tempValue.IndexOf("3B") > -1
            canPinchHit = canPinchHit And tempValue.IndexOf("SS") > -1
            If bolAmericanLeagueRules And Not gstrPostSeason = conWS Then
                canPinchHit = canPinchHit And tempValue.IndexOf("DH") > -1
            End If
            'Looking 3 OF instances
            position = tempValue.IndexOf("OF")
            canPinchHit = canPinchHit And position > -1

            tempValue = tempValue.Substring(position + 1)
            position = tempValue.IndexOf("OF")
            canPinchHit = canPinchHit And position > -1

            tempValue = tempValue.Substring(position + 1)
            position = tempValue.IndexOf("OF")
            canPinchHit = canPinchHit And position > -1
            Return canPinchHit
        Catch ex As Exception
            Call MsgBox("FillPositions " & ex.ToString, MsgBoxStyle.OkOnly)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' switches team and player objects between innings and halves of innings
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SwitchPlayers()
        Dim team As clsTeam
        Dim i As Integer
        Dim isFound As Boolean

        Try
            If lstAvailable.SelectedItems.Count < 1 Then
                Call MsgBox("Must select a player to add.", MsgBoxStyle.OkOnly)
                Exit Sub
            End If
            If Game.homeTeamBatting Then
                If bolDefensive Then
                    team = Visitor
                Else
                    'Offensive changes
                    team = Home
                End If
            Else
                If bolDefensive Then
                    team = Home
                Else
                    'Offensive changes
                    team = Visitor
                End If
            End If
            For i = 0 To 8
                If colLineupButtons.Item(i).Checked Then
                    isFound = True
                    Exit For
                End If
            Next i
            If Not isFound Then
                Call MsgBox("Must select a player to replace.", MsgBoxStyle.OkOnly)
                Exit Sub
            End If
            bolPromptSave = True
            If Not bolAmericanLeagueRules And team.GetPlayerNum(i + 1) >= team.hitters + 1 Then
                'If batter being removed is a pitcher
                lstAvailable.Items.Add(team.GetBatterPtr(team.hitters + team.nthPitcherUsed).player)
                bolPHPitchTemp = True
            Else
                lstUsed.Items.Add(team.GetBatterPtr(team.GetPlayerNum(i + 1)).player)
            End If
            colLineupButtons.Item(i).Text = lstAvailable.Items.Item(lstAvailable.SelectedIndex).ToString
            Call FillPositions(i + 1, team.GetIndexFromName(colLineupButtons.Item(i).Text), team, bolPHPitchTemp)
            colPosCombos.Item(i).SelectedIndex = 0 'Initialize to first position in list
            lstAvailable.Items.RemoveAt((lstAvailable.SelectedIndex))
        Catch ex As Exception
            Call MsgBox("SwitchPlayers " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' persists all objects based on the lineup changes
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function SaveChanges() As Boolean
        Dim team As clsTeam
        Dim isExiting As Boolean

        Try
            team = New clsTeam
            If Game.homeTeamBatting Then
                If bolDefensive Then
                    team = Visitor
                Else
                    'Offensive changes
                    team = Home
                End If
            Else
                If bolDefensive Then
                    team = Home
                Else
                    'Offensive changes
                    team = Visitor
                End If
            End If

            If Not ValidPinchHit(team) Then
                Call MsgBox("Cannot pinch hit. All positions cannot be covered. Board has been reset.", MsgBoxStyle.OkOnly, "Save Changes")
                BoardReset()
                Return isExiting
            End If
            'Save new positions
            For i As Integer = 0 To 8
                If colLineupButtons.Item(i).Text <> team.GetBatterPtr(team.GetPlayerNum(i + 1)).player Then
                    'Look for pinch run situation
                    If Not bolDefensive Then
                        If FirstBase.occupied And team.GetPlayerNum(i + 1) = FirstBase.runner Then
                            FirstBase.runner = team.GetIndexFromName(colLineupButtons.Item(i).Text)
                        ElseIf SecondBase.occupied And team.GetPlayerNum(i + 1) = SecondBase.runner Then
                            SecondBase.runner = team.GetIndexFromName(colLineupButtons.Item(i).Text)
                        ElseIf ThirdBase.occupied And team.GetPlayerNum(i + 1) = ThirdBase.runner Then
                            ThirdBase.runner = team.GetIndexFromName(colLineupButtons.Item(i).Text)
                        End If
                    End If
                    Call team.AddLineup(i + 1, team.GetIndexFromName(colLineupButtons.Item(i).Text))
                End If
                'New player has been inserted, get NEW batter object
                With team.GetBatterPtr(team.GetPlayerNum(i + 1))
                    If Not bolAmericanLeagueRules And team.GetPlayerNum(i + 1) >= team.hitters + 1 Then
                        If team.GetBatterPtr(team.GetPlayerNum(i + 1)).player = team.GetPitcherPtr(team.pitcherSel).player Then
                            bolPHPitchTemp = False
                        End If
                        .position = "P"
                    Else
                        If .position = "DH" And colPosCombos.Item(i).SelectedItem.ToString <> "DH" Then
                            'Once a player enters the game as a DH, they cannot be moved to another position
                            colPosCombos.Item(i).SelectedItem = "DH"
                            Call MsgBox("Once a player enters the game as a DH, they cannot be moved to another position.", MsgBoxStyle.OkOnly, "Save Changes")
                            Return isExiting
                        Else
                            .position = colPosCombos.Item(i).Text
                        End If

                        .errorRating = GetRating(.field, .position, "E")
                        .throwRating = GetRating(.field, .position, "T")
                        If .cd.IndexOf(.position) > -1 Or (.position.Substring(.position.Length - 1) = "F" And .cd.Substring(.cd.Length - 1) = "F") Then
                            .cdAct = (Val(.cd)).ToString
                        Else
                            .cdAct = "0"
                        End If
                    End If
                End With
            Next i
            If bolDefensive Then
                'Only validate defensive changes right away
                If Not ValidPositions(team) Then
                    Call MsgBox("Positions invalid. Recheck positions.", MsgBoxStyle.OkOnly, "Save Changes")
                    Return isExiting
                End If
            End If

            'Saving Data, Set .Available attribute based lstUsed
            For i As Integer = 0 To lstUsed.Items.Count - 1
                team.GetBatterPtr(team.GetIndexFromName(lstUsed.Items.Item(i).ToString)).available = False
            Next i
            team.pitchChange = team.pitchChange Or bolPHPitchTemp 'Save flag that pitcher needs to be changed
            bolPHPitchTemp = False
            If team.pitchChange Then
                'Make current pitcher unavailable
                Game.PTeam.GetPitcherPtr(Game.PTeam.pitcherSel).available = False
            End If
            'Changes are persisted here
            If Game.homeTeamBatting Then
                If bolDefensive Then
                    Visitor = team
                Else
                    'Offensive changes
                    Home = team
                End If
            Else
                If bolDefensive Then
                    Home = team
                Else
                    'Offensive changes
                    Visitor = team
                End If
            End If
            isExiting = True
        Catch ex As Exception
            Call MsgBox("SaveChanges " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return isExiting
    End Function

    Private Sub BoardReset()
        bolPromptSave = False
        lstUsed.Items.Clear()
        lstAvailable.Items.Clear()
        bolPHPitchTemp = False
        FillPlayers(True)
    End Sub

    ''' <summary>
    ''' determines whether a position changed in the lineup
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function PositionsChanged() As Boolean
        Dim isChange As Boolean
        Dim Team As clsTeam

        Try
            If Game.homeTeamBatting Then
                If bolDefensive Then
                    Team = Visitor
                Else
                    'Offensive changes
                    Team = Home
                End If
            Else
                If bolDefensive Then
                    Team = Home
                Else
                    'Offensive changes
                    Team = Visitor
                End If
            End If

            For i As Integer = 0 To 8
                With Team.GetBatterPtr(Team.GetPlayerNum(i + 1))
                    If .position <> colPosCombos.Item(i).Text Then
                        isChange = True
                    End If
                End With
            Next i
        Catch ex As Exception
            Call MsgBox("PositionChanged " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return isChange
    End Function

    ''' <summary>
    ''' persists the position and field rating to the Player objects
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SavePositions()
        Dim team As clsTeam

        Try
            If Game.HomeTeamBatting Then
                If bolDefensive Then
                    team = Visitor
                Else
                    'Offensive changes
                    team = Home
                End If
            Else
                If bolDefensive Then
                    team = Home
                Else
                    'Offensive changes
                    team = Visitor
                End If
            End If

            For i As Integer = 0 To 8
                With team.GetBatterPtr(team.GetPlayerNum(i + 1))
                    .position = colPosCombos.Item(i).Text
                    .errorRating = GetRating(.field, .position, "E")
                    .throwRating = GetRating(.field, .position, "T")
                    If .cd.IndexOf(.position) > 0 Or _
                            (.position.Substring(.position.Length - 1) = "F" And _
                            .cd.Substring(.cd.Length - 1) = "F") Then
                        .cdAct = (Val(.cd)).ToString
                    Else
                        .cdAct = "0"
                    End If
                End With
            Next i
            
        Catch ex As Exception
            Call MsgBox("SavePositions " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub
End Class