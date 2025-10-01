Option Strict On
Option Explicit On

Friend Class frmRoster
    Inherits System.Windows.Forms.Form

    Dim colLblIneligibles As clsLabelArray
    Dim colLblGames As clsLabelArray
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
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents cmdSubmit As System.Windows.Forms.Button
	Public WithEvents cmdMakeActiveP As System.Windows.Forms.Button
	Public WithEvents cmdMakeInactiveP As System.Windows.Forms.Button
	Public WithEvents cmdMakeActive As System.Windows.Forms.Button
	Public WithEvents cmdMakeInactive As System.Windows.Forms.Button
	Public WithEvents lstInactivePitchers As System.Windows.Forms.ListBox
	Public WithEvents lstActivePitchers As System.Windows.Forms.ListBox
	Public WithEvents lstInactivePP As System.Windows.Forms.ListBox
	Public WithEvents lstActivePP As System.Windows.Forms.ListBox
	Public WithEvents cmdDisplay As System.Windows.Forms.Button
	Public WithEvents lstTeams As System.Windows.Forms.ListBox
	Public WithEvents lblRosters As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
    'Public WithEvents lblGames As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    'Public WithEvents lblIneligible As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
    Friend WithEvents grpIneligibles As System.Windows.Forms.GroupBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmRoster))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdSubmit = New System.Windows.Forms.Button
        Me.cmdMakeActiveP = New System.Windows.Forms.Button
        Me.cmdMakeInactiveP = New System.Windows.Forms.Button
        Me.cmdMakeActive = New System.Windows.Forms.Button
        Me.cmdMakeInactive = New System.Windows.Forms.Button
        Me.lstInactivePitchers = New System.Windows.Forms.ListBox
        Me.lstActivePitchers = New System.Windows.Forms.ListBox
        Me.lstInactivePP = New System.Windows.Forms.ListBox
        Me.lstActivePP = New System.Windows.Forms.ListBox
        Me.cmdDisplay = New System.Windows.Forms.Button
        Me.lstTeams = New System.Windows.Forms.ListBox
        Me.lblRosters = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        'Me.lblGames = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        'Me.lblIneligible = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.grpIneligibles = New System.Windows.Forms.GroupBox
        'CType(Me.lblGames, System.ComponentModel.ISupportInitialize).BeginInit()
        'CType(Me.lblIneligible, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdSubmit
        '
        Me.cmdSubmit.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSubmit.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSubmit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSubmit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSubmit.Location = New System.Drawing.Point(616, 464)
        Me.cmdSubmit.Name = "cmdSubmit"
        Me.cmdSubmit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSubmit.Size = New System.Drawing.Size(81, 25)
        Me.cmdSubmit.TabIndex = 42
        Me.cmdSubmit.Text = "Submit"
        '
        'cmdMakeActiveP
        '
        Me.cmdMakeActiveP.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(0, Byte), CType(0, Byte))
        Me.cmdMakeActiveP.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdMakeActiveP.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdMakeActiveP.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdMakeActiveP.Image = CType(resources.GetObject("cmdMakeActiveP.Image"), System.Drawing.Image)
        Me.cmdMakeActiveP.Location = New System.Drawing.Point(432, 360)
        Me.cmdMakeActiveP.Name = "cmdMakeActiveP"
        Me.cmdMakeActiveP.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdMakeActiveP.Size = New System.Drawing.Size(33, 33)
        Me.cmdMakeActiveP.TabIndex = 14
        Me.cmdMakeActiveP.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'cmdMakeInactiveP
        '
        Me.cmdMakeInactiveP.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(0, Byte), CType(0, Byte))
        Me.cmdMakeInactiveP.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdMakeInactiveP.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdMakeInactiveP.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdMakeInactiveP.Image = CType(resources.GetObject("cmdMakeInactiveP.Image"), System.Drawing.Image)
        Me.cmdMakeInactiveP.Location = New System.Drawing.Point(432, 304)
        Me.cmdMakeInactiveP.Name = "cmdMakeInactiveP"
        Me.cmdMakeInactiveP.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdMakeInactiveP.Size = New System.Drawing.Size(33, 33)
        Me.cmdMakeInactiveP.TabIndex = 13
        Me.cmdMakeInactiveP.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'cmdMakeActive
        '
        Me.cmdMakeActive.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(0, Byte), CType(0, Byte))
        Me.cmdMakeActive.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdMakeActive.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdMakeActive.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdMakeActive.Image = CType(resources.GetObject("cmdMakeActive.Image"), System.Drawing.Image)
        Me.cmdMakeActive.Location = New System.Drawing.Point(432, 112)
        Me.cmdMakeActive.Name = "cmdMakeActive"
        Me.cmdMakeActive.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdMakeActive.Size = New System.Drawing.Size(33, 33)
        Me.cmdMakeActive.TabIndex = 12
        Me.cmdMakeActive.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'cmdMakeInactive
        '
        Me.cmdMakeInactive.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(0, Byte), CType(0, Byte))
        Me.cmdMakeInactive.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdMakeInactive.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdMakeInactive.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdMakeInactive.Image = CType(resources.GetObject("cmdMakeInactive.Image"), System.Drawing.Image)
        Me.cmdMakeInactive.Location = New System.Drawing.Point(432, 64)
        Me.cmdMakeInactive.Name = "cmdMakeInactive"
        Me.cmdMakeInactive.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdMakeInactive.Size = New System.Drawing.Size(33, 33)
        Me.cmdMakeInactive.TabIndex = 11
        Me.cmdMakeInactive.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'lstInactivePitchers
        '
        Me.lstInactivePitchers.BackColor = System.Drawing.SystemColors.Window
        Me.lstInactivePitchers.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstInactivePitchers.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstInactivePitchers.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstInactivePitchers.ItemHeight = 14
        Me.lstInactivePitchers.Location = New System.Drawing.Point(504, 280)
        Me.lstInactivePitchers.Name = "lstInactivePitchers"
        Me.lstInactivePitchers.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstInactivePitchers.Size = New System.Drawing.Size(129, 172)
        Me.lstInactivePitchers.Sorted = True
        Me.lstInactivePitchers.TabIndex = 10
        '
        'lstActivePitchers
        '
        Me.lstActivePitchers.BackColor = System.Drawing.SystemColors.Window
        Me.lstActivePitchers.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstActivePitchers.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstActivePitchers.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstActivePitchers.ItemHeight = 14
        Me.lstActivePitchers.Location = New System.Drawing.Point(272, 280)
        Me.lstActivePitchers.Name = "lstActivePitchers"
        Me.lstActivePitchers.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstActivePitchers.Size = New System.Drawing.Size(129, 172)
        Me.lstActivePitchers.Sorted = True
        Me.lstActivePitchers.TabIndex = 9
        '
        'lstInactivePP
        '
        Me.lstInactivePP.BackColor = System.Drawing.SystemColors.Window
        Me.lstInactivePP.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstInactivePP.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstInactivePP.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstInactivePP.ItemHeight = 14
        Me.lstInactivePP.Location = New System.Drawing.Point(504, 56)
        Me.lstInactivePP.Name = "lstInactivePP"
        Me.lstInactivePP.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstInactivePP.Size = New System.Drawing.Size(129, 172)
        Me.lstInactivePP.Sorted = True
        Me.lstInactivePP.TabIndex = 6
        '
        'lstActivePP
        '
        Me.lstActivePP.BackColor = System.Drawing.SystemColors.Window
        Me.lstActivePP.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstActivePP.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstActivePP.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstActivePP.ItemHeight = 14
        Me.lstActivePP.Location = New System.Drawing.Point(272, 56)
        Me.lstActivePP.Name = "lstActivePP"
        Me.lstActivePP.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstActivePP.Size = New System.Drawing.Size(129, 172)
        Me.lstActivePP.Sorted = True
        Me.lstActivePP.TabIndex = 5
        '
        'cmdDisplay
        '
        Me.cmdDisplay.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDisplay.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDisplay.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDisplay.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDisplay.Location = New System.Drawing.Point(168, 128)
        Me.cmdDisplay.Name = "cmdDisplay"
        Me.cmdDisplay.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDisplay.Size = New System.Drawing.Size(73, 33)
        Me.cmdDisplay.TabIndex = 3
        Me.cmdDisplay.Text = "Display"
        '
        'lstTeams
        '
        Me.lstTeams.BackColor = System.Drawing.SystemColors.Window
        Me.lstTeams.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstTeams.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstTeams.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstTeams.ItemHeight = 14
        Me.lstTeams.Location = New System.Drawing.Point(24, 128)
        Me.lstTeams.Name = "lstTeams"
        Me.lstTeams.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstTeams.Size = New System.Drawing.Size(121, 74)
        Me.lstTeams.Sorted = True
        Me.lstTeams.TabIndex = 2
        '
        'lblRosters
        '
        Me.lblRosters.BackColor = System.Drawing.SystemColors.Control
        Me.lblRosters.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblRosters.Font = New System.Drawing.Font("Arial", 13.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRosters.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRosters.Location = New System.Drawing.Point(192, 464)
        Me.lblRosters.Name = "lblRosters"
        Me.lblRosters.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblRosters.Size = New System.Drawing.Size(257, 25)
        Me.lblRosters.TabIndex = 41
        Me.lblRosters.Text = "Roster Size is"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(280, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(97, 33)
        Me.Label1.TabIndex = 39
        Me.Label1.Text = "Active Position Players"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(504, 256)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(121, 17)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "Inactive Pitchers"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(272, 256)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(113, 17)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Active Pitchers"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(512, 8)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(105, 33)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Inactive Postion Players"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(32, 96)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(65, 17)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Teams:"
        '
        'grpIneligibles
        '
        Me.grpIneligibles.Location = New System.Drawing.Point(24, 224)
        Me.grpIneligibles.Name = "grpIneligibles"
        Me.grpIneligibles.Size = New System.Drawing.Size(200, 232)
        Me.grpIneligibles.TabIndex = 43
        Me.grpIneligibles.TabStop = False
        Me.grpIneligibles.Text = "Ineligibles"
        '
        'frmRoster
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(708, 494)
        Me.Controls.Add(Me.grpIneligibles)
        Me.Controls.Add(Me.cmdSubmit)
        Me.Controls.Add(Me.cmdMakeActiveP)
        Me.Controls.Add(Me.cmdMakeInactiveP)
        Me.Controls.Add(Me.cmdMakeActive)
        Me.Controls.Add(Me.cmdMakeInactive)
        Me.Controls.Add(Me.lstInactivePitchers)
        Me.Controls.Add(Me.lstActivePitchers)
        Me.Controls.Add(Me.lstInactivePP)
        Me.Controls.Add(Me.lstActivePP)
        Me.Controls.Add(Me.cmdDisplay)
        Me.Controls.Add(Me.lstTeams)
        Me.Controls.Add(Me.lblRosters)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(4, 30)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmRoster"
        Me.Text = maxRosterSize.ToString & " Man Roster"
        'CType(Me.lblGames, System.ComponentModel.ISupportInitialize).EndInit()
        'CType(Me.lblIneligible, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As frmRoster
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmRoster
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmRoster()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	
	Private Sub cmdDisplay_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDisplay.Click
        Dim ds As DataSet = Nothing
        Dim teamName As String
        Dim sqlQuery As String = ""
        Dim injuryCount As Integer
        Dim hitTable As String
        Dim pitchTable As String
        Dim filDebug As Integer
        Dim DataAccess As New clsDataAccess(gstrSeason & "data")

        Try
            If Dir(gAppPath & "debug.bbb") <> Nothing Then
                filDebug = FreeFile()
                FileOpen(filDebug, gAppPath & "debug.out", OpenMode.Append)
                PrintLine(filDebug, "In cmdDisplayClick")
                FileClose(filDebug)
            End If

            'Clear form
            For i As Integer = 1 To 12
                colLblIneligibles.Item(i - 1).Text = ""
                colLblGames.Item(i - 1).Text = ""
            Next i
            lstActivePP.Items.Clear()
            lstInactivePP.Items.Clear()
            lstActivePitchers.Items.Clear()
            lstInactivePitchers.Items.Clear()

            If lstTeams.SelectedItems.Count < 1 Then
                Call MsgBox("Must select a team to display.", MsgBoxStyle.OkOnly)
                Exit Sub
            End If
            teamName = lstTeams.Items.Item(lstTeams.SelectedIndex).ToString

            If gbolNewPS Then
                'Grab regular season flags
                hitTable = "HITTINGSTATS"
                pitchTable = "PITCHINGSTATS"
            Else
                hitTable = gstrHittingTable
                pitchTable = gstrPitchingTable
            End If

            sqlQuery = "Select Name FROM " & hitTable & " " & "WHERE Team = '" & teamName & "' AND " & "Active = 1 AND " & "GamesP = 0"
            'If gAccessConnectStr = Nothing Then
            '    gAccessConnectStr = GetConnectString()
            'End If
            'Using accessConn As OleDb.OleDbConnection = New OleDb.OleDbConnection
            '    accessConn.ConnectionString = GetConnectString()
            '    accessConn.Open()
            ds = DataAccess.ExecuteDataSet(sqlQuery)
            For Each dr As DataRow In ds.Tables(0).Rows
                lstActivePP.Items.Add(dr.Item("name"))
            Next dr
            ds.Clear()
            sqlQuery = "Select Name, Inj FROM " & hitTable & " " & "WHERE Team = '" & teamName & "' AND " & "Active = 0 AND " & "GamesP = 0"
            ds = DataAccess.ExecuteDataSet(sqlQuery)
            For Each dr As DataRow In ds.Tables(0).Rows
                If CInt(dr.Item("Inj")) > 0 Then
                    colLblIneligibles.Item(injuryCount).Text = dr.Item("name").ToString
                    colLblGames.Item(injuryCount).Text = dr.Item("Inj").ToString
                    injuryCount += 1
                Else
                    lstInactivePP.Items.Add(dr.Item("name"))
                End If
            Next dr
            ds.Clear()

            sqlQuery = "Select Name FROM " & pitchTable & " " & "WHERE Team = '" & teamName & "' AND " & "Active = 1"
            ds = DataAccess.ExecuteDataSet(sqlQuery)
            For Each dr As DataRow In ds.Tables(0).Rows
                lstActivePitchers.Items.Add(dr.Item("name"))
            Next dr
            ds.Clear()

            sqlQuery = "Select Name, Inj FROM " & pitchTable & " " & "WHERE Team = '" & teamName & "' AND " & "Active = 0"
            ds = DataAccess.ExecuteDataSet(sqlQuery)
            For Each dr As DataRow In ds.Tables(0).Rows
                If CInt(dr.Item("Inj")) > 0 Then
                    colLblIneligibles.Item(injuryCount).Text = dr.Item("name").ToString
                    colLblGames.Item(injuryCount).Text = dr.Item("Inj").ToString
                    injuryCount += 1
                Else
                    lstInactivePitchers.Items.Add(dr.Item("name"))
                End If
            Next dr
            'End Using
        Catch ex As Exception
            Call MsgBox("cmdDisplay_Click " & ex.ToString)
        Finally
            If Not ds Is Nothing Then ds.Dispose()
        End Try
    End Sub

    ''' <summary>
    ''' activates a hitter
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    Private Sub cmdMakeActive_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdMakeActive.Click
        If lstInactivePP.SelectedIndex <> -1 Then
            lstActivePP.Items.Add(lstInactivePP.Items.Item(lstInactivePP.SelectedIndex).ToString)
            lstInactivePP.Items.RemoveAt(lstInactivePP.SelectedIndex)
            lblRosters.Text = "Current roster size is " & lstActivePP.Items.Count + lstActivePitchers.Items.Count
        End If
    End Sub

    ''' <summary>
    ''' activates a pitcher
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    Private Sub cmdMakeActiveP_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdMakeActiveP.Click
        If lstInactivePitchers.SelectedIndex <> -1 Then
            lstActivePitchers.Items.Add(lstInactivePitchers.Items.Item(lstInactivePitchers.SelectedIndex).ToString)
            lstInactivePitchers.Items.RemoveAt(lstInactivePitchers.SelectedIndex)
            lblRosters.Text = "Current roster size is " & lstActivePP.Items.Count + lstActivePitchers.Items.Count
        End If
    End Sub

    ''' <summary>
    ''' inactivates a hitter
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    Private Sub cmdMakeInactive_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdMakeInactive.Click
        If lstActivePP.SelectedIndex <> -1 Then
            lstInactivePP.Items.Add(lstActivePP.Items.Item(lstActivePP.SelectedIndex).ToString)
            lstActivePP.Items.RemoveAt(lstActivePP.SelectedIndex)
            lblRosters.Text = "Current roster size is " & lstActivePP.Items.Count + lstActivePitchers.Items.Count
        End If
    End Sub

    ''' <summary>
    ''' inactivates a pitcher
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    Private Sub cmdMakeInactiveP_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdMakeInactiveP.Click
        If lstActivePitchers.SelectedIndex <> -1 Then
            lstInactivePitchers.Items.Add(lstActivePitchers.Items.Item(lstActivePitchers.SelectedIndex).ToString)
            lstActivePitchers.Items.RemoveAt(lstActivePitchers.SelectedIndex)
            lblRosters.Text = "Current roster size is " & lstActivePP.Items.Count + lstActivePitchers.Items.Count
        End If
    End Sub

    ''' <summary>
    ''' saves all roster changes and exits roster form
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    Private Sub cmdSubmit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSubmit.Click
        Dim sqlQuery As String = ""
        Dim teamName As String
        Dim tableKey As String
        Dim sqlFields As String
        Dim sqlValues As String
        Dim filDebug As Integer
        Dim DataAccess As New clsDataAccess(gstrSeason & "data")

        Try
            If Dir(gAppPath & "debug.bbb") <> Nothing Then
                filDebug = FreeFile()
                FileOpen(filDebug, gAppPath & "debug.out", OpenMode.Append)
                PrintLine(filDebug, "In cmdSubmitClick")
                FileClose(filDebug)
            End If
            maxRosterSize = IIF(CInt(gstrSeason) >= 2021, 26, 25)
            twoWayPlayers = 0

            'Look for two-way players. They should not count against the roster max
            For i As Integer = 0 To lstActivePP.Items.Count - 1
                lstActivePP.Items.Item(i).ToString()
                If lstActivePitchers.FindStringExact(lstActivePP.Items.Item(i).ToString()) > 0 Then
                    twoWayPlayers += 1
                End If
            Next i

            If lstActivePP.Items.Count + lstActivePitchers.Items.Count - twoWayPlayers > maxRosterSize Then
                Call MsgBox("Roster size cannot exceed " & maxRosterSize.ToString & " players.", MsgBoxStyle.OkOnly)
                Exit Sub
            End If
            teamName = lstTeams.Items.Item(lstTeams.SelectedIndex).ToString
            'If gAccessConnectStr = Nothing Then
            '    gAccessConnectStr = GetConnectString()
            'End If
            'Using accessConn As OleDb.OleDbConnection = New OleDb.OleDbConnection
            '    accessConn.ConnectionString = GetConnectString()
            '    accessConn.Open()

            'Update Hitting stats
            'Actives
            For i As Integer = 0 To lstActivePP.Items.Count - 1
                tableKey = HandleQuotes(StripChar(lstActivePP.Items.Item(i).ToString & teamName, " "))
                If gbolNewPS Then
                    sqlFields = "ACTIVE,PLAYERID"
                    sqlValues = "1,'" & tableKey & "'"
                    sqlQuery = "INSERT INTO " & gstrHittingTable & " (" & sqlFields & ") VALUES (" & sqlValues & ")"
                Else
                    sqlQuery = "UPDATE " & gstrHittingTable & " SET ACTIVE = 1 WHERE " & "PLAYERID = '" & tableKey & "'"
                End If
                DataAccess.ExecuteScalar(sqlQuery)
            Next i

            'Inactives
            For i As Integer = 0 To lstInactivePP.Items.Count - 1
                tableKey = HandleQuotes(StripChar(lstInactivePP.Items.Item(i).ToString & teamName, " "))
                If gbolNewPS Then
                    sqlFields = "ACTIVE,INJ,PLAYERID"
                    sqlValues = "0,15,'" & tableKey & "'"
                    sqlQuery = "INSERT INTO " & gstrHittingTable & " (" & sqlFields & ") VALUES (" & sqlValues & ")"
                Else
                    sqlQuery = "UPDATE " & gstrHittingTable & " SET ACTIVE = 0, INJ = 15 WHERE " & "PLAYERID = '" & tableKey & "' AND " & "ACTIVE = 1 AND " & "INJ < 15"
                    DataAccess.ExecuteScalar(sqlQuery)
                    sqlQuery = "UPDATE " & gstrHittingTable & " SET ACTIVE = 0 WHERE " & "PLAYERID = '" & tableKey & "' AND " & "ACTIVE = 1 AND " & "INJ >= 15"
                End If
                DataAccess.ExecuteScalar(sqlQuery)
            Next i

            'Update Pitching stats
            'Actives
            For i As Integer = 0 To lstActivePitchers.Items.Count - 1
                tableKey = StripChar(lstActivePitchers.Items.Item(i).ToString & teamName, " ")
                tableKey = HandleQuotes(tableKey)
                If gbolNewPS Then
                    sqlFields = "ACTIVE,PLAYERID,LG1"
                    sqlValues = "1,'" & tableKey & "',0"
                    sqlQuery = "INSERT INTO " & gstrPitchingTable & " (" & sqlFields & ") VALUES (" & sqlValues & ")"
                Else
                    sqlQuery = "UPDATE " & gstrPitchingTable & " SET ACTIVE = 1 WHERE " & "PLAYERID = '" & tableKey & "'"
                End If
                DataAccess.ExecuteScalar(sqlQuery)
            Next i
            'Inactives
            For i As Integer = 0 To lstInactivePitchers.Items.Count - 1
                tableKey = StripChar(lstInactivePitchers.Items.Item(i).ToString & teamName, " ")
                tableKey = HandleQuotes(tableKey)
                If gbolNewPS Then
                    sqlFields = "INJ,ACTIVE,PLAYERID,LG1"
                    sqlValues = "15,0,'" & tableKey & "',0"
                    sqlQuery = "INSERT INTO " & gstrPitchingTable & " (" & sqlFields & ") VALUES (" & sqlValues & ")"
                Else
                    sqlQuery = "UPDATE " & gstrPitchingTable & " SET INJ = 15, ACTIVE = 0 WHERE " & "PLAYERID = '" & tableKey & "' AND " & "ACTIVE = 1 AND " & "INJ < 15"
                End If
                DataAccess.ExecuteScalar(sqlQuery)
            Next i
            'End Using
            Me.Close()
        Catch ex As Exception
            Call MsgBox("cmdSubmit_click " & ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' Loads the roster for each team
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    Private Sub frmRoster_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dim teamName As String

        Try
            colLblIneligibles = New clsLabelArray(Me.grpIneligibles)
            colLblGames = New clsLabelArray(Me.grpIneligibles)
            For j As Integer = 1 To 12
                colLblIneligibles.AddNewRosterLabel(0)
                colLblGames.AddNewRosterLabel(1)
            Next j
            teamName = Dir(gAppPath & "teams\" & gstrSeason & "\*.*", FileAttribute.Directory)
            While teamName <> Nothing
                If teamName.IndexOf(".") = -1 Then
                    lstTeams.Items.Add(teamName)
                End If
               teamName = Dir()
            End While
        Catch ex As Exception
            Call MsgBox("frmRoster_Load " & ex.ToString)
        End Try
    End Sub
End Class