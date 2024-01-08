Option Strict On
Option Explicit On

Imports system.text

Public Class frmSettings
    Inherits System.Windows.Forms.Form
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If Disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(Disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Public WithEvents cmdPlayBall As System.Windows.Forms.Button
    Public WithEvents cmdPitching As System.Windows.Forms.Button
    Public WithEvents cmdOtherTeam As System.Windows.Forms.Button
    Public WithEvents cmdRemove As System.Windows.Forms.Button
    Public WithEvents cmdAdd As System.Windows.Forms.Button
    Public WithEvents lstRoster As System.Windows.Forms.ListBox
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Friend WithEvents grpLineup As System.Windows.Forms.GroupBox
    Friend WithEvents pnlBatters As System.Windows.Forms.Panel
    Friend WithEvents BCard1 As StatisProBaseball.BCard
    Friend WithEvents BCard13 As StatisProBaseball.BCard
    Friend WithEvents BCard12 As StatisProBaseball.BCard
    Friend WithEvents BCard11 As StatisProBaseball.BCard
    Friend WithEvents BCard10 As StatisProBaseball.BCard
    Friend WithEvents BCard9 As StatisProBaseball.BCard
    Friend WithEvents BCard8 As StatisProBaseball.BCard
    Friend WithEvents BCard7 As StatisProBaseball.BCard
    Friend WithEvents BCard6 As StatisProBaseball.BCard
    Friend WithEvents BCard5 As StatisProBaseball.BCard
    Friend WithEvents BCard4 As StatisProBaseball.BCard
    Friend WithEvents BCard3 As StatisProBaseball.BCard
    Friend WithEvents BCard2 As StatisProBaseball.BCard
    Friend WithEvents BCard23 As StatisProBaseball.BCard
    Friend WithEvents BCard22 As StatisProBaseball.BCard
    Friend WithEvents BCard21 As StatisProBaseball.BCard
    Friend WithEvents BCard20 As StatisProBaseball.BCard
    Friend WithEvents BCard19 As StatisProBaseball.BCard
    Friend WithEvents BCard18 As StatisProBaseball.BCard
    Friend WithEvents BCard17 As StatisProBaseball.BCard
    Friend WithEvents BCard16 As StatisProBaseball.BCard
    Friend WithEvents BCard15 As StatisProBaseball.BCard
    Friend WithEvents BCard14 As StatisProBaseball.BCard
    Friend WithEvents BCard28 As StatisProBaseball.BCard
    Friend WithEvents BCard27 As StatisProBaseball.BCard
    Friend WithEvents BCard26 As StatisProBaseball.BCard
    Friend WithEvents BCard25 As StatisProBaseball.BCard
    Friend WithEvents BCard24 As StatisProBaseball.BCard
    Friend WithEvents dgUsage As System.Windows.Forms.DataGridView
    Public WithEvents cmdExit As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSettings))
        Me.cmdPlayBall = New System.Windows.Forms.Button
        Me.cmdPitching = New System.Windows.Forms.Button
        Me.cmdOtherTeam = New System.Windows.Forms.Button
        Me.cmdRemove = New System.Windows.Forms.Button
        Me.cmdAdd = New System.Windows.Forms.Button
        Me.lstRoster = New System.Windows.Forms.ListBox
        Me.grpLineup = New System.Windows.Forms.GroupBox
        Me.cmdExit = New System.Windows.Forms.Button
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
        Me.dgUsage = New System.Windows.Forms.DataGridView
        Me.pnlBatters.SuspendLayout()
        CType(Me.dgUsage, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdPlayBall
        '
        Me.cmdPlayBall.BackColor = System.Drawing.SystemColors.Control
        Me.cmdPlayBall.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdPlayBall.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPlayBall.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPlayBall.Location = New System.Drawing.Point(538, 581)
        Me.cmdPlayBall.Name = "cmdPlayBall"
        Me.cmdPlayBall.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdPlayBall.Size = New System.Drawing.Size(120, 32)
        Me.cmdPlayBall.TabIndex = 991
        Me.cmdPlayBall.Text = "Play Ball"
        Me.cmdPlayBall.UseVisualStyleBackColor = False
        '
        'cmdPitching
        '
        Me.cmdPitching.BackColor = System.Drawing.SystemColors.Control
        Me.cmdPitching.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdPitching.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPitching.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPitching.Location = New System.Drawing.Point(538, 533)
        Me.cmdPitching.Name = "cmdPitching"
        Me.cmdPitching.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdPitching.Size = New System.Drawing.Size(120, 32)
        Me.cmdPitching.TabIndex = 990
        Me.cmdPitching.Text = "Pitching"
        Me.cmdPitching.UseVisualStyleBackColor = False
        '
        'cmdOtherTeam
        '
        Me.cmdOtherTeam.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOtherTeam.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOtherTeam.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOtherTeam.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOtherTeam.Location = New System.Drawing.Point(538, 485)
        Me.cmdOtherTeam.Name = "cmdOtherTeam"
        Me.cmdOtherTeam.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOtherTeam.Size = New System.Drawing.Size(120, 32)
        Me.cmdOtherTeam.TabIndex = 989
        Me.cmdOtherTeam.Text = "Opponent Lineup"
        Me.cmdOtherTeam.UseVisualStyleBackColor = False
        '
        'cmdRemove
        '
        Me.cmdRemove.BackColor = System.Drawing.SystemColors.Control
        Me.cmdRemove.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdRemove.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRemove.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdRemove.Location = New System.Drawing.Point(282, 508)
        Me.cmdRemove.Name = "cmdRemove"
        Me.cmdRemove.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdRemove.Size = New System.Drawing.Size(41, 25)
        Me.cmdRemove.TabIndex = 11
        Me.cmdRemove.Text = "<--"
        Me.cmdRemove.UseVisualStyleBackColor = False
        '
        'cmdAdd
        '
        Me.cmdAdd.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAdd.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAdd.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAdd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAdd.Location = New System.Drawing.Point(282, 476)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAdd.Size = New System.Drawing.Size(41, 25)
        Me.cmdAdd.TabIndex = 10
        Me.cmdAdd.Text = "-->"
        Me.cmdAdd.UseVisualStyleBackColor = False
        '
        'lstRoster
        '
        Me.lstRoster.BackColor = System.Drawing.SystemColors.Window
        Me.lstRoster.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstRoster.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstRoster.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstRoster.ItemHeight = 14
        Me.lstRoster.Location = New System.Drawing.Point(50, 476)
        Me.lstRoster.Name = "lstRoster"
        Me.lstRoster.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstRoster.Size = New System.Drawing.Size(216, 200)
        Me.lstRoster.TabIndex = 0
        '
        'grpLineup
        '
        Me.grpLineup.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpLineup.Location = New System.Drawing.Point(338, 460)
        Me.grpLineup.Name = "grpLineup"
        Me.grpLineup.Size = New System.Drawing.Size(184, 240)
        Me.grpLineup.TabIndex = 993
        Me.grpLineup.TabStop = False
        Me.grpLineup.Text = "Lineup"
        '
        'cmdExit
        '
        Me.cmdExit.BackColor = System.Drawing.SystemColors.Control
        Me.cmdExit.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdExit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdExit.Location = New System.Drawing.Point(538, 629)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdExit.Size = New System.Drawing.Size(120, 32)
        Me.cmdExit.TabIndex = 992
        Me.cmdExit.Text = "Exit Application"
        Me.cmdExit.UseVisualStyleBackColor = False
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
        Me.pnlBatters.Location = New System.Drawing.Point(1, 1)
        Me.pnlBatters.Name = "pnlBatters"
        Me.pnlBatters.Size = New System.Drawing.Size(1026, 462)
        Me.pnlBatters.TabIndex = 994
        '
        'BCard28
        '
        Me.BCard28.Location = New System.Drawing.Point(864, 526)
        Me.BCard28.Name = "BCard28"
        Me.BCard28.Size = New System.Drawing.Size(144, 172)
        Me.BCard28.TabIndex = 27
        Me.BCard28.Tag = "28"
        '
        'BCard27
        '
        Me.BCard27.Location = New System.Drawing.Point(720, 526)
        Me.BCard27.Name = "BCard27"
        Me.BCard27.Size = New System.Drawing.Size(144, 172)
        Me.BCard27.TabIndex = 26
        Me.BCard27.Tag = "27"
        '
        'BCard26
        '
        Me.BCard26.Location = New System.Drawing.Point(576, 526)
        Me.BCard26.Name = "BCard26"
        Me.BCard26.Size = New System.Drawing.Size(144, 172)
        Me.BCard26.TabIndex = 25
        Me.BCard26.Tag = "26"
        '
        'BCard25
        '
        Me.BCard25.Location = New System.Drawing.Point(432, 526)
        Me.BCard25.Name = "BCard25"
        Me.BCard25.Size = New System.Drawing.Size(144, 172)
        Me.BCard25.TabIndex = 24
        Me.BCard25.Tag = "25"
        '
        'BCard24
        '
        Me.BCard24.Location = New System.Drawing.Point(288, 526)
        Me.BCard24.Name = "BCard24"
        Me.BCard24.Size = New System.Drawing.Size(144, 172)
        Me.BCard24.TabIndex = 23
        Me.BCard24.Tag = "24"
        '
        'BCard23
        '
        Me.BCard23.Location = New System.Drawing.Point(144, 526)
        Me.BCard23.Name = "BCard23"
        Me.BCard23.Size = New System.Drawing.Size(144, 172)
        Me.BCard23.TabIndex = 22
        Me.BCard23.Tag = "23"
        '
        'BCard22
        '
        Me.BCard22.Location = New System.Drawing.Point(0, 526)
        Me.BCard22.Name = "BCard22"
        Me.BCard22.Size = New System.Drawing.Size(144, 172)
        Me.BCard22.TabIndex = 21
        Me.BCard22.Tag = "22"
        '
        'BCard21
        '
        Me.BCard21.Location = New System.Drawing.Point(864, 354)
        Me.BCard21.Name = "BCard21"
        Me.BCard21.Size = New System.Drawing.Size(144, 172)
        Me.BCard21.TabIndex = 20
        Me.BCard21.Tag = "21"
        '
        'BCard20
        '
        Me.BCard20.Location = New System.Drawing.Point(720, 354)
        Me.BCard20.Name = "BCard20"
        Me.BCard20.Size = New System.Drawing.Size(144, 172)
        Me.BCard20.TabIndex = 19
        Me.BCard20.Tag = "20"
        '
        'BCard19
        '
        Me.BCard19.Location = New System.Drawing.Point(576, 354)
        Me.BCard19.Name = "BCard19"
        Me.BCard19.Size = New System.Drawing.Size(144, 172)
        Me.BCard19.TabIndex = 18
        Me.BCard19.Tag = "19"
        '
        'BCard18
        '
        Me.BCard18.Location = New System.Drawing.Point(432, 354)
        Me.BCard18.Name = "BCard18"
        Me.BCard18.Size = New System.Drawing.Size(144, 172)
        Me.BCard18.TabIndex = 17
        Me.BCard18.Tag = "18"
        '
        'BCard17
        '
        Me.BCard17.Location = New System.Drawing.Point(288, 354)
        Me.BCard17.Name = "BCard17"
        Me.BCard17.Size = New System.Drawing.Size(144, 172)
        Me.BCard17.TabIndex = 16
        Me.BCard17.Tag = "17"
        '
        'BCard16
        '
        Me.BCard16.Location = New System.Drawing.Point(144, 354)
        Me.BCard16.Name = "BCard16"
        Me.BCard16.Size = New System.Drawing.Size(144, 172)
        Me.BCard16.TabIndex = 15
        Me.BCard16.Tag = "16"
        '
        'BCard15
        '
        Me.BCard15.Location = New System.Drawing.Point(0, 354)
        Me.BCard15.Name = "BCard15"
        Me.BCard15.Size = New System.Drawing.Size(144, 172)
        Me.BCard15.TabIndex = 14
        Me.BCard15.Tag = "15"
        '
        'BCard14
        '
        Me.BCard14.Location = New System.Drawing.Point(864, 182)
        Me.BCard14.Name = "BCard14"
        Me.BCard14.Size = New System.Drawing.Size(144, 172)
        Me.BCard14.TabIndex = 13
        Me.BCard14.Tag = "14"
        '
        'BCard13
        '
        Me.BCard13.Location = New System.Drawing.Point(720, 182)
        Me.BCard13.Name = "BCard13"
        Me.BCard13.Size = New System.Drawing.Size(144, 172)
        Me.BCard13.TabIndex = 12
        Me.BCard13.Tag = "13"
        '
        'BCard12
        '
        Me.BCard12.Location = New System.Drawing.Point(576, 182)
        Me.BCard12.Name = "BCard12"
        Me.BCard12.Size = New System.Drawing.Size(144, 172)
        Me.BCard12.TabIndex = 11
        Me.BCard12.Tag = "12"
        '
        'BCard11
        '
        Me.BCard11.Location = New System.Drawing.Point(432, 182)
        Me.BCard11.Name = "BCard11"
        Me.BCard11.Size = New System.Drawing.Size(144, 172)
        Me.BCard11.TabIndex = 10
        Me.BCard11.Tag = "11"
        '
        'BCard10
        '
        Me.BCard10.Location = New System.Drawing.Point(288, 182)
        Me.BCard10.Name = "BCard10"
        Me.BCard10.Size = New System.Drawing.Size(144, 172)
        Me.BCard10.TabIndex = 9
        Me.BCard10.Tag = "10"
        '
        'BCard9
        '
        Me.BCard9.Location = New System.Drawing.Point(144, 182)
        Me.BCard9.Name = "BCard9"
        Me.BCard9.Size = New System.Drawing.Size(144, 172)
        Me.BCard9.TabIndex = 8
        Me.BCard9.Tag = "9"
        '
        'BCard8
        '
        Me.BCard8.Location = New System.Drawing.Point(0, 182)
        Me.BCard8.Name = "BCard8"
        Me.BCard8.Size = New System.Drawing.Size(144, 172)
        Me.BCard8.TabIndex = 7
        Me.BCard8.Tag = "8"
        '
        'BCard7
        '
        Me.BCard7.Location = New System.Drawing.Point(864, 10)
        Me.BCard7.Name = "BCard7"
        Me.BCard7.Size = New System.Drawing.Size(144, 172)
        Me.BCard7.TabIndex = 6
        Me.BCard7.Tag = "7"
        '
        'BCard6
        '
        Me.BCard6.Location = New System.Drawing.Point(720, 10)
        Me.BCard6.Name = "BCard6"
        Me.BCard6.Size = New System.Drawing.Size(144, 172)
        Me.BCard6.TabIndex = 5
        Me.BCard6.Tag = "6"
        '
        'BCard5
        '
        Me.BCard5.Location = New System.Drawing.Point(576, 10)
        Me.BCard5.Name = "BCard5"
        Me.BCard5.Size = New System.Drawing.Size(144, 172)
        Me.BCard5.TabIndex = 4
        Me.BCard5.Tag = "5"
        '
        'BCard4
        '
        Me.BCard4.Location = New System.Drawing.Point(432, 10)
        Me.BCard4.Name = "BCard4"
        Me.BCard4.Size = New System.Drawing.Size(144, 172)
        Me.BCard4.TabIndex = 3
        Me.BCard4.Tag = "4"
        '
        'BCard3
        '
        Me.BCard3.Location = New System.Drawing.Point(288, 10)
        Me.BCard3.Name = "BCard3"
        Me.BCard3.Size = New System.Drawing.Size(144, 172)
        Me.BCard3.TabIndex = 2
        Me.BCard3.Tag = "3"
        '
        'BCard2
        '
        Me.BCard2.Location = New System.Drawing.Point(144, 10)
        Me.BCard2.Name = "BCard2"
        Me.BCard2.Size = New System.Drawing.Size(144, 172)
        Me.BCard2.TabIndex = 1
        Me.BCard2.Tag = "2"
        '
        'BCard1
        '
        Me.BCard1.Location = New System.Drawing.Point(0, 10)
        Me.BCard1.Name = "BCard1"
        Me.BCard1.Size = New System.Drawing.Size(144, 172)
        Me.BCard1.TabIndex = 0
        Me.BCard1.Tag = "1"
        '
        'dgUsage
        '
        Me.dgUsage.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgUsage.Location = New System.Drawing.Point(688, 485)
        Me.dgUsage.Name = "dgUsage"
        Me.dgUsage.Size = New System.Drawing.Size(321, 191)
        Me.dgUsage.TabIndex = 995
        '
        'frmSettings
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1028, 713)
        Me.ControlBox = False
        Me.Controls.Add(Me.dgUsage)
        Me.Controls.Add(Me.pnlBatters)
        Me.Controls.Add(Me.grpLineup)
        Me.Controls.Add(Me.cmdPlayBall)
        Me.Controls.Add(Me.cmdPitching)
        Me.Controls.Add(Me.cmdOtherTeam)
        Me.Controls.Add(Me.cmdRemove)
        Me.Controls.Add(Me.cmdAdd)
        Me.Controls.Add(Me.lstRoster)
        Me.Controls.Add(Me.cmdExit)
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(4, 23)
        Me.Name = "frmSettings"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Home Lineup"
        Me.pnlBatters.ResumeLayout(False)
        CType(Me.dgUsage, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
#End Region

    Dim isLoadingPlayers As Boolean
    Dim arrBatCardLabelsCols(22) As clsLabelArray
    Dim colPlayerTextBoxes As clsTextBoxArray
    Dim colPosCombos As clsComboArray

    Private Sub cmdAdd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAdd.Click
        Me.AddPlayerToLineup(IIF(bolHomeActive, Home, Visitor))
    End Sub

    ''' <summary>
    ''' reset lineup view
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    Private Sub cmdClear_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Try
            For i As Integer = 0 To 8
                If colPlayerTextBoxes.Item(i).Text <> Nothing Then
                    colPlayerTextBoxes.Item(i).Text = ""
                    colPosCombos.Item(i).Items.Clear()
                    If bolHomeActive Then
                        Call Home.RemoveLineup(i + 1)
                    Else
                        Call Visitor.RemoveLineup(i + 1)
                    End If
                End If
            Next i
        Catch ex As Exception
            Call MsgBox("cmdClear_click " & ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' go to view other team
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    Private Sub cmdOtherTeam_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOtherTeam.Click

        Me.SavePositions(IIF(bolHomeActive, Home, Visitor))

        bolHomeActive = Not bolHomeActive
        'colPlayerTextBoxes = New clsTextBoxArray(Me.grpLineup)
        'colPosCombos = New clsComboArray(Me.grpLineup)
        If bolHomeActive Then
            Me.Text = Home.teamName & " Lineup"
            cmdOtherTeam.Text = Visitor.teamName & " Lineup"
            LoadTeam(Home, True)
        Else
            Me.Text = Visitor.teamName & " Lineup"
            cmdOtherTeam.Text = Home.teamName & " Lineup"
            LoadTeam(Visitor, False)
        End If
    End Sub

    ''' <summary>
    ''' view pitching roster
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    Private Sub cmdPitching_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPitching.Click
        Dim frmPitching As frmPitching

        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        Me.SavePositions(IIF(bolHomeActive, Home, Visitor))

        frmPitching = New frmPitching
        frmPitching.Show()
        Me.Cursor = System.Windows.Forms.Cursors.Arrow
        Me.Close()
    End Sub

    ''' <summary>
    ''' Resume or start game
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    Private Sub cmdPlayBall_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPlayBall.Click
        Dim frmMainForm As frmMain

        Me.SavePositions(IIF(bolHomeActive, Home, Visitor))

        If ValidData() Then
            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
            frmMainForm = New frmMain
            frmMainForm.Show()
            Me.Cursor = System.Windows.Forms.Cursors.Arrow
            Me.Close()
        Else
            Call MsgBox("Data invalid. Recheck lineups or select pitcher.", MsgBoxStyle.OkOnly)
        End If
    End Sub

    ''' <summary>
    ''' remove player from lineup
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    Private Sub cmdRemove_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdRemove.Click
        Dim i As Integer
        Dim bolFound As Boolean
        Dim pIndex As Integer

        Try
            While i < 9 And Not bolFound
                If colPlayerTextBoxes.Item(i).Text = Nothing Then
                    bolFound = True
                End If
                i += 1
            End While
            If Not bolFound Then i = 10
            If i > 1 Then
                If bolHomeActive Then
                    pIndex = Home.GetPlayerNum(i - 1)
                    lstRoster.Items.Add(Home.GetBatterPtr(pIndex).player)
                    Call Home.RemoveLineup(i - 1)
                Else
                    pIndex = Visitor.GetPlayerNum(i - 1)
                    lstRoster.Items.Add(Visitor.GetBatterPtr(pIndex).player)
                    Call Visitor.RemoveLineup(i - 1)
                End If
                colPlayerTextBoxes.Item(i - 2).Text = ""
                colPosCombos.Item(i - 2).Items.Clear()
            End If
        Catch ex As Exception
            Call MsgBox("cmdRemove_click " & ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' loads the Settings form
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    Private Sub frmSettings_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        bolHomeActive = True
        colPlayerTextBoxes = New clsTextBoxArray(Me.grpLineup)
        colPosCombos = New clsComboArray(Me.grpLineup)

        Me.Text = Home.teamName & " Lineup"
        cmdOtherTeam.Text = Visitor.teamName & " Lineup"
        Me.LoadTeam(Home, True)
        Me.Show()
    End Sub

    ''' <summary>
    ''' fills the player cards on the form
    ''' </summary>
    ''' <param name="Team"></param>
    ''' <param name="playerIndex"></param>
    ''' <param name="cardIndex"></param>
    ''' <remarks></remarks>
    Private Sub FillCards(ByRef Team As clsTeam, ByRef playerIndex As Integer, ByVal cardIndex As Integer)
        'Dim playerCardTag As Integer
        Dim objBCard As BCard = Nothing

        Try
            'playerCardTag = playerIndex + 1
            'If playerIndex >= 20 Then
            '    playerIndex += 0
            'End If
            For i As Integer = 0 To Me.pnlBatters.Controls.Count - 1
                'If Me.Controls(i).Tag.ToString = playerCardTag.ToString Then
                If Me.pnlBatters.Controls(i).Tag.ToString = cardIndex.ToString Then
                    objBCard = CType(Me.pnlBatters.Controls(i), BCard)
                    Exit For
                End If
            Next i
            With Team.GetBatterPtr(playerIndex + 1)
                If playerIndex <= Team.hitters - 1 Then
                    'If .player <> Nothing Then
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
                    If .BatStatPtr.GamesInj > 0 Then
                        .available = False
                        objBCard.BackColor = System.Drawing.Color.Red
                    ElseIf Not .BatStatPtr.Active Then
                        .available = False
                        objBCard.Enabled = False
                        objBCard.BackColor = System.Drawing.Color.Pink
                    ElseIf .hitLast And .playedLast Then
                        objBCard.BackColor = System.Drawing.Color.Cyan
                    ElseIf .playedLast Then
                        objBCard.BackColor = System.Drawing.Color.Yellow
                    Else
                        objBCard.BackColor = System.Drawing.Color.White
                    End If
                    If .available Then
                        lstRoster.Items.Add(.player)
                    End If
                    objBCard.Enabled = False
                Else
                    objBCard.Visible = False
                End If
            End With
        Catch ex As Exception
            Call MsgBox("FillCards " & ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' Loads the team roster cards
    ''' </summary>
    ''' <param name="Team"></param>
    ''' <remarks></remarks>
    Private Sub LoadTeam(ByRef Team As clsTeam, ByVal isHomeTeam As Boolean)
        Dim sqlQuery As New StringBuilder
        Dim hitTable As String
        Dim ds As New DataSet
        Dim bs As New BindingSource
        Dim DataAccess As New clsDataAccess(gstrSeason & "data")

        Try
            colPlayerTextBoxes.RemoveAll()
            colPosCombos.RemoveAll()
            For i As Integer = 1 To 9
                colPlayerTextBoxes.AddNewTextBoxSettings()
                colPosCombos.AddNewComboSettings()
            Next i
            lstRoster.Items.Clear()
            For i As Integer = 0 To conMaxBatters - 1
                Call FillCards(Team, i, i + 1)
            Next

            'Add Pitcher if NL
            If Not bolAmericanLeagueRules Then
                Call LoadBattingCard(Team)
                lstRoster.Items.Add(Team.GetBatterPtr(Team.hitters + Team.nthPitcherUsed).player)
            End If
            For i As Integer = 1 To 9
                If colPosCombos(i - 1).Items.Count > 0 Then
                    colPosCombos(i - 1).Items.Clear()
                End If
                If Team.GetPlayerNum(i) > 0 Then
                    colPlayerTextBoxes.Item(i - 1).Text = Team.GetBatterPtr(Team.GetPlayerNum(i)).player
                    Call FillPositions(i - 1, Team.GetPlayerNum(i), Team)
                    colPosCombos(i - 1).SelectedIndex = SetListIndex(colPosCombos(i - 1), (Team.GetBatterPtr(Team.GetPlayerNum(i)).position))
                Else
                    colPlayerTextBoxes.Item(i - 1).Text = ""
                End If
            Next i

            'Load usage data grid
            hitTable = IIF(gbolPostSeason, "PSHITTINGSTATS", "HITTINGSTATS")
            sqlQuery.Append("SELECT name, gamesc as C, games1b as '1B', games2b as '2B', gamesss as SS, games3b as '3B', " & _
                            "gamesof as OFD, gamesdh as DH ")
            sqlQuery.Append("FROM " & hitTable & " " & "WHERE team = '" & Team.teamName & "' and games > gamesP")
            ' and games > gamesP
            'If gAccessConnectStr = Nothing Then
            '    gAccessConnectStr = GetConnectString()
            'End If
            'Using accessConn As OleDb.OleDbConnection = New OleDb.OleDbConnection
            '    accessConn.ConnectionString = GetConnectString()
            '    accessConn.Open()
            ds = DataAccess.ExecuteDataSet(sqlQuery.ToString)
            'End Using
            bs.DataSource = ds
            bs.DataMember = ds.Tables(0).TableName
            dgUsage.DataSource = bs
            dgUsage.AutoGenerateColumns = True
            dgUsage.AutoResizeColumns()

        Catch ex As Exception
            Call MsgBox("FillCards " & ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' Fills position data into position combo boxes
    ''' </summary>
    ''' <param name="lineupIndex"></param>
    ''' <param name="playerIndex"></param>
    ''' <param name="Team"></param>
    ''' <remarks></remarks>
    Private Sub FillPositions(ByRef lineupIndex As Integer, ByRef playerIndex As Integer, ByRef Team As clsTeam)
        Dim indexPosition As Integer
        Dim fieldLine As String
        Dim fieldPosition As String = ""
        Dim hasDH As Boolean

        Try
            fieldLine = Team.GetBatterPtr(playerIndex).field
            fieldLine = "*" & fieldLine 'Add a dummy text buffer
            indexPosition = fieldLine.IndexOf("-")
            While indexPosition <> -1
                fieldPosition = fieldLine.Substring(indexPosition - 2, 2)
                Select Case fieldPosition.Substring(fieldPosition.Length - 1)
                    Case "C"
                        fieldPosition = "C"
                    Case "P"
                        fieldPosition = "P"
                End Select
                If fieldPosition = "OF" Then
                    colPosCombos.Item(lineupIndex).Items.Add("LF")
                    colPosCombos.Item(lineupIndex).Items.Add("CF")
                    colPosCombos.Item(lineupIndex).Items.Add("RF")
                Else
                    colPosCombos.Item(lineupIndex).Items.Add(fieldPosition)
                    If fieldPosition = "DH" Then
                        hasDH = True
                    End If
                End If
                fieldLine = fieldLine.Substring(indexPosition + 1)
                indexPosition = fieldLine.IndexOf("-")
            End While
            If Not bolAmericanLeagueRules And fieldPosition = Nothing Then
                'Hard code a P for batting pitcher
                colPosCombos.Item(lineupIndex).Items.Add("P")
            End If

            If gstrPostSeason = conWS And bolAmericanLeagueRules And Not hasDH Then
                'Allow DH for all players
                colPosCombos.Item(lineupIndex).Items.Add("DH")
            End If
        Catch ex As Exception
            Call MsgBox("FillPositions " & ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' checks to make sure all 9 fielding positions are accounted for
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ValidPositions() As Boolean
        Dim isValid As Boolean = True
        Dim positionList As String
        Dim stringIndex As Integer
        Dim tempValue As String
        Dim mustAbort As Boolean

        Try
            positionList = IIF(bolAmericanLeagueRules, conALPos, conNLPos)
            For i As Integer = 0 To 8
                If colPosCombos.Item(i).Text = Nothing Then
                    mustAbort = True
                End If
            Next i
            If Not mustAbort Then
                stringIndex = positionList.IndexOf("|")
                While stringIndex > -1 And isValid
                    tempValue = positionList.Substring(0, stringIndex)
                    isValid = False
                    For i As Integer = 0 To 8
                        If colPosCombos.Item(i).Text.Trim = tempValue.Trim Then
                            isValid = True
                        End If
                    Next i
                    positionList = positionList.Substring(stringIndex + 1)
                    stringIndex = positionList.IndexOf("|")
                End While
            End If
            Return isValid
        Catch ex As Exception
            Call MsgBox("ValidPositions " & ex.ToString)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' save position changes
    ''' </summary>
    ''' <param name="Team"></param>
    ''' <remarks></remarks>
    Private Sub SavePositions(ByRef Team As clsTeam)
        Try
            If colPlayerTextBoxes.Item(8).Text = "" Then
                'all players have not been entered, so do not save the positions
                Exit Sub
            End If

            For i As Integer = 0 To 8
                With Team.GetBatterPtr(Team.GetPlayerNum(i + 1))
                    If .position = "DH" And _
                            colPosCombos.Item(i).Text <> "DH" And _
                            bolStartGame Then
                        'Once a player enters the game as a DH, they 
                        'cannot be moved to another position
                        colPosCombos.Item(i).SelectedIndex = _
                            SetListIndex(colPosCombos.Item(i), "DH")
                    Else
                        .position = colPosCombos.Item(i).Text
                    End If
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
            Call MsgBox("SavePositions " & ex.ToString)
        End Try
    End Sub

    Private Sub AddPlayerToLineup(ByRef Team As clsTeam)
        Dim isFound As Boolean
        Dim i As Integer
        Dim playerName As String
        Dim playerIndex As Integer

        Try
            If lstRoster.SelectedIndex < 0 Then
                'do nothing
                Exit Sub
            End If
            isLoadingPlayers = True
            While i < 9 And Not isFound
                If colPlayerTextBoxes.Item(i).Text = Nothing Then
                    isFound = True
                End If
                i += 1
            End While
            If isFound Then
                playerName = lstRoster.Items(lstRoster.SelectedIndex).ToString()
                playerIndex = Team.GetIndexFromName(playerName)
                'With Team.GetBatterPtr(lstRoster.SelectedIndex + 1)
                With Team.GetBatterPtr(playerIndex)
                    If Not Team.InLineup(playerIndex) And .BatStatPtr.GamesInj = 0 And .BatStatPtr.Active Then
                        colPlayerTextBoxes.Item(i - 1).Text = .player
                        'Call FillPositions(i - 1, lstRoster.SelectedIndex + 1, Team)
                        Call FillPositions(i - 1, playerIndex, Team)
                        colPosCombos.Item(i - 1).SelectedIndex = 0
                        .position = colPosCombos.Item(i - 1).Text
                        .errorRating = GetRating(.field, .position, "E")
                        .throwRating = GetRating(.field, .position, "T")
                        If .position.Length > 0 And .cd.Length > 0 Then
                            If .cd.IndexOf(.position) > -1 Or (.position.Substring(.position.Length - 1) = "F" And .cd.Substring(.cd.Length - 1) = "F") Then
                                .cdAct = (Val(.cd)).ToString
                            Else
                                .cdAct = "0"
                            End If
                        End If
                        If Not bolAmericanLeagueRules And .position = "P" Then
                            Call Team.AddLineup(i, Team.hitters + Team.nthPitcherUsed)
                            colPlayerTextBoxes.Item(i - 1).Text = Team.GetBatterPtr(Team.hitters + Team.nthPitcherUsed).player
                        Else
                            Call Team.AddLineup(i, playerIndex)
                        End If
                    End If
                End With
            End If
            lstRoster.Items.RemoveAt((lstRoster.SelectedIndex))
            isLoadingPlayers = False
        Catch ex As Exception
            Call MsgBox("AddPlayerToLineup " & ex.ToString)
        End Try
    End Sub


    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        If MessageBox.Show("Exit Application?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            Application.Exit()
        End If
    End Sub

End Class