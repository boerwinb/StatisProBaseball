Option Strict On
Option Explicit On
Friend Class frmDebugCard
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
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents cmdGo As System.Windows.Forms.Button
	Public WithEvents txtCARDID As System.Windows.Forms.TextBox
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmDebugCard))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.cmdGo = New System.Windows.Forms.Button
		Me.txtCARDID = New System.Windows.Forms.TextBox
		Me.Label1 = New System.Windows.Forms.Label
		Me.Text = "FAC Card"
		Me.ClientSize = New System.Drawing.Size(185, 150)
		Me.Location = New System.Drawing.Point(4, 23)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmDebugCard"
		Me.cmdGo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdGo.Text = "Go"
		Me.cmdGo.Size = New System.Drawing.Size(41, 25)
		Me.cmdGo.Location = New System.Drawing.Point(80, 88)
		Me.cmdGo.TabIndex = 2
		Me.cmdGo.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdGo.BackColor = System.Drawing.SystemColors.Control
		Me.cmdGo.CausesValidation = True
		Me.cmdGo.Enabled = True
		Me.cmdGo.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdGo.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdGo.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdGo.TabStop = True
		Me.cmdGo.Name = "cmdGo"
		Me.txtCARDID.AutoSize = False
		Me.txtCARDID.Size = New System.Drawing.Size(41, 19)
		Me.txtCARDID.Location = New System.Drawing.Point(80, 48)
		Me.txtCARDID.Maxlength = 3
		Me.txtCARDID.TabIndex = 1
		Me.txtCARDID.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtCARDID.AcceptsReturn = True
		Me.txtCARDID.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtCARDID.BackColor = System.Drawing.SystemColors.Window
		Me.txtCARDID.CausesValidation = True
		Me.txtCARDID.Enabled = True
		Me.txtCARDID.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtCARDID.HideSelection = True
		Me.txtCARDID.ReadOnly = False
		Me.txtCARDID.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtCARDID.MultiLine = False
		Me.txtCARDID.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtCARDID.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtCARDID.TabStop = True
		Me.txtCARDID.Visible = True
		Me.txtCARDID.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtCARDID.Name = "txtCARDID"
		Me.Label1.Text = "Card ID:"
		Me.Label1.Size = New System.Drawing.Size(41, 17)
		Me.Label1.Location = New System.Drawing.Point(32, 48)
		Me.Label1.TabIndex = 0
		Me.Label1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.BackColor = System.Drawing.SystemColors.Control
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Controls.Add(cmdGo)
		Me.Controls.Add(txtCARDID)
		Me.Controls.Add(Label1)
	End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As frmDebugCard
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmDebugCard
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmDebugCard()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	
	Private Sub cmdGo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdGo.Click
        Dim cardIndex As Integer
        cardIndex = CType(txtCARDID.Text, Integer)
        If cardIndex < 1 Or cardIndex > 388 Or cardIndex = 147 Then
            Call MsgBox("Card ID is invalid. Please enter another.", MsgBoxStyle.Exclamation)
        Else
            gblDebugCardID = cardIndex
            Me.Close()
        End If
	End Sub
	
	Private Sub frmDebugCard_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		txtCARDID.Text = "162"
	End Sub
End Class