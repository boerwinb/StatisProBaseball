<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBaseThrow
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btn2nd = New System.Windows.Forms.Button()
        Me.btn3rd = New System.Windows.Forms.Button()
        Me.btnHome = New System.Windows.Forms.Button()
        Me.lblQuestion = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'btn2nd
        '
        Me.btn2nd.Location = New System.Drawing.Point(35, 62)
        Me.btn2nd.Name = "btn2nd"
        Me.btn2nd.Size = New System.Drawing.Size(75, 23)
        Me.btn2nd.TabIndex = 0
        Me.btn2nd.Text = "2nd"
        Me.btn2nd.UseVisualStyleBackColor = True
        '
        'btn3rd
        '
        Me.btn3rd.Location = New System.Drawing.Point(116, 62)
        Me.btn3rd.Name = "btn3rd"
        Me.btn3rd.Size = New System.Drawing.Size(75, 23)
        Me.btn3rd.TabIndex = 1
        Me.btn3rd.Text = "3rd"
        Me.btn3rd.UseVisualStyleBackColor = True
        '
        'btnHome
        '
        Me.btnHome.Location = New System.Drawing.Point(197, 62)
        Me.btnHome.Name = "btnHome"
        Me.btnHome.Size = New System.Drawing.Size(75, 23)
        Me.btnHome.TabIndex = 2
        Me.btnHome.Text = "Home"
        Me.btnHome.UseVisualStyleBackColor = True
        '
        'lblQuestion
        '
        Me.lblQuestion.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQuestion.Location = New System.Drawing.Point(55, 18)
        Me.lblQuestion.Name = "lblQuestion"
        Me.lblQuestion.Size = New System.Drawing.Size(217, 33)
        Me.lblQuestion.TabIndex = 3
        Me.lblQuestion.Text = "Choose which base to throw."
        '
        'frmBaseThrow
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 114)
        Me.Controls.Add(Me.lblQuestion)
        Me.Controls.Add(Me.btnHome)
        Me.Controls.Add(Me.btn3rd)
        Me.Controls.Add(Me.btn2nd)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmBaseThrow"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Defensive Option"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btn2nd As System.Windows.Forms.Button
    Friend WithEvents btn3rd As System.Windows.Forms.Button
    Friend WithEvents btnHome As System.Windows.Forms.Button
    Friend WithEvents lblQuestion As System.Windows.Forms.Label
End Class
