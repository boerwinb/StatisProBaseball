<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAdvanceFlyout
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
        Me.cmdYes = New System.Windows.Forms.Button()
        Me.cmdNo = New System.Windows.Forms.Button()
        Me.lblQuestion = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'cmdYes
        '
        Me.cmdYes.Location = New System.Drawing.Point(35, 159)
        Me.cmdYes.Name = "cmdYes"
        Me.cmdYes.Size = New System.Drawing.Size(75, 23)
        Me.cmdYes.TabIndex = 0
        Me.cmdYes.Text = "Yes"
        Me.cmdYes.UseVisualStyleBackColor = True
        '
        'cmdNo
        '
        Me.cmdNo.Location = New System.Drawing.Point(139, 159)
        Me.cmdNo.Name = "cmdNo"
        Me.cmdNo.Size = New System.Drawing.Size(75, 23)
        Me.cmdNo.TabIndex = 1
        Me.cmdNo.Text = "No"
        Me.cmdNo.UseVisualStyleBackColor = True
        '
        'lblQuestion
        '
        Me.lblQuestion.Location = New System.Drawing.Point(32, 66)
        Me.lblQuestion.Name = "lblQuestion"
        Me.lblQuestion.Size = New System.Drawing.Size(217, 33)
        Me.lblQuestion.TabIndex = 2
        Me.lblQuestion.Text = "Label1"
        '
        'frmAdvanceFlyout
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 261)
        Me.Controls.Add(Me.lblQuestion)
        Me.Controls.Add(Me.cmdNo)
        Me.Controls.Add(Me.cmdYes)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmAdvanceFlyout"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Advance on Flyout"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cmdYes As System.Windows.Forms.Button
    Friend WithEvents cmdNo As System.Windows.Forms.Button
    Friend WithEvents lblQuestion As System.Windows.Forms.Label
End Class
