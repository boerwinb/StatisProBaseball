<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAdvance2
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
        Me.lblQuestion = New System.Windows.Forms.Label()
        Me.cmdNo = New System.Windows.Forms.Button()
        Me.cmdYes = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lblQuestion
        '
        Me.lblQuestion.BackColor = System.Drawing.SystemColors.Control
        Me.lblQuestion.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblQuestion.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQuestion.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblQuestion.Location = New System.Drawing.Point(34, 44)
        Me.lblQuestion.Name = "lblQuestion"
        Me.lblQuestion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblQuestion.Size = New System.Drawing.Size(217, 33)
        Me.lblQuestion.TabIndex = 1
        '
        'cmdNo
        '
        Me.cmdNo.BackColor = System.Drawing.SystemColors.Control
        Me.cmdNo.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdNo.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdNo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdNo.Location = New System.Drawing.Point(162, 119)
        Me.cmdNo.Name = "cmdNo"
        Me.cmdNo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdNo.Size = New System.Drawing.Size(65, 25)
        Me.cmdNo.TabIndex = 4
        Me.cmdNo.Text = "No"
        Me.cmdNo.UseVisualStyleBackColor = False
        '
        'cmdYes
        '
        Me.cmdYes.BackColor = System.Drawing.SystemColors.Control
        Me.cmdYes.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdYes.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdYes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdYes.Location = New System.Drawing.Point(58, 119)
        Me.cmdYes.Name = "cmdYes"
        Me.cmdYes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdYes.Size = New System.Drawing.Size(65, 25)
        Me.cmdYes.TabIndex = 3
        Me.cmdYes.Text = "Yes"
        Me.cmdYes.UseVisualStyleBackColor = False
        '
        'frmAdvance2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 181)
        Me.Controls.Add(Me.cmdNo)
        Me.Controls.Add(Me.cmdYes)
        Me.Controls.Add(Me.lblQuestion)
        Me.Name = "frmAdvance2"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Advance Extra Base"
        Me.ResumeLayout(False)

    End Sub
    Public WithEvents lblQuestion As System.Windows.Forms.Label
    Public WithEvents cmdNo As System.Windows.Forms.Button
    Public WithEvents cmdYes As System.Windows.Forms.Button
End Class
