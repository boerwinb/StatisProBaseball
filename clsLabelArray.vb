Option Strict On
Option Explicit On

Public Class clsLabelArray
    Inherits System.Collections.CollectionBase
    Private ReadOnly HostGroupBox As System.Windows.Forms.GroupBox
    Private ReadOnly HostForm As System.Windows.Forms.Form


    ''' <summary>
    ''' Create a new instance of the Forms.Label class
    ''' These are for the labels on the 25 man roster form (frmRoster). 
    ''' </summary>
    ''' <param name="rosterType"> 0 - Ineligibles, 1 - Actives</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AddNewRosterLabel(ByVal rosterType As Integer) As System.Windows.Forms.Label
        Dim aLabel As New System.Windows.Forms.Label()

        Try
            ' Add the button to the collection's internal list.
            Me.List.Add(aLabel)
            ' Add the button to the controls collection of the form 
            ' referenced by the HostForm field.
            HostGroupBox.Controls.Add(aLabel)

            ' Set intial properties for the button object.

            Select Case rosterType
                Case 0
                    aLabel.Left = 8
                    aLabel.Top = (Me.Count - 1) * 16 + 32
                    aLabel.Tag = Me.Count
                    aLabel.Width = 97
                    aLabel.Height = 17
                    aLabel.Font = New System.Drawing.Font("Arial", 6, FontStyle.Regular)
                Case 1
                    aLabel.Left = 128
                    aLabel.Top = (Me.Count - 1) * 16 + 32
                    aLabel.Tag = Me.Count
                    aLabel.Width = 33
                    aLabel.Height = 17
                    aLabel.Font = New System.Drawing.Font("Arial", 6, FontStyle.Regular)
            End Select
        Catch ex As Exception
            Call MsgBox("Error in AddNewRosterLabel. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return aLabel
    End Function

    ''' <summary>
    ''' Adds new K label (strikeouts)
    ''' </summary>
    ''' <param name="homeIndicator"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AddNewKLabel(ByVal homeIndicator As Integer) As System.Windows.Forms.Label
        ' Create a new instance of the Button class.
        Dim aLabel As New System.Windows.Forms.Label()

        Try
            ' Add the button to the collection's internal list.
            Me.List.Add(aLabel)
            ' Add the button to the controls collection of the form 
            ' referenced by the HostForm field.
            HostForm.Controls.Add(aLabel)

            ' Set intial properties for the button object.

            Select Case homeIndicator
                Case conHome
                    aLabel.Left = (Me.Count - 1) * 16 + 152
                Case conVisitor
                    aLabel.Left = 816 - (Me.Count - 1) * 16
            End Select
            aLabel.Top = 344
            aLabel.Tag = Me.Count
            aLabel.Width = 15
            aLabel.Height = 17
            aLabel.Font = New System.Drawing.Font("Arial", 12, FontStyle.Bold)
            aLabel.BackColor = System.Drawing.Color.White
            aLabel.Text = "K"
            aLabel.Visible = False
        Catch ex As Exception
            Call MsgBox("Error in AddNewKLabel. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return aLabel
    End Function

    ''' <summary>
    ''' Add new position label
    ''' </summary>
    ''' <param name="homeIndicator"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AddNewPosLabel(ByVal homeIndicator As Integer) As System.Windows.Forms.Label
        Dim aLabel As New System.Windows.Forms.Label()

        Try
            ' Add the button to the collection's internal list.
            Me.List.Add(aLabel)
            ' Add the button to the controls collection of the form 
            ' referenced by the HostForm field.
            HostForm.Controls.Add(aLabel)

            ' Set intial properties for the button object.
            '136,920
            Select Case homeIndicator
                Case conHome
                    aLabel.Left = 136
                Case conVisitor
                    aLabel.Left = 920
            End Select
            aLabel.Top = (Me.Count - 1) * 22 + 412 '382
            aLabel.Tag = Me.Count
            aLabel.Width = 20
            aLabel.Height = 16
            aLabel.Font = New System.Drawing.Font("Arial", 8, FontStyle.Bold)
            aLabel.ForeColor = System.Drawing.Color.White
            aLabel.BackColor = System.Drawing.Color.Transparent
        Catch ex As Exception
            Call MsgBox("Error in AddNewPosLabel. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return aLabel
    End Function

    ''' <summary>
    ''' Add new batting side label
    ''' </summary>
    ''' <param name="homeIndicator"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AddNewBatSideLabel(ByVal homeIndicator As Integer) As System.Windows.Forms.Label
        Dim aLabel As New System.Windows.Forms.Label()

        Try
            ' Add the button to the collection's internal list.
            Me.List.Add(aLabel)
            ' Add the button to the controls collection of the form 
            ' referenced by the HostForm field.
            HostForm.Controls.Add(aLabel)

            ' Set intial properties for the button object.
            '136,920
            Select Case homeIndicator
                Case conHome
                    aLabel.Left = 162
                Case conVisitor
                    aLabel.Left = 946
            End Select
            aLabel.Top = (Me.Count - 1) * 22 + 412 '382
            aLabel.Tag = Me.Count
            aLabel.Width = 17
            aLabel.Height = 17
            aLabel.Font = New System.Drawing.Font("Arial", 8, FontStyle.Regular)
            aLabel.ForeColor = System.Drawing.Color.Yellow
            aLabel.BackColor = System.Drawing.Color.Transparent
        Catch ex As Exception
            Call MsgBox("Error in AddNewBatSideLabel. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return aLabel
    End Function

    ''' <summary>
    ''' Add new batter label
    ''' </summary>
    ''' <param name="homeIndicator"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AddNewBatterLabel(ByVal homeIndicator As Integer) As System.Windows.Forms.Label
        Dim aLabel As New System.Windows.Forms.Label

        Try
            ' Add the button to the collection's internal list.
            Me.List.Add(aLabel)
            ' Add the button to the controls collection of the form 
            ' referenced by the HostForm field.
            HostForm.Controls.Add(aLabel)

            ' Set intial properties for the button object.
            '136,920
            Select Case homeIndicator
                Case conHome
                    aLabel.Left = 8
                Case conVisitor
                    aLabel.Left = 792
            End Select
            aLabel.Top = (Me.Count - 1) * 22 + 412 '382
            aLabel.Tag = Me.Count
            aLabel.Width = 121
            aLabel.Height = 19
            aLabel.Font = New System.Drawing.Font("Arial", 8, FontStyle.Bold)
            aLabel.ForeColor = System.Drawing.Color.White
            aLabel.BackColor = System.Drawing.Color.Transparent
        Catch ex As Exception
            Call MsgBox("Error in AddNewBatterLabel. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return aLabel
    End Function

    ''' <summary>
    '''Remove the last button added to the array from the host form 
    '''controls collection. Note the use of the default property in 
    '''accessing the array.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Remove()
        Try
            ' Check to be sure there is a button to remove.
            If Me.Count > 0 Then
                HostGroupBox.Controls.Remove(Me(Me.Count - 1))
                Me.List.RemoveAt(Me.Count - 1)
            End If
        Catch ex As Exception
            Call MsgBox("Error in Remove. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Public Sub New(ByVal host As System.Windows.Forms.GroupBox)
        HostGroupBox = host
    End Sub

    Public Sub New(ByVal host As System.Windows.Forms.Form)
        HostForm = host
    End Sub

    Default Public ReadOnly Property Item(ByVal Index As Integer) As System.Windows.Forms.Label
        Get
            Return CType(Me.List.Item(Index), System.Windows.Forms.Label)
        End Get
    End Property

End Class
