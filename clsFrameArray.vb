Option Strict On
Option Explicit On

Public Class clsFrameArray
    Inherits System.Collections.CollectionBase
    'Private ReadOnly HostForm As System.Windows.Forms.Form
    Private ReadOnly hostform As System.Windows.Forms.Panel

    ''' <summary>
    ''' Create a new instance of the BCard class
    ''' </summary>
    ''' <param name="aGroupBox"></param>
    ''' <remarks></remarks>
    Public Sub AddNewBCard(ByRef aGroupBox As BCard)
        Try
            ' Add the button to the collection's internal list.
            Me.List.Add(aGroupBox)
            ' Add the button to the controls collection of the form 
            ' referenced by the HostForm field.
            HostForm.Controls.Add(aGroupBox)

            ' Set intial properties for the button object.

            Select Case Me.Count
                Case 1 To 7
                    aGroupBox.Top = 0
                    aGroupBox.Left = (Count - 1) * 136
                Case 8 To 14
                    aGroupBox.Top = 152
                    aGroupBox.Left = (Count - 8) * 136
                Case 15 To 21
                    aGroupBox.Top = 304
                    aGroupBox.Left = (Count - 15) * 136
                Case 22
                    aGroupBox.Top = 456
                    aGroupBox.Left = (Count - 22) * 136
            End Select
            aGroupBox.Tag = Me.Count
        Catch ex As Exception
            Call MsgBox("Error in AddNewBCard. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' Create a new instance of the PCard class
    ''' </summary>
    ''' <param name="objPCard"></param>
    ''' <remarks></remarks>
    Public Sub AddNewPCard(ByVal objPCard As PCard)
        Try
            ' Add the button to the collection's internal list.
            Me.List.Add(objPCard)
            ' Add the button to the controls collection of the form 
            ' referenced by the HostForm field.
            HostForm.Controls.Add(objPCard)

            ' Set intial properties for the button object.

            Select Case Me.Count
                Case 1 To 5
                    objPCard.Top = 24
                    objPCard.Left = (Count - 1) * 136 + 16
                Case 6 To 10
                    objPCard.Top = 176
                    objPCard.Left = (Count - 6) * 136 + 16
                Case 11 To 15
                    objPCard.Top = 328
                    objPCard.Left = (Count - 11) * 136 + 16
                Case 16 To 20
                    objPCard.Top = 480
                    objPCard.Left = (Count - 16) * 136 + 16
                Case 21 To 25
                    objPCard.Top = 632
                    objPCard.Left = (Count - 21) * 136 + 16
                Case 26 To 30
                    objPCard.Top = 784
                    objPCard.Left = (Count - 26) * 136 + 16
                Case 31 To 35
                    objPCard.Top = 936
                    objPCard.Left = (Count - 31) * 136 + 16
                    'Case 16
                    '    objPCard.Top = 480
                    '    objPCard.Left = 360
            End Select
            objPCard.Tag = Me.Count
        Catch ex As Exception
            Call MsgBox("Error in AddNewPCard. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

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
                HostForm.Controls.Remove(Me(Me.Count - 1))
                Me.List.RemoveAt(Me.Count - 1)
            End If
        Catch ex As Exception
            Call MsgBox("Error in Remove. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    '''  Remove the last button added to the array from the host form 
    '''  controls collection. Note the use of the default property in 
    '''  accessing the array.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub RemoveAll()
        Try
            ' Check to be sure there is a button to remove.
            While Me.Count > 0
                HostForm.Controls.Remove(Me(Me.Count - 1))
                Me.List.RemoveAt(Me.Count - 1)
            End While
        Catch ex As Exception
            Call MsgBox("Error in RemoveAll. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    'Public Sub New(ByVal host As System.Windows.Forms.Form)
    '    HostForm = host
    'End Sub
    Public Sub New(ByVal host As System.Windows.Forms.Panel)
        hostform = host
    End Sub

    Default Public ReadOnly Property Item(ByVal Index As Integer) As System.Windows.Forms.UserControl
        Get
            Return CType(Me.List.Item(Index), System.Windows.Forms.UserControl)
        End Get
    End Property
End Class
