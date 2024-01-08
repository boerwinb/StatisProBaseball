Option Strict On
Option Explicit On

Public Class clsRadioButtonArray
    Inherits System.Collections.CollectionBase
    Private ReadOnly HostGroupBox As System.Windows.Forms.GroupBox


    ''' <summary>
    ''' Create a new instance of the RadioButton class
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AddNewRadioButton() As System.Windows.Forms.RadioButton
        Dim aRadioButton As New System.Windows.Forms.RadioButton()

        Try
            ' Add the button to the collection's internal list.
            Me.List.Add(aRadioButton)
            ' Add the button to the controls collection of the form 
            ' referenced by the HostForm field.
            HostGroupBox.Controls.Add(aRadioButton)

            ' Set intial properties for the button object.

            aRadioButton.Left = 16
            aRadioButton.Top = (Me.Count - 1) * 19 + 14
            aRadioButton.Width = 120
            aRadioButton.Height = 17
            aRadioButton.Tag = Me.Count
        Catch ex As Exception
            Call MsgBox("Error in AddNewRadioButton. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return aRadioButton
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

    Default Public ReadOnly Property Item(ByVal Index As Integer) As System.Windows.Forms.RadioButton
        Get
            Return CType(Me.List.Item(Index), System.Windows.Forms.RadioButton)
        End Get
    End Property

    ''' <summary>
    ''' Remove the last control added to the array from the host form 
    ''' controls collection. Note the use of the default property in 
    ''' accessing the array.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub RemoveAll()
        Try
            ' Check to be sure there is a button to remove.
            While Me.Count > 0
                HostGroupBox.Controls.Remove(Me(Me.Count - 1))
                Me.List.RemoveAt(Me.Count - 1)
            End While
        Catch ex As Exception
            Call MsgBox("Error in RemoveAll. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub
End Class
