Option Strict On
Option Explicit On

Public Class clsComboArray
    Inherits System.Collections.CollectionBase
    Private ReadOnly HostGroupBox As System.Windows.Forms.GroupBox

    ''' <summary>
    ''' Create a new instance of the ComboBox class.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AddNewCombo() As System.Windows.Forms.ComboBox
        Dim aCombo As New System.Windows.Forms.ComboBox()

        Try
            ' Add the ComboBox to the collection's internal list.
            Me.List.Add(aCombo)
            ' Add the ComboBox to the controls collection of the form 
            ' referenced by the HostForm field.
            HostGroupBox.Controls.Add(aCombo)

            ' Set intial properties for the button object.

            aCombo.Left = 144
            aCombo.Top = (Me.Count - 1) * 19 + 14
            aCombo.Width = 41
            aCombo.Height = 22
            aCombo.Tag = Me.Count
            aCombo.DropDownStyle = ComboBoxStyle.DropDownList
        Catch ex As Exception
            Call MsgBox("Error in AddNewCombo. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return aCombo
    End Function

    ''' <summary>
    ''' Create a new instance of the ComboBox class.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AddNewComboSettings() As System.Windows.Forms.ComboBox
        Dim aCombo As New System.Windows.Forms.ComboBox()

        Try
            ' Add the ComboBox to the collection's internal list.
            Me.List.Add(aCombo)
            ' Add the ComboBox to the controls collection of the form 
            ' referenced by the HostForm field.
            HostGroupBox.Controls.Add(aCombo)

            ' Set intial properties for the button object.

            aCombo.Left = 138
            aCombo.Width = 41
            aCombo.Height = 22
            aCombo.Top = (Me.Count - 1) * 21 + 27
            aCombo.Tag = Me.Count
            aCombo.DropDownStyle = ComboBoxStyle.DropDownList
        Catch ex As Exception
            Call MsgBox("Error in AddNewComboSetting. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return aCombo
    End Function

    ''' <summary>
    ''' Remove the last button added to the array from the host form 
    ''' controls collection. Note the use of the default property in 
    ''' accessing the array.
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
            Call MsgBox("Error in Remove(). " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Public Sub New(ByVal host As System.Windows.Forms.GroupBox)
        HostGroupBox = host
    End Sub

    Default Public ReadOnly Property Item(ByVal Index As Integer) As System.Windows.Forms.ComboBox
        Get
            Return CType(Me.List.Item(Index), System.Windows.Forms.ComboBox)
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
            Call MsgBox("Error in RemoveAll(). " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub
End Class
