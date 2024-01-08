Option Strict On

Public Class clsTextBoxArray
    Inherits System.Collections.CollectionBase
    Private ReadOnly HostGroupBox As System.Windows.Forms.GroupBox
    Private ReadOnly HostForm As System.Windows.Forms.Form

    ''' <summary>
    ''' Create a new instance of the TextBox class.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AddNewTextBoxSettings() As System.Windows.Forms.TextBox
        Dim aTextBox As New System.Windows.Forms.TextBox()

        Try
            ' Add the button to the collection's internal list.
            Me.List.Add(aTextBox)
            ' Add the button to the controls collection of the form 
            ' referenced by the HostForm field.
            HostGroupBox.Controls.Add(aTextBox)

            ' Set intial properties for the button object.
            aTextBox.Left = 8
            aTextBox.Width = 121
            aTextBox.Height = 19
            aTextBox.Top = (Me.Count - 1) * 21 + 27
            aTextBox.Tag = Me.Count
        Catch ex As Exception
            Call MsgBox("Error in AddNewTextBoxSettings" & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return aTextBox
    End Function

    ''' <summary>
    ''' Create a new instance of the TextBox class for Line Scores.
    ''' </summary>
    ''' <param name="isHomeLineScore"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AddNewTextBoxLS(ByVal isHomeLineScore As Boolean) As System.Windows.Forms.TextBox
        ' Create a new instance of the TextBox class.
        Dim aTextBox As New System.Windows.Forms.TextBox()
        Dim offset As Integer = 0

        Try
            ' Add the button to the collection's internal list.
            Me.List.Add(aTextBox)
            ' Add the button to the controls collection of the form 
            ' referenced by the HostForm field.
            HostForm.Controls.Add(aTextBox)

            If isHomeLineScore Then
                'Set offset for Y-location value of text boxes
                offset = 23
            End If
            ' Set intial properties for the button object.
            If Me.Count > 10 Then
                'Handle Runs, Hits, Errors boxes
                aTextBox.Left = (Me.Count - 11) * 25 + 485
                aTextBox.Width = 25
            Else
                'Handle LineScore boxes
                aTextBox.Left = (Me.Count - 1) * 16 + 300
                aTextBox.Width = 16
            End If
            aTextBox.Height = 19
            aTextBox.Top = 27 + offset
            aTextBox.Tag = Me.Count
            aTextBox.Font = New System.Drawing.Font("MS Sans Serif", 8, FontStyle.Bold)
            aTextBox.TextAlign = HorizontalAlignment.Center
        Catch ex As Exception
            Call MsgBox("Error in AddNewTextBoxLS" & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return aTextBox
    End Function

    ''' <summary>
    ''' create a new instance of the TextBox class for bases objects
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AddNewTextBoxBase() As System.Windows.Forms.TextBox
        Dim aTextBox As New System.Windows.Forms.TextBox()

        Try
            ' Add the button to the collection's internal list.
            Me.List.Add(aTextBox)
            ' Add the button to the controls collection of the form 
            ' referenced by the HostForm field.
            HostForm.Controls.Add(aTextBox)

            Select Case Me.Count
                Case 1
                    aTextBox.Left = 528
                    aTextBox.Top = 392
                Case 2
                    aTextBox.Left = 472
                    aTextBox.Top = 344
                Case 3
                    aTextBox.Left = 416
                    aTextBox.Top = 392
            End Select
            aTextBox.Width = 17
            aTextBox.Height = 19
            aTextBox.Tag = Me.Count
        Catch ex As Exception
            Call MsgBox("Error in AddNewTextBoxBase" & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return aTextBox
    End Function

    ''' <summary>
    ''' Remove the last button added to the array from the host form controls collection.
    ''' Note the use of the default property in accessing the array.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Remove()
        ' Check to be sure there is a button to remove.
        If Me.Count > 0 Then
            HostForm.Controls.Remove(Me(Me.Count - 1))
            Me.List.RemoveAt(Me.Count - 1)
        End If
    End Sub

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

    'Public Sub ClickHandler(ByVal sender As Object, ByVal e As _
    '            System.EventArgs)
    '    MessageBox.Show("you have clicked button " & CType(CType(sender, _
    '                System.Windows.Forms.Button).Tag, String))
    'End Sub

    Public Sub New(ByVal host As System.Windows.Forms.GroupBox)
        HostGroupBox = host
    End Sub

    Public Sub New(ByVal host As System.Windows.Forms.Form)
        HostForm = host
    End Sub

    Default Public ReadOnly Property Item(ByVal Index As Integer) As _
        System.Windows.Forms.TextBox
        Get
            Return CType(Me.List.Item(Index), System.Windows.Forms.TextBox)
        End Get
    End Property
End Class
