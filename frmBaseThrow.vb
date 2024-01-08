Public Class frmBaseThrow


    Private Sub btn2nd_Click(sender As Object, e As EventArgs) Handles btn2nd.Click
        Game.GetResultPtr.baseThrow = 2
        Me.Close()
    End Sub

    Private Sub btn3rd_Click(sender As Object, e As EventArgs) Handles btn3rd.Click
        Game.GetResultPtr.baseThrow = 3
        Me.Close()
    End Sub

    Private Sub btnHome_Click(sender As Object, e As EventArgs) Handles btnHome.Click
        Game.GetResultPtr.baseThrow = 4
        Me.Close()
    End Sub
End Class