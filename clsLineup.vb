Option Strict On
Option Explicit On
Friend Class clsLineup
    Private _Hitters(9) As String
	
    Public Sub Add(ByVal index As Integer, ByVal playerName As String)
        _Hitters(index) = playerName
    End Sub
    Public Sub RemoveLineup(ByVal index As Integer)
        _Hitters(index) = ""
    End Sub
End Class