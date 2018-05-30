Public Class Device

    Public Property UnitID As String
    Public Property MaxRT As Integer
    Public Property RoomsServed As List(Of Room) = New List(Of Room)

    Public Sub New(uID As String)
        UnitID = uID
    End Sub
    Public Sub New(uID As String, mxRT As Integer)

        UnitID = uID
        MaxRT = mxRT

    End Sub

End Class
