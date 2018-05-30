Public Class Room

    Public Property RoomID As String
    Public Property RoomDevices As List(Of String)
    Public Property Aliases As List(Of String)


    Public Sub New(rID As String, devList As List(Of String), aList As List(Of String))

        RoomID = rID
        SetDeviceList(devList)
        SetAliases(aList)

    End Sub
    Public Sub New(rM As Room)
        RoomID = rM.RoomID
        RoomDevices = rM.RoomDevices
        Aliases = rM.Aliases
    End Sub

    Sub SetAliases(aList As List(Of String))

        Aliases = New List(Of String)
        For Each item As String In aList
            Aliases.Add(item)
        Next

    End Sub

    Sub SetDeviceList(devList As List(Of String))

        RoomDevices = New List(Of String)
        For Each item As String In devList
            RoomDevices.Add(item)
        Next
    End Sub

End Class
