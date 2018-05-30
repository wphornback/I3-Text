Public Class DeviceDictionary
    Public Property DRoomsList As RoomsList
    Public Property DeviceList As DeviceList
    Public Property DevDictionary As Dictionary(Of String, Device) = New Dictionary(Of String, Device)
    Public Property DevDictionaryCount As Integer

    Public Sub New(aliasRFP As String, maxRTRFP As String)
        DRoomsList = New RoomsList(aliasRFP)
        DeviceList = New DeviceList(maxRTRFP)
        MergeList()
        BuildDevDictionary()
        GetDevDictionaryCount()
    End Sub
    Sub MergeList()
        Dim foundDevice As Boolean = False
        For Each Room In DRoomsList.RList
            For Each item In Room.RoomDevices
                For i = 0 To DeviceList.DListCount - 1
                    If item = DeviceList.DList(i).UnitID Then
                        DeviceList.DList(i).RoomsServed.Add(Room)
                        foundDevice = True
                    End If
                    If Room.RoomDevices.Count = 1 And foundDevice = True Then
                        GoTo nextRoom
                    End If
                Next
            Next
nextRoom:
            foundDevice = False
        Next

    End Sub
    Sub BuildDevDictionary()

        For Each unit In DeviceList.DList
            DevDictionary.Add(unit.UnitID, unit)
        Next
        DRoomsList.RList.Clear()
        DeviceList.DList.Clear()

    End Sub
    Sub GetDevDictionaryCount()
        DevDictionaryCount = DevDictionary.Count
    End Sub



End Class
