Public Class ScheduleEventList
    Public Property ScheduleEList As List(Of ScheduleEvent) = New List(Of ScheduleEvent)
    Public Property scheduleImport As MSExcelFileHandler
    Public Property ScheduleDeviceDictionary As DeviceDictionary
    Public Property NonMatchedLocations As List(Of String) = New List(Of String)


    Public Sub New(aliasR As String, mxR As String, schR As String)
        ScheduleDeviceDictionary = New DeviceDictionary(aliasR, mxR)
        scheduleImport = New MSExcelFileHandler(schR)
        BuildScheduleEList()
        scheduleImport.CloseFile()
        CreateScheduleDeviceList()

    End Sub

    Sub BuildScheduleEList()
        Dim evName As String
        Dim sDate As DateTime
        Dim eDate As DateTime
        Dim locations As String
        Dim locationsArray() As String
        Dim tempLocList As List(Of String) = New List(Of String)
        Dim tempLocString As String
        'Dim charTest As Char
        For i = 2 To scheduleImport.LastRow
            evName = scheduleImport.Worksheet.Cells(i, 1).Value
            sDate = scheduleImport.Worksheet.Cells(i, 2).Value
            eDate = scheduleImport.Worksheet.Cells(i, 3).Value
            locations = scheduleImport.Worksheet.Cells(i, 4).Value
            locationsArray = locations.Split(",")
            For j = 0 To locationsArray.Length - 1
                tempLocString = locationsArray(j)
                If tempLocString(0) = " " Then
                    tempLocString = tempLocString.Remove(0, 1)
                End If
                tempLocList.Add(tempLocString.Replace(vbCr, ""))
            Next
            CreateAndBuildEventScheduleList(evName, sDate, eDate, tempLocList)
            tempLocList.Clear()
        Next

    End Sub
    Sub CreateAndBuildEventScheduleList(eN As String, sD As DateTime, eD As DateTime, lL As List(Of String))
        Dim tempEvent As ScheduleEvent = New ScheduleEvent(eN, sD, eD, lL)
        ScheduleEList.Add(tempEvent)
    End Sub

    Sub CreateScheduleDeviceList()
        Dim locCounter As Integer = 0
        Dim tempDev As String = ""
        Dim devCount As Integer = ScheduleDeviceDictionary.DevDictionaryCount

        For Each schEvent In ScheduleEList ' This is the event
            For Each location In schEvent.LocationList 'This is the Event location List
                For Each dev In ScheduleDeviceDictionary.DevDictionary 'This is a device in the Dictionary
                    For Each room In dev.Value.RoomsServed 'Go through each room in the devices room list
                        If tempDev = "" Or Not tempDev = dev.Value.UnitID Then
                            If location = room.RoomID Then 'Check to see if room Name =  location Name
                                schEvent.ScheduleDeviceList.Add(dev.Value.UnitID) 'If True store the Device Name
                                tempDev = dev.Value.UnitID
                                GoTo nextLocation 'Check Next Location in the Schedule Event locations list
                            Else
                                For Each rmAlias In room.Aliases 'If Room Name doesn't match check Aliases
                                    If location = rmAlias Then 'Check to see if location matches Alias
                                        schEvent.ScheduleDeviceList.Add(dev.Value.UnitID) 'Store the Device ID
                                        tempDev = dev.Value.UnitID
                                        GoTo nextLocation 'Check Next Location in the Schedule Event locations list
                                    End If
                                Next
                            End If
                        End If
                    Next
                Next
nextLocation:
            Next
            tempDev = ""
            schEvent.ScheduleDeviceList.Sort()
            deleteRepeatingUnitID(schEvent.ScheduleDeviceList)
        Next

    End Sub

    Sub deleteRepeatingUnitID(uIDList As List(Of String))

        For i = uIDList.Count - 1 To 1 Step -1
            If uIDList(i) = uIDList(i - 1) Then
                uIDList.RemoveAt(i)
            End If
        Next

    End Sub
    Sub SortScheduleEventList()
        ScheduleEList.Sort(Function(x, y) x.StartDate.CompareTo(y.StartDate))
    End Sub



End Class
