Public Class ScheduleEvent
    Public Property EventName As String
    Public Property StartDate As DateTime
    Public Property StartTime As DateTime
    Public Property EndDate As DateTime
    Public Property EndTime As DateTime
    Public Property LocationList As List(Of String) = New List(Of String)
    Public Property ScheduleDeviceList As List(Of String) = New List(Of String)

    Public Sub New(eN As String, sDaT As DateTime, eDaT As DateTime, lL As List(Of String))
        EventName = eN
        StartDate = sDaT
        EndDate = eDaT
        AddLocationsToList(lL)
    End Sub

    Public Sub New(eN As String, sD As DateTime, sT As DateTime, eD As DateTime, eT As DateTime, lL As List(Of String))
        EventName = eN
        StartDate = sD
        StartTime = sT
        EndDate = eD
        EndTime = eT
        AddLocationsToList(lL)
    End Sub
    Sub AddLocationsToList(tempList As List(Of String))
        For Each item In tempList
            LocationList.Add(item)
        Next
    End Sub



End Class
