Public Class _4CPScheduleAdjust
    Public Property _4CPNewEventsList As List(Of ScheduleEvent) = New List(Of ScheduleEvent)
    Public Property MxRT As Integer
    Sub Calc4CPEvent(Schedule As ScheduleEventList)
        Dim tempDevStr As String
        For Each sEvnt In Schedule.ScheduleEList
            If Is4CPMonth(sEvnt) = True Then
                If IsWeekDayEvent(sEvnt) = True Then
                    If Is4CPImpactedEvent(sEvnt) = True Then
                        tempDevStr = sEvnt.ScheduleDeviceList(0)
                        MxRT = Schedule.ScheduleDeviceDictionary.DevDictionary(tempDevStr).MaxRT
                        AdjustEventSchedule(sEvnt.StartDate, sEvnt.EndDate, sEvnt)
                    End If
                End If
            End If
        Next
        Add4CPAdjustedEventsToSchedule(Schedule)
    End Sub
    Function Is4CPMonth(sE As ScheduleEvent)
        Dim result As Boolean = False
        If sE.StartDate.Month > 5 And sE.StartDate.Month < 10 Then
            result = True
        End If
        Return result
    End Function
    Function IsWeekDayEvent(sE As ScheduleEvent)
        Dim result As Boolean = False
        If sE.StartDate.DayOfWeek > 0 And sE.StartDate.DayOfWeek < 6 Then
            result = True
        End If
        Return result
    End Function

    Function Is4CPImpactedEvent(sE As ScheduleEvent)
        Dim result As Boolean
        Dim sTA4CPPlusMxRT As Date = sE.StartDate.AddMinutes(MxRT * -1) 'start time after 4CP Plus Max Ramp Time
        Dim _4CPS As Date = Date.Parse(sE.StartDate.ToShortDateString + " 4:00 PM")
        Dim _4CPE As Date = Date.Parse(sE.StartDate.ToShortDateString + " 5:00 PM")
        If (sE.StartDate >= _4CPS And sE.StartDate <= _4CPE) Or
           (sE.EndDate > _4CPS And sE.EndDate <= _4CPE) Or
           (sE.StartDate < _4CPS And sE.EndDate > _4CPS) Or
           (sTA4CPPlusMxRT < _4CPE And sE.StartDate > _4CPE) Then
            result = True
        Else
            result = False
        End If
        Return result
    End Function
    Sub AdjustEventSchedule(sD As Date, eD As Date, sE As ScheduleEvent)
        Dim _4CPAdj As Integer
        Dim tempDate As Date
        Dim tempDate2 As Date
        Dim nEvntStrtTime As Date
        Dim nEvntEndTime As Date
        _4CPAdj = Determine4CPAdjustment(sD, eD)
        Dim _4CPSTime As String = " 4:00 PM"
        Dim _4CPETime As String = " 5:00 PM"
        Dim eB4_4CPSTime As String = " 3:45 PM"
        Select Case _4CPAdj
            Case 1 'Start time and End time are in 4CP window
                sE.StartDate = Date.Parse(sE.StartDate.ToShortDateString + eB4_4CPSTime)
                sE.EndDate = Date.Parse(sE.EndDate.ToShortDateString + _4CPSTime)
                sE.EventName = sE.EventName + " 4CP_Adj Ein4CP...C1"
            Case 2 'Start time in 4CP window and endtime after 4cp window
                tempDate = Date.Parse(sE.EndDate.ToShortDateString + _4CPETime).AddMinutes(MxRT)
                tempDate2 = sE.EndDate
                sE.StartDate = Date.Parse(sE.StartDate.ToShortDateString + eB4_4CPSTime)
                sE.EndDate = Date.Parse(sE.StartDate.ToShortDateString + _4CPSTime)
                If tempDate < tempDate2 Then
                    nEvntStrtTime = tempDate
                    nEvntEndTime = tempDate2
                Else
                    nEvntStrtTime = tempDate
                    nEvntEndTime = tempDate.AddMinutes(15)
                End If

                AddNewEventToTempScheduleList(sE, nEvntStrtTime, nEvntEndTime)
                sE.EventName = sE.EventName + " 4CP_Adj ST&ET...C2"
            Case 3 'Start time before 4CP and end time during 4CP window
                tempDate = sE.StartDate.AddMinutes(15)
                tempDate2 = Date.Parse(sE.StartDate.ToShortDateString + _4CPSTime)
                If tempDate < tempDate2 Then
                    sE.EndDate = tempDate2
                Else
                    sE.StartDate = Date.Parse(sE.StartDate.ToShortDateString + eB4_4CPSTime)
                    sE.EndDate = tempDate2
                End If
                sE.EventName = sE.EventName + " 4CP_Adj ET...C3"
            Case 4 'Start time before 4cp window and End Time after 4CP window
                nEvntEndTime = sE.EndDate
                tempDate = sE.StartDate.AddMinutes(15)
                tempDate2 = Date.Parse(sE.StartDate.ToShortDateString + _4CPSTime)
                If tempDate < tempDate2 Then
                    sE.EndDate = tempDate2
                Else
                    sE.StartDate = Date.Parse(sE.StartDate.ToShortDateString + eB4_4CPSTime)
                    sE.EndDate = tempDate2
                End If
                tempDate = Date.Parse(sE.EndDate.ToShortDateString + _4CPETime).AddMinutes(MxRT)
                If tempDate < nEvntEndTime Then
                    nEvntStrtTime = tempDate
                Else
                    nEvntStrtTime = tempDate
                    nEvntEndTime = tempDate.AddMinutes(15)
                End If
                AddNewEventToTempScheduleList(sE, nEvntStrtTime, nEvntEndTime)
                sE.EventName = sE.EventName + " 4CP_Adj ET...C4"
            Case 5 'Start time plus ramptime will push schedule start into 4CP
                sE.StartDate = Date.Parse(sE.StartDate.ToShortDateString + _4CPETime).AddMinutes(MxRT)
                If sE.StartDate > sE.EndDate Then 'In case new start time is after original end time
                    sE.EndDate = sE.StartDate.AddMinutes(15)
                End If
                sE.EventName = sE.EventName + " 4CP_Adj ST...C5"
        End Select
    End Sub
    Function Determine4CPAdjustment(sD As Date, eD As Date)
        Dim _4CPS As Date
        Dim _4CPE As Date
        Dim sTIn4CPT As Boolean
        Dim eTIn4CPT As Boolean
        Dim sTB4_4CP As Boolean
        Dim eTAft_4CP As Boolean
        Dim rTPushIn4CP As Boolean
        Dim result As Integer
        _4CPS = Date.Parse(sD.ToShortDateString + " 4:00 PM")
        _4CPE = Date.Parse(eD.ToShortDateString + " 5:00 PM")
        If sD >= _4CPS And sD <= _4CPE Then
            sTIn4CPT = True
        End If
        If eD >= _4CPS And eD <= _4CPE Then
            eTIn4CPT = True
        End If
        If sD < _4CPS Then
            sTB4_4CP = True
        End If
        If eD > _4CPE Then
            eTAft_4CP = True
        End If
        If sD.AddMinutes(MxRT * -1) < _4CPE Then
            rTPushIn4CP = True
        End If

        If sTIn4CPT And eTIn4CPT Then
            result = 1
        ElseIf sTIn4CPT And eTAft_4CP Then
            result = 2
        ElseIf sTB4_4CP And eTIn4CPT Then
            result = 3
        ElseIf sTB4_4CP And eTAft_4CP Then
            result = 4
        ElseIf rTPushIn4CP Then
            result = 5
        Else
            Console.WriteLine("4CP Case could not be determined...SOMETHING WENT WRONG")
        End If
        Return result
    End Function
    Sub AddNewEventToTempScheduleList(sE As ScheduleEvent, sD As Date, eD As Date)
        Dim tempE As ScheduleEvent = New ScheduleEvent(sE.EventName + " 4CP_Adj NE", sD, eD, sE.LocationList)
        _4CPNewEventsList.Add(tempE)

    End Sub
    Sub Add4CPAdjustedEventsToSchedule(s As ScheduleEventList)
        Dim tempEvent As ScheduleEvent
        For Each sch In _4CPNewEventsList
            tempEvent = New ScheduleEvent(sch.EventName, sch.StartDate, sch.EndDate, sch.LocationList)
            s.ScheduleEList.Add(tempEvent)
        Next
        s.SortScheduleEventList()
        _4CPNewEventsList.Clear()
    End Sub
End Class
