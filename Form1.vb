Imports System.IO

Public Class Form1
    Dim maxRampTimeReport As String
    Dim roomAliasesReport As String
    Dim weeklyScheduleExport As String
    Dim files As List(Of String) = New List(Of String)
    Dim lastDir As String = ""
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim fd As New OpenFileDialog
        If lastDir = "" Then Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)

        With fd
            .Title = "Select required Files"
            .InitialDirectory = lastDir
            .Multiselect = True
        End With

        If fd.ShowDialog() = DialogResult.OK Then
            lastDir = Path.GetDirectoryName(fd.FileName)
            For Each mFile As String In fd.FileNames
                files.Add(mFile)
            Next
        End If

        maxRampTimeReport = files(2)
        roomAliasesReport = files(0)
        weeklyScheduleExport = files(1)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim testSchedule As ScheduleEventList = New ScheduleEventList(roomAliasesReport, maxRampTimeReport, weeklyScheduleExport)
        Dim _4CP As _4CPScheduleAdjust = New _4CPScheduleAdjust
        Dim locations As String = ""

        'Prepare to write 4CP adjusted Schedule to New File
        Dim _4CPAdjustedFile As MSExcelFileHandler = New MSExcelFileHandler
        _4CPAdjustedFile.Workbook = _4CPAdjustedFile.App.Workbooks.Add()
        _4CPAdjustedFile.Worksheet = _4CPAdjustedFile.Workbook.Sheets("Sheet1")
        _4CPAdjustedFile.Worksheet.Cells(1, 1) = "Event Name"
        _4CPAdjustedFile.Worksheet.Cells(1, 2) = "Start Date"
        _4CPAdjustedFile.Worksheet.Cells(1, 3) = "Start Time"
        _4CPAdjustedFile.Worksheet.Cells(1, 4) = "End Date"
        _4CPAdjustedFile.Worksheet.Cells(1, 5) = "End Time"
        _4CPAdjustedFile.Worksheet.Cells(1, 6) = "Locations"
        _4CPAdjustedFile.Worksheet.Cells(1, 1).EntireRow.Font.Bold = True

        _4CP.Calc4CPEvent(testSchedule)
        Dim first As Integer = 0
        Dim i As Integer = 2
        For Each evnt In testSchedule.ScheduleEList
            For Each room In evnt.LocationList
                If first = 0 Then
                    locations = room
                    first = 1
                Else
                    locations = locations + ", " + room
                End If
            Next
            _4CPAdjustedFile.Worksheet.Cells(i, 1) = evnt.EventName
            _4CPAdjustedFile.Worksheet.Cells(i, 2) = evnt.StartDate.ToShortDateString
            _4CPAdjustedFile.Worksheet.Cells(i, 3) = evnt.StartDate.ToShortTimeString
            _4CPAdjustedFile.Worksheet.Cells(i, 4) = evnt.EndDate.ToShortDateString
            _4CPAdjustedFile.Worksheet.Cells(i, 5) = evnt.EndDate.ToShortTimeString
            _4CPAdjustedFile.Worksheet.Cells(i, 6) = locations
            first = 0
            i += 1
            locations = ""
        Next
        Dim adjFN() As String = weeklyScheduleExport.Split(".")
        _4CPAdjustedFile.Workbook.SaveAs(adjFN(0) + " 4CP Adjusted")
        _4CPAdjustedFile.CloseFile()
    End Sub
End Class
