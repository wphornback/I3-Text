Public Class DeviceList
    Public Property CORampTimeReport As MSExcelFileHandler
    Public Property DList As List(Of Device) = New List(Of Device)
    Public Property DListCount As Integer

    Public Sub New(fPath As String)
        CORampTimeReport = New MSExcelFileHandler(fPath)
        BuildDeviceList()
        GetDeviceListCount()
        CORampTimeReport.CloseFile()
    End Sub
    Sub BuildDeviceList()
        Dim dName As String = ""
        Dim today As Date = Now()
        Dim deviceRT As Integer
        Dim tDate1 As Date
        Dim tDate2 As Date
        Dim i As Integer = 1
        Dim tString As String = ""
        Dim foundURT As Boolean = False


        While i <= CORampTimeReport.LastRow
            tString = CORampTimeReport.Worksheet.Cells(i, 1).value
            If Integer.TryParse(tString, deviceRT) Then
                tDate1 = CORampTimeReport.Worksheet.Cells(i, 2).Value
                tDate2 = CORampTimeReport.Worksheet.Cells(i, 3).Value
                While tDate1 <= today And today >= tDate2
                    i += 1
                    tDate1 = CORampTimeReport.Worksheet.Cells(i, 2).Value
                    tDate2 = CORampTimeReport.Worksheet.Cells(i, 3).Value
                End While
                deviceRT = CORampTimeReport.Worksheet.Cells(i, 1).Value
                foundURT = True
            ElseIf tString = "" Then
                i += 1
            Else
                dName = tString
            End If

            If foundURT = True Then
                CreateAndAddDeviceToList(dName, deviceRT)
                tString = CORampTimeReport.Worksheet.Cells(i, 1).Value
                While Not tString = ""
                    i += 1
                    tString = CORampTimeReport.Worksheet.Cells(i, 1).Value
                End While

            End If
            foundURT = False
            i += 1
        End While

    End Sub

    Sub CreateAndAddDeviceToList(dN As String, dRT As Integer)
        Dim tempD As Device = New Device(dN, dRT)
        DList.Add(tempD)
    End Sub

    Sub GetDeviceListCount()
        DListCount = DList.Count
    End Sub
End Class
