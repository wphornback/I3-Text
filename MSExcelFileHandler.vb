Imports Excel = Microsoft.Office.Interop.Excel

Public Class MSExcelFileHandler
    Public Property App As New Excel.Application
    Public Property Worksheet As Excel.Worksheet
    Public Property Workbook As Excel.Workbook
    Public Property LastRow As Integer
    Public Sub New()

    End Sub

    Public Sub New(fPath As String)

        Workbook = App.Workbooks.Open(fPath)
        Worksheet = App.Sheets(1)

        'Set the last row for later use
        LastRow = Worksheet.UsedRange.Rows.Count

    End Sub


    Sub CloseFile()

        'Me.Workbook.SaveAs()
        Me.App.Workbooks.Close()
        Me.App.Quit()
        releaseObject(Me.App)

    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

End Class
