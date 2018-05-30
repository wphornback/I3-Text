Public Class RoomsList
    'Create a list of Rooms from the CO Alias Report 
    Public Property RList As List(Of Room) = New List(Of Room)
    'Create an instance of the MSExcelClass to Open the CO Alias
    Public Property COAliasReport As MSExcelFileHandler
    'Set a counter to use for varification that the rooms list has every room from the CO Alias Report
    Public Property RoomsListCount As Integer
    'This sub routine is the one that builds the room list from the CO Alias report
    Public Sub New(fPath As String)
        'Create an instance of the MS Excel Interop class and open the excel file
        COAliasReport = New MSExcelFileHandler(fPath)
        'Build the room list from the COAliasReport file
        GetRoomInfo()
        'Invoke the method in the MSExcelClass that closes file
        CloseExcelFile()
    End Sub

    Sub GetRoomInfo()
        'Declare counters to walk through the excel file 
        Dim i As Integer = 1
        Dim j As Integer = 1
        'Set temp variables to hold information
        Dim tempRoom As Room 'tempRoom is intialized after the alias list is built and FoundAlias test as true
        Dim roomName As String = ""
        'Because some rooms have multiple units, the unit ID's are stored in a list of strings -
        'they are converted to Device objects when creating the Room Objects
        Dim roomDeviceName As List(Of String) = New List(Of String)
        'Temporary list to hold all the Aliases for each room
        Dim tempAliasList As List(Of String) = New List(Of String)

        Dim tString As String
        Dim aliasString As String
        Dim FoundAlias As Boolean = False
        'Builds the room list several test used to find the correct information
        While i <= COAliasReport.LastRow ' End of used range check
            'Get the value of curent cell, will always start with the room name(assumes all top info removed from file)
            tString = COAliasReport.Worksheet.Cells(i, j).Value
            'check to see if the value is empty
            If Not tString = "" Then
                If tString = "Aliases" Then
                    'Builds the Aliases list in pre for creating a Room Object
                    FoundAlias = True 'Causes program to create a Rooms object
                    j += 1
                    i += 1
                    aliasString = COAliasReport.Worksheet.Cells(i, j).Value
                    'Some Alias list are on multiple rows so need to ensure that all are being captured
                    While Not aliasString = ""
                        tempAliasList.Add(aliasString)
                        If j < 5 Then
                            j += 1
                            aliasString = COAliasReport.Worksheet.Cells(i, j).Value
                        Else 'Go to next row of Aliases if needed
                            i += 1
                            j = 2
                            aliasString = COAliasReport.Worksheet.Cells(i, j).Value
                        End If

                    End While
                    j = 1 'reset j to read the first column
                Else 'Since it's the first column and it's not "Aliases" its the room name
                    roomName = tString
                End If
            End If
            'This is the search for devices
            If tString = "" And COAliasReport.Worksheet.Cells(i, j + 1).Value = "Device" Then
                i += 1
                j += 1
                'rooms can have multiple devices need to make sure we are getting them all
                'Test to make sure that we are getting every row with a device name
                While Not COAliasReport.Worksheet.Cells(i, 1).Value = "Aliases"
                    tString = COAliasReport.Worksheet.Cells(i, j).Value
                    roomDeviceName.Add(tString) 'Add the Device to the list in prep from Creating the Rooms object
                    i += 1
                End While
                j = 1 'Reset J to column 1
                i -= 1 'Becouse we test for the string "Aliases" and have found it we need to go back 1 row 
            End If
            'Build a room object using the room name, Device List, and Alias list
            If FoundAlias = True Then
                'Pass the info to the constructor for creation of a Rooms object
                '(string, list of strings, list of strings)
                tempRoom = New Room(roomName, roomDeviceName, tempAliasList)
                'Call method to add the room to the list
                BuildRoomsList(tempRoom)
                'Reset list and boolean in prep for creation of next Room Object
                FoundAlias = False
                tempAliasList.Clear()
                roomDeviceName.Clear()
            End If

            i += 1

        End While
        'Call method to calculate number of Rooms objects in the Rooms list
        GetRoomsListCount()

    End Sub
    Sub BuildRoomsList(room As Room)
        RList.Add(room)
    End Sub

    Public Sub GetRoomsListCount()
        RoomsListCount = RList.Count()
    End Sub
    Sub CloseExcelFile()
        COAliasReport.CloseFile()
    End Sub

End Class
