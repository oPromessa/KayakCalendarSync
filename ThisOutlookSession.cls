VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim newCalFolder As Outlook.Folder
Public f As Object

Private Sub Application_Startup()
   Dim NS As Outlook.NameSpace
   Set NS = Application.GetNamespace("MAPI")
   
    Call initLog
    PrintLog "Start"
    PrintLog "-----"

    Set newCalFolder = Nothing

    Set newCalFolder = NS.GetDefaultFolder(olFolderCalendar)

    Set NS = Nothing
    
    MsgBox "APP_STARTUP"
   
End Sub
  
Public Sub SyncCalendar()
Dim i, j As Integer
Dim iFinished, jFinished As Boolean
Dim cAppt As AppointmentItem
Dim moveCal As AppointmentItem

' We can use this function to SYNC... instead of event driven ADD/CHANGE/REMOVE actions...
' MsgBox "SyncCalendar"
 
Debug.Print "Sync Calendar"
PrintLog "Sync Calendar"

' Get Restricted List of Outlook Calendar with Categories = moved
    Dim OlmyNamespace As Outlook.NameSpace
    Dim OlmyFolder As Outlook.Folder
    Dim OlmyItems As Outlook.Items
    Dim OlmyRestrictItems As Outlook.Items
    Dim OlmyItem As Outlook.AppointmentItem
    Dim OlstrEntryID As String

    Set OlmyNamespace = Application.GetNamespace("MAPI")
    Set OlmyFolder = OlmyNamespace.GetDefaultFolder(olFolderCalendar)
    Set OlmyItems = OlmyFolder.Items
    ' First run there is no 'moved' categories. Comment next line.
    Set OlmyRestrictItems = OlmyItems.Restrict("[Categories] = 'moved'")
    If OlmyRestrictItems.Count = 0 Then
        Debug.Print "Sync Calendar: No Ol appts found!"
        Debug.Print "Sync Calendar: Possibly continuing..."
        ' Maybe remove this line for 'First run' situations.
        Exit Sub
    Else
        OlmyRestrictItems.Sort "[Start]", True
    End If
        
' Get list of appointments on Kayak Internet Calendars
    Dim myNamespace As Outlook.NameSpace
    Dim myFolder As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim myRestrictItems As Outlook.Items
    Dim myItem As Outlook.AppointmentItem
    Dim mystrEntryID As String
    

    Set myNamespace = Application.GetNamespace("MAPI")
    ' Calendar added via Accounts.Internet Calendar option
    Set myFolder = GetFolderPath("\\Internet Calendars\Kayak Trips Calendar")
    ' Calendar added by direct click on URL .ics from Web
    Set myFolder = GetFolderPath("\\YourEmail@YourCompany.com\Calendar\Kayak Trips Calendar")
    Set myItems = myFolder.Items
    ' Set myRestrictItems = myItems.Restrict("[Categories] = 'moved'")
    Set myRestrictItems = myItems
    If myRestrictItems.Count = 0 Then
        Debug.Print "Remove: No my appts found!"
        ' Exit Sub
    Else
        myRestrictItems.Sort "[Start]", True
    End If
    
' Sorted lists by Start... to the lack of another topic.. could use Name... have to remove Copied: text from beginning.

    i = OlmyRestrictItems.Count
    j = myRestrictItems.Count
    iFinished = False
    jFinished = False
    
' Treat cycle when any of the indexes reaches 0 and the other one still not finished.
' Entries exist beyond the other one.
    Do Until (i = 0) And (j = 0)
        Debug.Print "-------------------- Next: " & "MainCal=" & i & "(" & iFinished & ")" & vbTab & "SyncCal=" & j & "(" & jFinished & ")"
        PrintLog "-------------------- Next: " & "MainCal=" & i & "(" & iFinished & ")" & vbTab & "SyncCal=" & j & "(" & jFinished & ")"
        If i > 0 Then
            OlstrEntryID = (Left(Right(OlmyRestrictItems(i).Subject, 49), 48))
            ' OlstrEntryID = (Left(Right(OlmyRestrictItems(i).Subject, 141), 140))
            Debug.Print ("MainCal: [" & OlmyRestrictItems(i).Subject & "] EntryID: [" & OlstrEntryID & "]")
            PrintLog ("MainCal: [" & OlmyRestrictItems(i).Subject & "] EntryID: [" & OlstrEntryID & "]")
        Else
            OlstrEntryID = "999"
            iFinished = True
            ' In this case, in fact, the Original list is larger (j) than the Copied list (i)
            ' Correct: Copy Calendar list is finished so I can back out! Could eventuall ADD remaining items from the Original list. Check below under "Found Later"
            ' Exit Sub
        End If
        
        If j > 0 Then
            mystrEntryID = myRestrictItems(j).EntryID
            Debug.Print ("SyncCal: [" & myRestrictItems(j).Subject & "] EntryID: [" & mystrEntryID & "]")
            PrintLog ("SyncCal: [" & myRestrictItems(j).Subject & "] EntryID: [" & mystrEntryID & "]")
        Else
            strEntryID = "888"
            jFinished = True
        End If
        
        Debug.Print ("MainCal.EntryID: [" & OlstrEntryID & "]")
        Debug.Print ("SyncCal.EntryID: [" & mystrEntryID & "]")
             
        If OlstrEntryID = mystrEntryID Then
            Debug.Print "-------------------- Both EQUAL. Do nothing. Continuing!"
            PrintLog "-------------------- Both EQUAL. Do nothing. Continuing!"
            Debug.Print ("Found on both systems: [" & OlmyRestrictItems(i).Subject & "] = [" & myRestrictItems(j).Subject & "]")
            i = i - 1
            j = j - 1
        Else
            If jFinished Then
                Debug.Print ("SyncCal list finised. Found Item on MainCal (originally sync'd) beyond SyncCal list. Deleting...")
                Debug.Print ("Remove/Deleting MainCal entry" & OlmyRestrictItems(i).Subject)
                Set OlmyItem = OlmyRestrictItems(i)
                OlmyItem.Delete
                i = i - 1
            Else
                If Not iFinished Then
                    If OlmyRestrictItems(i).Start < myRestrictItems(j).Start Then
                        Debug.Print "-------------------- Found Earlier"
                        PrintLog "-------------------- Found Earlier"
                        Debug.Print ("Found Item on MainCal (originally sync'd) not in SyncCal list. Remove entry: " & OlmyRestrictItems(i).Subject)
                        Set OlmyItem = OlmyRestrictItems(i)
                        OlmyItem.Delete
                        ' Cycle from top to bottom to avoid issues when deleting the entry itself is removed and the list of the array is reduced
                        ' i = i + 1
                        i = i - 1
                    Else
                        Debug.Print ("-------------------- Found Later... CHECK TO ADDING TO MainCal")
                        PrintLog ("-------------------- Found Later... CHECK TO ADDING TO MainCal: " & myRestrictItems(j).Start)
                        ' Could eventually ADD this item into Ol Calendar...
                        If myRestrictItems(j).BusyStatus = olTentative Or myRestrictItems(j).BusyStatus = olBusy Then
                            Debug.Print ("Found Later... ADDING TO MainCal")
                            Set cAppt = Application.CreateItem(olAppointmentItem)
                            With cAppt
                                .Subject = "Trip: " & myRestrictItems(j).Subject & " [" & myRestrictItems(j).EntryID & "]"
                                ' CODING: Debug
                                Debug.Print (cAppt.Subject)
                                .Start = myRestrictItems(j).Start
                                .Duration = myRestrictItems(j).Duration
                                .Location = myRestrictItems(j).Location
                                .Body = myRestrictItems(j).Body
                                ' CODING: Added Busy Status to adhere to KAYAK's logic
                                .BusyStatus = olOutOfOffice
                                .ReminderSet = False
                            End With
                        
                         
                            ' set the category after it's moved to force EAS to sync changes
                            Set moveCal = cAppt.Move(newCalFolder)
                            moveCal.Categories = "moved"
                            moveCal.Save
                            Debug.Print ("-------------------- Found Later... ADDED TO OlCalendar")
                            PrintLog ("-------------------- Found Later... ADDED TO OlCalendar")
                        End If
                        j = j - 1
                    End If
                Else
                    ' Repeat code from above...
                    Debug.Print ("-------------------- Found Later Beyond List... CHECK TO ADDING TO MainCal")
                    PrintLog ("-------------------- Found Later Beyond List... CHECK TO ADDING TO MainCal: " & myRestrictItems(j).Start)
                    ' Could eventually ADD this item into Ol Calendar...
                    If myRestrictItems(j).BusyStatus = olTentative Or myRestrictItems(j).BusyStatus = olBusy Then
                        Debug.Print ("Found Later Beyong List... ADDING TO OlCalendar")
                        Set cAppt = Application.CreateItem(olAppointmentItem)
                        With cAppt
                            .Subject = "Trip: " & myRestrictItems(j).Subject & " [" & myRestrictItems(j).EntryID & "]"
                            ' CODING: Debug
                            Debug.Print (cAppt.Subject)
                            .Start = myRestrictItems(j).Start
                            .Duration = myRestrictItems(j).Duration
                            .Location = myRestrictItems(j).Location
                            .Body = myRestrictItems(j).Body
                            ' CODING: Added Busy Status to adhere to KAYAK's logic
                            .BusyStatus = olOutOfOffice
                            .ReminderSet = False
                        End With
                    
                     
                        ' set the category after it's moved to force EAS to sync changes
                        Set moveCal = cAppt.Move(newCalFolder)
                        moveCal.Categories = "moved"
                        moveCal.Save
                        Debug.Print ("-------------------- Found Later... ADDED TO OlCalendar")
                        PrintLog ("-------------------- Found Later... ADDED TO OlCalendar")
                    End If
                    j = j - 1
                End If
            End If
        End If
    Loop

    Debug.Print "Sync Calendar: Finished"
    PrintLog "Sync Calendar: Finished"

    Exit Sub
   

End Sub


Function GetFolderPath(ByVal FolderPath As String) As Outlook.Folder
    Dim oFolder As Outlook.Folder
    Dim FoldersArray As Variant
    Dim i As Integer
         
    On Error GoTo GetFolderPath_Error
    If Left(FolderPath, 2) = "\\" Then
        FolderPath = Right(FolderPath, Len(FolderPath) - 2)
    End If
    'Convert folderpath to array
    FoldersArray = Split(FolderPath, "\")
    Set oFolder = Application.Session.folders.Item(FoldersArray(0))
    If Not oFolder Is Nothing Then
        For i = 1 To UBound(FoldersArray, 1)
            Dim SubFolders As Outlook.folders
            Set SubFolders = oFolder.folders
            Set oFolder = SubFolders.Item(FoldersArray(i))
            If oFolder Is Nothing Then
                Set GetFolderPath = Nothing
            End If
        Next
    End If
    'Return the oFolder
    Set GetFolderPath = oFolder
    Exit Function
         
GetFolderPath_Error:
    Set GetFolderPath = Nothing
    Exit Function
End Function


Public Sub initLog()

    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    
    On Error GoTo initLogError
    
    Dim prompt As VbMsgBoxResult
    Dim fs As Object
    
    prompt = MsgBox("Would you like to log events for this session?", vbYesNo, "Log Events?")

    If prompt = vbYes Then
        Set fs = CreateObject("Scripting.FileSystemObject")
        ' Set f = fs.CreateTextFile("C:\Logs\KayakCalendarSync.txt", False)
        ' Set f = fs.OpenTextFile("C:\Logs\KayakCalendarSync.txt", ForAppending, True)
        Set f = fs.OpenTextFile(CStr(Environ("USERPROFILE") & "\Documents\Logs\KayakCalendarSync.txt"), ForAppending, True)
        f.WriteLine "Start of Log"
    End If
    
Done:
    Exit Sub
    
initLogError:
    MsgBox "The following error occurred: " & Err.Description

End Sub

Public Sub PrintLog(argument As String)

On Error GoTo PrintLogError

    If Not f Is Nothing Then
        f.WriteLine argument
    End If
    
Done:
    Exit Sub
    
PrintLogError:
    MsgBox "The following error occurred: " & Err.Description

End Sub

