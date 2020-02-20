VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim WithEvents curCal As Items
Attribute curCal.VB_VarHelpID = -1
Dim WithEvents DeletedItems As Items
Attribute DeletedItems.VB_VarHelpID = -1
Dim newCalFolder As Outlook.Folder
Dim curCalFolder As Outlook.Folder
Public f As Object

Private Sub Application_Startup()
   Dim NS As Outlook.NameSpace
   Set NS = Application.GetNamespace("MAPI")
   
Call initLog
PrintLog "Start"
PrintLog "-----"

MsgBox "APP_STARTUP"
' calendar to watch for new items
'   Set curCal = NS.GetDefaultFolder(olFolderCalendar).Items
' watch deleted folder
    Set DeletedItems = Nothing
'Set DeletedItems = NS.GetDefaultFolder(olFolderDeletedItems).Items
' calendar moving copy to
    Set newCalFolder = Nothing
'   Set newCalFolder = GetFolderPath("data-file-name\calendar")
   
'Set curCalFolder = GetFolderPath("\\YourEmail@YourCompany.com\Calendar\Kayak Trips")
    Set curCalFolder = GetFolderPath("\\Internet Calendars\Kayak Trips Calendar")
    Set curCal = curCalFolder.Items

    Set newCalFolder = NS.GetDefaultFolder(olFolderCalendar)

    Set NS = Nothing
   
End Sub
  
 
 
Private Sub curCal_ItemChange(ByVal Item As Object)
Dim cAppt As AppointmentItem
Dim objAppointment As AppointmentItem
Dim strFilter, strEntryID As String
Dim filteredItems As Outlook.Items
     
    
On Error Resume Next
   
' MsgBox "ITEMCHANGE"
Debug.Print "ITEMCHANGE Event"

' use 2 + the length of the GUID
' strBody = Right(Item.Body, 38)
' Debug.Print ("strBody=" & strBody)

    Set cAppt = Application.CreateItem(olAppointmentItem)
    
    strEntryID = Item.EntryID
    Debug.Print ("strEntryID=" & strEntryID)

    Dim myNamespace As Outlook.NameSpace
    Dim myFolder As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim myRestrictItems As Outlook.Items
    Dim myItem As Outlook.AppointmentItem

    Set myNamespace = Application.GetNamespace("MAPI")
    Set myFolder = myNamespace.GetDefaultFolder(olFolderCalendar)
    Set myItems = myFolder.Items
    Set myRestrictItems = myItems.Restrict("[Categories] = 'moved'")
    If myRestrictItems.Count = 0 Then
        Debug.Print "Change: No appts found!"
    Else
        For i = myRestrictItems.Count To 1 Step -1
            Debug.Print ("Restricted items:" & myRestrictItems(i).Subject)
            If InStr(1, myRestrictItems(i).Subject, strEntryID) Then
                Set cAppt = myRestrictItems(i)
                ' CODING: Added for debugging purposes...
                Debug.Print ("Found APPT via EntryID:" & cAppt.Subject)
                If myRestrictItems(i).Categories = "moved" Then
                    Debug.Print (myRestrictItems(i).Subject & "CATEGORIES=" & myRestrictItems(i).Categories)
                    Debug.Print ("EXITING FOR WITH cAppt found")
                    Exit For
                End If
            End If
        Next
        Debug.Print ("FINISHED FOR WITH.")
    End If

' CODING: Added for debugging purposes...
Debug.Print (cAppt.Subject)

With cAppt
    .Subject = "Trip: " & Item.Subject & " [" & Item.EntryID & "]"
    .Start = Item.Start
    .Duration = Item.Duration
    .Location = Item.Location
    .Body = Item.Body
    ' CODING: Added Busy Status to adhere to KAYAK's logic
    .BusyStatus = olOutOfOffice
    .Categories = "moved"
    .ReminderSet = False
    .Save
End With
  
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
    ' First run there is no 'moved' categories. Commetn next line.
    Set OlmyRestrictItems = OlmyItems.Restrict("[Categories] = 'moved'")
    Set OlmyRestrictItems = OlmyItems
    If OlmyRestrictItems.Count = 0 Then
        Debug.Print "Remove: No Ol appts found!"
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
    Set myFolder = GetFolderPath("\\Internet Calendars\Kayak Trips Calendar")
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
        Debug.Print "-------------------- Next: " & "Ol=" & i & "(" & iFinished & ")" & vbTab & "my=" & j & "(" & jFinished & ")"
        PrintLog "-------------------- Next: " & "Ol=" & i & "(" & iFinished & ")" & vbTab & "my=" & j & "(" & jFinished & ")"
        If i > 0 Then
            OlstrEntryID = (Left(Right(OlmyRestrictItems(i).Subject, 49), 48))
            ' OlstrEntryID = (Left(Right(OlmyRestrictItems(i).Subject, 141), 140))
            Debug.Print ("OlRestricted: [" & OlmyRestrictItems(i).Subject & "] OlstrEntryID: [" & OlstrEntryID & "]")
            PrintLog ("OlRestricted: [" & OlmyRestrictItems(i).Subject & "] OlstrEntryID: [" & OlstrEntryID & "]")
        Else
            OlstrEntryID = "999"
            iFinished = True
            ' In this case, in fact, the Original list is larger (j) than the Copied list (i)
            ' Correct: Copy Calendar list is finished so I can back out! Could eventuall ADD remaining itmes from the Original list. Check below under "Found Later"
            ' Exit Sub
        End If
        
        If j > 0 Then
            mystrEntryID = myRestrictItems(j).EntryID
            Debug.Print ("Restricted: [" & myRestrictItems(j).Subject & "] mystrEntryID: [" & mystrEntryID & "]")
            PrintLog ("Restricted: [" & myRestrictItems(j).Subject & "] mystrEntryID: [" & mystrEntryID & "]")
        Else
            strEntryID = "888"
            jFinished = True
        End If
        
        Debug.Print ("OlstrEntryID: [" & OlstrEntryID & "]")
        Debug.Print ("mystrEntryID: [" & mystrEntryID & "]")
             
        If OlstrEntryID = mystrEntryID Then
            Debug.Print "-------------------- Both EQUAL"
            PrintLog "-------------------- Both EQUAL"
            Debug.Print ("Found on both systems: [" & OlmyRestrictItems(i).Subject & "] = [" & myRestrictItems(j).Subject & "]")
            i = i - 1
            j = j - 1
        Else
            If jFinished Then
                Debug.Print ("Found Items on Destination beyond Original list")
                Debug.Print ("Remove entry" & OlmyRestrictItems(i).Subject)
                Set OlmyItem = OlmyRestrictItems(i)
                OlmyItem.Delete
                i = i - 1
            Else
                If Not iFinished Then
                    If OlmyRestrictItems(i).Start < myRestrictItems(j).Start Then
                        Debug.Print "-------------------- Found Earlier"
                        PrintLog "-------------------- Found Earlier"
                        Debug.Print ("Remove entry" & OlmyRestrictItems(i).Subject)
                        Set OlmyItem = OlmyRestrictItems(i)
                        OlmyItem.Delete
                        ' Cycle from top to bottom to avoid issues when deleting the entry itself is removed and the list of the array is reduced
                        ' i = i + 1
                        i = i - 1
                    Else
                        Debug.Print ("-------------------- Found Later... CHECK TO ADDING TO OlCalendar")
                        PrintLog ("-------------------- Found Later... CHECK TO ADDING TO OlCalendar: " & myRestrictItems(j).Start)
                        ' Could eventually ADD this item into Ol Calendar...
                        If myRestrictItems(j).BusyStatus = olTentative Or myRestrictItems(j).BusyStatus = olBusy Then
                            Debug.Print ("Found Later... ADDING TO OlCalendar")
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
                    ' Repeat code form above...
                    Debug.Print ("-------------------- Found Later Beyond List... CHECK TO ADDING TO OlCalendar")
                    PrintLog ("-------------------- Found Later Beyond List... CHECK TO ADDING TO OlCalendar: " & myRestrictItems(j).Start)
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
        f.WriteLine "yadayada"
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

