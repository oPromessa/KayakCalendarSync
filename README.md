# KayakCalendarSync
Sync your Kayak Trips Internet Calendar with your main Outlook Calendar to block travel times.

# To setup

* Enable Macros on Outlook
* Go into VBA Editor
* Import file ThisOutlookSession.cls
* Adjust the following lines accordingly
```
    ' Calendar added via Accounts.Internet Calendar option
    Set myFolder = GetFolderPath("\\Internet Calendars\Kayak Trips Calendar")
    ' Calendar added by direct click on URL .ics from Web	
    Set myFolder = GetFolderPath("\\<<Your USERNAME>>\Calendar\Kayak Trips Calendar")
```