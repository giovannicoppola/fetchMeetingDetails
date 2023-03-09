### Tuesday, September 14, 2021, 8:38 AM
## – Sunny ☀️ 🌡️+64°F (feels +64°F, 87%) 🌬️↘4mph 🌓

## second version of the script with 2 main changes:
# 1. deal with recurring events (but only works if run the same day
# 2. copy to clipboard instead of a new DEVONthink note


### Monday, October 12, 2020, 10:57 AM
### script to automate the creation of a markdown note in devonthink for a new meeting. 
### procedure would be to focus on a calendar event, then run the script via alfred or hotkey
## need to extract fromn the calendar event:
# 1) meeting title
# 2) meeting day, time 
# 3) meeting participants	

tell application "Microsoft Outlook"
	activate
	
	
	
	if view of the first main window is equal to calendar view then
		set calendarEvent to selection -- grabbing the selected event
		
		
		-- if there are no events selected, warn the user and then quit
		if calendarEvent is missing value then
			display dialog "Please select a calendar event first and then run this script! 😀" with icon 1
			return
		end if
		
		
		-- "disconnecting" recurring events. as far as I can tell, there is no way to extract information (date start, end etc) about individual recurring event, as all the properties returned are from the series. This workaround takes advantage of a feature (bug?) by which just the actio of getting the occurrence of a recurring calendar event makes it no longer recurring. 
		-- because I can't get the date of a particular instance of a recurring event, this currently only works if I run the script the same day of the (recurring) event. If an istance of a recurring event does not occur on the current day, a warning is presented. 
		
		-- in the future, one could set an interval (for example one month) and 'disconnect' all occurring events within that range
		
		set isRec to is recurring of calendarEvent
		if isRec = true then
			set tdate to current date
			set sdate to start time of calendarEvent # because current date will have also current time, extract the time from the event (series)
			tell me to set time of tdate to time of sdate
			set myID to id of calendarEvent
			
			
			try
				set calendarEvent to get occurrence of calendar event id myID at tdate
				#log thisCalEvent
				set eventStart to start time of calendarEvent
				log {"starting event after conversion", eventStart}
				#set thisCalEvent to get occurrence of calendar event id myID at sdate
			on error
				#log sdate
				#set myRec to recurrence of calendarEvent
				#log myRec
				#set thisCalEvent to get occurrence of calendar event id myID at sdate
				
				display dialog "❗This is a recurrent event, but I see no occurrences on " & tdate with icon 1
				
			end try
		end if
		
		
		
		set eventTitle to subject of calendarEvent -- grabbing the event title
		set eventStart to start time of calendarEvent
		set startFormat to date string of eventStart
		
		
		set noteTitle to eventTitle & ", " & startFormat
		log noteTitle
		#	set calendarAttendees to get attendees of calendarEvent
		
		set emailList to ""
		set nameList to ""
		set emailList to get every email address of every attendee of calendarEvent
		
		set ind to 0
		repeat with theName in emailList
			set ind to (ind + 1)
			if ind = 1 then
				set nameList to name of theName
			else
				set nameList to nameList & ", " & name of theName
			end if
		end repeat
		
		set finalString to "# " & eventTitle & "
" & eventStart & "
" & tab & nameList & "

"
		
		set the clipboard to finalString
		
		
	else
		display dialog "This is not calendar view! 😀" with icon 1
		return
	end if
	
end tell