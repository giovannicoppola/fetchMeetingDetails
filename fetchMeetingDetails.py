#!/usr/bin/python3 
# 
# Partly cloudy ⛅️  🌡️+44°F (feels +39°F, 51%) 🌬️↓12mph 🌕 Thu Mar  9 15:22:34 2023
# W10Q1 – 68 ➡️ 296 – 302 ❇️ 62

"""
FETCH OUTLOOK MEETING DETAILS
A script to use applescript to grab meeting details without focusing on the app or selecting the event (as I had in the previous versions)

"""

import os
import time
import sys
import json
import requests
from config import log
from subprocess import Popen, PIPE, run

myTimeStart = round(time.time())


scpt = '''
    on run {minInterval}

# Original script from user https://www.reddit.com/user/scrutinizer1/ from this thread https://www.reddit.com/r/applescript/comments/mep684/is_it_easy_to_pull_content_from_an_outlook/


with timeout of (2800 * minutes) seconds
	
	set now to current date
	set now5 to now + (minInterval * minutes) #will work for the current event until 5 min from the end, after which will work for the next event
	
	tell application "Microsoft Outlook"
		set the clipboard to ""
		
		-- Create and add items
		set theItems to {}
        set theTitles to ""
		set CalEvProperties to ""
        set totalList to (get every calendar event whose start time is less than or equal to now5 and end time is greater than now5)
        set totalNumber to the number of items in totalList
        repeat with myCaz in totalList
            set theTitles to theTitles & "--" & (subject of myCaz as text) 
        end repeat
		repeat with CalEv in (get every calendar event whose start time is less than or equal to now5 and end time is greater than now5)
			set the clipboard to ""
			tell CalEv
                set theTitles to theTitles & "--" & (subject as text) 
				if (all day flag) then
					set AllDay to "Yes"
				else
					set AllDay to "No"
				end if
				if has reminder then
					set HasRem to "Yes"
					set RemOn to "Minutes until fires up: " & (reminder time) as text
				else
					set HasRem to "No"
					set RemOn to ""
				end if
				if is recurring then
					set IsRecur to "Yes"
					set Recur to recurrence
					set RecurID to recurrence id
				else
					set IsRecur to "No"
					set Recur to ""
					set RecurID to "N/A"
				end if
				if is occurrence then
					set IsOcur to "Yes"
				else
					set IsOcur to "No"
				end if
				if request responses then
					set Requests to "Yes"
				else
					set Requests to "No"
				end if
				
				
				
				set |Attendees| to ""
				set _Attendees to attendees
				if _Attendees is not {} then
					repeat with i from 1 to the number of items in _Attendees
						set PersonAttends to _Attendees's item i
						tell PersonAttends
							set |Attendees| to (((|Attendees| & "Email: " & (email address) as text) & ", " & "Attendee type: " & (type) as text) & ", " & "Status: " & status as text) & return
						end tell
					end repeat
				end if
				
				set CalEvProperties to (((((("Event" & id as text) & " info" & return & return & "Subject: " & (subject as text) & "," & return & "Starts: " & (start time) as text) & "," & space & "Ends: " & (end time) as text) & "," & space & "All day: " & AllDay & "," & return & "Free-busy: " & (free busy status) as text) & ", " & return & "Reminder set: " & HasRem & ", " & RemOn & return & "Organized by: " & organizer & ", " & return & "Recurring: " & IsRecur & ", " & "Recurrence: " & Recur & ", " & "Recurrence ID: " & RecurID & ", " & return & "Is event occurrence: " & IsOcur & ", " & return & "Requests responses: " & Requests & ", " & "Event description: " & (plain text content) & ", " & return & "Last modified: " & (modification date) as text) & ", " & return & "Location: " & location & "TMZ: " & (timezone) as text) & "," & return & return & "_______________" & return & return & "Attendees: " & return & return & |Attendees|
				set the clipboard to CalEvProperties
				repeat until (the clipboard) contains ("Event" & id as text) & " info"
					delay 0.1
				end repeat
				
                
                
                set mySubject to (subject as text)
				set eventStart to start time
                set eventEnd to end time
                set startFormat to date string of eventStart
                set startTimeFormat to time string of eventStart
                set endTimeFormat to time string of eventEnd
                

                set emailList to ""
		        set nameList to ""
		        set emailList to get every email address of every attendee
                
                set ind to 0
                repeat with theName in emailList
                    set ind to (ind + 1)
                    if ind = 1 then
                        set nameList to name of theName
                    else
                        set nameList to nameList & ", " & name of theName
                    end if
                end repeat
		
                
				#return {CalEvProperties, mySubject}
                return startFormat & " – " & startTimeFormat & "-" & endTimeFormat & "|||---|||" & mySubject & "|||---|||" & nameList & "|||---|||" & theTitles & "|||---|||" & totalNumber & "|||---|||"

			end tell
			#delay 0.5
		end repeat
 


	end tell
end timeout

end run
'''

#args = [SPRINT_DUR, Email_Start, Email_StartF, sprintDurSec]
args = ["-383"]        
p = Popen(['osascript', '-'] + args, stdin=PIPE, stdout=PIPE, stderr=PIPE, universal_newlines=True)
stdout, stderr = p.communicate(scpt)
print (stdout)

myResults = stdout.split("|||---|||")
myStartTime = myResults[0]
myTitle = myResults[1]
myAttendees = myResults[2]

myFinalString = f"# {myTitle}\n{myStartTime}\n\t{myAttendees}\n\n"

print (myFinalString)

myTimeEnd = round(time.time())

main_timeElapsed = round (myTimeEnd - myTimeStart)
print (f"time elapsed: {main_timeElapsed}")


myTimeStartF = time.strftime('%Y-%m-%d-%a, %I:%M', time.localtime(myTimeStart))
myTimeEndF = time.strftime('%I:%M %p', time.localtime(myTimeEnd))


# Copy the string to the clipboard
run('pbcopy', universal_newlines=True, input=myFinalString)














 

