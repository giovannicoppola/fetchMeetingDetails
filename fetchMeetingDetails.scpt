on run argv

# Original script from user https://www.reddit.com/user/scrutinizer1/ from this thread https://www.reddit.com/r/applescript/comments/mep684/is_it_easy_to_pull_content_from_an_outlook/

-- import JSON library
tell application "Finder"
	set json_path to file "emitting-json.scpt" of folder of (path to me)
end tell
set json to load script (json_path as alias)



with timeout of (2800 * minutes) seconds
	
	set now to current date
	set now5 to now + (5 * minutes) #will work for the current event until 5 min from the end, after which will work for the next event
	
	tell application "Microsoft Outlook"
		set the clipboard to ""
		
		-- Create and add items
		set theItems to {}

		set CalEvProperties to ""
		repeat with CalEv in (get every calendar event whose start time is less than or equal to now5 and end time is greater than now5)
			set the clipboard to ""
			tell CalEv
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
				log (mySubject)
				set end of theItems to json's createDictWith({{"title", (mySubject)}, {"uid", 1}})

			end tell
			#delay 0.5
		end repeat
 
set theString to {}
set end of theItems to json's createDictWith({{"title", "bar"}, {"uid", 2}})
set theString to json's createDictWith({{"myString", CalEvProperties}})

set theVar to {}
set theVar to json's createDictWith({{"variables", theString}})

-- Create root items object and encode to JSON
set itemDict to json's createDict()
set varDict to json's createDict()
itemDict's setkv("items", theItems)
varDict's setkv("alfredworkflow", theVar)
do shell script "python3 JSONparse.py"
return json's encode(varDict)
#return json's encode(itemDict)


	end tell
end timeout

end run
