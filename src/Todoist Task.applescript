(*
	Create Todoist task from email

*)

on replace_chars(this_text, search_string, replacement_string)
	set AppleScript's text item delimiters to the search_string
	set the item_list to every text item of this_text
	set AppleScript's text item delimiters to the replacement_string
	set this_text to the item_list as string
	set AppleScript's text item delimiters to ""
	return this_text
end replace_chars

-- get todoist token from text file
set todoistToken to my replace_chars(read file ((path to home folder as text) & "Library:Application Support:Microsoft:Office:Outlook Script Menu Items:todoist-token.txt"), "\n", "")

tell application "Microsoft Outlook"
	
	-- get the currently selected message or messages
	set selectedMessages to current messages
	
	-- if there are no messages selected, warn the user and then quit
	if selectedMessages is {} then
		display dialog "Please select a message first and then run this script." with icon 1
		return
	end if
	
	repeat with theMessage in selectedMessages
		
		-- get the information from the message, and store it in variables
		set theName to subject of theMessage
		set theSender to sender of theMessage
		set thePriority to priority of theMessage
		if thePriority is priority high then
			set thePriority to "4"
		else
			set thePriority to "0"
		end if
		
		-- create a new task with the information from the message
		do shell script "curl -X POST -d 'content=" & theName & " for " & name of theSender & "' -d 'token=" & todoistToken & "'  -d 'priority=" & thePriority & "' https://todoist.com/API/additem"
		
	end repeat
	
end tell