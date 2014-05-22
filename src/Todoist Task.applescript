(*
	Create Todoist task from email

*)

tell application "Microsoft Outlook"
	
	-- get the currently selected message or messages
	set selectedMessages to current messages
	
	-- if there are no messages selected, warn the user and then quit
	if selectedMessages is {} then
		display dialog "Please select a message first and then run this script." with icon 1
		return
	end if
	
	-- get todoist token from text file
	set todoistToken to (read file ((path to home folder as text) & "Library:Application Support:Microsoft:Office:Outlook Script Menu Items:todoist-token.txt"))
	
	repeat with theMessage in selectedMessages
		
		-- get the information from the message, and store it in variables
		set theName to subject of theMessage
		set theSender to sender of theMessage
		set thePriority to priority of theMessage
		set theContent to content of theMessage
		
		-- create a new task with the information from the message
		do shell script "curl -X POST -d 'content=" & theName & " for " & name of theSender & "' -d 'token=" & todoistToken & "'  -d 'priority=0' https://todoist.com/API/additem"
		
	end repeat
	
end tell