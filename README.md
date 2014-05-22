#todoist-outlook-2011
Simple Applescript to add a Todoist task from email messages in Outlook 2011 for Mac.

### Installation

Add a copy of the script bundle to your Outlook 2011 scripts folder which is located at:  
[YOUR HOME FOLDER]/Library/Application Support/Microsoft/Office/Outlook Script Menu Items  
You can confirm the location of this folder by opening the Applescript menu in Outlook 2011 and clicking "About This Menu..."  
In the same folder, create a text file named todoist-token.txt that contains your Todoist API token.

### Usage

Highlight one or more messages in Outlook, then select Todoist Task from the Applescript menu.  
A new task will be created in your Todoist inbox in the following format:  
[SUBJECT LINE OF MESSAGE] for [NAME OF MESSAGE SENDER]

