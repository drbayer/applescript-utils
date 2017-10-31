set trashName to "Deleted Items"
set myName to "Empty Deleted Items"

tell application "Microsoft Outlook"
	
	set allFolders to every mail folder of default account
	set trashFolder to null
	repeat with theFolder in allFolders
		if name of theFolder is trashName then
			set trashFolder to theFolder
		end if
	end repeat
	
	if trashFolder is null then
		display alert myName message ("Trash folder \"" & trashName & "\" not found") as critical
		return 0
	end if
	
	set deletedMessages to (messages of trashFolder)
	set messageCount to 0
	repeat with theMessage in deletedMessages
		permanently delete theMessage
		set messageCount to (messageCount + 1)
	end repeat
	display notification "Permanently deleted " & messageCount & " messages from folder " & trashName with title myName

end tell