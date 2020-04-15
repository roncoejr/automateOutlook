set DELIM to {","}

my getInviteeList()



on getInviteeList()
	set DELIM to {","}
	
	set customerName to {}
	set customerEmail to {}
	set customerCompany to {}
	
	set customerLine to {customerName, customerEmail, customerCompany}
	
	-- set emailTemplate to (choose file)
	set arecipientList to (choose file with prompt "Select the List of Invitees:" of type {"csv"})
	set recipientList to read arecipientList using delimiter linefeed
	-- set theCopyRecipient to "eric.schuster@nutanix.com"
	log recipientList
	
	-- set customerRecord to read aCSV using delimiter linefeed
	
	-- log customerRecord
	
	set {TID, AppleScript's text item delimiters} to {AppleScript's text item delimiters, DELIM}
	set tTempCompany to ""
	
	set theMessage to read (choose file with prompt "Select the Template:" of type {"emltpl"})
	
	repeat with aLine in recipientList
		set lineItems to text items of aLine
		-- repeat with i from 1 to ((count of lineItems) - 1)
		--	copy (item i of lineItems) to the end of (item i of customerLine)
		-- end repeat
		tell application "Microsoft Outlook"
			-- repeat with aInvitee in Invitees
			set theRecipient to item 2 of lineItems
			if item 3 of lineItems is not equal to tTempCompany then
				set tTempCompany to item 3 of lineItems
				log tTempCompany
				set aMessage to make new outgoing message with properties {subject:"Greetings, Team " & tTempCompany & " You're Invited!  Compliments of Nutanix", source:theMessage}
				make new recipient with properties {email address:{address:theRecipient}} at end of to recipients of aMessage
			else
				set tTempCompany to item 3 of lineItems
				make new recipient with properties {email address:{address:theRecipient}} at end of to recipients of aMessage
			end if
			set theRecipient to item 2 of lineItems
			log theRecipient
			-- end repeat
		end tell
	end repeat
	
	set AppleScript's text item delimiters to TID
	-- log item 2 of lineItems
	return customerLine
end getInviteeList
