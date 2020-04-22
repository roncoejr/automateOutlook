set DELIM to {","}

my getInviteeList()



on getInviteeList()
	set DELIM to {","}
	set DELIMCOLON to {":"}
	set DELIMPERIOD to {"."}
	
	set customerName to {}
	set customerEmail to {}
	set customerCompany to {}
	set customerGiftCard to {}
	
	set customerLine to {customerName, customerEmail, customerCompany, customerGiftCard}
	
	set AppleScript's text item delimiters to DELIM
	set mytime to item 1 of time string of (current date) & "_" & item 2 of time string of (current date) & "_" & item 8 of time string of (current date)
	set mydate to month of (current date) & day of (current date)
	
	--display dialog mytime & "-" & mydate
	
	
	-- set emailTemplate to (choose file)
	set arecipientList to (choose file with prompt "Select the List of Invitees:" of type {"csv"})
	set recipientList to read arecipientList using delimiter linefeed
	--set cmdReadFile to quoted form of "(sed -E 's/^\\"\([A-Za-z]+\), \([A-Za-z]+[ A-Z\\.| A-Za-z]*\)\\"/\\\2 \\\1/g' " & arecipientList & | awk -F \\",\\" {'printf \\"%s,%s,%s\n\\", $1, $2, $3'} | sort -k3d -t '\t')"
	
	set name_recipientList to arecipientList as string
	set name_recipientListNew to name_recipientList
	set name_recipientList to quoted form of (POSIX path of name_recipientList)
	
	set AppleScript's text item delimiters to DELIMPERIOD
	--set listComp to every text item of arecipientList
	set name_recipientListNew to text items of name_recipientListNew
	--& "-" & mytime & mydate as string
	--set AppleScript's text item delimiters to ""
	--display dialog item 1 of name_recipientListNew
	--display dialog item 2 of name_recipientListNew
	
	set theFile to "~/" & "attendee-list-" & mydate & "-" & mytime & ".csv"
	--display dialog theFile
	
	
	-- "
	--display dialog name_recipientList
	--set AppleScript's text item delimiters to DELIM
	
	set recipientList to do shell script "(sed -E 's/^\"([A-Za-z]+), ([A-Za-z]+[ A-Z\\.| A-Za-z]*)\"/\\2 \\1/g' " & name_recipientList & " | awk -F \"\\,\" {'printf \"%s,%s,%s\\n\", $1,$2,$3'} | sort -k3 -t ',' -d >" & theFile & ")"
	--set theOpenedFile to open file theFile
	set recipientList to read getDiskItem(theFile) using delimiter linefeed
	-- & " | awk -F \"\\,\" {'printf \"%s,%s,%s\"\\, $1\\, $2\\, $3'} | sort -k3d -t '	')"
	--set recipientList to do shell script "ls -lth"
	-- set theCopyRecipient to "eric.schuster@nutanix.com"
	--set AppleScript's text item delimiters to DELIM
	--set recipientListtmp to text items of recipientList as string
	log recipientList
	
	-- set customerRecord to read aCSV using delimiter linefeed
	
	-- log customerRecord
	--display dialog recipientList
	set {TID, AppleScript's text item delimiters} to {AppleScript's text item delimiters, DELIM}
	set tTempCompany to ""
	
	set theMessage to read (choose file with prompt "Select the Template:" of type {"emltpl"})
	
	repeat with aLine in recipientList
		--display dialog aLine
		set lineItems to text items of aLine
		--repeat with i from 1 to ((count of lineItems) - 1)
		--	copy (item i of lineItems) to the end of (item i of customerLine)
		--end repeat
		set countOfItems to count of lineItems
		set strCountOfItems to countOfItems as string
		--display dialog strCountOfItems
		if strCountOfItems is greater than 3 then
			set item 1 of lineItems to item 1 of lineItems & ", " & item 2 of lineItems as string
			set item 2 of lineItems to item 3 of lineItems
			set item 3 of lineItems to item 4 of lineItems
			-- delete item 4 of lineItems
			set stringOfLineItems to item 1 of lineItems & item 2 of lineItems as string
			-- display dialog stringOfLineItems
			--display dialog item 1 of lineItems & item 2 of lineItems & item 3 of lineItems
			-- display dialog strCountOfItems
		end if
		tell application "Microsoft Outlook"
			-- repeat with aInvitee in Invitees
			set theRecipient to item 2 of lineItems as string
			--display dialog theRecipient
			if item 3 of lineItems is not equal to tTempCompany then
				set savedDelims to AppleScript's text item delimiters
				-- set AppleScript's text item delimiters to "\""
				
				set tTempCompany to item 3 of lineItems
				set AppleScript's text item delimiters to "\""
				set itemsTempCompany to the text items of tTempCompany
				set AppleScript's text item delimiters to ""
				set tTempCompany to the text items of itemsTempCompany as string
				log tTempCompany
				-- set AppleScript's text item delimiters to savedDelims
				set aMessage to make new outgoing message with properties {subject:"Greetings, Team " & tTempCompany & " You're Invited!  Compliments of Nutanix", source:theMessage}
				make new recipient with properties {email address:{address:theRecipient}} at end of to recipients of aMessage
			else
				set savedDelims to AppleScript's text item delimiters
				-- set AppleScript's text item delimiters to "\""
				
				set tTempCompany to item 3 of lineItems
				set AppleScript's text item delimiters to "\""
				set itemsTempCompany to the text items of tTempCompany
				set AppleScript's text item delimiters to ""
				set tTempCompany to the text items of itemsTempCompany as string
				-- set AppleScript's text item delimiters to "\""
				-- set tTempCompany to text item tTempCompany
				-- set tTempCompany to item 3 of lineItems
				make new recipient with properties {email address:{address:theRecipient}} at end of to recipients of aMessage
			end if
			set theRecipient to item 2 of lineItems
			
			log theRecipient
			-- end repeat
		end tell
		set AppleScript's text item delimiters to savedDelims
	end repeat
	
	set AppleScript's text item delimiters to TID
	-- log item 2 of lineItems
	return lineItems
end getInviteeList

on getDiskItem(theItem)
	tell application "System Events"
		return (path of disk item (theItem as string) as alias)
		
	end tell
end getDiskItem
