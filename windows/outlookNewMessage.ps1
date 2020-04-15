$olFolderDrafts = 16
# $outl = New-Object -comObject "Outlook.Application"

$fileInvitees = "~/Downloads/test.csv"
$fileTemplate = "~/Documents/04022020-Cheers with Engineers.emltpl"

$inviteeList = Import-csv $fileInvitees


$tTempCompany = ""
foreach ($invitee in $inviteeList) {
    if ($tTempCompany -ne $invitee.contactCompany) {
        $tTempCompany = $invitee.contactCompany
 #       $mailMessage = $outl.createItemFromTemplate($fileTemplate)
        Write-Host "| MAIL MESSAGE CREATION | "
 #       $mailMessage.Recipients.Add($invitee.contactEmail)
        Write-Host "Would Add Recipient: " $invitee.contactEmail
 #       $mailMessage.subject = "Greetings, Team " + $invitee.contactCompany + " You've been invited!"
        Write-Host "| MAIL: SUBJECT: " $invitee.contactCompany
 #       $mailMessage.body = " | - - - Test Message - - - | "
        Write-Host "| MAIL: BODY"
 #       $mailMessage.save()
        Write-Host "| MAIL: SAVE TO DRAFTS"
    }
    else {
        $tTempCompany = $invitee.contactCompany
 #       $mailMessage.Recipients.Add($invitee.contactEmail)
        Write-Host "Would Add Recipient: " $invitee.contactEmail
    }
}