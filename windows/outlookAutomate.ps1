param(
    [Parameter(Mandatory=$false)][String]$giftcardpattern="{{gift_card}}",
    [Parameter(Mandatory=$false)][String]$followuppattern="attendee_name",
    [Parameter(Mandatory=$false)][String]$companyreplaceuppattern="company_name",
    [Parameter(Mandatory=$false)][String]$pinassign="N",
    [Parameter(Mandatory=$false)][String]$followup="N",
    [Parameter(Mandatory=$false)][String]$subjectPre="",
    [Parameter(Mandatory=$false)][String]$subjectPost="",
    [Parameter(Mandatory=$false)][String]$sendMail="N"
)

$olFolderDrafts = 16
$outl = New-Object -comObject "Outlook.Application"

Function Get-FileName($initialDirectory, $fileFilterType)
{   
 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
 Out-Null
 $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
 $OpenFileDialog.initialDirectory = $initialDirectory
 $OpenFileDialog.filter = "" + $fileFilterType + ""
 $OpenFileDialog.ShowDialog() | Out-Null
 $OpenFileDialog.filename
} #end function Get-FileName

$fileInvitees = Get-FileName -initialDirectory ".\" -fileFilterType "All CSV Files (*.csv)|*.csv"
$fileTemplate = Get-FileName -initialDirectory ".\" -fileFilterType "All E-mail Templates (*.oft)|*.oft"
$fileHTML = Get-FileName -initialDirectory ".\" -fileFilterType "All HTM Files (*.htm)|*.htm|All HTML Files (*.html)|*.html"

$inviteeListRaw = Import-csv $fileInvitees
$inviteeList = $inviteeListRaw | Sort-Object -Property contactCompany
$htmlMail = [IO.File]::ReadAllText($fileHTML)

$tTempCompany = ""
foreach ($invitee in $inviteeList) {
    if (($tTempCompany -ne $invitee.contactCompany) -or ($pinassign -eq "Y") -or ($followup -eq "Y")) {
        $tTempCompany = $invitee.contactCompany
         $mailMessage = $outl.createItemFromTemplate($fileTemplate)
        # $mailMessage = $outl.createItem(0)
        Write-Host "| MAIL MESSAGE CREATION | "
        $mailMessage.Recipients.Add($invitee.contactEmail)
        Write-Host "Would Add Recipient: " $invitee.contactEmail ($invitee.contactName)
#        $mailMessage.subject = $subjectPre + $invitee.contactCompany + $subjectPost
#        $mailMessage.subject = $subjectPre + $invitee.contactCompany + $subjectPost
#        $mailMessage.subject = $subjectPre + $mailMessage.subject.Replace($followuppattern, $invitee.contactName) + $subjectPost
        Write-Host "| MAIL: SUBJECT: " $invitee.contactCompany
        #$mailMessage.body = " | - - - Test Message - - - | "
        $contactFirst = $invitee.contactName -split " +"
        # $mailMessage.Body = $mailMessage.Body.Replace($followuppattern, $contactFirst[0])
        $tmpSubject = $mailMessage.subject
        # $tmpSubject = $tmpSubject.Replace($companyreplaceuppattern, $contactFirst[0])
        $tmpSubject = $tmpSubject.Replace($followuppattern, $contactFirst[0])
        $mailMessage.subject = $tmpSubject
        $mailMessage.HTMLBody = $htmlMail.Replace($followuppattern, $contactFirst[0])
        if ($pinassign -eq "Y") {
            $mailMessage.Body = $mailMessage.Body.Replace($giftcardpattern, $invitee.contactGiftCard)
            write-host $giftcardpattern ":" $invitee.contactGiftCard
        }

        if ($followup -eq "Y") {
            $mailMessage.HTMLBody = $mailMessage.HTMLBody.Replace($followuppattern, $contactFirst[0])
#            $mailMessage.HTMLBody = $mailMessage.HTMLBody.Replace($followuppattern, $contactFirst[0])
            write-host $followuppattern ":" $invitee.contactName
        }

        Write-Host "| MAIL: BODY"
        if($sendMail -eq "Y") {
            $mailMessage.send()
        }
        else {
            $mailMessage.save()
        }
        #$mailMessage.send()
        Write-Host "| MAIL: SAVE TO DRAFTS"
        
    }
    else {
        $tTempCompany = $invitee.contactCompany
        $mailMessage.subject.Replace($followuppattern, $invitee.contactName)

        $mailMessage.Recipients.Add($invitee.contactEmail)
        Write-Host "Would Add Recipient: " $invitee.contactEmail
    }
}