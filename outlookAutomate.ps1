param(
    [Parameter(Mandatory=$false)][String]$giftcardpattern="{{gift_card}}",
    [Parameter(Mandatory=$false)][String]$followuppattern="attendee_name",
    [Parameter(Mandatory=$false)][String]$pinassign="N",
    [Parameter(Mandatory=$false)][String]$followup="N",
    [Parameter(Mandatory=$false)][String]$subjectPre="",
    [Parameter(Mandatory=$false)][String]$subjectPost=""
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

$inviteeListRaw = Import-csv $fileInvitees
$inviteeList = $inviteeListRaw | Sort-Object -Property contactCompany

$tTempCompany = ""
foreach ($invitee in $inviteeList) {
    if (($tTempCompany -ne $invitee.contactCompany) -or ($pinassign -eq "Y") -or ($followup -eq "Y")) {
        $tTempCompany = $invitee.contactCompany
         $mailMessage = $outl.createItemFromTemplate($fileTemplate)
        # $mailMessage = $outl.createItem(0)
        Write-Host "| MAIL MESSAGE CREATION | "
        $mailMessage.Recipients.Add($invitee.contactEmail)
        Write-Host "Would Add Recipient: " $invitee.contactEmail
#        $mailMessage.subject = $subjectPre + $invitee.contactCompany + $subjectPost
#        $mailMessage.subject = $subjectPre + $invitee.contactCompany + $subjectPost
        $mailMessage.subject = $subjectPre + $mailMessage.subject.Replace($followuppattern, $invitee.contactName) + $subjectPost
        Write-Host "| MAIL: SUBJECT: " $invitee.contactCompany
        #$mailMessage.body = " | - - - Test Message - - - | "
        if ($pinassign -eq "Y") {
            $mailMessage.Body = $mailMessage.Body.Replace($giftcardpattern, $invitee.contactGiftCard)
            write-host $giftcardpattern ":" $invitee.contactGiftCard
        }

        if ($followup -eq "Y") {
#            $mailMessage.Body = $mailMessage.Body.Replace($followuppattern, $invitee.contactName)
            $mailMessage.HTMLBody = $mailMessage.HTMLBody.Replace($followuppattern, $invitee.contactName)
            write-host $followuppattern ":" $invitee.contactName
        }

        Write-Host "| MAIL: BODY"
        $mailMessage.save()
        Write-Host "| MAIL: SAVE TO DRAFTS"
    }
    else {
        $tTempCompany = $invitee.contactCompany
        $mailMessage.Recipients.Add($invitee.contactEmail)
        Write-Host "Would Add Recipient: " $invitee.contactEmail
    }
}