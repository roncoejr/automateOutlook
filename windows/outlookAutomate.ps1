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
    if ($tTempCompany -ne $invitee.contactCompany) {
        $tTempCompany = $invitee.contactCompany
         $mailMessage = $outl.createItemFromTemplate($fileTemplate)
        # $mailMessage = $outl.createItem(0)
        Write-Host "| MAIL MESSAGE CREATION | "
        $mailMessage.Recipients.Add($invitee.contactEmail)
        Write-Host "Would Add Recipient: " $invitee.contactEmail
        $mailMessage.subject = "Greetings, Team " + $invitee.contactCompany + " You've been invited!"
        Write-Host "| MAIL: SUBJECT: " $invitee.contactCompany
        #$mailMessage.body = " | - - - Test Message - - - | "
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