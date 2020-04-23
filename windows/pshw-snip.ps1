param(
    [Parameter(Mandatory=$false)][String]$giftcardpattern="{{gift_card}}",
    [Parameter(Mandatory=$false)][String]$pinassign="N"
)



        if ($pinassign -eq "Y") {
            $mailMessage.HTMLBody = $mailMessage.HTMLBody.Replace($giftcardpattern, $invitee.contactGiftCard)
            write-host $giftcardpattern ":" $invitee.contactGiftCard
        }


contactName,contactEmail,contactCompany,contactGiftCard