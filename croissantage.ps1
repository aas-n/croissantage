param(
    [string[]]$Recipients
)

function Send-MailViaOutlook {
    param(
        [string[]]$Recipients
    )

    $Outlook = New-Object -ComObject Outlook.Application
    $Mail = $Outlook.CreateItem(0) # 0 = Mail Item

    foreach ($Recipient in $Recipients) {
        $Mail.Recipients.Add($Recipient) | Out-Null
    }

    $Mail.Subject = "Croissants pour tout le monde !"
    $Mail.Body = "Coucou, je vous apporte les croissants ASAP. Remerciez-moi chaleureusement !"
    $Mail.Send()
}

Start-Process "https://www.croissantage.fr"
