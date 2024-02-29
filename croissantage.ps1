# .\croissantage.ps1 -invites "mail1@test.fr","mail2@test.fr"

param(
    [string[]]$invites
)

 

function Get-NextMonday {
    $today = Get-Date
    $dayOfWeek = $today.DayOfWeek.value__
    $dayUntilNextMonday = (8 - $dayOfWeek) % 8
    $nextMonday = $today.AddDays($dayUntilNextMonday)
    return $nextMonday.Date.Addhours(8)
}

$Outlook = New-Object -ComObject Outlook.Application
$Appointement = $Outlook.CreateItem(1)
$Appointement.Subject = "Croissants pour tous !"
$Appointement.Body = "Des croissants seront a disposition de petits et grands ce Lundi, pour votre plus grand plaisir. Ne pas hesiter a me remercier chaleureusement."
$Appointement.Start = Get-NextMonday
$Appointement.End = $Appointement.Start.AddHours(1)
foreach ($invite in $invites) { $Appointement.Recipients.Add($invite) }
$Appointement.Save()
$Appointement.Send()

Start-Process "https://www.croissantage.fr"
