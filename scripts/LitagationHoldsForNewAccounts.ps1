#Find mailboxes created in the last 60 days and set them to litigation hold
#Handy for setting mail retention on newly created accounts
#created by ChristopherWStevenson
#Requires ExchangeOnlineManagement PS module: https://www.powershellgallery.com/packages/ExchangeOnlineManagement/
#
$days = "[Number of days to set lit hold]"
#
Connect-ExchangeOnline -UserPrincipalName "[Enter Exchange Admin UPN Here]"
#
#You can change how many days back to look at accounts on the next two lines (change the 60)
Write-Host "--Setting $days day litigation hold for all mailboxes created in the last 60 days--"
$Date = (Get-Date).AddDays(-60)
#
#You can comment out the next line, and remove comment from the second line down, to set the lit hold only on user mailboxes
$Mailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object {$_.WhenCreated -gt $Date}
#$Mailboxes = Get-Mailbox -ResultSize Unlimited -Filter {(RecipientTypeDetails -eq 'UserMailbox') -and (WhenCreated -ge $Date) -and (LitigationHoldEnabled -eq $false)}
#
$Mailboxes | Set-Mailbox -LitigationHoldEnabled $true -LitigationHoldDuration $days


