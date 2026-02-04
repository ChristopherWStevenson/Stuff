#Search Exchange online for disabled accounts that have an out of office auto reply enabled
#Displays the results in the PS session
#created by ChristopherWStevenson
#requires ExchangeOnlineManagement PS module: https://www.powershellgallery.com/packages/ExchangeOnlineManagement/

$ExchangeAdminUPN = "[Your Exchange Admin UPN here]"

#Connect to Exchange Online
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -UserPrincipalName $ExchangeAdminUPN
 
#Get all disabled mailboxes with out of office auto reply enabled
$disabledMailboxesWithOOO = Get-Mailbox -ResultSize Unlimited | Where-Object { $_.AccountDisabled -eq $true } | ForEach-Object {
	$mailbox = $_
	$oooSettings = Get-MailboxAutoReplyConfiguration -Identity $mailbox.Identity
	if ($oooSettings.AutoReplyState -ne 'Disabled') {
		return $mailbox
	}
}
#Output the results
$disabledMailboxesWithOOO | Select-Object DisplayName,PrimarySmtpAddress