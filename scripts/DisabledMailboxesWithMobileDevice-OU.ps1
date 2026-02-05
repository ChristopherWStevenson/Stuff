#Gets disabled accounts from defined OU and searches their exchange online account 
#to see if there are mobile device associated, and exports the results to csv.
#
#Also uses out-gridview to display results immediately.
#
#Requires PS modules ActiveDirectory & ExchangeOnlineManagement
#
#ChristopherWStevenson

#Enter User Principal Name for connection:
$ExchangeAdminAccount = "Exchange Admin UPN"
#Define the OU to search for disabled accounts:
$OU = "OU=DisabledUsers,DC=domain,DC=net"
#path to save csv file:
$CsvFilePath = "C:\temp\DisabledUsersWithMobileDevicesSpecificOU.csv"

Write-Host "Pulling data from AD (OU: $OU)..."
#Import Active Directory module
Import-Module ActiveDirectory
#Get disabled user accounts from the specified OU
$DisabledUsers = Get-ADUser -Filter {Enabled -eq $false} -SearchBase $OU -Properties EmailAddress
$disabledCount = if ($DisabledUsers) { $DisabledUsers.Count } else { 0 }
Write-Host "Found $disabledCount disabled user(s) in $OU."

Write-Host "Connecting to Exchange Online as $ExchangeAdminAccount..."
# Connect to Exchange Online
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -UserPrincipalName $ExchangeAdminAccount
Write-Host "Connected to Exchange Online."

function Get-FirstPropertyValue {
	param (
		[Parameter(Mandatory = $true)] $Object,
		[Parameter(Mandatory = $true)] [string[]] $PropertyNames
	)
	foreach ($name in $PropertyNames) {
		if ($Object -and $Object.PSObject.Properties.Name -contains $name) {
			$value = $Object.$name
			if ($null -ne $value -and $value -ne '') {
				return $value
			}
		}
	}
	return $null
}

$results = [System.Collections.Generic.List[psobject]]::new()

# Prepare list of users that actually have an email address
$UsersWithEmail = $DisabledUsers | Where-Object { $_.EmailAddress -and $_.EmailAddress.Trim() -ne '' }
$totalToProcess = if ($UsersWithEmail) { $UsersWithEmail.Count } else { 0 }
Write-Host "Processing $totalToProcess user(s) with email addresses to collect mobile device info."

$index = 0
foreach ($User in $UsersWithEmail) {
	$index++
	$UserEmail = $User.EmailAddress

	$percent = [int](($index / $totalToProcess) * 100)
	Write-Progress -Activity "Collecting mobile device info" -Status "Processing $UserEmail ($index of $totalToProcess)" -PercentComplete $percent

	# Prefer statistics cmdlet (returns sync timestamps)
	$MobileDevices = Get-MobileDeviceStatistics -Mailbox $UserEmail -ErrorAction SilentlyContinue

	# fall back to Get-MobileDevice if absolutely needed
	if (-not $MobileDevices) {
		$MobileDevices = Get-MobileDevice -Mailbox $UserEmail -ErrorAction SilentlyContinue
	}

	if ($MobileDevices) {
		foreach ($Device in $MobileDevices) {
			# If timestamps are missing, try re-querying the device by Identity to get full stats
			$needsRequery = ($null -eq (Get-FirstPropertyValue -Object $Device -PropertyNames @('LastSuccessSync','LastSyncAttemptTime'))) -and
						   ($Device.PSObject.Properties.Name -contains 'Identity' -and $Device.Identity)
			if ($needsRequery) {
				try {
					$ref = Get-MobileDeviceStatistics -Identity $Device.Identity -ErrorAction Stop
					if ($ref) { $Device = $ref }
				} catch {
					# ignore and continue with whatever we have
				}
			}

			$deviceId        = Get-FirstPropertyValue -Object $Device -PropertyNames @('DeviceId','DeviceIdentity','Identity','Id')
			$deviceOS        = Get-FirstPropertyValue -Object $Device -PropertyNames @('DeviceOS','OperatingSystem')
			$deviceType      = Get-FirstPropertyValue -Object $Device -PropertyNames @('DeviceType','ClientType')
			$deviceModel     = Get-FirstPropertyValue -Object $Device -PropertyNames @('DeviceModel','DeviceModelString','Model')
			$firstSync       = Get-FirstPropertyValue -Object $Device -PropertyNames @('FirstSyncTime','FirstSuccessfulSync')
			$lastSuccessSync = Get-FirstPropertyValue -Object $Device -PropertyNames @('LastSuccessSync','LastSuccessfulSync')
			$lastSyncAttempt = Get-FirstPropertyValue -Object $Device -PropertyNames @('LastSyncAttemptTime','LastAttemptTime')

			$results.Add([pscustomobject]@{
				Mailbox             = $UserEmail
				UserDisplayName     = $User.Name
				DeviceId            = $deviceId
				DeviceOS            = $deviceOS
				DeviceType          = $deviceType
				DeviceModel         = $deviceModel
				ClientType          = $deviceType
				FirstSyncTime       = $firstSync
				LastSuccessSync     = $lastSuccessSync
				LastSyncAttemptTime = $lastSyncAttempt
			})
		}
	}
}

# Complete progress bar
Write-Progress -Activity "Collecting mobile device info" -Completed
Write-Host "Collection complete. Found $($results.Count) device entry(ies)."

if ($results.Count -gt 0) {
	Write-Host "Opening Out-GridView (may be behind other windows)..."
	$results | Out-GridView -Title "Disabled Users with Mobile Devices in $OU"

	# Ensure output directory exists
	$dir = Split-Path -Path $CsvFilePath -Parent
	if (-not (Test-Path $dir)) {
		New-Item -Path $dir -ItemType Directory -Force | Out-Null
	}

	# Prepare export header and write CSV data safely
	$exportHeader = "Disabled Users with Mobile Devices from OU: $OU`nExport Date/Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n"

	$tempCsv = [IO.Path]::GetTempFileName()
	try {
		Write-Host "Exporting results to CSV..."
		# Export results to a proper CSV (this creates valid CSV header row)
		$results | Export-Csv -Path $tempCsv -NoTypeInformation -Encoding UTF8

		# Write metadata header then append the CSV content
		$exportHeader | Out-File -FilePath $CsvFilePath -Encoding UTF8
		Get-Content -Path $tempCsv | Out-File -FilePath $CsvFilePath -Encoding UTF8 -Append
		Write-Host "Exported $($results.Count) device entries to $CsvFilePath"
	} finally {
		Remove-Item -Path $tempCsv -ErrorAction SilentlyContinue
	}
} else {
	Write-Host "No disabled users with mobile devices found in $OU. No CSV created."
}

# Disconnect from Exchange Online
Write-Host "Disconnecting from Exchange Online..."
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Disconnected."

