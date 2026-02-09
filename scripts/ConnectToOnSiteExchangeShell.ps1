#
# Connect to Onsite Exchange Management Shell - Interactive Script
# Created by ChristopherWStevenson
#

# Verify PowerShell version. Require Windows PowerShell 5.x.
Write-Host "PowerShell v5 required. Detecting version..."
$psVersion = $PSVersionTable.PSVersion
if ($psVersion -and $psVersion.Major -eq 5) {
    Write-Host "PowerShell $psVersion detected. Continuing..." -ForegroundColor Green
}
elseif ($psVersion -and $psVersion.Major -ge 7) {
    Write-Host "PowerShell $psVersion detected. This script requires Windows PowerShell 5.x and is not supported on PowerShell 7+" -ForegroundColor Red
    Write-Host "Press any key to exit..."
    try {
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    } catch {
        Read-Host -Prompt "Press Enter to end script"
    }
    exit 1
}
else {
    Write-Host "PowerShell $psVersion detected. This script expects Windows PowerShell 5.x. Exiting." -ForegroundColor Red
    Write-Host "Press any key to exit..."
    try {
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    } catch {
        Read-Host -Prompt "Press Enter to end script"
    }
    exit 1
}

# Check to see if the Exchange management tools are installed
Write-Host "Results of Exchange connection test:"
$exchangeInstallPath = $env:ExchangeInstallPath -as [string]

if ([string]::IsNullOrWhiteSpace($exchangeInstallPath)) {
    # Environment variable missing or empty — don't call Join-Path or Test-Path
    $hasExchangeTools = $false
    $remoteExchange = $null
} else {
    try {
        # Guard Join-Path in try/catch in case path is malformed
        $remoteExchange = Join-Path -Path $exchangeInstallPath -ChildPath 'bin\RemoteExchange.ps1'
    } catch {
        $remoteExchange = $null
    }

    # Only call Test-Path when we have a non-null candidate path; suppress errors
    if ($remoteExchange) {
        $hasExchangeTools = Test-Path -Path $remoteExchange -PathType Leaf -ErrorAction SilentlyContinue
    } else {
        $hasExchangeTools = $false
    }
}

Write-Host "RemoteExchange.ps1 present: $hasExchangeTools"
Write-Host "If this returns False, the tools are missing."

if ($hasExchangeTools) {
    Write-Host "Using local Exchange Management Shell (dot-sourcing RemoteExchange.ps1)..."
    try {
        . $remoteExchange
        Connect-ExchangeServer -Auto -ClientApplication:ManagementShell -ErrorAction Stop
        Write-Host "Connected using local Exchange Management Shell."
    } catch {
        # Fail quietly for initialization and fall back to remote method
        Write-Host "Failed to initialize local Exchange Management Shell." -ForegroundColor Red
        $hasExchangeTools = $false
    }
}

if (-not $hasExchangeTools) {
    
    Write-Host "Exchange Management Tools not found locally." -ForegroundColor Red

    Write-Host "Enter Exchange server FQDN or host name for remote PowerShell (example: exch01.contoso.local)" -ForegroundColor Green
    $serverName = Read-Host

    if ([string]::IsNullOrWhiteSpace($serverName)) {
        Write-Host "No server name provided. Aborting Exchange connection step."
    } else {
        $connectionUri = "http://$serverName/PowerShell/"
        Write-Host "Attempting remote connection to $connectionUri"

        try {
            $session = New-PSSession -ConfigurationName Microsoft.Exchange `
                -ConnectionUri $connectionUri `
                -Authentication Kerberos -AllowRedirection -ErrorAction Stop

            Import-PSSession $session -DisableNameChecking -AllowClobber -ErrorAction Stop
            Write-Host "Connected to Exchange via remote session on $serverName." -ForegroundColor Green

            # -- informational cleanup hint (shows session Id and computer name)
            if ($session) {
                Write-Host (" --- cleanup connection when complete by running: Remove-PSSession -Id {0} (or Remove-PSSession -Session `$session) ---" -f $session.Id) -ForegroundColor Yellow
                Write-Host (" --- session target: {0} ---" -f $session.ComputerName) -ForegroundColor Yellow
            } else {
                Write-Host " --- no session object available to remove ---" -ForegroundColor Yellow
            }
        } catch {
            # Diagnostics and primary result message
            Write-Host "Failed to create remote Exchange session to ${serverName}." -ForegroundColor Red

            Write-Host "Error summary:" -ForegroundColor Yellow
            if ($_.Exception) {
                Write-Host $_.Exception.Message -ForegroundColor Yellow
                if ($_.Exception.InnerException) {
                    Write-Host "Inner exception:" -ForegroundColor Yellow
                    Write-Host $_.Exception.InnerException.Message -ForegroundColor Yellow
                }
            } else {
                Write-Host $_.ToString() -ForegroundColor Yellow
            }

            if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
                Write-Host "Error details:" -ForegroundColor Yellow
                Write-Host $_.ErrorDetails.Message -ForegroundColor Yellow
            }

            # Optional: show the full error record (string) and script stack trace if present
            Write-Host "Full error record (for debugging):" -ForegroundColor DarkYellow
            Write-Host $_.ToString()

            if ($_.ScriptStackTrace) {
                Write-Host "Script stack trace:" -ForegroundColor DarkYellow
                Write-Host $_.ScriptStackTrace
            }

            Write-Host "Suggested checks:" -ForegroundColor Cyan
            Write-Host " - Can you reach the server? Run: Test-WSMan -ComputerName $serverName" -ForegroundColor Cyan
            Write-Host " - Check network port: Test-NetConnection -ComputerName $serverName -Port 5985" -ForegroundColor Cyan
            Write-Host " - Verify authentication and RBAC rights for the account used to run this script." -ForegroundColor Cyan

            Write-Host "Press any key to continue..."
            try {
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            } catch {
                Read-Host -Prompt "Press Enter to continue"
            }
        }
    }
}

