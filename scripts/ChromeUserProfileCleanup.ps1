# ================================
# Chrome User Profile Cleanup
# Run this after you have disabled creation of Chrome profiles
# Computer Shutdown Script will run as SYSTEM
# Created by ChristopherWStevenson
# ================================

$marker = 'C:\ProgramData\ChromeProfileCleanup.done'

# Exit if we've already run
if (Test-Path $marker) {
    exit 0
}

# Make sure Chrome is not running
Get-Process chrome -ErrorAction SilentlyContinue | Stop-Process -Force

# User profile folders to skip
$excludedUsers = @(
    'Public',
    'Default',
    'Default User',
    'All Users',
    'Administrator'
)

# Loop through each local user profile
Get-ChildItem 'C:\Users' -Directory |
Where-Object { $excludedUsers -notcontains $_.Name } |
ForEach-Object {

    $chromePath = Join-Path $_.FullName 'AppData\Local\Google\Chrome\User Data'

    if (Test-Path $chromePath) {

        # Remove extra Chrome profile directories (Profile 1, Profile 2, etc.)
        Get-ChildItem $chromePath -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^Profile \d+$' } |
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

        # Remove Local State so Chrome rebuilds clean
        $localState = Join-Path $chromePath 'Local State'
        Remove-Item $localState -Force -ErrorAction SilentlyContinue
    }
}

# Create marker file so this only runs once
New-Item $marker -ItemType File -Force | Out-Null

exit 0