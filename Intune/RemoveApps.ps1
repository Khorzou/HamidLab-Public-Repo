<#
.SYNOPSIS
  Remove selected Windows 10/11 inbox apps (installed + provisioned).

.DESCRIPTION
  - Intended to run as SYSTEM via Intune Management Extension.
  - Safe to run post-Autopilot / Device Preparation (after ESP / first logon).
  - Uses an explicit allowlist of packages to remove.
  - Logs all actions to C:\Windows\Logs\AUM-AppRemoval.log
#>

$ErrorActionPreference = 'Stop'

# region: logging
$logDir  = 'C:\Windows\Logs'
$logPath = Join-Path $logDir 'AUM-AppRemoval.log'

if (-not (Test-Path $logDir)) {
    New-Item -Path $logDir -ItemType Directory -Force | Out-Null
}

function Write-Log {
    param([string]$Message)
    "$(Get-Date -Format s)  $Message" | Tee-Object -FilePath $logPath -Append
}

Write-Log "==== AUM App Removal script starting ===="

# region: app patterns to remove (explicit allowlist)
# Use full or partial PackageName (the -like operator is used)
# Adjust this list to your environment.
$AppPatterns = @(
    'Microsoft.Microsoft3DViewer',
    'Clipchamp.Clipchamp',
    'Microsoft.MicrosoftSolitaireCollection',
    'Microsoft.BingNews',
    'Microsoft.BingWeather',
    'Microsoft.BingSports',
    'Microsoft.WindowsMaps',
    'Microsoft.MixedReality.Portal',
    'Microsoft.People',
    'Microsoft.MicrosoftOfficeHub',
    'Microsoft.XboxApp',
    'Microsoft.GamingApp',
    'Microsoft.Xbox.TCUI',
    'Microsoft.XboxGameOverlay',
    'Microsoft.XboxGamingOverlay',
    'Microsoft.XboxIdentityProvider',
    'Microsoft.XboxSpeechToTextOverlay',
    'SpotifyAB.SpotifyMusic',
    'Netflix.Netflix',
    'Microsoft.PowerAutomateDesktop'
)

Write-Log "App patterns to remove:`n$($AppPatterns -join ', ')"

# region: helper to remove installed Appx
function Remove-InstalledAppx {
    param(
        [Parameter(Mandatory)]
        [string]$Pattern
    )

    $apps = Get-AppxPackage -AllUsers | Where-Object { $_.Name -like "*$Pattern*" }

    foreach ($app in $apps) {
        try {
            Write-Log "Removing installed appx: $($app.Name) | PackageFullName: $($app.PackageFullName)"
            Remove-AppxPackage -Package $app.PackageFullName -AllUsers -ErrorAction Stop
            Write-Log "Successfully removed installed appx: $($app.Name)"
        }
        catch {
            Write-Log "FAILED to remove installed appx $($app.Name): $($_.Exception.Message)"
        }
    }

    if (-not $apps) {
        Write-Log "No installed appx found for pattern '$Pattern'"
    }
}

# region: helper to remove provisioned Appx
function Remove-ProvisionedAppx {
    param(
        [Parameter(Mandatory)]
        [string]$Pattern
    )

    $provApps = Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -like "*$Pattern*" }

    foreach ($prov in $provApps) {
        try {
            Write-Log "Removing provisioned appx: $($prov.DisplayName) | PackageName: $($prov.PackageName)"
            Remove-AppxProvisionedPackage -Online -PackageName $prov.PackageName -ErrorAction Stop | Out-Null
            Write-Log "Successfully removed provisioned appx: $($prov.DisplayName)"
        }
        catch {
            Write-Log "FAILED to remove provisioned appx $($prov.DisplayName): $($_.Exception.Message)"
        }
    }

    if (-not $provApps) {
        Write-Log "No provisioned appx found for pattern '$Pattern'"
    }
}

# region: main loop
foreach ($pattern in $AppPatterns) {
    Write-Log "Processing pattern: $pattern"

    # Remove provisioned first (so it doesn't get re-installed for new profiles)
    Remove-ProvisionedAppx -Pattern $pattern

    # Then remove installed apps for existing users
    Remove-InstalledAppx -Pattern $pattern
}

Write-Log "==== AUM App Removal script completed ===="
exit 0
