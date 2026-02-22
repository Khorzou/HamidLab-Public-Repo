<#
.SYNOPSIS
  Rename device to AUM-%SERIAL% after enrollment.

.NOTES
  - Run as SYSTEM (Intune Devices > Scripts or Proactive Remediation).
  - Log: C:\Windows\Logs\AUM-Rename.log
  - NetBIOS limit: 15 chars (AUM- + 11-char suffix).
#>

# region: prep & logging
$ErrorActionPreference = 'Stop'
$logDir  = 'C:\Windows\Logs'
$logPath = Join-Path $logDir 'AUM-Rename.log'
if (-not (Test-Path $logDir)) { New-Item -Path $logDir -ItemType Directory -Force | Out-Null }

function Write-Log {
    param([string]$Message)
    "$(Get-Date -Format s)  $Message" | Tee-Object -FilePath $logPath -Append
}

try {
    Write-Log "==== AUM rename script starting ===="

    # region: get serial (prefer BIOS SerialNumber; fallback to UUID)
    $serial = try { (Get-CimInstance -ClassName Win32_BIOS).SerialNumber } catch { $null }
    if ([string]::IsNullOrWhiteSpace($serial)) {
        $serial = try { (Get-CimInstance -ClassName Win32_ComputerSystemProduct).UUID } catch { $null }
    }
    if ([string]::IsNullOrWhiteSpace($serial)) { $serial = 'UNKNOWN' }

    # sanitize serial: uppercase, alphanumeric only
    $sanitized = ($serial.ToUpper() -replace '[^A-Z0-9]', '')
    if ([string]::IsNullOrWhiteSpace($sanitized)) { $sanitized = 'UNKNOWN' }

    # NetBIOS length: AUM- (4 chars) + up to 11 serial chars
    $suffix = if ($sanitized.Length -gt 11) {
        $sanitized.Substring($sanitized.Length - 11)  # rightmost 11
    } else {
        $sanitized
    }
    $targetName = "AUM-$suffix"
    if ($targetName.Length -gt 15) { $targetName = $targetName.Substring(0,15) }

    $current = $env:COMPUTERNAME
    Write-Log "Current name: $current | Serial (raw): '$serial' | Target: $targetName"

    if ($current -eq $targetName) {
        Write-Log "Device already named correctly. No action taken."
        exit 0
    }

    # region: rename
    Write-Log "Renaming computer to: $targetName"
    try {
        Rename-Computer -NewName $targetName -Force -ErrorAction Stop
        Write-Log "Rename succeeded. Scheduling reboot in 30 seconds to apply."
        shutdown.exe /r /t 30 /c "Renamed to $targetName. Rebooting to apply name."
        exit 0
    }
    catch {
        Write-Log "Rename failed: $($_.Exception.Message)"
        throw
    }
}
catch {
    Write-Log "Fatal error: $($_.Exception.Message)"
    exit 1
}
