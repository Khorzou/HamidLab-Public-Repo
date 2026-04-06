<#
.SYNOPSIS
Find and optionally delete unassigned Win32 apps from Intune via Microsoft Graph.

.DESCRIPTION
- Ensures required Microsoft Graph PowerShell modules are installed
- Imports Graph modules explicitly for VSCode compatibility
- Connects to Microsoft Graph
- Retrieves Intune mobile apps
- Filters to Win32 apps only
- Checks assignments for each app
- Exports a CSV report
- Dry-run by default
- Deletes only unassigned Win32 apps when -Delete is specified

.NOTES
Recommended to run in:
- Windows PowerShell 5.1 or PowerShell 7
- VSCode integrated terminal

Required Graph delegated scope:
- DeviceManagementApps.ReadWrite.All

Examples:
.\Remove-Unassigned-IntuneWin32Apps.ps1
.\Remove-Unassigned-IntuneWin32Apps.ps1 -DisplayNameLike "OLD-*"
.\Remove-Unassigned-IntuneWin32Apps.ps1 -Delete -Confirm:$false
.\Remove-Unassigned-IntuneWin32Apps.ps1 -DisplayNameLike "OLD-*" -Delete -Confirm:$false
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$DisplayNameLike = "*",
    [switch]$Delete,
    [string]$ReportPath = ".\Intune-Unassigned-Win32Apps-Report.csv",
    [string]$LogPath = ".\Intune-Unassigned-Win32Apps-Log.txt"
)

# ----------------------------
# Initial session hardening
# ----------------------------

$ErrorActionPreference = 'Stop'

try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
} catch {
    # Ignore if not applicable
}

# ----------------------------
# Helper functions
# ----------------------------

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet("INFO","WARN","ERROR")]
        [string]$Level = "INFO"
    )

    $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Write-Host $line

    $logFolder = Split-Path -Path $LogPath -Parent
    if ($logFolder -and -not (Test-Path -LiteralPath $logFolder)) {
        New-Item -Path $logFolder -ItemType Directory -Force | Out-Null
    }

    Add-Content -Path $LogPath -Value $line
}

function Ensure-Module {
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName
    )

    $available = Get-Module -ListAvailable -Name $ModuleName

    if (-not $available) {
        Write-Log "Module '$ModuleName' is not installed. Installing for current user..."
        try {
            Install-Module -Name $ModuleName -Scope CurrentUser -Repository PSGallery -Force -AllowClobber
            Write-Log "Installed module '$ModuleName'."
        }
        catch {
            throw "Failed to install module '$ModuleName'. $($_.Exception.Message)"
        }
    }
    else {
        Write-Log "Module '$ModuleName' is already installed."
    }

    try {
        Import-Module $ModuleName -Force -ErrorAction Stop
        Write-Log "Imported module '$ModuleName'."
    }
    catch {
        throw "Failed to import module '$ModuleName'. $($_.Exception.Message)"
    }
}

function Ensure-GraphPrerequisites {
    Write-Log "Checking Microsoft Graph PowerShell prerequisites..."

    if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
        Write-Log "NuGet package provider not found. Installing..."
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null
    }

    try {
        $repo = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
        if (-not $repo) {
            throw "PSGallery repository is missing."
        }

        if ($repo.InstallationPolicy -ne 'Trusted') {
            Write-Log "Setting PSGallery as Trusted."
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
        }
    }
    catch {
        throw "Failed to prepare PSGallery. $($_.Exception.Message)"
    }

    Ensure-Module -ModuleName "Microsoft.Graph.Authentication"
}

function Ensure-GraphConnection {
    param(
        [string[]]$Scopes = @("DeviceManagementApps.ReadWrite.All")
    )

    try {
        $ctx = $null

        if (Get-Command -Name Get-MgContext -ErrorAction SilentlyContinue) {
            $ctx = Get-MgContext -ErrorAction SilentlyContinue
        }

        if (-not $ctx -or -not $ctx.Account) {
            Write-Log "Connecting to Microsoft Graph..."
            Connect-MgGraph -Scopes $Scopes -NoWelcome | Out-Null
            $ctx = Get-MgContext
        }

        if (-not $ctx -or -not $ctx.Account) {
            throw "Graph context was not established."
        }

        Write-Log ("Connected to Graph. TenantId={0} Account={1}" -f $ctx.TenantId, $ctx.Account)
    }
    catch {
        throw "Failed to connect to Microsoft Graph. $($_.Exception.Message)"
    }
}

function Invoke-GraphGetAllPages {
    param(
        [Parameter(Mandatory)]
        [string]$Uri
    )

    $allItems = @()
    $next = $Uri

    while ($next) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject

        if ($response.value) {
            $allItems += @($response.value)
        }

        $next = $response.'@odata.nextLink'
    }

    return $allItems
}

function Split-IntoChunks {
    param(
        [Parameter(Mandatory)]
        [array]$Items,

        [int]$ChunkSize = 20
    )

    if (-not $Items -or $Items.Count -eq 0) {
        return @()
    }

    $chunks = @()

    for ($i = 0; $i -lt $Items.Count; $i += $ChunkSize) {
        $end = [Math]::Min($i + $ChunkSize - 1, $Items.Count - 1)
        $chunks += ,($Items[$i..$end])
    }

    return $chunks
}

function Get-Win32Apps {
    Write-Log "Retrieving all Intune mobile apps..."

    $apps = Invoke-GraphGetAllPages -Uri "https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps"

    $win32Apps = @($apps | Where-Object {
        $_.'@odata.type' -eq '#microsoft.graph.win32LobApp'
    })

    Write-Log ("Retrieved {0} total mobile apps; {1} are Win32 apps." -f @($apps).Count, $win32Apps.Count)

    return $win32Apps
}

function Get-AppAssignmentCountsBatch {
    param(
        [Parameter(Mandatory)]
        [array]$Apps
    )

    $results = @{}
    $chunks = Split-IntoChunks -Items $Apps -ChunkSize 20
    $chunkNumber = 0

    foreach ($chunk in $chunks) {
        $chunkNumber++
        Write-Log ("Checking assignments for chunk {0}/{1}..." -f $chunkNumber, $chunks.Count)

        $requests = @()
        $counter = 1

        foreach ($app in $chunk) {
            $requests += @{
                id     = [string]$counter
                method = "GET"
                url    = "/deviceAppManagement/mobileApps/$($app.id)/assignments"
            }
            $counter++
        }

        $batchBody = @{
            requests = $requests
        } | ConvertTo-Json -Depth 10

        $batchResponse = Invoke-MgGraphRequest `
            -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/`$batch" `
            -Body $batchBody `
            -ContentType "application/json" `
            -OutputType PSObject

        foreach ($resp in $batchResponse.responses) {
            $reqId = [int]$resp.id
            $app = $chunk[$reqId - 1]

            if ($resp.status -eq 200) {
                $assignmentCount = 0

                if ($resp.body -and $resp.body.value) {
                    $assignmentCount = @($resp.body.value).Count
                }

                $results[$app.id] = $assignmentCount
            }
            else {
                Write-Log ("Failed to query assignments for app '{0}' ({1}). HTTP {2}" -f $app.displayName, $app.id, $resp.status) "ERROR"
                $results[$app.id] = $null
            }
        }
    }

    return $results
}

function Remove-AppsBatch {
    param(
        [Parameter(Mandatory)]
        [array]$AppsToDelete
    )

    $deleteResults = @()
    $chunks = Split-IntoChunks -Items $AppsToDelete -ChunkSize 20
    $chunkNumber = 0

    foreach ($chunk in $chunks) {
        $chunkNumber++
        Write-Log ("Deleting chunk {0}/{1}..." -f $chunkNumber, $chunks.Count)

        $requests = @()
        $counter = 1

        foreach ($app in $chunk) {
            $requests += @{
                id     = [string]$counter
                method = "DELETE"
                url    = "/deviceAppManagement/mobileApps/$($app.id)"
            }
            $counter++
        }

        $batchBody = @{
            requests = $requests
        } | ConvertTo-Json -Depth 10

        $batchResponse = Invoke-MgGraphRequest `
            -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/`$batch" `
            -Body $batchBody `
            -ContentType "application/json" `
            -OutputType PSObject

        foreach ($resp in $batchResponse.responses) {
            $reqId = [int]$resp.id
            $app = $chunk[$reqId - 1]
            $success = ($resp.status -ge 200 -and $resp.status -lt 300)

            if ($success) {
                Write-Log ("Deleted app '{0}' ({1})" -f $app.displayName, $app.id)
            }
            else {
                Write-Log ("Failed to delete app '{0}' ({1}). HTTP {2}" -f $app.displayName, $app.id, $resp.status) "ERROR"
            }

            $deleteResults += [pscustomobject]@{
                DisplayName = $app.displayName
                Id          = $app.id
                Publisher   = $app.publisher
                DeleteHttp  = $resp.status
                Deleted     = $success
            }
        }
    }

    return $deleteResults
}

# ----------------------------
# Main
# ----------------------------

try {
    $resolvedLogPath = [System.IO.Path]::GetFullPath($LogPath)
    $resolvedReportPath = [System.IO.Path]::GetFullPath($ReportPath)

    $script:LogPath = $resolvedLogPath
    $script:ReportPath = $resolvedReportPath

    if (Test-Path -LiteralPath $script:LogPath) {
        Remove-Item -LiteralPath $script:LogPath -Force -ErrorAction SilentlyContinue
    }

    Write-Log "Starting Intune unassigned Win32 app discovery."
    Write-Log ("PowerShell edition: {0} Version: {1}" -f $PSVersionTable.PSEdition, $PSVersionTable.PSVersion)
    Write-Log ("Running from: {0}" -f (Get-Location).Path)

    Ensure-GraphPrerequisites
    Ensure-GraphConnection

    $win32Apps = Get-Win32Apps

    if (-not $win32Apps -or $win32Apps.Count -eq 0) {
        Write-Log "No Win32 apps found."
        return
    }

    $filteredApps = @($win32Apps | Where-Object {
        $_.displayName -like $DisplayNameLike
    })

    Write-Log ("Filtered Win32 apps count: {0}" -f $filteredApps.Count)

    if (-not $filteredApps -or $filteredApps.Count -eq 0) {
        Write-Log "No Win32 apps matched the supplied display name filter."
        return
    }

    $assignmentCounts = Get-AppAssignmentCountsBatch -Apps $filteredApps

    $report = foreach ($app in $filteredApps) {
        $count = $assignmentCounts[$app.id]

        [pscustomobject]@{
            DisplayName      = $app.displayName
            Id               = $app.id
            Publisher        = $app.publisher
            IsFeatured       = $app.isFeatured
            AssignmentCount  = $count
            IsUnassigned     = ($count -eq 0)
            CreatedDateTime  = $app.createdDateTime
            LastModifiedDate = $app.lastModifiedDateTime
        }
    }

    $report = @($report | Sort-Object DisplayName)
    $report | Export-Csv -Path $script:ReportPath -NoTypeInformation -Encoding UTF8
    Write-Log ("Report exported to: {0}" -f $script:ReportPath)

    $unassignedApps = @($report | Where-Object { $_.IsUnassigned -eq $true })

    Write-Log ("Unassigned Win32 apps found: {0}" -f $unassignedApps.Count)

    if (-not $Delete) {
        Write-Log "Dry run only. No apps were deleted."

        if ($unassignedApps.Count -gt 0) {
            Write-Host ""
            Write-Host "Unassigned Win32 apps:"
            $unassignedApps | Format-Table DisplayName, Publisher, AssignmentCount, Id -AutoSize
        }
        else {
            Write-Host ""
            Write-Host "No unassigned Win32 apps found."
        }

        return
    }

    if ($unassignedApps.Count -eq 0) {
        Write-Log "Nothing to delete."
        return
    }

    $appsToDelete = foreach ($entry in $unassignedApps) {
        $filteredApps | Where-Object { $_.id -eq $entry.Id }
    }

    $appsToDelete = @($appsToDelete)

    Write-Log ("Apps queued for deletion: {0}" -f $appsToDelete.Count)

    if ($PSCmdlet.ShouldProcess(("Win32 apps ({0})" -f $appsToDelete.Count), "Delete unassigned Intune Win32 apps")) {
        $deleteResults = Remove-AppsBatch -AppsToDelete $appsToDelete

        $deleteReportPath = [System.IO.Path]::ChangeExtension($script:ReportPath, ".delete-results.csv")
        $deleteResults | Export-Csv -Path $deleteReportPath -NoTypeInformation -Encoding UTF8

        Write-Log ("Delete results exported to: {0}" -f $deleteReportPath)
        Write-Log "Completed."
    }
}
catch {
    try {
        Write-Log $_.Exception.Message "ERROR"
    }
    catch {
        Write-Host ("ERROR: {0}" -f $_.Exception.Message)
    }

    throw
}