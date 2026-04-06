<#
.SYNOPSIS
    Reports direct Intune / Microsoft Graph assignments for a specified Entra ID group
    or across the tenant.

.DESCRIPTION
    Queries multiple Intune workload endpoints and returns items that are directly assigned
    to a specified Entra ID group, or all groups when -AllGroups is used.

    Features:
    - Supports GroupId or GroupName
    - Optional interactive group picker
    - Supports tenant-wide mode without selecting a single group
    - Follows Graph paging
    - CSV / HTML / Excel export
    - Broader Intune workload coverage
    - Returns clean objects to pipeline

.NOTES
Connect to MS Graph first
Connect-MgGraph -Scopes "Group.Read.All,DeviceManagementApps.Read.All,DeviceManagementConfiguration.Read.All,DeviceManagementManagedDevices.Read.All"

    Notes:
    - This reports direct group-targeted assignments.
    - It does not fully calculate effective targeting across nested groups, policy sets,
      or all include/exclude evaluation paths.
    - Excel export requires the ImportExcel module.
    - HTML export is self-contained and does not require extra modules.

.PARAMETER GroupId
    Entra group object ID.

.PARAMETER GroupName
    Entra group display name.

.PARAMETER AllGroups
    Tenant-wide mode. Returns assignments for all directly targeted groups.

.PARAMETER ExportPath
    Base output path. Extension is added automatically per format if needed.

.PARAMETER OutputFormat
    Console, Csv, Html, Excel, or All.

.PARAMETER UseBeta
    Uses Microsoft Graph beta. Default is $true because some Intune workloads are better exposed there.

.PARAMETER SkipInteractivePicker
    Prevents Out-GridView if no group input is supplied.

.EXAMPLE
    .\Get-IntuneAssignmentsByGroup.ps1 -GroupId "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee"

.EXAMPLE
    .\Get-IntuneAssignmentsByGroup.ps1 -GroupName "Intune - Pilot Users" -OutputFormat Csv -ExportPath "C:\Temp\PilotAssignments"

.EXAMPLE
    .\Get-IntuneAssignmentsByGroup.ps1 -AllGroups -OutputFormat Html -ExportPath "C:\Temp\TenantAssignments"

.EXAMPLE
    .\Get-IntuneAssignmentsByGroup.ps1 -AllGroups -OutputFormat All -ExportPath "C:\Temp\TenantAssignments"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidatePattern('^[0-9a-fA-F-]{36}$')]
    [string]$GroupId,

    [Parameter(Mandatory = $false)]
    [string]$GroupName,

    [Parameter(Mandatory = $false)]
    [switch]$AllGroups,

    [Parameter(Mandatory = $false)]
    [string]$ExportPath = "C:\Temp\IntuneAssignments",

    [Parameter(Mandatory = $false)]
    [ValidateSet("Console","Csv","Html","Excel","All")]
    [string]$OutputFormat = "Console",

    [Parameter(Mandatory = $false)]
    [bool]$UseBeta = $true,

    [Parameter(Mandatory = $false)]
    [switch]$SkipInteractivePicker
)

begin {
    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    function Write-Info {
        param([string]$Message)
        Write-Host "[INFO] $Message" -ForegroundColor Cyan
    }

    function Write-WarnMsg {
        param([string]$Message)
        Write-Warning $Message
    }

    function Test-CommandExists {
        param([Parameter(Mandatory = $true)][string]$Name)
        return [bool](Get-Command -Name $Name -ErrorAction SilentlyContinue)
    }

    function Ensure-GraphConnection {
        if (-not (Test-CommandExists -Name 'Get-MgContext')) {
            throw "Microsoft Graph PowerShell module is not available in this session. Install/import Microsoft.Graph."
        }

        $ctx = Get-MgContext
        if (-not $ctx) {
            throw @"
Not connected to Microsoft Graph.

Connect first, for example:
Connect-MgGraph -Scopes "Group.Read.All,DeviceManagementApps.Read.All,DeviceManagementConfiguration.Read.All,DeviceManagementManagedDevices.Read.All"
"@
        }

        Write-Info "Connected to Microsoft Graph tenant [$($ctx.TenantId)] as [$($ctx.Account)]."
    }

    function Invoke-GraphGetPaged {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Uri
        )

        $allItems = [System.Collections.Generic.List[object]]::new()
        $nextUri = $Uri

        do {
            Write-Info "GET $nextUri"
            $response = Invoke-MgGraphRequest -Method GET -Uri $nextUri -OutputType PSObject

            if ($response.PSObject.Properties['value']) {
                foreach ($item in @($response.value)) {
                    [void]$allItems.Add($item)
                }
            }
            else {
                [void]$allItems.Add($response)
            }

            $nextLinkProperty = $response.PSObject.Properties['@odata.nextLink']
            if ($null -ne $nextLinkProperty -and -not [string]::IsNullOrWhiteSpace([string]$nextLinkProperty.Value)) {
                $nextUri = [string]$nextLinkProperty.Value
            }
            else {
                $nextUri = $null
            }
        }
        while ($nextUri)

        return $allItems
    }

    function Resolve-TargetGroup {
        param(
            [string]$InputGroupId,
            [string]$InputGroupName
        )

        $graphVersion = if ($UseBeta) { "beta" } else { "v1.0" }

        if ($InputGroupId) {
            $uri = "https://graph.microsoft.com/$graphVersion/groups/$InputGroupId?`$select=id,displayName,description"
            try {
                $group = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject
                return [pscustomobject]@{
                    Id          = $group.id
                    DisplayName = $group.displayName
                    Description = $group.description
                }
            }
            catch {
                throw "Failed to resolve group by GroupId [$InputGroupId]. $($_.Exception.Message)"
            }
        }

        if ($InputGroupName) {
            $escapedName = $InputGroupName.Replace("'", "''")
            $uri = "https://graph.microsoft.com/$graphVersion/groups?`$select=id,displayName,description&`$filter=displayName eq '$escapedName'"
            $matches = Invoke-GraphGetPaged -Uri $uri

            if (-not $matches -or $matches.Count -eq 0) {
                throw "No group found with display name [$InputGroupName]."
            }

            if ($matches.Count -eq 1) {
                return [pscustomobject]@{
                    Id          = $matches[0].id
                    DisplayName = $matches[0].displayName
                    Description = $matches[0].description
                }
            }

            if (Test-CommandExists -Name 'Out-GridView') {
                $selected = $matches |
                    Select-Object @{n='DisplayName';e={$_.displayName}}, @{n='Id';e={$_.id}}, @{n='Description';e={$_.description}} |
                    Out-GridView -Title "Multiple groups matched [$InputGroupName] - select one" -PassThru

                if (-not $selected) {
                    throw "No group selected."
                }

                return [pscustomobject]@{
                    Id          = $selected.Id
                    DisplayName = $selected.DisplayName
                    Description = $selected.Description
                }
            }

            throw "Multiple groups matched [$InputGroupName]. Please rerun with -GroupId."
        }

        if ($SkipInteractivePicker) {
            throw "No GroupId or GroupName supplied, and interactive picker was disabled."
        }

        if (-not (Test-CommandExists -Name 'Out-GridView')) {
            throw "Out-GridView is not available. Supply -GroupId, -GroupName, or use -AllGroups."
        }

        $uri = "https://graph.microsoft.com/$graphVersion/groups?`$select=id,displayName,description"
        $groups = Invoke-GraphGetPaged -Uri $uri

        $selectedGroup = $groups |
            Select-Object @{n='DisplayName';e={$_.displayName}}, @{n='Id';e={$_.id}}, @{n='Description';e={$_.description}} |
            Sort-Object DisplayName |
            Out-GridView -Title "Select target Entra ID group" -PassThru

        if (-not $selectedGroup) {
            throw "No group selected."
        }

        return [pscustomobject]@{
            Id          = $selectedGroup.Id
            DisplayName = $selectedGroup.DisplayName
            Description = $selectedGroup.Description
        }
    }

    function Get-AllGroupsLookup {
        param([string]$GraphVersion)

        $uri = "https://graph.microsoft.com/$GraphVersion/groups?`$select=id,displayName,description"
        $groups = Invoke-GraphGetPaged -Uri $uri

        $lookup = @{}
        foreach ($g in @($groups)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$g.id)) {
                $lookup[[string]$g.id] = [pscustomobject]@{
                    Id          = [string]$g.id
                    DisplayName = [string]$g.displayName
                    Description = [string]$g.description
                }
            }
        }

        return $lookup
    }

    function Get-AssignmentTargetDetails {
        param(
            [Parameter(Mandatory = $true)]
            $Assignment
        )

        $target = $Assignment.target
        $odataType = $null
        $groupId = $null
        $filterId = $null
        $filterType = $null
        $assignmentType = "Unknown"

        if ($target) {
            $odataTypeProp = $target.PSObject.Properties['@odata.type']
            if ($odataTypeProp) { $odataType = [string]$odataTypeProp.Value }

            $groupIdProp = $target.PSObject.Properties['groupId']
            if ($groupIdProp) { $groupId = [string]$groupIdProp.Value }

            $filterIdProp = $target.PSObject.Properties['deviceAndAppManagementAssignmentFilterId']
            if ($filterIdProp) { $filterId = [string]$filterIdProp.Value }

            $filterTypeProp = $target.PSObject.Properties['deviceAndAppManagementAssignmentFilterType']
            if ($filterTypeProp) { $filterType = [string]$filterTypeProp.Value }
        }

        if ($odataType) {
            switch -Regex ($odataType) {
                'exclusion' { $assignmentType = 'Exclude'; break }
                'allLicensedUsers' { $assignmentType = 'All Users'; break }
                'allDevices' { $assignmentType = 'All Devices'; break }
                'groupAssignmentTarget' { $assignmentType = 'Include'; break }
                default { $assignmentType = $odataType }
            }
        }

        [pscustomobject]@{
            AssignmentType = $assignmentType
            TargetODataType = $odataType
            GroupId = $groupId
            FilterId = $filterId
            FilterType = $filterType
        }
    }

    function Get-ItemName {
        param(
            [Parameter(Mandatory = $true)]
            $Item,
            [string[]]$PreferredProperties
        )

        foreach ($prop in $PreferredProperties) {
            $p = $Item.PSObject.Properties[$prop]
            if ($p -and -not [string]::IsNullOrWhiteSpace([string]$p.Value)) {
                return [string]$p.Value
            }
        }

        return "<Unnamed>"
    }

    function Get-AssignmentsFromEndpoint {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Uri,

            [Parameter(Mandatory = $true)]
            [string]$Workload,

            [Parameter(Mandatory = $true)]
            [string[]]$NameProperties,

            [Parameter(Mandatory = $false)]
            [string]$TargetGroupId,

            [Parameter(Mandatory = $false)]
            [hashtable]$GroupLookup,

            [Parameter(Mandatory = $false)]
            [switch]$TenantWide
        )

        $results = [System.Collections.Generic.List[object]]::new()

        try {
            $items = Invoke-GraphGetPaged -Uri $Uri
        }
        catch {
            Write-WarnMsg "Failed to query workload [$Workload]. $($_.Exception.Message)"
            return $results
        }

        foreach ($item in @($items)) {
            $itemId = if ($item.PSObject.Properties['id']) { [string]$item.id } else { $null }
            $itemName = Get-ItemName -Item $item -PreferredProperties $NameProperties

            if (-not $item.PSObject.Properties['assignments']) {
                continue
            }

            foreach ($assignment in @($item.assignments)) {
                $details = Get-AssignmentTargetDetails -Assignment $assignment

                if ([string]::IsNullOrWhiteSpace($details.GroupId)) {
                    continue
                }

                if (-not $TenantWide -and $details.GroupId -ne $TargetGroupId) {
                    continue
                }

                $intent = $null
                if ($assignment.PSObject.Properties['intent']) {
                    $intent = [string]$assignment.intent
                }

                $resolvedGroupName = $null
                if ($GroupLookup -and $GroupLookup.ContainsKey($details.GroupId)) {
                    $resolvedGroupName = $GroupLookup[$details.GroupId].DisplayName
                }

                [void]$results.Add([pscustomobject]@{
                    GroupId            = $details.GroupId
                    GroupName          = $resolvedGroupName
                    Workload           = $Workload
                    ItemName           = $itemName
                    ItemId             = $itemId
                    AssignmentType     = $details.AssignmentType
                    AssignmentIntent   = $intent
                    TargetODataType    = $details.TargetODataType
                    FilterId           = $details.FilterId
                    FilterType         = $details.FilterType
                    SourceUri          = $Uri
                })
            }
        }

        return $results
    }

    function Get-AllWorkloadDefinitions {
        param([string]$GraphVersion)

        $base = "https://graph.microsoft.com/$GraphVersion"

        return @(
            @{ Workload = "Mobile Apps";                         Uri = "$base/deviceAppManagement/mobileApps?`$expand=assignments"; NameProperties = @("displayName","name") }
            @{ Workload = "Device Configuration Profiles";      Uri = "$base/deviceManagement/deviceConfigurations?`$expand=assignments"; NameProperties = @("displayName","name") }
            @{ Workload = "Settings Catalog / Config Policies"; Uri = "$base/deviceManagement/configurationPolicies?`$expand=assignments"; NameProperties = @("name","displayName") }
            @{ Workload = "Compliance Policies";                Uri = "$base/deviceManagement/deviceCompliancePolicies?`$expand=assignments"; NameProperties = @("displayName","name") }
            @{ Workload = "Device Management Scripts";          Uri = "$base/deviceManagement/deviceManagementScripts?`$expand=assignments"; NameProperties = @("displayName","name") }
            @{ Workload = "Proactive Remediations";             Uri = "$base/deviceManagement/deviceHealthScripts?`$expand=assignments"; NameProperties = @("displayName","name") }
            @{ Workload = "Administrative Templates";           Uri = "$base/deviceManagement/groupPolicyConfigurations?`$expand=assignments"; NameProperties = @("displayName","name") }
            @{ Workload = "Update Rings";                       Uri = "$base/deviceManagement/deviceConfigurations?`$filter=isof('microsoft.graph.windowsUpdateForBusinessConfiguration')&`$expand=assignments"; NameProperties = @("displayName","name") }
            @{ Workload = "Feature Update Profiles";            Uri = "$base/deviceManagement/windowsFeatureUpdateProfiles?`$expand=assignments"; NameProperties = @("displayName","name") }
            @{ Workload = "Quality Update Profiles";            Uri = "$base/deviceManagement/windowsQualityUpdateProfiles?`$expand=assignments"; NameProperties = @("displayName","name") }
            @{ Workload = "Driver Update Profiles";             Uri = "$base/deviceManagement/windowsDriverUpdateProfiles?`$expand=assignments"; NameProperties = @("displayName","name") }
            @{ Workload = "Enrollment Configurations";          Uri = "$base/deviceManagement/deviceEnrollmentConfigurations?`$expand=assignments"; NameProperties = @("displayName","name") }
            @{ Workload = "Autopilot Deployment Profiles";      Uri = "$base/deviceManagement/windowsAutopilotDeploymentProfiles?`$expand=assignments"; NameProperties = @("displayName","name") }
            @{ Workload = "Endpoint Security Policies";         Uri = "$base/deviceManagement/intents?`$expand=assignments"; NameProperties = @("displayName","name") }

            # App protection / MAM
            @{ Workload = "App Protection - Android Managed Apps"; Uri = "$base/deviceAppManagement/androidManagedAppProtections?`$expand=assignments"; NameProperties = @("displayName","name") }
            @{ Workload = "App Protection - iOS Managed Apps";     Uri = "$base/deviceAppManagement/iosManagedAppProtections?`$expand=assignments"; NameProperties = @("displayName","name") }
            @{ Workload = "App Protection - Windows Managed Apps"; Uri = "$base/deviceAppManagement/windowsManagedAppProtections?`$expand=assignments"; NameProperties = @("displayName","name") }
            @{ Workload = "App Protection - MDM Windows Info Protection"; Uri = "$base/deviceAppManagement/mdmWindowsInformationProtectionPolicies?`$expand=assignments"; NameProperties = @("displayName","name") }
            @{ Workload = "App Protection - WIP Without Enrollment";      Uri = "$base/deviceAppManagement/windowsInformationProtectionPolicies?`$expand=assignments"; NameProperties = @("displayName","name") }

            # App configuration
            @{ Workload = "App Configuration - Managed Devices"; Uri = "$base/deviceAppManagement/mobileAppConfigurations?`$expand=assignments"; NameProperties = @("displayName","name") }
            @{ Workload = "App Configuration - Managed Apps";    Uri = "$base/deviceAppManagement/targetedManagedAppConfigurations?`$expand=assignments"; NameProperties = @("displayName","name") }

            # Platform-specific script workloads
            @{ Workload = "macOS Shell Scripts";                 Uri = "$base/deviceManagement/deviceShellScripts?`$expand=assignments"; NameProperties = @("displayName","name") }
            @{ Workload = "macOS Custom Attribute Shell Scripts"; Uri = "$base/deviceManagement/deviceCustomAttributeShellScripts?`$expand=assignments"; NameProperties = @("displayName","name") }

            # Policy sets
            @{ Workload = "Policy Sets";                         Uri = "$base/deviceAppManagement/policySets?`$expand=assignments"; NameProperties = @("displayName","name") }
        )
    }

    function Export-ResultsCsv {
        param(
            [Parameter(Mandatory = $true)]$Data,
            [Parameter(Mandatory = $true)][string]$PathBase
        )

        $csvPath = if ($PathBase -match '\.csv$') { $PathBase } else { "$PathBase.csv" }
        $dir = Split-Path -Path $csvPath -Parent
        if ($dir -and -not (Test-Path -Path $dir)) {
            New-Item -Path $dir -ItemType Directory -Force | Out-Null
        }

        $Data | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-Info "Exported CSV to [$csvPath]."
    }

    function Export-ResultsHtml {
        param(
            [Parameter(Mandatory = $true)]$Data,
            [Parameter(Mandatory = $true)][string]$PathBase,
            [Parameter(Mandatory = $true)][string]$Title
        )

        $htmlPath = if ($PathBase -match '\.html$') { $PathBase } else { "$PathBase.html" }
        $dir = Split-Path -Path $htmlPath -Parent
        if ($dir -and -not (Test-Path -Path $dir)) {
            New-Item -Path $dir -ItemType Directory -Force | Out-Null
        }

        $style = @"
<style>
body { font-family: Segoe UI, Arial, sans-serif; font-size: 12px; margin: 20px; }
h1, h2 { color: #1f4e79; }
table { border-collapse: collapse; width: 100%; }
th, td { border: 1px solid #d9d9d9; padding: 6px 8px; text-align: left; vertical-align: top; }
th { background-color: #f2f2f2; }
tr:nth-child(even) { background-color: #fafafa; }
.summary { margin-bottom: 20px; }
</style>
"@

        $summary = [pscustomobject]@{
            GeneratedOn = (Get-Date)
            TotalRows   = @($Data).Count
            TotalGroups = (@($Data | Where-Object { $_.GroupId } | Select-Object -ExpandProperty GroupId -Unique)).Count
            TotalWorkloads = (@($Data | Select-Object -ExpandProperty Workload -Unique)).Count
        }

        $html = @()
        $html += "<html><head><title>$Title</title>$style</head><body>"
        $html += "<h1>$Title</h1>"
        $html += "<div class='summary'>"
        $html += ($summary | ConvertTo-Html -Fragment)
        $html += "</div>"
        $html += "<h2>Assignments</h2>"
        $html += ($Data | ConvertTo-Html -Fragment)
        $html += "</body></html>"

        [System.IO.File]::WriteAllText($htmlPath, ($html -join [Environment]::NewLine), [System.Text.Encoding]::UTF8)
        Write-Info "Exported HTML to [$htmlPath]."
    }

    function Export-ResultsExcel {
        param(
            [Parameter(Mandatory = $true)]$Data,
            [Parameter(Mandatory = $true)][string]$PathBase,
            [Parameter(Mandatory = $true)]$SummaryData
        )

        if (-not (Test-CommandExists -Name 'Export-Excel')) {
            throw "Excel export requires the ImportExcel module. Install-Module ImportExcel"
        }

        $xlsxPath = if ($PathBase -match '\.xlsx$') { $PathBase } else { "$PathBase.xlsx" }
        $dir = Split-Path -Path $xlsxPath -Parent
        if ($dir -and -not (Test-Path -Path $dir)) {
            New-Item -Path $dir -ItemType Directory -Force | Out-Null
        }

        $Data | Export-Excel -Path $xlsxPath -WorksheetName 'Assignments' -TableName 'Assignments' -AutoSize -FreezeTopRow -ClearSheet
        $SummaryData | Export-Excel -Path $xlsxPath -WorksheetName 'Summary' -TableName 'Summary' -AutoSize -FreezeTopRow
        Write-Info "Exported Excel to [$xlsxPath]."
    }

    function Get-SummaryRows {
        param([Parameter(Mandatory = $true)]$Data)

        return $Data |
            Group-Object GroupName, Workload |
            Sort-Object Name |
            ForEach-Object {
                $first = $_.Group[0]
                [pscustomobject]@{
                    GroupName   = $first.GroupName
                    GroupId     = $first.GroupId
                    Workload    = $first.Workload
                    AssignmentCount = $_.Count
                }
            }
    }
}

process {
    Ensure-GraphConnection

    if ($AllGroups -and ($GroupId -or $GroupName)) {
        throw "Use either -AllGroups or a specific group selector, not both."
    }

    if (-not $AllGroups -and -not $GroupId -and -not $GroupName -and $SkipInteractivePicker) {
        throw "No GroupId or GroupName supplied, interactive picker disabled, and -AllGroups not selected."
    }

    $graphVersion = if ($UseBeta) { "beta" } else { "v1.0" }
    Write-Info "Using Microsoft Graph API version [$graphVersion]."

    $groupLookup = Get-AllGroupsLookup -GraphVersion $graphVersion
    Write-Info "Loaded [$($groupLookup.Count)] groups into lookup."

    $targetGroup = $null
    if (-not $AllGroups) {
        $targetGroup = Resolve-TargetGroup -InputGroupId $GroupId -InputGroupName $GroupName
        Write-Info "Resolved target group: [$($targetGroup.DisplayName)] [$($targetGroup.Id)]"
    }
    else {
        Write-Info "Tenant-wide mode enabled. Collecting assignments for all directly targeted groups."
    }

    $allResults = [System.Collections.Generic.List[object]]::new()
    $workloads = Get-AllWorkloadDefinitions -GraphVersion $graphVersion

    foreach ($workload in $workloads) {
        Write-Info "Processing workload [$($workload.Workload)]..."

        $matches = Get-AssignmentsFromEndpoint `
            -Uri $workload.Uri `
            -Workload $workload.Workload `
            -NameProperties $workload.NameProperties `
            -TargetGroupId $(if ($targetGroup) { $targetGroup.Id } else { $null }) `
            -GroupLookup $groupLookup `
            -TenantWide:$AllGroups

        foreach ($match in @($matches)) {
            [void]$allResults.Add($match)
        }
    }

    $dedupedResults = $allResults |
        Sort-Object GroupName, Workload, ItemName, ItemId, AssignmentType, AssignmentIntent |
        Select-Object GroupName, GroupId, Workload, ItemName, ItemId, AssignmentType, AssignmentIntent, TargetODataType, FilterId, FilterType, SourceUri -Unique

    if (-not $dedupedResults) {
        if ($AllGroups) {
            Write-Info "No direct group assignments found across queried workloads."
        }
        else {
            Write-Info "No direct assignments found for group [$($targetGroup.DisplayName)] [$($targetGroup.Id)]."
        }
        return
    }

    $summaryRows = Get-SummaryRows -Data $dedupedResults

    switch ($OutputFormat) {
        "Console" {
            $dedupedResults | Sort-Object GroupName, Workload, ItemName | Format-Table -AutoSize
        }
        "Csv" {
            Export-ResultsCsv -Data $dedupedResults -PathBase $ExportPath
        }
        "Html" {
            $title = if ($AllGroups) { "Intune Assignments - Tenant Wide" } else { "Intune Assignments - $($targetGroup.DisplayName)" }
            Export-ResultsHtml -Data $dedupedResults -PathBase $ExportPath -Title $title
        }
        "Excel" {
            Export-ResultsExcel -Data $dedupedResults -PathBase $ExportPath -SummaryData $summaryRows
        }
        "All" {
            Export-ResultsCsv -Data $dedupedResults -PathBase $ExportPath
            $title = if ($AllGroups) { "Intune Assignments - Tenant Wide" } else { "Intune Assignments - $($targetGroup.DisplayName)" }
            Export-ResultsHtml -Data $dedupedResults -PathBase $ExportPath -Title $title
            Export-ResultsExcel -Data $dedupedResults -PathBase $ExportPath -SummaryData $summaryRows
            $dedupedResults | Sort-Object GroupName, Workload, ItemName | Format-Table -AutoSize
        }
    }

    $dedupedResults
}