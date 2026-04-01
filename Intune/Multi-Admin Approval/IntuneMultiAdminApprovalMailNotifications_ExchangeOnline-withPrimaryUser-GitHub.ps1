#Requires -Version 5.1

$ErrorActionPreference = 'Stop'

# =========================================
# BOOTSTRAP
# =========================================
$MAAFolder = "C:\Temp\MAA"
$BootstrapFolder = $MAAFolder
$BootstrapLog = Join-Path $BootstrapFolder "BootstrapLog.txt"
$TranscriptPath = Join-Path $BootstrapFolder "Transcript.txt"

New-Item -Path $BootstrapFolder -ItemType Directory -Force | Out-Null
Add-Content -Path $BootstrapLog -Value "$(Get-Date -Format 'dd-MM-yyyy HH:mm:ss') [INFO] Script bootstrap started."

try {
    Start-Transcript -Path $TranscriptPath -Force | Out-Null
}
catch {
    Add-Content -Path $BootstrapLog -Value "$(Get-Date -Format 'dd-MM-yyyy HH:mm:ss') [WARN] Failed to start transcript: $($_.Exception.Message)"
}

try {
    # =========================================
    # User Input Section
    # =========================================
    $tenantId        = "c8b9f88a-afab-4e3c-a57d-0c1506b66757"
    $clientId        = "f588b292-5fad-44df-937b-a0f43b9a4de5"
    $clientSecret    = "REPLACE SECRET"

    # Email configuration
    $From            = "IntuneMultiAdminApprovers@mpc3.onmicrosoft.com"
    $ApprovalMembers = @(
        "JohannaL@mpc3.onmicrosoft.com",
        "AdeleV@mpc3.onmicrosoft.com",
        "DiegoS@mpc3.onmicrosoft.com"
    )

    # Audit event reporting duration
    $ReportingDuration = 4   # days

    # Limits & filters for mail communication
    $MaxPerStatus = @{
        needsapproval = 3
        approved      = 3
        completed     = 1
        rejected      = 1
        cancelled     = 1
    }

    # In case MailTracking.csv does not exist
    $RebuildHours     = 48   # hours
    $HoursBackToCheck = 48   # hours

    # =========================================
    # ENSURE DIRECTORIES EXIST
    # =========================================
    $StatusesToNotify    = @("needsapproval","approved","completed","rejected","cancelled")
    $ActionTypesToNotify = @("Retire","Wipe","Delete","Delete/Retire/Wipe")

    if ([string]::IsNullOrWhiteSpace($MAAFolder)) {
        $MAAFolder = "C:\Temp"
    }

    $CsvPath      = Join-Path $MAAFolder "AdministrativeTasks.csv"
    $TrackingFile = Join-Path $MAAFolder "MailTracking.csv"
    $LogFile      = Join-Path $MAAFolder "MailLog.txt"
    $ExportPath   = Join-Path $MAAFolder "IntuneApprovalAuditReport.csv"

    $paths = @($CsvPath, $TrackingFile, $LogFile, $ExportPath)

    foreach ($path in $paths) {
        $folder = Split-Path $path -Parent

        if (-not [string]::IsNullOrWhiteSpace($folder) -and -not (Test-Path $folder)) {
            New-Item -Path $folder -ItemType Directory -Force | Out-Null
        }

        if (-not (Test-Path $path)) {
            New-Item -Path $path -ItemType File -Force | Out-Null
        }
    }

    # =========================================
    # HELPER FUNCTIONS
    # =========================================
    function Write-RunLog {
        param(
            [Parameter(Mandatory)] [string]$Message,
            [ValidateSet("INFO","WARN","ERROR")] [string]$Level = "INFO"
        )

        $line = "{0} [{1}] {2}" -f (Get-Date -Format "dd-MM-yyyy HH:mm:ss"), $Level, $Message
        Add-Content -Path $LogFile -Value $line
        Write-Host $line
    }

    function Get-GraphAccessToken {
        param(
            [Parameter(Mandatory)] [string]$TenantId,
            [Parameter(Mandatory)] [string]$ClientId,
            [Parameter(Mandatory)] [string]$ClientSecret
        )

        $body = @{
            grant_type    = "client_credentials"
            scope         = "https://graph.microsoft.com/.default"
            client_id     = $ClientId
            client_secret = $ClientSecret
        }

        $tokenResponse = Invoke-RestMethod -Method POST `
            -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
            -Body $body `
            -ContentType "application/x-www-form-urlencoded"

        return $tokenResponse.access_token
    }

    function Get-AuthHeader {
        param(
            [Parameter(Mandatory)] [string]$AccessToken
        )

        return @{
            Authorization = "Bearer $AccessToken"
        }
    }

    function Get-StatusAction {
        param($displayName)
        switch ($displayName) {
            "Create ApprovalRequest"             { return "NeedsApproval" }
            "Request Approved ApprovalRequest"   { return "Approved" }
            "Approved"                           { return "Approved" }
            "Completed"                          { return "Completed" }
            "Request Rejected ApprovalRequest"   { return "Rejected" }
            "Request Cancelled ApprovalRequest"  { return "Cancelled" }
            default                              { return $displayName }
        }
    }

    function Get-ValidEmailAddresses {
        param(
            [string[]]$Addresses
        )

        $emailPattern = '^[^@\s]+@[^@\s]+\.[^@\s]+$'

        return @(
            $Addresses |
            ForEach-Object { [string]$_ } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ -match $emailPattern } |
            Sort-Object -Unique
        )
    }

    function ConvertTo-GraphRecipients {
        param(
            [Parameter(Mandatory)] [string[]]$Addresses
        )

        $clean = Get-ValidEmailAddresses -Addresses $Addresses

        $recipients = foreach ($addr in $clean) {
            [PSCustomObject]@{
                emailAddress = [PSCustomObject]@{
                    address = [string]$addr
                }
            }
        }

        return @($recipients)
    }

    function Send-GraphMail {
        param(
            [Parameter(Mandatory)] [string]$AccessToken,
            [Parameter(Mandatory)] [string]$FromUser,
            [Parameter(Mandatory)] [string[]]$To,
            [string[]]$Cc,
            [Parameter(Mandatory)] [string]$Subject,
            [Parameter(Mandatory)] [string]$BodyHtml
        )

        $toRecipients = ConvertTo-GraphRecipients -Addresses $To
        $ccRecipients = @()

        if ($Cc) {
            $ccRecipients = ConvertTo-GraphRecipients -Addresses $Cc
        }

        if ($toRecipients.Count -eq 0) {
            throw "No valid To recipients were supplied."
        }

        $payload = @{
            message = @{
                subject = [string]$Subject
                body = @{
                    contentType = "HTML"
                    content     = [string]$BodyHtml
                }
                toRecipients = @($toRecipients)
            }
            saveToSentItems = $true
        }

        if ($ccRecipients.Count -gt 0) {
            $payload.message.ccRecipients = @($ccRecipients)
        }

        $jsonBody = $payload | ConvertTo-Json -Depth 10 -Compress
        Write-RunLog -Message "Graph payload for From=$FromUser : $jsonBody"

        $mailUri = "https://graph.microsoft.com/v1.0/users/$([uri]::EscapeDataString($FromUser))/sendMail"

        try {
            Invoke-RestMethod -Method POST `
                -Uri $mailUri `
                -Headers (Get-AuthHeader -AccessToken $AccessToken) `
                -Body $jsonBody `
                -ContentType "application/json" | Out-Null
        }
        catch {
            $responseText = $null

            if ($_.Exception.Response) {
                try {
                    $stream = $_.Exception.Response.GetResponseStream()
                    if ($stream) {
                        $reader = New-Object System.IO.StreamReader($stream)
                        $responseText = $reader.ReadToEnd()
                        $reader.Close()
                    }
                }
                catch {}
            }

            if ($responseText) {
                throw "Graph sendMail failed. Response body: $responseText"
            }
            else {
                throw "Graph sendMail failed. $($_.Exception.Message)"
            }
        }
    }

    function Get-EmptyManagedDeviceResult {
        return [PSCustomObject]@{
            Found        = $false
            DeviceName   = $null
            Platform     = $null
            OSVersion    = $null
            Manufacturer = $null
            Model        = $null
            SerialNumber = $null
        }
    }

    function Get-ModifiedPropertyValue {
        param(
            [Parameter(Mandatory)] $ModifiedProperties,
            [Parameter(Mandatory)] [string]$PropertyName
        )

        $match = @($ModifiedProperties) | Where-Object { $_.displayName -eq $PropertyName } | Select-Object -First 1
        if ($match) {
            return [string]$match.newValue
        }

        return $null
    }

    function Get-ManagedDeviceDetailsById {
        param(
            [Parameter(Mandatory)] [string]$AccessToken,
            [Parameter(Mandatory)] [string]$ManagedDeviceId
        )

        $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$([uri]::EscapeDataString($ManagedDeviceId))"

        try {
            $response = Invoke-RestMethod -Method GET `
                -Uri $uri `
                -Headers (Get-AuthHeader -AccessToken $AccessToken)

            return [PSCustomObject]@{
                Found        = $true
                DeviceName   = $response.deviceName
                Platform     = $response.operatingSystem
                OSVersion    = $response.osVersion
                Manufacturer = $response.manufacturer
                Model        = $response.model
                SerialNumber = $response.serialNumber
            }
        }
        catch {
            $statusCode = $null
            try {
                if ($_.Exception.Response -and $null -ne $_.Exception.Response.StatusCode) {
                    $statusCode = [int]$_.Exception.Response.StatusCode
                }
            }
            catch {}

            if ($statusCode -eq 404) {
                Write-RunLog -Level INFO -Message ("Managed device ID not found for payload ID [{0}]. Will try fallback lookup." -f $ManagedDeviceId)
            }
            else {
                Write-RunLog -Level WARN -Message ("Managed device lookup by ID failed for [{0}]. {1}" -f $ManagedDeviceId, $_.Exception.Message)
            }

            return Get-EmptyManagedDeviceResult
        }
    }

    function Get-ManagedDevicePrimaryUser {
        param(
            [Parameter(Mandatory)] [string]$AccessToken,
            [Parameter(Mandatory)] [string]$ManagedDeviceId
        )

        $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$([uri]::EscapeDataString($ManagedDeviceId))/users?`$select=id,displayName,userPrincipalName"

        try {
            $response = Invoke-RestMethod -Method GET `
                -Uri $uri `
                -Headers (Get-AuthHeader -AccessToken $AccessToken)

            $user = @($response.value) | Select-Object -First 1
            if ($user) {
                Write-RunLog -Message ("Primary user matched for managed device [{0}] -> [{1}] / [{2}]" -f `
                    $ManagedDeviceId, $user.displayName, $user.userPrincipalName)

                return [PSCustomObject]@{
                    Found             = $true
                    PrimaryUserName   = [string]$user.displayName
                    PrimaryUserUPN    = [string]$user.userPrincipalName
                    PrimaryUserId     = [string]$user.id
                }
            }

            Write-RunLog -Level INFO -Message ("No primary user returned for managed device [{0}]." -f $ManagedDeviceId)
            return [PSCustomObject]@{
                Found             = $false
                PrimaryUserName   = $null
                PrimaryUserUPN    = $null
                PrimaryUserId     = $null
            }
        }
        catch {
            Write-RunLog -Level WARN -Message ("Primary user lookup failed for managed device [{0}]. {1}" -f `
                $ManagedDeviceId, $_.Exception.Message)

            return [PSCustomObject]@{
                Found             = $false
                PrimaryUserName   = $null
                PrimaryUserUPN    = $null
                PrimaryUserId     = $null
            }
        }
    }

    function Find-ManagedDeviceDetails {
        param(
            [Parameter(Mandatory)] [string]$AccessToken,
            [string]$ManagedDeviceId,
            [string]$SerialNumber,
            [string]$DeviceName
        )

        Write-RunLog -Message ("Device lookup input -> ManagedDeviceId=[{0}] SerialNumber=[{1}] DeviceName=[{2}]" -f `
            $ManagedDeviceId, $SerialNumber, $DeviceName)

        if (-not [string]::IsNullOrWhiteSpace($ManagedDeviceId)) {
            $byId = Get-ManagedDeviceDetailsById -AccessToken $AccessToken -ManagedDeviceId $ManagedDeviceId
            if ($byId.Found) {
                Write-RunLog -Message ("Managed device matched by ID [{0}] -> Name=[{1}] Platform=[{2}]" -f `
                    $ManagedDeviceId, $byId.DeviceName, $byId.Platform)
                return $byId
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($SerialNumber)) {
            $escapedSerial = $SerialNumber.Replace("'", "''")
            $filter = [uri]::EscapeDataString("serialNumber eq '$escapedSerial'")
            $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=$filter&`$top=5"

            try {
                $response = Invoke-RestMethod -Method GET `
                    -Uri $uri `
                    -Headers (Get-AuthHeader -AccessToken $AccessToken)

                $matches = @($response.value)
                Write-RunLog -Message ("Serial lookup returned {0} result(s) for [{1}]." -f $matches.Count, $SerialNumber)

                if ($matches.Count -ge 1) {
                    $match = $matches | Select-Object -First 1
                    Write-RunLog -Message ("Managed device matched by serial number [{0}] -> Name=[{1}] Platform=[{2}]" -f `
                        $SerialNumber, $match.deviceName, $match.operatingSystem)

                    return [PSCustomObject]@{
                        Found        = $true
                        DeviceName   = $match.deviceName
                        Platform     = $match.operatingSystem
                        OSVersion    = $match.osVersion
                        Manufacturer = $match.manufacturer
                        Model        = $match.model
                        SerialNumber = $match.serialNumber
                    }
                }
            }
            catch {
                Write-RunLog -Level WARN -Message ("Managed device lookup by serial number failed for [{0}]. {1}" -f $SerialNumber, $_.Exception.Message)
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($DeviceName)) {
            $escapedName = $DeviceName.Replace("'", "''")
            $filter = [uri]::EscapeDataString("deviceName eq '$escapedName'")
            $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=$filter&`$top=10"

            try {
                $response = Invoke-RestMethod -Method GET `
                    -Uri $uri `
                    -Headers (Get-AuthHeader -AccessToken $AccessToken)

                $matches = @($response.value)
                Write-RunLog -Message ("Device name exact lookup returned {0} result(s) for [{1}]." -f $matches.Count, $DeviceName)

                if ($matches.Count -ge 1) {
                    $match = $matches | Select-Object -First 1
                    Write-RunLog -Message ("Managed device matched by exact device name [{0}] -> Platform=[{1}]" -f `
                        $DeviceName, $match.operatingSystem)

                    return [PSCustomObject]@{
                        Found        = $true
                        DeviceName   = $match.deviceName
                        Platform     = $match.operatingSystem
                        OSVersion    = $match.osVersion
                        Manufacturer = $match.manufacturer
                        Model        = $match.model
                        SerialNumber = $match.serialNumber
                    }
                }
            }
            catch {
                Write-RunLog -Level WARN -Message ("Managed device lookup by exact device name failed for [{0}]. {1}" -f $DeviceName, $_.Exception.Message)
            }

            try {
                $escapedSearch = [uri]::EscapeDataString("`"$DeviceName`"")
                $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$search=$escapedSearch&`$top=25"

                $response = Invoke-RestMethod -Method GET `
                    -Uri $uri `
                    -Headers @{
                        Authorization    = "Bearer $AccessToken"
                        ConsistencyLevel = "eventual"
                    }

                $matches = @($response.value)
                Write-RunLog -Message ("Device name search returned {0} result(s) for [{1}]." -f $matches.Count, $DeviceName)

                $match = $matches |
                    Where-Object { $_.deviceName -eq $DeviceName } |
                    Select-Object -First 1

                if (-not $match) {
                    $match = $matches | Select-Object -First 1
                }

                if ($match) {
                    Write-RunLog -Message ("Managed device matched by broader search [{0}] -> Name=[{1}] Platform=[{2}]" -f `
                        $DeviceName, $match.deviceName, $match.operatingSystem)

                    return [PSCustomObject]@{
                        Found        = $true
                        DeviceName   = $match.deviceName
                        Platform     = $match.operatingSystem
                        OSVersion    = $match.osVersion
                        Manufacturer = $match.manufacturer
                        Model        = $match.model
                        SerialNumber = $match.serialNumber
                    }
                }
            }
            catch {
                Write-RunLog -Level WARN -Message ("Managed device lookup by broader device name search failed for [{0}]. {1}" -f $DeviceName, $_.Exception.Message)
            }
        }

        Write-RunLog -Level INFO -Message "No managed device match found by payload ID, serial number, or device name. Using audit event values."
        return Get-EmptyManagedDeviceResult
    }

    # =========================================
    # START LOGGING
    # =========================================
    Write-RunLog -Message "Script main execution started."

    # =========================================
    # GET GRAPH API TOKEN
    # =========================================
    $accessToken = Get-GraphAccessToken -TenantId $tenantId -ClientId $clientId -ClientSecret $clientSecret
    Write-RunLog -Message "Successfully obtained Microsoft Graph access token."

    # =========================================
    # FETCH AUDIT EVENTS
    # =========================================
    $startDate = (Get-Date).AddDays(-$ReportingDuration).ToString("o")
    $rawFilter = "activityDateTime ge $startDate and componentName eq 'ApprovalRequest'"
    $filter    = [uri]::EscapeDataString($rawFilter)
    $uri       = "https://graph.microsoft.com/beta/deviceManagement/auditEvents?`$filter=$filter&`$orderby=activityDateTime desc"

    Write-RunLog -Message "Fetching audit events using URI: $uri"

    $allResults = @()
    while ($uri) {
        $response = Invoke-RestMethod -Method GET -Uri $uri -Headers (Get-AuthHeader -AccessToken $accessToken)
        if ($response.value) {
            $allResults += $response.value
        }
        $uri = $response.'@odata.nextLink'
    }

    Write-RunLog -Message "Fetched audit event count: $(@($allResults).Count)"

    # =========================================
    # PROCESS AUDIT EVENTS INTO EXPORT DATA
    # =========================================
    $exportData = @(
        $allResults |
        Where-Object { $_.componentName -eq "ApprovalRequest" } |
        Group-Object { $_.resources[0].resourceId } |
        ForEach-Object {
            $events = $_.Group | Sort-Object activityDateTime
            $createEvent = $events | Where-Object { $_.displayName -eq "Create ApprovalRequest" } | Select-Object -First 1
            if (-not $createEvent) { $createEvent = $events | Select-Object -First 1 }

            if (-not $createEvent) {
                Write-RunLog -Level WARN -Message "Skipping grouped item with no createEvent."
                return
            }

            $approveEvent = $events |
                Where-Object {
                    $_.activityDateTime -gt $createEvent.activityDateTime -and
                    $_.displayName -ne "Create ApprovalRequest"
                } |
                Select-Object -First 1

            $latestEvent = $events | Select-Object -Last 1
            if (-not $latestEvent) {
                Write-RunLog -Level WARN -Message "Skipping grouped item with no latestEvent."
                return
            }

            $status      = Get-StatusAction -displayName $latestEvent.displayName
            $requestedBy = if ($createEvent.actor.userPrincipalName) { [string]$createEvent.actor.userPrincipalName } else { "" }
            $approvedBy  = if ($approveEvent -and $approveEvent.actor.userPrincipalName) { [string]$approveEvent.actor.userPrincipalName } else { "" }

            $approvalResourceId = [string]$latestEvent.resources[0].resourceId
            $modifiedProps      = @($latestEvent.resources[0].modifiedProperties)

            $payloadId         = Get-ModifiedPropertyValue -ModifiedProperties $modifiedProps -PropertyName "PayloadId"
            $payloadType       = Get-ModifiedPropertyValue -ModifiedProperties $modifiedProps -PropertyName "PayloadType"
            $payloadOperation  = Get-ModifiedPropertyValue -ModifiedProperties $modifiedProps -PropertyName "PayloadOperation"
            $approverUpn       = Get-ModifiedPropertyValue -ModifiedProperties $modifiedProps -PropertyName "ApproverUpn"

            if (-not [string]::IsNullOrWhiteSpace($approverUpn)) {
                $approvedBy = $approverUpn
            }

            $auditDeviceName   = [string]$latestEvent.resources[0].displayName
            $auditSerialNumber = [string]$latestEvent.resources[0].serialNumber

            Write-RunLog -Message ("Audit event values -> ApprovalResourceId=[{0}] PayloadId=[{1}] PayloadType=[{2}] PayloadOperation=[{3}] DisplayName=[{4}] SerialNumber=[{5}]" -f `
                $approvalResourceId, $payloadId, $payloadType, $payloadOperation, $auditDeviceName, $auditSerialNumber)

            $managedDeviceIdToUse = $null
            if ($payloadType -eq "Microsoft.Management.Services.Api.ManagedDevice" -and -not [string]::IsNullOrWhiteSpace($payloadId)) {
                $managedDeviceIdToUse = $payloadId
            }

            $deviceDetails = Find-ManagedDeviceDetails `
                -AccessToken $accessToken `
                -ManagedDeviceId $managedDeviceIdToUse `
                -SerialNumber $auditSerialNumber `
                -DeviceName $auditDeviceName

            $primaryUser = [PSCustomObject]@{
                Found           = $false
                PrimaryUserName = $null
                PrimaryUserUPN  = $null
                PrimaryUserId   = $null
            }

            if (-not [string]::IsNullOrWhiteSpace($managedDeviceIdToUse)) {
                $primaryUser = Get-ManagedDevicePrimaryUser `
                    -AccessToken $accessToken `
                    -ManagedDeviceId $managedDeviceIdToUse
            }

            [PSCustomObject]@{
                Id                    = $approvalResourceId
                ManagedDeviceId       = $managedDeviceIdToUse
                PayloadType           = $payloadType
                PayloadOperation      = $payloadOperation
                Status                = $status
                Action                = "Delete/Retire/Wipe"
                RequestedBy           = $requestedBy
                ApprovedBy            = $approvedBy
                createdDateTime       = [datetime]$createEvent.activityDateTime
                dueDateTime           = $null
                DeviceName            = if ($deviceDetails.DeviceName) { $deviceDetails.DeviceName } else { $auditDeviceName }
                Platform              = if ($deviceDetails.Platform) { $deviceDetails.Platform } else { $null }
                OSVersion             = $deviceDetails.OSVersion
                Manufacturer          = $deviceDetails.Manufacturer
                Model                 = $deviceDetails.Model
                SerialNumber          = if ($deviceDetails.SerialNumber) { $deviceDetails.SerialNumber } else { $auditSerialNumber }
                PrimaryUserName       = $primaryUser.PrimaryUserName
                PrimaryUserUPN        = $primaryUser.PrimaryUserUPN
                RequestJustification  = [string]$latestEvent.resources[0].requestJustification
                ApprovalJustification = [string]$latestEvent.resources[0].approvalJustification
            }
        }
    )

    $finalData = @($exportData)
    $finalData | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
    $finalData | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-RunLog -Message "Final data prepared for email processing. Record count: $(@($finalData).Count)"

    # =========================================
    # LOAD OR REBUILD TRACKING
    # =========================================
    $TrackingData = @()
    $IsRebuild = $false

    if ((Test-Path $TrackingFile) -and ((Get-Item $TrackingFile).Length -gt 0)) {
        $TrackingData = @(Import-Csv $TrackingFile)
        Write-RunLog -Message "Loaded existing tracking file. Record count: $(@($TrackingData).Count)"
    }
    else {
        Write-RunLog -Level WARN -Message "Tracking file missing or empty - rebuilding..."
        $IsRebuild = $true
        $Cutoff = (Get-Date).AddHours(-$RebuildHours)

        foreach ($item in $finalData | Where-Object {
            $_.Action -in $ActionTypesToNotify -and
            $_.Status.ToLower() -in $StatusesToNotify -and
            $_.createdDateTime -ge $Cutoff
        }) {
            $TrackingData += [PSCustomObject]@{
                TaskID             = $item.Id
                RequestedBy        = $item.RequestedBy
                ApprovedBy         = $item.ApprovedBy
                Action             = $item.Action
                DeviceName         = $item.DeviceName
                LastStatus         = $item.Status.ToLower()
                NeedsApprovalCount = 0
                ApprovedCount      = 0
                CompletedCount     = 0
                RejectedCount      = 0
                CancelledCount     = 0
                TotalMailCount     = 0
            }
        }

        Write-RunLog -Message "Tracking rebuild completed. Record count: $(@($TrackingData).Count)"
    }

    # =========================================
    # FILTER DATA FOR EMAIL
    # =========================================
    $CutoffTime = if ($IsRebuild) {
        (Get-Date).AddHours(-$RebuildHours)
    }
    else {
        (Get-Date).AddHours(-$HoursBackToCheck)
    }

    $Data = @(
        $finalData | Where-Object {
            $_.Action -in $ActionTypesToNotify -and
            $_.Status.ToLower() -in $StatusesToNotify -and
            $_.createdDateTime -ge $CutoffTime
        }
    )

    Write-RunLog -Message "Total actionable requests: $(@($Data).Count)"

    # =========================================
    # PROCESS EMAILS
    # =========================================
    foreach ($item in $Data) {

        $TaskID = $item.Id
        $Status = $item.Status.ToLower()

        $record = $TrackingData | Where-Object { $_.TaskID -eq $TaskID } | Select-Object -First 1

        if (-not $record) {
            $record = [PSCustomObject]@{
                TaskID             = $TaskID
                RequestedBy        = $item.RequestedBy
                ApprovedBy         = $item.ApprovedBy
                Action             = $item.Action
                DeviceName         = $item.DeviceName
                LastStatus         = $Status
                NeedsApprovalCount = 0
                ApprovedCount      = 0
                CompletedCount     = 0
                RejectedCount      = 0
                CancelledCount     = 0
                TotalMailCount     = 0
            }
        }

        $currentCountProperty = switch ($Status) {
            "needsapproval" { "NeedsApprovalCount" }
            "approved"      { "ApprovedCount" }
            "completed"     { "CompletedCount" }
            "rejected"      { "RejectedCount" }
            "cancelled"     { "CancelledCount" }
            default         { $null }
        }

        if (-not $currentCountProperty) {
            Write-RunLog -Level WARN -Message "Skipping Task ID $TaskID because status '$Status' is not actionable."
            continue
        }

        $limit = $MaxPerStatus[$Status]
        if ([int]$record.$currentCountProperty -ge $limit) {
            Write-RunLog -Level WARN -Message "Mail limit reached for Task ID: $TaskID ($Status)"
            continue
        }

        $record.$currentCountProperty = [int]$record.$currentCountProperty + 1
        $record.TotalMailCount        = [int]$record.TotalMailCount + 1
        $record.LastStatus            = $Status

        $GreetingName = if (-not [string]::IsNullOrWhiteSpace($item.RequestedBy)) { $item.RequestedBy } else { "User" }
        $Greeting = "Hi $GreetingName,"

        $Message = switch ($Status) {
            "needsapproval" {
                "Your administrative request has been submitted and is awaiting approval. If it is urgent, please contact any approver listed in CC to expedite the process."
            }
            "approved" {
                "Your administrative request has been approved by: $($item.ApprovedBy). Please log in to the Intune portal to complete the necessary actions and move your request to Completed."
            }
            "completed" {
                "Your administrative request has been completed by: $($item.ApprovedBy). No further action is needed."
            }
            "rejected" {
                "Your administrative request has been rejected by: $($item.ApprovedBy). If required, review the approval comment and submit a new request with proper justification and the relevant change, incident, or task number."
            }
            "cancelled" {
                "Your administrative request has been cancelled by: $($item.ApprovedBy). If required, review the approval comment and submit a new request with proper justification and the relevant change, incident, or task number."
            }
            default {
                "Your administrative request status is $Status."
            }
        }

        $CreatedDate      = (Get-Date $item.createdDateTime).ToString("dd-MM-yyyy HH:mm")
        $Platform         = if ([string]::IsNullOrWhiteSpace($item.Platform)) { "Not available from audit/device lookup" } else { $item.Platform }
        $DeviceName       = if ([string]::IsNullOrWhiteSpace($item.DeviceName)) { "Not available from audit/device lookup" } else { $item.DeviceName }
        $OSVersion        = if ([string]::IsNullOrWhiteSpace($item.OSVersion)) { "Not available from audit/device lookup" } else { $item.OSVersion }
        $Manufacturer     = if ([string]::IsNullOrWhiteSpace($item.Manufacturer)) { "Not available from audit/device lookup" } else { $item.Manufacturer }
        $Model            = if ([string]::IsNullOrWhiteSpace($item.Model)) { "Not available from audit/device lookup" } else { $item.Model }
        $SerialNumber     = if ([string]::IsNullOrWhiteSpace($item.SerialNumber)) { "Not available from audit/device lookup" } else { $item.SerialNumber }
        $PrimaryUserName  = if ([string]::IsNullOrWhiteSpace($item.PrimaryUserName)) { "Not available from Intune" } else { $item.PrimaryUserName }
        $PrimaryUserUPN   = if ([string]::IsNullOrWhiteSpace($item.PrimaryUserUPN)) { "Not available from Intune" } else { $item.PrimaryUserUPN }

        $BodyM = @"
<html>
<body style='font-family:Segoe UI; font-size:12px;'>
<p>$Greeting</p>
<p>$Message</p>
<table border='1' cellpadding='6' cellspacing='0'>
<tr><td><b>Device action</b></td><td>$($item.Action)</td></tr>
<tr><td><b>Device name</b></td><td>$DeviceName</td></tr>
<tr><td><b>Platform</b></td><td>$Platform</td></tr>
<tr><td><b>OS version</b></td><td>$OSVersion</td></tr>
<tr><td><b>Manufacturer</b></td><td>$Manufacturer</td></tr>
<tr><td><b>Model</b></td><td>$Model</td></tr>
<tr><td><b>Serial number</b></td><td>$SerialNumber</td></tr>
<tr><td><b>Primary user</b></td><td>$PrimaryUserName</td></tr>
<tr><td><b>Primary user UPN</b></td><td>$PrimaryUserUPN</td></tr>
<tr><td><b>Requested by</b></td><td>$($item.RequestedBy)</td></tr>
<tr><td><b>Approved by</b></td><td>$($item.ApprovedBy)</td></tr>
<tr><td><b>Requested on</b></td><td>$CreatedDate</td></tr>
<tr><td><b>Status</b></td><td><b>$Status</b></td></tr>
</table>
<br>
<p>Regards,<br>Intune Automation Team</p>
</body>
</html>
"@

        $Prefix = "[Intune]"
        $Subject = switch ($Status) {
            "needsapproval" { "$Prefix Action Required: Approval Pending - $($item.Action) - $DeviceName ($Platform)" }
            "approved"      { "$Prefix Action Required: Approval Granted - $($item.Action) - $DeviceName ($Platform)" }
            "completed"     { "$Prefix Completed Successfully - $($item.Action) - $DeviceName ($Platform)" }
            "rejected"      { "$Prefix Rejected - $($item.Action) - $DeviceName ($Platform)" }
            "cancelled"     { "$Prefix Cancelled - $($item.Action) - $DeviceName ($Platform)" }
            default         { "$Prefix Status Update - $($item.Action) - $DeviceName ($Platform)" }
        }

        $rawRequestedBy = [string]$item.RequestedBy
        Write-RunLog -Message ("RequestedBy raw value for Task ID {0}: [{1}] Length={2}" -f `
            $TaskID,
            $rawRequestedBy,
            $rawRequestedBy.Length)

        $toRecipients = Get-ValidEmailAddresses -Addresses @($rawRequestedBy)
        $ccRecipients = Get-ValidEmailAddresses -Addresses @($ApprovalMembers) |
            Where-Object { $_ -notin $toRecipients }

        if ($toRecipients.Count -eq 0) {
            Write-RunLog -Level WARN -Message ("RequestedBy invalid or blank for Task ID {0}: '{1}'. Redirecting mail to approvers only." -f $TaskID, $rawRequestedBy)
            $toRecipients = Get-ValidEmailAddresses -Addresses @($ApprovalMembers)
            $ccRecipients = @()
        }

        Write-RunLog -Message ("Preparing mail for Task ID {0}. From={1}; To={2}; Cc={3}" -f `
            $TaskID,
            $From,
            ($toRecipients -join ';'),
            ($ccRecipients -join ';'))

        try {
            Send-GraphMail `
                -AccessToken $accessToken `
                -FromUser $From `
                -To $toRecipients `
                -Cc $ccRecipients `
                -Subject $Subject `
                -BodyHtml $BodyM

            $TrackingData = @($TrackingData | Where-Object { $_.TaskID -ne $TaskID })
            $TrackingData += $record

            Write-RunLog -Message "Mail sent for Task ID: $TaskID ($Status)"
        }
        catch {
            Write-RunLog -Level ERROR -Message "Failed to send mail for Task ID: $TaskID ($Status). $($_.Exception.Message)"
            continue
        }
    }

    # =========================================
    # SAVE TRACKING
    # =========================================
    $TrackingData | Export-Csv -Path $TrackingFile -NoTypeInformation -Encoding UTF8
    Write-RunLog -Message "Script execution completed successfully."
}
catch {
    $fatal = "FATAL: $($_.Exception.Message)"
    Add-Content -Path $BootstrapLog -Value "$(Get-Date -Format 'dd-MM-yyyy HH:mm:ss') [ERROR] $fatal"
    Write-Host $fatal -ForegroundColor Red

    if ($_.ScriptStackTrace) {
        Add-Content -Path $BootstrapLog -Value "$(Get-Date -Format 'dd-MM-yyyy HH:mm:ss') [ERROR] Stack: $($_.ScriptStackTrace)"
        Write-Host $_.ScriptStackTrace -ForegroundColor Red
    }

    throw
}
finally {
    try {
        Stop-Transcript | Out-Null
    }
    catch {}
}