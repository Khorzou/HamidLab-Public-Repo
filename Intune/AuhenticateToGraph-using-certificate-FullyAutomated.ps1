<#
.SYNOPSIS
    Authenticates to Microsoft Graph using certificate-based authentication and retrieves Entra device information.

.DESCRIPTION
    This script implements a production-ready workflow for Microsoft Graph app-only authentication using a self-signed certificate.

    It performs the following actions:
    - Creates or reuses a self-signed certificate in the LocalMachine store
    - Exports the certificate (CER and optionally PFX)
    - Optionally pauses for manual upload of the CER to the target Entra app registration
    - Uses the certificate for app-only authentication via MSAL
    - Queries Microsoft Graph (v1.0 by default) for device information
    - Handles Graph pagination safely
    - Logs all actions and errors to a log file

    This approach avoids storing client secrets and aligns with recommended security practices for service-to-service authentication.

.PREREQUISITES
    - Run PowerShell as Administrator
    - MSAL.PS module installed (for Get-MsalToken)
    - App registration created in Entra ID
    - App registration has required Microsoft Graph Application permissions (e.g. Device.Read.All)
    - Admin consent granted for those permissions

    IMPORTANT:
    The certificate must be uploaded to the app registration before certificate-based authentication will work.

.PARAMETER ApplicationId
    The Client ID (App ID) of the Entra app registration.

.PARAMETER TenantId
    The Azure AD / Entra tenant ID.

.PARAMETER CreateNewCertificate
    Forces creation of a new self-signed certificate instead of reusing an existing one.

.PARAMETER ExportCer
    Exports the public certificate (.cer) for upload to the app registration.

.PARAMETER ExportPfx
    Exports the certificate including private key (.pfx) with password protection.

.PARAMETER WaitForManualCerUploadConfirmation
    Pauses execution and prompts the user to confirm the CER has been uploaded to the app registration.

.PARAMETER SkipGraphQuery
    Skips the Microsoft Graph query step (useful for setup/validation only).

.PARAMETER SkipCertificateReuse
    Forces creation of a new certificate even if a valid one already exists.

.EXAMPLES

    First-time setup (recommended):

    This will:
    - Create a new certificate
    - Export CER and PFX
    - Pause for manual upload to app registration
    - Then authenticate to Graph

    .\Script.ps1 `
        -ApplicationId "<APP_ID>" `
        -TenantId "<TENANT_ID>" `
        -CreateNewCertificate `
        -ExportCer `
        -ExportPfx `
        -WaitForManualCerUploadConfirmation


    Steps during first run:
    1. Script creates and exports the certificate
    2. Upload the CER file to:
       Entra Admin Center → App registrations → Your App → Certificates & secrets → Certificates → Upload certificate
    3. Return to the script and type YES to continue
    4. Script proceeds with Graph authentication


    Subsequent runs:

    This will:
    - Reuse existing certificate
    - Authenticate directly to Graph
    - No manual steps required

    .\Script.ps1 `
        -ApplicationId "<APP_ID>" `
        -TenantId "<TENANT_ID>"


.NOTES
    - Certificates are stored in: Cert:\LocalMachine\My
    - Log file location: C:\Temp\Certs\AzureAuthCert.log
    - Default certificate validity: 24 months
    - Graph endpoint: https://graph.microsoft.com/v1.0/devices

    Recommended future improvements:
    - Certificate rotation before expiry
    - Cleanup of expired certificates
    - Validation of app permissions before execution

#>

#requires -RunAsAdministrator

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [string]$SubjectName = "AzureAuthCert",

    [Parameter(Mandatory = $false)]
    [string]$CertStoreLocation = "Cert:\LocalMachine\My",

    [Parameter(Mandatory = $true)]
    [string]$ApplicationId,

    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [string]$ExportPath = "C:\Temp\Certs",

    [Parameter(Mandatory = $false)]
    [int]$ValidityMonths = 24,

    [Parameter(Mandatory = $false)]
    [ValidateRange(2048, 8192)]
    [int]$KeyLength = 2048,

    [Parameter(Mandatory = $false)]
    [ValidateSet("v1.0", "beta")]
    [string]$GraphApiVersion = "v1.0",

    [Parameter(Mandatory = $false)]
    [switch]$CreateNewCertificate,

    [Parameter(Mandatory = $false)]
    [switch]$ExportCer,

    [Parameter(Mandatory = $false)]
    [switch]$ExportPfx,

    [Parameter(Mandatory = $false)]
    [switch]$SkipGraphQuery,

    [Parameter(Mandatory = $false)]
    [switch]$DisablePfxExportablePrivateKey,

    [Parameter(Mandatory = $false)]
    [switch]$WaitForManualCerUploadConfirmation,

    [Parameter(Mandatory = $false)]
    [switch]$SkipCertificateReuse,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = "C:\Temp\Certs\AzureAuthCert.log"
)

begin {
    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    function Write-Log {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Message,

            [Parameter(Mandatory = $false)]
            [ValidateSet("INFO", "WARN", "ERROR", "DEBUG")]
            [string]$Level = "INFO"
        )

        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $line = "[$timestamp] [$Level] $Message"

        switch ($Level) {
            "ERROR" { Write-Host $line -ForegroundColor Red }
            "WARN"  { Write-Warning $Message }
            default { Write-Host $line }
        }

        try {
            $logDir = Split-Path -Path $LogPath -Parent
            if ($logDir -and -not (Test-Path -Path $logDir)) {
                New-Item -Path $logDir -ItemType Directory -Force | Out-Null
            }
            Add-Content -Path $LogPath -Value $line
        }
        catch {
            Write-Warning "Failed to write to log file [$LogPath]. $($_.Exception.Message)"
        }
    }

    function Test-IsAdministrator {
        $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = [Security.Principal.WindowsPrincipal]::new($currentIdentity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }

    function Assert-Prerequisites {
        Write-Log "Validating prerequisites..."

        if (-not (Test-IsAdministrator)) {
            throw "This script must be run in an elevated PowerShell session as Administrator."
        }

        foreach ($cmd in @(
            "New-SelfSignedCertificate",
            "Export-Certificate",
            "Export-PfxCertificate",
            "Invoke-RestMethod",
            "Get-ChildItem",
            "Read-Host"
        )) {
            if (-not (Get-Command -Name $cmd -ErrorAction SilentlyContinue)) {
                throw "Required command [$cmd] is not available."
            }
        }

        if (-not (Get-Command -Name Get-MsalToken -ErrorAction SilentlyContinue)) {
            throw "Get-MsalToken is not available. Install and import the MSAL.PS module first."
        }

        if (-not (Test-Path -Path $CertStoreLocation)) {
            throw "Certificate store path [$CertStoreLocation] does not exist."
        }

        if (-not (Test-Path -Path $ExportPath)) {
            Write-Log "Creating export path [$ExportPath]..."
            New-Item -Path $ExportPath -ItemType Directory -Force | Out-Null
        }

        if ($WaitForManualCerUploadConfirmation -and -not $ExportCer) {
            throw "When using -WaitForManualCerUploadConfirmation, you must also specify -ExportCer."
        }

        if ($ExportPfx -and $DisablePfxExportablePrivateKey) {
            throw "You cannot use -ExportPfx together with -DisablePfxExportablePrivateKey."
        }
    }

    function Get-ExistingCertificate {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Subject
        )

        Write-Log "Searching for existing valid certificate with subject [CN=$Subject] in [$CertStoreLocation]..."

        $certs = Get-ChildItem -Path $CertStoreLocation |
            Where-Object {
                $_.Subject -eq "CN=$Subject" -and
                $_.NotAfter -gt (Get-Date) -and
                $_.HasPrivateKey
            } |
            Sort-Object NotAfter -Descending

        if ($certs) {
            $selected = $certs | Select-Object -First 1
            Write-Log "Found existing certificate with thumbprint [$($selected.Thumbprint)] expiring [$($selected.NotAfter)]."
            return $selected
        }

        Write-Log "No existing valid certificate found." "WARN"
        return $null
    }

    function New-ProductionSelfSignedCertificate {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Subject
        )

        $keyExportPolicy = if ($DisablePfxExportablePrivateKey) { "NonExportable" } else { "Exportable" }

        $certParams = @{
            Subject           = "CN=$Subject"
            CertStoreLocation = $CertStoreLocation
            KeyAlgorithm      = "RSA"
            KeyLength         = $KeyLength
            HashAlgorithm     = "SHA256"
            KeySpec           = "Signature"
            KeyExportPolicy   = $keyExportPolicy
            Provider          = "Microsoft Enhanced RSA and AES Cryptographic Provider"
            NotAfter          = (Get-Date).AddMonths($ValidityMonths)
            FriendlyName      = $Subject
            Type              = "Custom"
            TextExtension     = @(
                "2.5.29.37={text}1.3.6.1.5.5.7.3.2"
            )
            KeyUsage          = "DigitalSignature"
        }

        Write-Log "Creating new self-signed certificate for subject [CN=$Subject]..."
        $cert = New-SelfSignedCertificate @certParams
        Write-Log "Created certificate with thumbprint [$($cert.Thumbprint)] expiring [$($cert.NotAfter)]."

        return $cert
    }

    function Export-CertificateFiles {
        param(
            [Parameter(Mandatory = $true)]
            [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
        )

        $result = [ordered]@{
            CerPath = $null
            PfxPath = $null
        }

        if ($ExportCer) {
            $cerPath = Join-Path -Path $ExportPath -ChildPath "$SubjectName.cer"
            Write-Log "Exporting CER to [$cerPath]..."
            Export-Certificate -Cert $Certificate -FilePath $cerPath -Force | Out-Null
            $result.CerPath = $cerPath
        }

        if ($ExportPfx) {
            $pfxPath = Join-Path -Path $ExportPath -ChildPath "$SubjectName.pfx"
            $certPw = Read-Host -Prompt "Enter password for PFX export" -AsSecureString

            Write-Log "Exporting PFX to [$pfxPath]..."
            Export-PfxCertificate `
                -Cert ("{0}\{1}" -f $CertStoreLocation, $Certificate.Thumbprint) `
                -FilePath $pfxPath `
                -ChainOption EndEntityCertOnly `
                -Password $certPw `
                -Force | Out-Null

            $result.PfxPath = $pfxPath
        }

        return [pscustomobject]$result
    }

    function Confirm-CertificateUploadedToAppRegistration {
        param(
            [Parameter(Mandatory = $true)]
            [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,

            [Parameter(Mandatory = $true)]
            [string]$CerPath,

            [Parameter(Mandatory = $true)]
            [string]$ApplicationId
        )

        Write-Host ""
        Write-Host "==================================================================" -ForegroundColor Yellow
        Write-Host "MANUAL ACTION REQUIRED" -ForegroundColor Yellow
        Write-Host "==================================================================" -ForegroundColor Yellow
        Write-Host "Exported CER file : $CerPath"
        Write-Host "Application ID    : $ApplicationId"
        Write-Host "Subject           : $($Certificate.Subject)"
        Write-Host "Thumbprint        : $($Certificate.Thumbprint)"
        Write-Host "Expires           : $($Certificate.NotAfter)"
        Write-Host ""
        Write-Host "Upload this CER file to the target Entra app registration"
        Write-Host "before the script continues."
        Write-Host ""
        Write-Host "Type YES once complete."
        Write-Host "Type NO to cancel."
        Write-Host "==================================================================" -ForegroundColor Yellow
        Write-Host ""

        do {
            $response = Read-Host "Have you uploaded the CER to the app registration? (YES/NO)"
            $response = $response.Trim().ToUpperInvariant()

            switch ($response) {
                "YES" {
                    Write-Log "User confirmed the certificate has been uploaded to the app registration."
                    return
                }
                "NO" {
                    throw "Operation cancelled by user. The certificate must be uploaded to the app registration before continuing."
                }
                default {
                    Write-Warning "Invalid response. Please type YES or NO."
                }
            }
        }
        while ($true)
    }

    function Get-GraphAccessToken {
        param(
            [Parameter(Mandatory = $true)]
            [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
        )

        Write-Log "Requesting Microsoft Graph token using certificate authentication..."

        try {
            $tokenResponse = Get-MsalToken `
                -ClientId $ApplicationId `
                -TenantId $TenantId `
                -ClientCertificate $Certificate `
                -Scopes "https://graph.microsoft.com/.default"
        }
        catch {
            throw @"
Failed to obtain token using certificate authentication.

Likely causes:
- The public key of this certificate has not been uploaded to the Entra app registration.
- The app registration does not have the required Microsoft Graph application permissions.
- Admin consent has not been granted.
- ApplicationId or TenantId is incorrect.

Underlying error:
$($_.Exception.Message)
"@
        }

        if (-not $tokenResponse.AccessToken) {
            throw "Token request completed but no access token was returned."
        }

        Write-Log "Access token acquired successfully."
        return $tokenResponse.AccessToken
    }

    function Invoke-GraphGetPaged {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Uri,

            [Parameter(Mandatory = $true)]
            [hashtable]$Headers
        )

        $allItems = [System.Collections.Generic.List[object]]::new()
        $nextUri = $Uri

        do {
            Write-Log "Calling Graph endpoint [$nextUri]..."
            $response = Invoke-RestMethod -Method Get -Uri $nextUri -Headers $Headers -ContentType "application/json"

            if ($response.PSObject.Properties['value']) {
                foreach ($item in @($response.value)) {
                    [void]$allItems.Add($item)
                }
            }
            else {
                [void]$allItems.Add($response)
            }

            $nextLinkProperty = $response.PSObject.Properties['@odata.nextLink']
            if ($null -ne $nextLinkProperty) {
                $nextUri = [string]$nextLinkProperty.Value
            }
            else {
                $nextUri = $null
            }
        }
        while ($nextUri)

        return $allItems
    }
}

process {
    try {
        Assert-Prerequisites

        $certificate = $null

        if (-not $CreateNewCertificate -and -not $SkipCertificateReuse) {
            $certificate = Get-ExistingCertificate -Subject $SubjectName
        }

        if (-not $certificate) {
            if ($PSCmdlet.ShouldProcess("Certificate store [$CertStoreLocation]", "Create certificate [$SubjectName]")) {
                $certificate = New-ProductionSelfSignedCertificate -Subject $SubjectName
            }
        }

        if (-not $certificate) {
            throw "No certificate is available for use."
        }

        $exportResult = Export-CertificateFiles -Certificate $certificate

        Write-Log "Certificate summary:"
        Write-Log "  Subject    : $($certificate.Subject)"
        Write-Log "  Thumbprint : $($certificate.Thumbprint)"
        Write-Log "  Expires    : $($certificate.NotAfter)"
        if ($exportResult.CerPath) { Write-Log "  CER Export : $($exportResult.CerPath)" }
        if ($exportResult.PfxPath) { Write-Log "  PFX Export : $($exportResult.PfxPath)" }

        if ($WaitForManualCerUploadConfirmation) {
            if (-not $exportResult.CerPath) {
                throw "CER export is required when using -WaitForManualCerUploadConfirmation. Use -ExportCer."
            }

            Confirm-CertificateUploadedToAppRegistration `
                -Certificate $certificate `
                -CerPath $exportResult.CerPath `
                -ApplicationId $ApplicationId
        }

        if (-not $SkipGraphQuery) {
            $accessToken = Get-GraphAccessToken -Certificate $certificate

            $headers = @{
                Authorization = "Bearer $accessToken"
            }

            $devicesUri = "https://graph.microsoft.com/$GraphApiVersion/devices?`$select=id,displayName,operatingSystem,operatingSystemVersion,trustType,approximateLastSignInDateTime"
            $devices = Invoke-GraphGetPaged -Uri $devicesUri -Headers $headers

            Write-Log "Retrieved [$($devices.Count)] device(s) from Microsoft Graph."

            $devices |
                Sort-Object displayName |
                Select-Object id, displayName, operatingSystem, operatingSystemVersion, trustType, approximateLastSignInDateTime |
                Format-Table -AutoSize
        }
        else {
            Write-Log "Graph query was skipped by parameter."
        }
    }
    catch {
        Write-Log "Script failed. $($_.Exception.Message)" "ERROR"
        exit 1
    }
}

end {
    Write-Log "Script completed."
}