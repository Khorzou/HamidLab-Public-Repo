<#
.SYNOPSIS
    Retrieves an OAuth 2.0 access token from Azure AD using the 
    client credentials flow (client_id + client_secret) for Microsoft Graph.

.DESCRIPTION
    This script authenticates against Azure AD by sending a POST request 
    to the v2.0 token endpoint. It constructs the required request body, 
    including client_id, client_secret, grant_type, and the .default scope 
    for Microsoft Graph. The script returns a valid Bearer token that can 
    be used in subsequent Graph API calls.

    It is designed to be reusable for automation scenarios such as:
        - Intune device management
        - Graph-based reporting
        - Azure AD automation
        - Background services or scheduled tasks

    Update the ClientID, ClientSecret, and TenantID values as needed 
    when reusing this script in other environments.

.NOTES
    Author: Hamid
    Purpose: Reusable Graph API authentication helper
    Requirements:
        - App Registration with Application permissions
        - Admin consent granted
        - Valid client secret
#>


$ClientID     = "bd6b5575-4fde-4a4a-9cd3-37b4eb9e5438"
$ClientSecret = "SECRET"
$TenantID     = "c8b9f88a-afab-4e3c-a57d-0c1506b66757"
$Resource     = "https://graph.microsoft.com" 

$TokenEndpoint = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"

$body = @{
    client_id     = $ClientID
    client_secret = $ClientSecret
    scope         = "$Resource/.default" 
    grant_type    = "client_credentials"
}

$tokenResponse = Invoke-RestMethod `
    -Uri $TokenEndpoint `
    -Method POST `
    -Body $body `
    -ContentType "application/x-www-form-urlencoded"

$AccessToken = $tokenResponse.access_token

$headers = @{
    Authorization = "Bearer $AccessToken"
    ContentType = "application/json"
}

Invoke-RestMethod -Method GET -Uri "https://graph.microsoft.com/beta/devicemanagement/managedDevices" -Headers $headers
