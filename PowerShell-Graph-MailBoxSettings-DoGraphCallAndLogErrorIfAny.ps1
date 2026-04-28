# PowerShell-Graph-MailBoxSettings-DoGraphCallAndLogErrorIfAny.ps1
# This sample script demonstrates how to make a call to the Microsoft Graph API to get mailbox settings for a user and 
# log detailed information about the call, especially in case of errors. It uses app-only authentication (client credentials flow) 
# and requires the MailboxSettings.Read application permission to be granted to the app registration in Azure AD. 
# 
# 	MailboxSettings.Read application permission is required to run this script successfully.
#   https://learn.microsoft.com/en-us/graph/api/user-get-mailboxsettings?view=graph-rest-1.0&tabs=http
# 	Make sure to replace the hardcoded values with actual values before running the script.
 
# Set logging type and if your logigng failures or successes or both.
$DetailedLogging = $false  # Set to $false for more concise logging (only log failures with basic info)
$SummeryLogging = $true   # Set to $true to log a one-liner for each call with basic info (good for quick monitoring of successes vs failures)
$LogFailure = $true   # Set to $true to log failures, set to $false to not log failures (not recommended)
$LogSuccess = $false # Set to $false if you only want to log failures, set to $true to log all calls (both successes and failures)   

# ====================================
# Hardcoded Configuration
# ====================================
 
$TenantId     = "<TENANT_ID>"
$ClientId     = "<CLIENT_ID>"
$ClientSecret = "<CLIENT_SECRET>"

$UserPrincipalName = "user@contoso.com"

$OutputFolder = "C:\GraphFailures"
$OutputFile   = Join-Path $OutputFolder "MailboxSettingsFailures.log"
 
# ============================================
# Check output folder exists, if not create it
# ============================================

if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}

# ===========================
# Get access token (app-only)
# ===========================
try {
    $tokenUri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $tokenBody = @{
    client_id     = $ClientId
    client_secret = $ClientSecret
    scope         = "https://graph.microsoft.com/.default"
    grant_type    = "client_credentials"
    }

    "[$(Get-Date -Format o)] Requesting token..." | Out-File -FilePath $LogPath -Encoding utf8
    $tokenResponse = Invoke-RestMethod -Method POST -Uri $tokenUri -Body $tokenBody -ContentType "application/x-www-form-urlencoded"
    $accessToken   = $tokenResponse.access_token 
}
catch {
    Write-Error "Failed to acquire access token: $_"
    exit 1
}
  

# =================================================================
# Get granted permissions for the app (for logging purposes)
# =================================================================

$AccessToken = $TokenResponse.access_token

$TokenParts = $AccessToken.Split('.')
$Payload = $TokenParts[1]

# Fix base64 padding
switch ($Payload.Length % 4) {
    2 { $Payload += '==' }
    3 { $Payload += '=' }
}


$DecodedPayload = [System.Text.Encoding]::UTF8.GetString(
    [System.Convert]::FromBase64String($Payload)
) | ConvertFrom-Json

# Display granted permissions

$GrantedApplicationPermissions = @()
$GrantedDelegatedPermissions  = @()
 

if ($DecodedPayload.roles) {
    $GrantedApplicationPermissions = $DecodedPayload.roles
 
}

if ($DecodedPayload.scp) {
    $GrantedDelegatedPermissions = $DecodedPayload.scp.Split(' ')
    
}

# Optional: combined list
#$GrantedPermissions = "Applicaiton Permissions: " + $GrantedApplicationPermissions + ", Delegated Permissions: " + $GrantedDelegatedPermissions

 
# ===========================
# Call mailboxSettings
# ===========================

$headers = @{
    Authorization = "Bearer $accessToken"
    "Content-Type"= "application/json"
    Accept  = "application/json"
    "client-request-id"         = $ClientRequestId
    "return-client-request-id"  = "true"
    }

$HadError = $false
$ClientRequestId = [guid]::NewGuid().ToString()
$RequestId = "N/A"
$AgsDiagnostic = "N/A"
$StartCallTimeUtc = (Get-Date).ToUniversalTime().ToString("o")
$Uri = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/mailboxSettings"
$StatusCode  = ""
$ResponseBody = ""
$StatusCode  = "Unknown"
$HadError = $true

try {
    $response =Invoke-RestMethod `
        -Method GET `
        -Uri $Uri `
        -Headers $Headers `
        -ErrorAction Stop `
        -ResponseHeadersVariable responseHeaders `
        -StatusCodeVariable StatusCode  

    $RequestId = $responseHeaders['request-id']
    $ClientRequestId = $responseHeaders['client-request-id']
    $AgsDiagnostic = $ResponseHeaders["x-ms-ags-diagnostic"]
    $HadError = $false
}
catch {
    $StatusCode  = "Unknown"
    $ResponseBody = ""
    $HadError = $true

    if ($_.Exception.Response) {
        $StatusCode = $_.Exception.Response.StatusCode.value__
 
        try {
            $RequestId = $ResponseHeaders["request-id"]
        }
        catch {
            $RequestId = "N/A"
        }

        try {
            $AgsDiagnostic = $ResponseHeaders["x-ms-ags-diagnostic"]
        }
        catch {
            $AgsDiagnostic = "N/A"
        }


        try { 
            $Stream = $_.Exception.Response.GetResponseStream()
            if ($Stream) {
                $Reader = New-Object System.IO.StreamReader($Stream)
                $ResponseBody = $Reader.ReadToEnd()
                $Reader.Close()
            }
        }
        catch {
            #$ResponseBody = "Failed to read response body: $_"
            $ResponseBody = "<< Response body may not be available for this error. >>`n$ResponseBody"
            # Note: If there is no response stream then this is the error text: Method invocation failed because [System.Net.Http.HttpResponseMessage] does not contain a method named 'GetResponseStream'.
        }   
    }
 

}

$EndCallTimeUtc = (Get-Date).ToUniversalTime().ToString("o")

# --------------------------------------
# Log failure   
# --------------------------------------
    # Simple logging with just basic info for failures - suggest using in Excellent for quick monitoring of failures without too much detail (can be used in combination with detailed logging or on its own by setting $DetailedLogging to $false)

    #$GrantedPermissions = "Applicaiton Permissions: " + $GrantedApplicationPermissions + ", Delegated Permissions: " + $GrantedDelegatedPermissions


# Log detailed error information 
if ($DetailedLogging -and $HadError -and $LogFailure) {
@"
Call Status: Failed
Graph URL: $Uri
HTTPStatus: $StatusCode
SMTP: $UserPrincipalName
Start Call Time: $StartCallTimeUtc 
End Call Time: $EndCallTimeUtc
Client-Request-Id: $ClientRequestId  
RequestId: $RequestId
AGS-Diagnostic: $AgsDiagnostic
Application Permissions: $($GrantedApplicationPermissions -join ", ")
Delegated Permissions: $($GrantedDelegatedPermissions -join ", ")   

ResponseBody:
$ResponseBody 
----------------------------
"@ | Out-File -FilePath $OutputFile -Append -Encoding UTF8
}

if ($SummeryLogging -and $HadError -and $LogFailure) {
@"
Failed, $StatusCode, $UserPrincipalName, $StartCallTimeUtc, $EndCallTimeUtc, $ClientRequestId, $RequestId
"@ | Out-File -FilePath $OutputFile -Append -Encoding UTF8
}


 # --------------------------------------
# Log success  
# --------------------------------------
if ($LogSuccess -and -not $HadError -and $DetailedLogging) {
     @"
Call Status: Success
Graph URL: $Uri
HTTPStatus: $StatusCode
SMTP: $UserPrincipalName
Start Call Time: $StartCallTimeUtc 
End Call Time: $EndCallTimeUtc
Client-Request-Id: $ClientRequestId  
RequestId: $RequestId
AGS-Diagnostic: $AgsDiagnostic
Application Permissions: $($GrantedApplicationPermissions -join ", ")
Delegated Permissions: $($GrantedDelegatedPermissions -join ", ")   

ResponseBody:
$ResponseBody
----------------------------
"@ | Out-File -FilePath $OutputFile -Append -Encoding UTF8
}
  
if ($LogSuccess -and -not $HadError -and $SummeryLogging) {
         @"
Success, $StatusCode, $UserPrincipalName, $StartCallTimeUtc, $EndCallTimeUtc, $ClientRequestId, $RequestId
"@ | Out-File -FilePath $OutputFile -Append -Encoding UTF8
  
}

$StatusCode
