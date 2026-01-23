<#
AppRegistrations.ps1

Exports all Microsoft Graph App Registrations for a given tenant,
including secret expiry information, and uploads to FTP.

--------------------------------------
GitHub Actions configuration:

Variables (non-sensitive):
- CLIENT_ID
- CLIENT_SECRET
- TENANT_ID
- CUSTOMER_NAME
- RUN_APPREGISTRATIONS (true/false)

--------------------------------------
Permissions required in Azure AD (mandatory!):
- Microsoft Graph
- Application.Read.All

--------------------------------------
Script parameters:
- ClientID
- ClientSecret
- TenantId
- CustomerName

Usage:
- Typically executed from GitHub Actions if RUN_APPREGISTRATIONS=true
- Output CSV file: %TEMP%\AppRegistrations.csv
- FTP upload included

#>

param (
    [Parameter(Mandatory=$true)][string]$ClientID,
    [Parameter(Mandatory=$true)][string]$ClientSecret,
    [Parameter(Mandatory=$true)][string]$TenantId,
    [Parameter(Mandatory=$true)][string]$CustomerName
)

# ===============================
# Constants
# ===============================
$FtpUrl  = "ftp://13.80.26.17/${CustomerName}_AllApps.txt"

# ===============================
# Connect to Microsoft Graph
# ===============================
$secureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
$credential   = New-Object System.Management.Automation.PSCredential($ClientID, $secureSecret)

Connect-MgGraph `
    -TenantId $TenantId `
    -ClientSecretCredential $credential `
    -NoWelcome

# ===============================
# Retrieve all app registrations
# ===============================
$apps = Get-MgApplication -All
$now  = Get-Date

$allAppsWithExpiry = foreach ($app in $apps) {
    if ($app.PasswordCredentials -and $app.PasswordCredentials.Count -gt 0) {
        foreach ($secret in $app.PasswordCredentials) {
            [PSCustomObject]@{
                AppName     = $app.DisplayName
                AppId       = $app.AppId
                CreatedDate = $app.CreatedDateTime
                Expiry      = $secret.EndDateTime
                DaysLeft    = ($secret.EndDateTime - $now).Days
            }
        }
    } else {
        [PSCustomObject]@{
            AppName     = $app.DisplayName
            AppId       = $app.AppId
            CreatedDate = $app.CreatedDateTime
            Expiry      = "N/A"
            DaysLeft    = "N/A"
        }
    }
}

# ===============================
# Upload Function as CSV (semicolon)
# ===============================
function Upload-AppRegistrationsToFtp {
    param(
        [Parameter(Mandatory=$true)][array]$Apps,
        [Parameter(Mandatory=$true)][string]$TenantId,
        [Parameter(Mandatory=$true)][string]$FtpUrl,
        [Parameter(Mandatory=$true)][string]$FTPuser,
        [Parameter(Mandatory=$true)][string]$FTPpass
    )

    if (-not $Apps -or $Apps.Count -eq 0) {
        Write-Host "No app registrations found"
        return
    }

    # Create semicolon-separated CSV
    $localFile = Join-Path $env:TEMP "AppRegistrations.csv"
    $Apps | Select-Object AppName, AppId, CreatedDate, Expiry, DaysLeft |
        Export-Csv -Path $localFile -NoTypeInformation -Delimiter ';' -Encoding UTF8

    # Add TenantId as first line
    $csvLines = Get-Content $localFile
    $csvLines = @("TenantId=$TenantId") + $csvLines
    Set-Content -Path $localFile -Value $csvLines -Encoding UTF8

    try {
        $webclient = New-Object System.Net.WebClient
        $webclient.Credentials = New-Object System.Net.NetworkCredential($FTPuser, $FTPpass)
        $webclient.UploadFile($FtpUrl, $localFile)
        Write-Host "CSV file uploaded to FTP: $FtpUrl"
    }
    catch {
        Write-Error ("FTP upload failed: " + $_.Exception.Message)
    }
    finally {
        Remove-Item $localFile -ErrorAction SilentlyContinue
    }
}

# ===============================
# Execute Upload
# ===============================
Upload-AppRegistrationsToFtp `
    -Apps $allAppsWithExpiry `
    -TenantId $TenantId `
    -FtpUrl $FtpUrl `
    -FTPuser 'SecretExpire' `
    -FTPpass 'SecretExpire'
