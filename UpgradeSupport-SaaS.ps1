<#
UpgradeSupport_SaaS.ps1

Collects installed Business Central extensions for a given customer
and environment, exports the data to CSV, and uploads to FTP.

--------------------------------------
GitHub Actions configuration:

Settings, Secret and Variables, Actions
Secrets (sensitive):
- CLIENT_ID
- CLIENT_SECRET
- TENANT_ID

Settings, Secret and Variables, Actions
Variables (non-sensitive):
- CUSTOMER_NAME      e.g. (DK) ABC SaaS
- ENVIRONMENTS       e.g. Production,TestRelease,TestUpgrade

--------------------------------------
Script parameters:
- ClientID
- ClientSecret
- TenantId
- CustomerName
- Environment

Usage:
- Typically executed in a loop per environment from GitHub Actions
- CSV output: C:\BCSaaS Play\<CustomerName>.<Environment>.csv
- FTP upload included

#>

param(
    [Parameter(Mandatory=$true)] $ClientID,
    [Parameter(Mandatory=$true)] $ClientSecret,
    [Parameter(Mandatory=$true)] $TenantId,
    [Parameter(Mandatory=$true)] $CustomerName,
    [Parameter(Mandatory=$true)] $Environment
)

# --- Konstanter ---
$Scopes       = "https://api.businesscentral.dynamics.com/.default"
$LoginURL     = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
$BaseUrl      = "https://api.businesscentral.dynamics.com/v2.0/$Environment/api/microsoft/automation/v1.0"
$ExtensionUrl = "https://api.businesscentral.dynamics.com/v2.0/$TenantId/$Environment/api/microsoft/automation/v1.0"
$Path         = "C:\BCSaaS Play\"
If (!(Test-Path $Path)) { New-Item -Path $Path -ItemType Directory -Force }

# Debug info uden at afsløre secrets
Write-Host "CUSTOMER_NAME: $env:CUSTOMER_NAME"
Write-Host "ENVIRONMENT: $envName"
Write-Host "LoginURL (uden secrets): $LoginURL"


# --- Funktioner ---
function Get-AccessToken {
    param($ClientID, $ClientSecret)
    $Body = @{
        grant_type    = "client_credentials"
        scope         = $Scopes
        client_id     = $ClientID
        client_secret = $ClientSecret
    }
    Invoke-RestMethod -Method Post -Uri $LoginURL -Body $Body
}

function Invoke-BCAPI {
    param($Uri, $Token)
    Invoke-RestMethod -Method Get -Uri $Uri -Headers @{ Authorization = "Bearer $Token" }
}

function Upload-FTP {
    param($LocalFile, $RemoteFile)
    $webclient = New-Object System.Net.WebClient
    $uri = New-Object System.Uri $RemoteFile
    Write-Host "Uploading $LocalFile to $RemoteFile..."
    $webclient.UploadFile($uri, $LocalFile)
}

# --- Hovedlogik ---
$Token = (Get-AccessToken -ClientID $ClientID -ClientSecret $ClientSecret).access_token

# Hent første company
$Company = Invoke-BCAPI -Uri "$BaseUrl/companies" -Token $Token
$CompanyId = $Company.value[0].id
Write-Host "First companyID: $CompanyId"

# Hent extensions
$Extensions = Invoke-BCAPI -Uri "$ExtensionUrl/companies($CompanyId)/extensions" -Token $Token

# Rens og formater CSV direkte i memory
$Date = Get-Date -Format "yyyy-MM-dd HH:mm"
$CsvData = $Extensions.value | ForEach-Object {
    [PSCustomObject]@{
        APPName      = $_.displayName
        Publisher    = $_.publisher
        Version      = "$($_.versionMajor).$($_.versionMinor).$($_.versionBuild).$($_.versionRevision)"
        IsInstalled  = $_.isInstalled
        publishedAs  = $_.publishedAs
        packageId    = $_.id
        CustomerName = $CustomerName
        Environment  = $Environment
        DateTime     = $Date
    }
} | Sort-Object Publisher, APPName

$CsvFile = "$Path\$CustomerName.$Environment.csv"

# Eksporter CSV med komma og citater omkring alle felter
$CsvData | Export-Csv -Path $CsvFile -NoTypeInformation -Delimiter ',' -Force -Encoding UTF8

# FTP upload
$FtpUrl = "ftp://dxc`:-2~4bVX3~ffq<%wBa3Q9n2+)@13.80.26.17/$CustomerName.$Environment.csv"
Upload-FTP -LocalFile $CsvFile -RemoteFile $FtpUrl

Write-Host "Done!"
