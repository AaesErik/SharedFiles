<#
UpgradeSupport_OnPrem.ps1

Collects installed Business Central extensions for an on-prem environment,
exports the data to CSV, and optionally uploads to FTP.

Parameters:
- CustomerName   : Customer display name
- Environment    : BC environment name
- FTPpass        : FTP password
- SaveDir        : Local folder for CSV files

Note:
- Designed for on-prem BC installations
- Requires Microsoft.Dynamics.Nav.Apps.Management module (imported automatically)
#>

param(
    [parameter(Mandatory=$true)] $CustomerName,
    [parameter(Mandatory=$true)] $Environment,
    [parameter(Mandatory=$true)] $FTPpass,
    [parameter(Mandatory=$true)] $SaveDir
)

# -----------------------------
# Prepare environment
# -----------------------------
function PrepareEnvironment {
    if (-not (Get-Module -Name Microsoft.Dynamics.Nav.Apps.Management -ListAvailable -ErrorAction Ignore)) {
        $paths = @(
            "C:\Program Files\LS Retail\Update Service\Instances\*\Service\Management\Microsoft.Dynamics.Nav.Apps.Management.psd1",
            "C:\Program Files (x86)\LS Retail\Update Service\Instances\*\Service\Management\Microsoft.Dynamics.Nav.Apps.Management.psd1",
            "C:\Program Files\Microsoft Dynamics *\*\Service\Management\Microsoft.Dynamics.Nav.Apps.Management.psd1",
            "C:\Program Files\Microsoft Dynamics *\*\Service\Management\Management\Microsoft.Dynamics.Nav.Apps.Management.psd1",
            "C:\Program Files\Microsoft Dynamics 365 Business Central\*\Service\Admin\NavAdminTool.ps1"
        )

        $modulePath = $paths | ForEach-Object { Get-ChildItem -Path $_ -ErrorAction SilentlyContinue } | Select-Object -First 1
        if (-not $modulePath) { throw "Microsoft.Dynamics.Nav.Apps.Management module not found!" }
        Import-Module $modulePath.FullName
    }
}
PrepareEnvironment

# -----------------------------
# Prepare CSV folder
# -----------------------------
$CsvPath = Join-Path $SaveDir "BCDeploymentLog"
if (-not (Test-Path $CsvPath)) { New-Item -Path $CsvPath -ItemType Directory -Force }

# -----------------------------
# Export App Info
# -----------------------------
$Apps = Get-NAVAppInfo $Environment -Tenant default -TenantSpecificProperties |
    Select-Object Name, publisher, version, isInstalled, Scope, AppId |
    Sort-Object publisher, Name

$CsvFile = Join-Path $CsvPath "$CustomerName.$Environment.csv"
$Apps | Export-Csv -Path $CsvFile -NoTypeInformation -Delimiter ','

# Rename columns
(Get-Content $CsvFile -Raw) -replace 'Name','APPName' |
    Set-Content $CsvFile
(Get-Content $CsvFile -Raw) -replace 'Scope','publishedAs' |
    Set-Content $CsvFile
(Get-Content $CsvFile -Raw) -replace 'AppId','packageId' |
    Set-Content $CsvFile

# Add metadata columns
$Date = Get-Date -Format "yyyy-MM-dd HH:mm"
$CsvData = Import-Csv -Path $CsvFile -Delimiter ';'
$CsvData | ForEach-Object {
    $_ | Add-Member -MemberType NoteProperty -Name 'CustomerName' -Value $CustomerName
    $_ | Add-Member -MemberType NoteProperty -Name 'Environment' -Value $Environment
    $_ | Add-Member -MemberType NoteProperty -Name 'DateTime' -Value $Date
}
$CsvData | Export-Csv -Path $CsvFile -NoTypeInformation

# -----------------------------
# Upload to FTP
# -----------------------------
function Upload-FTP {
    param($LocalFile, $FtpHost, $FilePath, $FtpUser, $FtpPass)

    if (-not $FtpUser -or -not $FtpPass) { return }

    $uri = "ftp://$($FtpUser):$($FtpPass)@$($FtpHost)/$($FilePath)"
    Write-Host "Uploading $LocalFile to $uri..."
    $webclient = New-Object System.Net.WebClient
    $webclient.UploadFile($uri, $LocalFile)
}

# Conditional FTP handling for special customers
if ($CustomerName -in @('(DK) NRGi','(DK) NRGi BC13')) {
    Write-Host "Skipping FTP upload for NRGi customer"
} elseif ($CustomerName -eq "(DK) Evapco") {
    Set-NetFirewallProfile -Profile Domain, Public, Private -Enabled $false
    $RemoteFile = "13.80.26.17/$CustomerName.$Environment.csv"
    $FTPRequest = [System.Net.FtpWebRequest]::Create($RemoteFile)
    $FTPRequest.Method = [System.Net.WebRequestMethods+Ftp]::UploadFile
    $FTPRequest.Credentials = New-Object System.Net.NetworkCredential('dxc', $FTPpass)
    $FTPRequest.UseBinary = $false
    $FTPRequest.UsePassive = $false

    $FileContent = Get-Content -Path $CsvFile -Encoding Byte
    $FTPRequest.ContentLength = $FileContent.Length
    $Stream = $FTPRequest.GetRequestStream()
    $Stream.Write($FileContent, 0, $FileContent.Length)
    $Stream.Close()
    Set-NetFirewallProfile -Profile Domain, Public, Private -Enabled $true
} else {
    Upload-FTP -LocalFile $CsvFile -FtpHost '13.80.26.17' -FilePath "$CustomerName.$Environment.csv" -FtpUser 'dxc' -FtpPass $FTPpass

}

Write-Host "Done!"
