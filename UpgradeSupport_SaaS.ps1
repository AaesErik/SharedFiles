###################################
# PowerShell
###################################
param
(
    [parameter(Mandatory=$true)]
    $ClientID,
    [parameter(Mandatory=$true)]
    $ClientSecret,
    [parameter(Mandatory=$true)]
    $TenantId,
    [parameter(Mandatory=$true)]
    $CustomerName,
    [parameter(Mandatory=$true)]
    $Environment,
    [parameter(Mandatory=$true)]
    $FTPserver,
    [parameter(Mandatory=$true)]
    $FTPuser,
    [parameter(Mandatory=$true)]
    $FTPpass
)

$scopes       = "https://api.businesscentral.dynamics.com/.default"
$loginURL     = "https://login.microsoftonline.com"
$baseUrl      = "https://api.businesscentral.dynamics.com/v2.0/$Environment/api/microsoft/automation/v1.0"
$extensionUrl = "https://api.businesscentral.dynamics.com/v2.0/$TenantId/$Environment/api/microsoft/automation/v1.0"
Write-Host $extensionUrl

$path = "C:\BCSaaS Play\"
If(!(test-path $path))
{
      New-Item -ItemType Directory -Force -Path $path
}

# Get access token 
$body = @{grant_type="client_credentials";scope=$scopes;client_id=$ClientID;client_secret=$ClientSecret}
$URI = "$loginURL/$TenantId/oauth2/v2.0/token"
Write-Host $URI
$oauth = Invoke-RestMethod -Method Post -Uri $URI -Body $body -Verbose

# Get companies
$URI = "$baseurl/companies"
Write-Host $URI
$companies = Invoke-RestMethod `
             -Method Get `
             -Uri $URI `
             -Headers @{Authorization='Bearer ' + $oauth.access_token} -Verbose

#$companies.value | Format-Table -AutoSize
$companyId = $companies.value[0].id
Write-Host "First companyID: $companyId"

# Get extensions
$URI = "$extensionUrl/companies($companyId)/extensions"
Write-Host $URI
$extensions = Invoke-RestMethod `
             -Method Get `
             -Uri $URI `
             -Headers @{Authorization='Bearer ' + $oauth.access_token} -Verbose

# Format and export extensions
#write-host $extensions.value# | Where-Object {$_.publisher -ne 'Microsoft'} | Select-Object -Property displayName, publisher, versionMajor, versionMinor,versionBuild, versionRevision, isInstalled | Sort-Object -Property publisher,displayName > 'C:\BCSaaS play\Output.txt'


$extensions.value | 
#Where-Object {$_.publisher -ne 'Microsoft'} |
Select-Object -Property displayName, publisher, versionMajor, versionMinor,versionBuild, versionRevision, isInstalled, publishedAs, id |
Sort-Object -Property publisher,displayName |
Export-Csv -Path "$path\$CustomerName.$Environment.csv" -NoTypeInformation -Delimiter ';'

#Modify version
$DB = Get-Content "$path\$CustomerName.$Environment.csv"
Clear-Content "$path\$CustomerName.$Environment.csv"
foreach ($Data in $DB) {
  $v1, $v2, $v3, $v4, $v5, $v6, $v7, $v8, $v9 = $Data -split ';' -replace '^\s*|\s*$'
  $v3 = $v3 -replace '"' -replace ''
  $v4 = $v4 -replace '"' -replace ''
  $v5 = $v5 -replace '"' -replace ''
  $v6 = $v6 -replace '"' -replace ''
  $v31 = """$v3"
  $v41 = $v3
  $v51 = $v3
  $v61 = "$v6"""
  $WriteThis ="$v1;$v2;$v31.$v4.$v5.$v61;$v7;$v8;$v9"
  add-content -path "$path\$CustomerName.$Environment.csv" -value $WriteThis
}

$date = Get-Date -Format "yyyy-MM-dd HH:mm"
#Add-Content "$path\$CustomerName.$Environment.csv" $date

((Get-Content -path "$path\$CustomerName.$Environment.csv" -Raw) -replace 'displayName','APPName') | Set-Content -Path "$path\$CustomerName.$Environment.csv"
((Get-Content -path "$path\$CustomerName.$Environment.csv" -Raw) -replace 'publisher','Publisher') | Set-Content -Path "$path\$CustomerName.$Environment.csv"
((Get-Content -path "$path\$CustomerName.$Environment.csv" -Raw) -replace 'isInstalled','IsInstalled') | Set-Content -Path "$path\$CustomerName.$Environment.csv"
((Get-Content -path "$path\$CustomerName.$Environment.csv" -Raw) -replace 'id','packageId') | Set-Content -Path "$path\$CustomerName.$Environment.csv"
((Get-Content -path "$path\$CustomerName.$Environment.csv" -Raw) -replace 'versionMajor.versionMinor.versionBuild.versionRevision','Version') | Set-Content -Path "$path\$CustomerName.$Environment.csv"

$b = Import-CSV -Path "$path\$CustomerName.$Environment.csv" -Delimiter ";"
$b | Add-Member -MemberType NoteProperty -Name 'CustomerName' -Value $CustomerName
$b | Add-Member -MemberType NoteProperty -Name 'Environment' -Value $Environment
$b | Add-Member -MemberType NoteProperty -Name 'DateTime' -Value $date
$b | Export-Csv "$path\$CustomerName.$Environment.csv" -NoTypeInformation

    # FTP upload outout
    $File = "$path\$CustomerName.$Environment.csv";

    Get-Content -Path $File
    $ftp = "ftp://"+$FTPuser+":"+$FTPpass+"@"+$FTPserver+"/$CustomerName.$Environment.csv";

    Write-Host -Object "ftp url: $ftp";
    $webclient = New-Object -TypeName System.Net.WebClient;
    $uri = New-Object -TypeName System.Uri -ArgumentList $ftp;
    Write-Host -Object "Uploading $File...";
    $webclient.UploadFile($uri, $File);



<#
# USERS
$URI = "$extensionUrl/companies($companyId)/users"
Write-Host $URI
$users = Invoke-RestMethod `
             -Method Get `
             -Uri $URI `
             -Headers @{Authorization='Bearer ' + $oauth.access_token} -Verbose

$users.value | 
#Where-Object {$_.publisher -ne 'Microsoft'} |
Select-Object -Property userName, displayName, userSecurityId, state, expiryDate |
Sort-Object -Property userName |
Export-Csv -Path "$path\Users.$CustomerName.$Environment.csv" -NoTypeInformation -Delimiter ';'

    # FTP upload outout
    $File = "$path\Users.$CustomerName.$Environment.csv";

    Get-Content -Path $File
    $ftp = "ftp://"+$FTPuser+":"+$FTPpass+"@"+$FTPserver+"/Users.$CustomerName.$Environment.csv";

    Write-Host -Object "ftp url: $ftp";
    $webclient = New-Object -TypeName System.Net.WebClient;
    $uri = New-Object -TypeName System.Uri -ArgumentList $ftp;
    Write-Host -Object "Uploading $File...";
    $webclient.UploadFile($uri, $File);
# USERS
#>