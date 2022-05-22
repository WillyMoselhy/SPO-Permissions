[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]
    $SiteCollectionURL
)
Write-PSFMessage -Level Host -Message  "Starting scan for: $SiteCollectionURL"

# get information about the tenant 
$azTenant = Get-AzTenant
$tenantId = $azTenant.TenantId
$tenantFQDN = $env:_AZTenantDefaultDomain  
Write-PSFMessage -Message  "Got tenant information: $tenantId - $tenantFQDN"

# Get PnP certificate from Keyvault
$certBase64 = Get-AzKeyVaultSecret -ResourceId $env:_PnPPowerShell_KeyVaultId -Name $env:_PnPApplicationName -AsPlainText -ErrorAction Stop
Write-PSFMessage -Message  "Got PnP Application certificate as Base64"

# Connecto to PnP
Connect-PnPOnline -Url "https://$env:_SharePointDomain" -ClientId $env:_PnPClientId -Tenant $tenantFQDN -CertificateBase64Encoded $certBase64 -ErrorAction Stop
Write-PSFMessage -Message  "Connected to PNP"

# Prepare folder to store output
$null = New-Item -Path .\temp -ItemType Directory -Force
$timeStamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$tempFolder = New-Item -Path ".\temp\$timeStamp" -ItemType Directory
Write-PSFMessage -Message "CSVs will be stored in $tempFolder"

# Prepare site object with URL and file path
$site = [PSCustomObject]@{
    URL      = $SiteCollectionURL
    FileName = "$($SiteCollectionURL.Replace('https://','').Replace('/','_')).CSV" 
}
$filename = $site.FileName
$reportFile = Join-Path -Path $tempFolder.FullName -ChildPath $filename
Write-PSFMessage -Level Host -Message  "Report will be stored temporarily as: $reportFile"

#Connect to site collection
$SiteConn = Connect-PnPOnline -Url $Site.Url -ClientId $env:_PnPClientId -Tenant $tenantFQDN -CertificateBase64Encoded $certBase64
#Call the Function for site collection
Start-SPOPermissionCollection -SiteURL $Site.URL -ReportFile $reportFile -Recursive -ScanItemLevel -GraphApiToken $env:mgToken -Verbose # -IncludeInheritedPermissions
Disconnect-PnPOnline -Connection $SiteConn

# Pickup the stored CSV for upload
$body = Get-Content -Path $reportFile | Out-String -Width 9999

# Get storage access token and headers then upload file
$headers = Get-SPOPermissionStorageAccessHeaders
$url = "https://$env:_StorageAccountName.blob.core.windows.net/$env:_CSVBlobContainerName/$filename" 
Invoke-RestMethod -Method PUT -Uri $url -Headers $headers -Body $body 


Write-PSFMessage -Level Host  "Uploaded file to Blob storage: $reportFile"
Remove-Item -Path $reportFile -Force -Confirm:$false 

Write-PSFMessage -Level Host "Completed scan for: $SiteCollectionURL"