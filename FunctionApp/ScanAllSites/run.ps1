using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-PSFMessage -Message  "PowerShell HTTP trigger function processed a request."

$azTenant = Get-AzTenant
$tenantId = $azTenant.TenantId
$tenantFQDN = $env:_AZTenantDefaultDomain  
Write-PSFMessage -Message  "Got tenant information: $tenantId - $tenantFQDN"


$certBase64 = Get-AzKeyVaultSecret -ResourceId $env:_PnPPowerShell_KeyVaultId -Name $env:_PnPApplicationName -AsPlainText -ErrorAction Stop
Write-PSFMessage -Message  "Got PnP Application certificate as Base64"

Connect-PnPOnline -Url "https://$env:_SharePointDomain" -ClientId $env:_PnPClientId -Tenant $tenantFQDN -CertificateBase64Encoded $certBase64 -ErrorAction Stop

Write-PSFMessage -Message  "Connected to PNP"


Write-PSFMessage -Message  "Getting list of SharePoint site collections."
$skippedTempaltes = @(
    "SRCHCEN#0", "SPSMSITEHOST#0", "APPCATALOG#0", "POINTPUBLISHINGHUB#0", "EDISC#0", "STS#-1", "OINTPUBLISHINGTOPIC#0", "TEAMCHANNEL#0", "TEAMCHANNEL#1"
)
$SitesCollections = Get-PnPTenantSite | Where-Object -Property Template -NotIn $skippedTempaltes
# Calculate file name for each site
$SitesCollections | ForEach-Object {
    $_ | Add-Member -MemberType NoteProperty -Name FileName -Value "$($_.URL.Replace('https://','').Replace('/','_')).CSV" 
}
Write-PSFMessage -Message  "Found $($sitesCollections.Count) sites"
# upload list of site collections found to blob storage - used by Power BI to ensure we scanned all sites
$headers = Get-SPOPermissionStorageAccessHeaders
$body = $sitesCollections | Select-Object -Property Url, Template, FileName | ConvertTo-Csv | Out-String -Width 9999
$url = "https://$env:_StorageAccountName.blob.core.windows.net/$env:_CSVBlobContainerName/SiteCollections.csv" 
Invoke-RestMethod -Method PUT -Uri $url -Headers $headers -Body $body 

Write-PSFMessage -Message  "Uploaded list of Site Collections"

#endregion

#Loop through each site collection

ForEach ($site in $SitesCollections) {
    Push-OutputBinding -Name SiteCollectionURL -Value $site.Url
}
# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body       = 'Function completed successfully.'
    })