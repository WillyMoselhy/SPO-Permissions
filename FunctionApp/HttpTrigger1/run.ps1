using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."
# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body       = 'Function Started successfully. Collection will now begin.'
    })


$azTenant = Get-AzTenant
$tenantId = $azTenant.TenantId
$tenantFQDN = $env:_AZTenantDefaultDomain  
Write-Host "Got tenant information: $tenantId - $tenantFQDN"


$certBase64 = Get-AzKeyVaultSecret -ResourceId $env:_PnPPowerShell_KeyVaultId -Name $env:_PnPApplicationName -AsPlainText -ErrorAction Stop
Write-Host "Got PnP Application certificate as Base64"

Connect-PnPOnline -Url "https://$env:_SharePointDomain" -ClientId $env:_PnPClientId -Tenant $tenantFQDN -CertificateBase64Encoded $certBase64 -ErrorAction Stop

Write-Host "Connected to PNP"





$null = New-Item -Path .\temp -ItemType Directory -Force
$timeStamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$tempFolder = New-Item -Path ".\temp\$timeStamp" -ItemType Directory

#region: Decide to scan all sites or just a provided URL
$targetURL = $Request.Query.URL
if (-not $targetURL) {
    $targetURL = $Request.Body.URL
}

if ($targetURL) {
    $SitesCollections = [PSCustomObject]@{URL = $targetURL}
    Write-Host "Will scan only against: $targetURL"
}
else {
    Write-Host "URL not provided, scanning all SharePoint sites."
    $skippedTempaltes = @(
        "SRCHCEN#0", "SPSMSITEHOST#0", "APPCATALOG#0", "POINTPUBLISHINGHUB#0", "EDISC#0", "STS#-1", "OINTPUBLISHINGTOPIC#0", "TEAMCHANNEL#0", "TEAMCHANNEL#1"
    )
    $SitesCollections = Get-PnPTenantSite | Where-Object -Property Template -NotIn $skippedTempaltes
    Write-Host "Found $($sitesCollections.Count) sites"
    # upload list of site collections found to blob storage - used by Power BI to ensure we scanned all sites
    $headers = Get-SPOPermissionStorageAccessHeaders
    $body = $sitesCollections | Select-Object -Property Url,Template,LocaleID | ConvertTo-Csv
    $url = "https://$env:_StorageAccountName.blob.core.windows.net/$env:_CSVBlobContainerName/SiteCollections.csv" 
    Invoke-RestMethod -Method PUT -Uri $url -Headers $headers -Body $body 

    Write-Host "Found $($sitesCollections.Count) sites"
}
#endregion

#Loop through each site collection
$i = 0
ForEach ($Site in $SitesCollections) {
    $i++  # Counter for logs
    #Connect to site collection
    Write-Host "Generating Report for Site ($i):$($Site.Url)"
    $filename = "$($Site.URL.Replace('https://','').Replace('/','_')).CSV"
    $reportFile = Join-Path -Path $tempFolder.FullName -ChildPath $filename
    Write-Host "Report will be stored temporarily as: $reportFile"

    $SiteConn = Connect-PnPOnline -Url $Site.Url -ClientId $env:_PnPClientId -Tenant $tenantFQDN -CertificateBase64Encoded $certBase64


    #Call the Function for site collection
    Start-SPOPermissionCollection -SiteURL $Site.URL -ReportFile $reportFile -Recursive -ScanItemLevel -GraphApiToken $env:mgToken -Verbose # -IncludeInheritedPermissions
    Disconnect-PnPOnline -Connection $SiteConn

    $csv = Get-Content -Path $reportFile | Out-String -Width 9999
    $body = $csv


    # Get storage access token and headers then upload file
    $headers = Get-SPOPermissionStorageAccessHeaders
    $url = "https://$env:_StorageAccountName.blob.core.windows.net/$env:_CSVBlobContainerName/$filename" 
    Invoke-RestMethod -Method PUT -Uri $url -Headers $headers -Body $body 


    Write-Host "Uploaded file to Blob storage: $reportFile"


}
Remove-Item -Path $tempFolder -Force -Confirm:$false