using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)



# Write to the Azure Functions log stream.
Write-PSFMessage -Message  "PowerShell HTTP trigger function processed a request."

#region: get list of all site collections
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

#endregion

#region: Check if we are targeting specific URL(s) from query or body
$targetURLs = $Request.Query.URL
if (-not $targetURLs) {
    $targetURLs = $Request.Body.URL
}

if ($targetURLs) {
    $targetURLs = $targetURLs -split ','
    Write-PSFMessage -Message "Got request to scan specific URL(s):"
    $body = 'Scanning URL(s):'
    $targetURLs | ForEach-Object {
        Write-PSFMessage -Message $_
        $body += "`r`n    $_"
    }

    # Validate Supplied URLs
    foreach ($url in $targetURLs ) {
        if ($url -notin $SitesCollections.Url) {
            $body += "INVALID URL: $url"
            $badURLFound = 1

        }
    }

    
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::OK
            Body       = $body
        }) 
    if ($badURLFound) { Stop-PSFFunction -Message 'Bad URLs supplied' -EnableException $true }
    
    $scanList = $SitesCollections | Where-Object { $_ -in $targetURLs }
    $scanList #TODO Remove this line    
}
else {
    $body = "No URL defind in query or body. Will scan all sites."
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::OK
            Body       = $body
        })    
    # upload list of site collections found to blob storage - used by Power BI to ensure we scanned all sites
    $headers = Get-SPOPermissionStorageAccessHeaders
    $body = $sitesCollections | Select-Object -Property Url, Template, FileName | ConvertTo-Csv | Out-String -Width 9999
    $url = "https://$env:_StorageAccountName.blob.core.windows.net/$env:_CSVBlobContainerName/SiteCollections.csv" 
    Invoke-RestMethod -Method PUT -Uri $url -Headers $headers -Body $body 

    Write-PSFMessage -Message  "Uploaded list of Site Collections"

    $scanList = $SitesCollections
}
#endregion

#Loop through each site collection to scan

ForEach ($site in $scanList) {
    Push-OutputBinding -Name SiteCollectionURL -Value $site.Url
}

Write-PSFMessage -Level Host -Message "Function executed without errors"