using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-PSFMessage -Message  "PowerShell HTTP trigger function processed a request."
# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body       = 'Function Started successfully. Collection will now begin.'
    })


$azTenant = Get-AzTenant
$tenantId = $azTenant.TenantId
$tenantFQDN = $env:_AZTenantDefaultDomain  
Write-PSFMessage -Message  "Got tenant information: $tenantId - $tenantFQDN"


$certBase64 = Get-AzKeyVaultSecret -ResourceId $env:_PnPPowerShell_KeyVaultId -Name $env:_PnPApplicationName -AsPlainText -ErrorAction Stop
Write-PSFMessage -Message  "Got PnP Application certificate as Base64"

Connect-PnPOnline -Url "https://$env:_SharePointDomain" -ClientId $env:_PnPClientId -Tenant $tenantFQDN -CertificateBase64Encoded $certBase64 -ErrorAction Stop

Write-PSFMessage -Message  "Connected to PNP"





$null = New-Item -Path .\temp -ItemType Directory -Force
$timeStamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$tempFolder = New-Item -Path ".\temp\$timeStamp" -ItemType Directory

#region: Decide to scan all sites or just a provided URL
$targetURL = $Request.Query.URL
if (-not $targetURL) {
    $targetURL = $Request.Body.URL
}

if ($targetURL) {
    $SitesCollections = [PSCustomObject]@{
        URL = $targetURL
        FileName = "$($targetURL.Replace('https://','').Replace('/','_')).CSV" 
    }
    Write-PSFMessage -Message  "Will scan only against: $targetURL"
}
else {
    Write-PSFMessage -Message  "URL not provided, scanning all SharePoint sites."
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
    $body = $sitesCollections | Select-Object -Property Url,Template,FileName | ConvertTo-Csv | Out-String -Width 9999
    $url = "https://$env:_StorageAccountName.blob.core.windows.net/$env:_CSVBlobContainerName/SiteCollections.csv" 
    Invoke-RestMethod -Method PUT -Uri $url -Headers $headers -Body $body 

    Write-PSFMessage -Message  "Uploaded list of Site Collections"
}
#endregion

#Loop through each site collection
$i = 0
ForEach ($site in $SitesCollections) {
    $i++  # Counter for logs
    #Connect to site collection
    Write-PSFMessage -Level Host -Message  "Generating Report for Site ($i):$($Site.Url)"
    $filename = $site.FileName
    $reportFile = Join-Path -Path $tempFolder.FullName -ChildPath $filename
    Write-PSFMessage -Level Host -Message  "Report will be stored temporarily as: $reportFile"

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


    Write-PSFMessage -Level Host  "Uploaded file to Blob storage: $reportFile"
    Remove-Item -Path $reportFile -Force -Confirm:$false 

}
