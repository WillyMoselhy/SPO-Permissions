[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]
    $URL,

    [switch] $CalledByHTTP #This is to avoid errors when returning HTTP codes to timer trigger
)


#region: get list of all site collections
$azTenant = Get-AzTenant
$tenantId = $azTenant.TenantId
$tenantFQDN = $env:_AZTenantDefaultDomain
Write-PSFMessage -Message "Got tenant information: $tenantId - $tenantFQDN"


$certBase64 = Get-AzKeyVaultSecret -ResourceId $env:_PnPPowerShell_KeyVaultId -Name $env:_PnPApplicationName -AsPlainText -ErrorAction Stop
Write-PSFMessage -Message "Got PnP Application certificate as Base64"

Connect-PnPOnline -Url "https://$env:_SharePointDomain" -ClientId $env:_PnPClientId -Tenant $tenantFQDN -CertificateBase64Encoded $certBase64 -ErrorAction Stop

Write-PSFMessage -Message "Connected to PNP"


Write-PSFMessage -Message "Getting list of SharePoint site collections."
$skippedTempaltes = @(
    "SRCHCEN#0", "SPSMSITEHOST#0", "APPCATALOG#0", "POINTPUBLISHINGHUB#0", "EDISC#0", "STS#-1", "OINTPUBLISHINGTOPIC#0", "TEAMCHANNEL#0", "TEAMCHANNEL#1"
)
$SitesCollections = Get-PnPTenantSite | Where-Object -Property Template -NotIn $skippedTempaltes
# Calculate file name for each site
$SitesCollections | ForEach-Object {
    $_ | Add-Member -MemberType NoteProperty -Name FileName -Value "$($_.URL.Replace('https://','').Replace('/','_')).CSV"
}
Write-PSFMessage -Message "Found $($sitesCollections.Count) sites"

#endregion

if ($URL) {
    $targetURLs = $URL -split ','
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
            Write-PSFMessage -Level Error -Message "URL not in site collection list: $url"
        }
    }

    if ($badURLFound ) {
        $statusCode = [HttpStatusCode]::BadRequest
        #Stop-PSFFunction -Message 'Bad URLs supplied' -EnableException $true
    }
    else {
        $statusCode = [HttpStatusCode]::OK
        $scanList = $SitesCollections | Where-Object { $_.Url -in $targetURLs }
    }
}
else {
    # If no URL is defined we scan all site collections and update the Site Collections CSV list

    if ($CalledByHTTP) {
        $body = "No URL defined in query or body. Will scan all sites."
        $body += $sitesCollections | Select-Object -Property Url, Template, FileName | ConvertTo-Csv | Out-String -Width 9999
        $statusCode = [HttpStatusCode]::OK
    }

    # upload list of site collections found to blob storage - used by Power BI to ensure we scanned all sites
    $headers = Get-SPOPermissionStorageAccessHeaders
    $url = "https://$env:_StorageAccountName.blob.core.windows.net/$env:_CSVBlobContainerName/SiteCollections.csv"
    Invoke-RestMethod -Method PUT -Uri $url -Headers $headers -Body $body

    Write-PSFMessage -Message "Uploaded list of Site Collections"

    $scanList = $SitesCollections
}
#endregion

if (-not $badURLFound) {
    #Loop through each site collection to scan
    Write-PSFMessage "Pushing queue message to scan site collections(s)"
    ForEach ($site in $scanList) {
        Push-OutputBinding -Name SiteCollectionURL -Value $site.Url
        Write-PSFMessage ("Pushed message for: {0}" -f $site.Url)
    }
}

if ($CalledByHTTP) {
    # return body and HTTP code for manual call
    [HttpResponseContext]@{
        StatusCode = $statusCode
        Body       = $body
    }
}