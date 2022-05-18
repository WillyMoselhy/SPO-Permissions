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

# Interact with query parameters or the body of the request.
$name = $Request.Query.Name
if (-not $name) {
    $name = $Request.Body.Name

}

$body = "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response."
$body += (Get-Module -ListAvailable | Out-String -Width 999)
$body += $env:PSModulePath
$body += "r`n" + (Get-Location)

$azTenant = Get-AzTenant
$tenantId = $azTenant.TenantId
$tenantFQDN = $env:_AZTenantDefaultDomain  
Write-Host "Got tenant information: $tenantId - $tenantFQDN"


$certBase64 = Get-AzKeyVaultSecret -ResourceId $env:_PnPPowerShell_KeyVaultId -Name $env:_PnPApplicationName -AsPlainText -ErrorAction Stop
Write-Host "Got PnP Application certificate as Base64"

Connect-PnPOnline -Url "https://$env:_SharePointDomain" -ClientId $env:_PnPClientId -Tenant $tenantFQDN -CertificateBase64Encoded $certBase64 -ErrorAction Stop

Write-Host "Connected to PNP"

Write-Host "Getting all site collections"
$skippedTempaltes = @(
    "SRCHCEN#0", "SPSMSITEHOST#0", "APPCATALOG#0", "POINTPUBLISHINGHUB#0", "EDISC#0", "STS#-1","OINTPUBLISHINGTOPIC#0", "TEAMCHANNEL#0", "TEAMCHANNEL#1"
)

$SitesCollections = Get-PnPTenantSite | Where-Object -Property Template -NotIn $skippedTempaltes

$null = New-Item -Path .\temp -ItemType Directory -Force
$timeStamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$tempFolder = New-Item -Path ".\temp\$timeStamp" -ItemType Directory
Write-Host "Found $($sitesCollections.Count) sites"

#Loop through each site collection
$i=0
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

    # Storage Token
    Write-Host "Getting Storage Token"
    $resourceURI = "https://storage.azure.com"
    $tokenAuthURI = $env:IDENTITY_ENDPOINT + "?resource=$resourceURI&api-version=2019-08-01"
    $tokenResponse = Invoke-RestMethod -Method Get -Headers @{"X-IDENTITY-HEADER" = "$env:IDENTITY_HEADER" } -Uri $tokenAuthURI
    $storageToken = $tokenResponse.access_token

    $headers = @{
        Authorization    = "Bearer $storageToken"
        'x-ms-version'   = '2021-04-10'
        'x-ms-date'      = '{0:R}' -f (Get-Date).ToUniversalTime()
        'x-ms-blob-type' = 'BlockBlob'
    }

    $url = "https://$env:_StorageAccountName.blob.core.windows.net/$env:_CSVBlobContainerName/$filename" # TODO: Change this to an environment variable

    Invoke-RestMethod -Method PUT -Uri $url -Headers $headers -Body $body 


    Write-Host "Uploaded file to Blob storage: $reportFile"


}
Remove-Item -Path $tempFolder -Force -Confirm:$false