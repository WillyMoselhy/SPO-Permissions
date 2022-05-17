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



#$LoginInfo = [PSCustomObject]@{
    #TenantID        = $env:_TenantID
    #TenantName      = $env:_TenantName
    #AppID           = '2e1fee6b-7fe5-48ac-b51a-da35e149f1c5'# $env:_AppID
    #CertificatePath = 'C:\home\site\wwwroot\Cert\PnP Rocks2.pfx' #This can be EncodedBase64
    #BlobFunctionKey = $env:_SaveBlobFunction
#}

$azTenant = Get-AzTenant
$tenantId = $azTenant.TenantId
$tenantFQDN = $env:_AZTenantDefaultDomain  
Write-Host "Got tenant information: $tenantId - $tenantFQDN"

#$cert = Get-AzKeyVaultSecret -ResourceId $env:PnPPowerShell_KeyVaultId -Name PnPPowerShell -AsPlainText
#Write-Host "Cert Obtained from keyvault"

$certBase64 = Get-AzKeyVaultSecret -ResourceId $env:_PnPPowerShell_KeyVaultId -Name $env:_PnPApplicationName -AsPlainText
Write-Host "Got PnP Application certificate as Base64"

#$MsalToken = Get-MsalToken -ClientId $ClientId  -ClientCertificate $certBase64 -TenantId $tenantId -ForceRefresh
#Write-Host "Graph API token valid to: $($MSALToken.ExpiresOn)"


Connect-PnPOnline -Url "https://$env:_SharePointDomain" -ClientId $env:_PnPClientId -Tenant $tenantFQDN -CertificateBase64Encoded $certBase64 -ErrorAction Stop

Write-Host "Connected to PNP"

Write-Host "Getting all site collections"
$SitesCollections = Get-PnPTenantSite | Where-Object -Property Template -NotIn ("SRCHCEN#0", "SPSMSITEHOST#0", "APPCATALOG#0", "POINTPUBLISHINGHUB#0", "EDISC#0", "STS#-1")

$null = New-Item -Path .\temp -ItemType Directory -Force
$timeStamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$tempFolder = New-Item -Path ".\temp\$timeStamp" -ItemType Directory


#Loop through each site collection
ForEach ($Site in $SitesCollections) {
    #Connect to site collection
    Write-Host "Generating Report for Site:$($Site.Url)"
    $filename = "$($Site.URL.Replace('https://','').Replace('/','_')).CSV"
    $reportFile = Join-Path -Path $tempFolder.FullName -ChildPath $filename
    Write-Host "Report will be stored temporarily as: $reportFile"

    $SiteConn = Connect-PnPOnline -Url $Site.Url -ClientId $env:_PnPClientId -Tenant $tenantFQDN -CertificateBase64Encoded $certBase64


    #Call the Function for site collection
    Start-SPOPermissionCollection -SiteURL $Site.URL -ReportFile $reportFile -Recursive -ScanItemLevel -GraphApiToken $mgToken -Verbose # -IncludeInheritedPermissions
    Disconnect-PnPOnline -Connection $SiteConn

    $csv = Get-Content -Path $reportFile | Out-String -Width 9999
    $body = $csv
    #$body = ConvertTo-Json -InputObject @{
    #    csv = $csv
    #}

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

    Invoke-RestMethod -Method PUT -Uri $url -Headers $headers -Body $body #-ContentType "application/json"

    #$null = Invoke-RestMethod -Uri $LoginInfo.BlobFunctionKey -Headers @{filename = $filename } -Body $body -ContentType "application/json" -Method POST

    Write-Host "Uploaded file to Blob storage: $reportFile"


}
Remove-Item -Path $tempFolder -Force -Confirm:$false
