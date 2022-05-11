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



$LoginInfo = [PSCustomObject]@{
    #TenantID        = $env:_TenantID
    #TenantName      = $env:_TenantName
    AppID           = '2e1fee6b-7fe5-48ac-b51a-da35e149f1c5'# $env:_AppID
    CertificatePath = 'C:\home\site\wwwroot\Cert\PnP Rocks2.pfx' #This can be EncodedBase64
    BlobFunctionKey = $env:_SaveBlobFunction
}

$azTenant = Get-AzTenant
$tenantId = $azTenant.TenantId
$tenantFQDN = 'M365x21720695.onmicrosoft.com'  # $azTenant.DefaultDomain
$tenantName = $tenantFQDN -replace "(.+)\..+\..+", '$1'
Write-Host "Got tenant information: $tenantId - $tenantName - $tenantFQDN"

#$cert = Get-AzKeyVaultSecret -ResourceId $env:PnPPowerShell_KeyVaultId -Name PnPPowerShell -AsPlainText
#Write-Host "Cert Obtained from keyvault"

$certBase64 = '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'
$ClientId = '8df3c7d6-2adf-4e42-a549-2f8f665c80e5'

#$MsalToken = Get-MsalToken -ClientId $ClientId  -ClientCertificate $certBase64 -TenantId $tenantId -ForceRefresh
#Write-Host "Graph API token valid to: $($MSALToken.ExpiresOn)"

#Connect-PnPOnline -Url "https://$($tenantName).Sharepoint.com" -ClientId $LoginInfo.AppID -Tenant $tenantFQDN -CertificatePath $LoginInfo.CertificatePath -ErrorAction Stop
Connect-PnPOnline -Url "https://$($tenantName).Sharepoint.com" -ClientId $ClientId -Tenant $tenantFQDN -CertificateBase64Encoded $certBase64 -ErrorAction Stop

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

    $SiteConn = Connect-PnPOnline -Url $Site.Url -ClientId $ClientId -Tenant $tenantFQDN -CertificateBase64Encoded $certBase64


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
        'x-ms-blob-type' = 'BlockBlob '
    }

    $url = "https://safuncspopermissions01.blob.core.windows.net/output/$filename" # TODO: Change this to an environment variable

    Invoke-RestMethod -Method PUT -Uri $url -Headers $headers -Body $body #-ContentType "application/json"

    #$null = Invoke-RestMethod -Uri $LoginInfo.BlobFunctionKey -Headers @{filename = $filename } -Body $body -ContentType "application/json" -Method POST

    Write-Host "Uploaded file to Blob storage: $reportFile"


}
Remove-Item -Path $tempFolder -Force -Confirm:$false
