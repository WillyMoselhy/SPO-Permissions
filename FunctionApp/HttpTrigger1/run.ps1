using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."
# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = 'Function Started successfully. Collection will now begin.'
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
$tenantFQDN = $azTenant.DefaultDomain
$tenantName = $tenantFQDN -replace "(.+)\..+\..+",'$1'
Write-Host "Got tenant information: $tenantId - $tenantName - $tenantFQDN"

$cert = Get-AzKeyVaultSecret -ResourceId $env:PnPPowerShell_KeyVaultId -Name PnPPowerShell -AsPlainText
Write-Host "Cert Obtained from keyvault"

#$MsalToken = Get-MsalToken -ClientId 2e1fee6b-7fe5-48ac-b51a-da35e149f1c5 -ClientCertificate $cert -TenantId $tenantId -ForceRefresh
#Write-Host "Graph API token valid to: $($MSALToken.ExpiresOn)"

#Connect-PnPOnline -Url "https://$($tenantName).Sharepoint.com" -ClientId $LoginInfo.AppID -Tenant $tenantFQDN -CertificatePath $LoginInfo.CertificatePath -ErrorAction Stop
Connect-PnPOnline -Url "https://$($tenantName).Sharepoint.com" -ClientId $LoginInfo.AppID -Tenant $tenantFQDN -CertificateBase64Encoded $cert -ErrorAction Stop

Write-Host "Connected to PNP"

Write-Host "Getting all site collections"
$SitesCollections = Get-PnPTenantSite | Where-Object -Property Template -NotIn ("SRCHCEN#0", "SPSMSITEHOST#0", "APPCATALOG#0", "POINTPUBLISHINGHUB#0", "EDISC#0", "STS#-1")

$null = New-Item -Path .\temp -ItemType Directory -Force
$timeStamp =  Get-Date -Format 'yyyyMMdd-HHmmss'
$tempFolder = New-Item -Path ".\temp\$timeStamp" -ItemType Directory


#Loop through each site collection
ForEach ($Site in $SitesCollections) {
    #Connect to site collection
    Write-Host "Generating Report for Site:$($Site.Url)"
    $filename = "$($Site.URL.Replace('https://','').Replace('/','_')).CSV"
    $reportFile = Join-Path -Path $tempFolder.FullName -ChildPath $filename
    Write-Host "Report will be stored temporarily as: $reportFile"

    $SiteConn = Connect-PnPOnline -Url $Site.Url -ClientId $LoginInfo.AppID -Tenant $azTenant.DefaultDomain -CertificatePath $LoginInfo.CertificatePath


    #Call the Function for site collection
    Start-SPOPermissionCollection -SiteURL $Site.URL -ReportFile $reportFile -Recursive -ScanItemLevel -GraphApiToken $msalToken.AccessToken -Verbose # -IncludeInheritedPermissions
    Disconnect-PnPOnline -Connection $SiteConn

    $csv = Get-Content -Path $reportFile | Out-String -Width 9999

    $body =convertto-json -inputObject @{
        csv = $csv
    }

    $null = Invoke-RestMethod -URI  $LoginInfo.BlobFunctionKey -Headers @{filename = $filename} -Body $body -ContentType "application/json" -Method POST

    Write-Host "Uploaded file to Blob storage: $reportFile"


}
Remove-Item -Path $tempFolder -Force -Confirm:$false
