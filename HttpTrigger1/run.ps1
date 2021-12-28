using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

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
    TenantID        = '1aeaebf6-dfc4-49c8-a843-cc2b8d54a9b1'
    TenantName      = 'm365x252065'
    AppID           = '9ce25227-4018-427e-8f8d-cbc3c0d19657'
    CertificatePath = 'C:\home\site\wwwroot\Cert\PnP Rocks2.pfx' #This can be EncodedBase64
    BlobFunctionKey = 'https://saveblobfile.azurewebsites.net/api/HttpTrigger1?code=Sc2Cq8SCuWEC/7oBY0oVPqygpAwMILqXxPws2bOeXDmQzh5MavtcfA=='
}


$Cert = new-object security.cryptography.x509certificates.x509certificate2 -ArgumentList $LoginInfo.CertificatePath
Write-Host "Cert Converted"

$script:MSALToken = Get-MsalToken -ClientId 9ce25227-4018-427e-8f8d-cbc3c0d19657 -ClientCertificate $cert -TenantId 1aeaebf6-dfc4-49c8-a843-cc2b8d54a9b1 -ForceRefresh
Write-Host "Graph API token valid to: $($script:MSALToken.ExpiresOn)"

Connect-PnPOnline -Url "https://$($LoginInfo.TenantName).Sharepoint.com" -ClientId $LoginInfo.AppID -Tenant "$($LoginInfo.TenantName).OnMicrosoft.com" -CertificatePath $LoginInfo.CertificatePath -ErrorAction Stop
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

    $SiteConn = Connect-PnPOnline -Url $Site.Url -ClientId $LoginInfo.AppID -Tenant "$($LoginInfo.TenantName).OnMicrosoft.com" -CertificatePath $LoginInfo.CertificatePath


    #Call the Function for site collection
    Start-SPOPermissionCollection -SiteURL $Site.URL -ReportFile $reportFile -Recursive -ScanItemLevel -BlobFunctionKey $LoginInfo.BlobFunctionKey -Verbose # -IncludeInheritedPermissions
    Disconnect-PnPOnline -Connection $SiteConn

    $csv = Get-Content -Path $reportFile | Out-String -Width 9999

    $null = Invoke-Webrequest -URI  $LoginInfo.BlobFunctionKey -Headers @{filename = $filename} -Body @{CSV = $csv }
    Write-Host "Uploaded file to Blob storage: $reportFile"


}
Remove-Item -Path $tempFolder




Write-Host "Finished in: $duration"


if ($name) {
    $body = "Hello, $name. This HTTP triggered function executed successfully."
}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
})