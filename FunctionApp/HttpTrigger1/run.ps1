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
    TenantID        = $env:_TenantID
    TenantName      = $env:_TenantName
    AppID           = $env:_AppID
    CertificatePath = 'C:\home\site\wwwroot\Cert\PnP Rocks2.pfx' #This can be EncodedBase64
    BlobFunctionKey = $env:_SaveBlobFunction
}


$Cert = new-object security.cryptography.x509certificates.x509certificate2 -ArgumentList $LoginInfo.CertificatePath
Write-Host "Cert Converted"

$MsalToken = Get-MsalToken -ClientId 9ce25227-4018-427e-8f8d-cbc3c0d19657 -ClientCertificate $cert -TenantId 1aeaebf6-dfc4-49c8-a843-cc2b8d54a9b1 -ForceRefresh
Write-Host "Graph API token valid to: $($MSALToken.ExpiresOn)"

Connect-PnPOnline # -Url "https://$($LoginInfo.TenantName).Sharepoint.com" -ClientId $LoginInfo.AppID -Tenant "$($LoginInfo.TenantName).OnMicrosoft.com" -CertificatePath $LoginInfo.CertificatePath -ErrorAction Stop
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
