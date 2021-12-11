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

#$duration = Measure-Command -Expression {$result = Start-SPOPermissionCollection}

$LoginInfo = [PSCustomObject]@{
    TenantID        = '1aeaebf6-dfc4-49c8-a843-cc2b8d54a9b1'
    TenantName      = 'm365x252065'
    AppID           = '9ce25227-4018-427e-8f8d-cbc3c0d19657'
    CertificatePath = 'C:\home\site\wwwroot\Cert\PnP Rocks2.pfx' #This can be EncodedBase64
}


$Cert = new-object security.cryptography.x509certificates.x509certificate2 -ArgumentList $LoginInfo.CertificatePath
write-host "Cert Converted"

$script:MSALToken = Get-MsalToken -ClientId 9ce25227-4018-427e-8f8d-cbc3c0d19657 -ClientCertificate $cert -TenantId 1aeaebf6-dfc4-49c8-a843-cc2b8d54a9b1 -ForceRefresh
Write-Host $script:MSALToken.AccessToken
Connect-PnPOnline -Url "https://$($LoginInfo.TenantName).Sharepoint.com" -ClientId $LoginInfo.AppID -Tenant "$($LoginInfo.TenantName).OnMicrosoft.com" -CertificatePath $LoginInfo.CertificatePath -ErrorAction Stop
write-host "Connected to PNP"




Write-Host "Finished! That took only $duration"


if ($name) {
    $body = "Hello, $name. This HTTP triggered function executed successfully."
}


# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
})