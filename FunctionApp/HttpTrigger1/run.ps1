using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)


Write-PSFMessage -Message  "PowerShell HTTP trigger function processed a request."


$targetURL = $Request.Query.URL
if (-not $targetURL) {
    $targetURL = $Request.Body.URL
}

if($targetURL) {
    Write-PSFMessage -Message "Got request to scan: $targetURL"
    .\ScanSiteCollection.ps1 -SiteCollectionURL $targetURL
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body       = "Scan complete for $targetURL"
    })    
}
else{
    $body = "No URL defind in query or body."
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::BadRequest
        Body       = $body
    })    
    
    Stop-PSFFunction -Message $body -EnableException
}