using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)


Write-PSFMessage -Message "PowerShell HTTP trigger function processed a request."


$targetURL = $Request.Query.URL
if (-not $targetURL) {
    $targetURL = $Request.Body.URL
}

if (-not $targetURL) {
    $body = "No URL defined in query or body."
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::BadRequest
            Body       = $body
        })

    Stop-PSFFunction -Message $body -EnableException $true
}
Write-PSFMessage -Message "Got request to scan: $targetURL"

try{
    .\ScanSiteCollection.ps1 -SiteCollectionURL $targetURL
}
catch {
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::InternalServerError
        Body       = $_
    })
    return
}


Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body       = "Scan complete for $targetURL"
    })
