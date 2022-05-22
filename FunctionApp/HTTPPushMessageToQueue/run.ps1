using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-PSFMessage -Message  "PowerShell HTTP trigger function processed a request."
#region: Check if we are targeting specific URL(s) from query or body
$targetURLs = $Request.Query.URL
if (-not $targetURLs) {
    $targetURLs = $Request.Body.URL
}

if($targetURLs){
    $result = .\PushMessageToQueue.ps1 -URL $targetURLs
}
else{
    $result = .\PushMessageToQueue.ps1 
}


Push-OutputBinding -Name Response -Value $result 

Write-PSFMessage -Level Host -Message "Function executed without errors"