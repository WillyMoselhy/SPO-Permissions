using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

# Interact with query parameters or the body of the request.


    $psCommand = $Request.Body.PowerShell



if ($psCommand) {
    Write-PSFMessage "Invoking command: $psCommand"
    $body = (Invoke-Expression -Command $psCommand)
}
else{
    $body = "No PowerShell Command passed in request body."
}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
})
