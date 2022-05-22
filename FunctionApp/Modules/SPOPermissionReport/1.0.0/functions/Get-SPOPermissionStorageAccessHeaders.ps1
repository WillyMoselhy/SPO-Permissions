function Get-SPOPermissionStorageAccessHeaders {
    [CmdletBinding()]
    param (
        
    )
    Write-PSFMessage -Message  "Getting Storage Token"
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
    $headers  
}