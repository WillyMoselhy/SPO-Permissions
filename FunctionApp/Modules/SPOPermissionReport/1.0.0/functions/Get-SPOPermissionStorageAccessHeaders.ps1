function Get-SPOPermissionStorageAccessHeaders {
    [CmdletBinding()]
    param (

    )
    if ($env:MSI_SECRET){
        Write-PSFMessage -Message  "Getting Storage Token using MSI"
        $resourceURI = "https://storage.azure.com"
        $tokenAuthURI = $env:IDENTITY_ENDPOINT + "?resource=$resourceURI&api-version=2019-08-01"
        $tokenResponse = Invoke-RestMethod -Method Get -Headers @{"X-IDENTITY-HEADER" = "$env:IDENTITY_HEADER" } -Uri $tokenAuthURI
        $storageToken = $tokenResponse.access_token
    }
    else{ #We are running locally
        Write-PSFMessage -Message  "Getting Storage Token using service principal"
        $body = @{
            'tenant'        = $env:LOCAL_TenantId
            'client_id'     = $env:LOCAL_ClientId
            'scope'         = 'https://storage.azure.com/.default'
            'client_secret' = $env:LOCAL_ClientSecret
            'grant_type'    = 'client_credentials'
        }
        $params = @{
            'Uri'         = "https://login.microsoftonline.com/$env:LOCAL_TenantId/oauth2/v2.0/token"
            'Method'      = 'Post'
            'Body'        = $body
            'ContentType' = 'application/x-www-form-urlencoded'
        }
        $storageToken = (Invoke-RestMethod @Params).access_token
    }

    $headers = @{
        Authorization    = "Bearer $storageToken"
        'x-ms-version'   = '2021-04-10'
        'x-ms-date'      = '{0:R}' -f (Get-Date).ToUniversalTime()
        'x-ms-blob-type' = 'BlockBlob'
    }
    $headers
}