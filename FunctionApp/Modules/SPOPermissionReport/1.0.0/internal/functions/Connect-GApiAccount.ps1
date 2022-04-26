function Connect-GApiAccount {
    [CmdletBinding()]
    param (
        [string] $TenantId,
        [string] $ClientId,
        [string] $ClientSecret


    )
    $body = @{
        'tenant'        = $TenantId
        'client_id'     = $ClientId
        'scope'         = 'https://graph.microsoft.com/.default'
        'client_secret' = $ClientSecret
        'grant_type'    = 'client_credentials'
    }

    $params = @{
        'Uri'         = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
        'Method'      = 'Post'
        'Body'        = $Body
        'ContentType' = 'application/x-www-form-urlencoded'
    }

    $authResponse = Invoke-RestMethod @Params

    $Script:GraphApiToken = $authResponse.access_token
}