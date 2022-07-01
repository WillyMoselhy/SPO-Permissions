function Update-SPOPermissionGraphAPIToken {
    <#
    .SYNOPSIS
        This function should check for the Graph API token validity and (re)connect to Microsoft Graph
    #>
    [CmdletBinding()]
    param (

    )
    $currentTime = Get-Date
    if ($env:mgTokenExpiryTimeStamp) {
        $timeToExpire = New-TimeSpan -Start $currentTime -End $env:mgTokenExpiryTimeStamp
    }
    else {
        $timeToExpire = New-TimeSpan -Start $currentTime -End $currentTime
    }

    if ($timeToExpire.TotalSeconds -le 0) {
        # Token has expired and should be renewed
        Write-PSFMessage -Message "Graph API token expired. Getting new token."
        if ($env:MSI_SECRET) {
            Write-PSFMessage -Message "Getting Microsoft Graph Token as MSI"
            $resourceURI = "https://graph.microsoft.com"
            $tokenAuthURI = $env:IDENTITY_ENDPOINT + "?resource=$resourceURI&api-version=2019-08-01"
            $tokenResponse = Invoke-RestMethod -Method Get -Headers @{"X-IDENTITY-HEADER" = "$env:IDENTITY_HEADER" } -Uri $tokenAuthURI
            $env:mgToken = $tokenResponse.access_token
        }
        else {
            #We are running locally
            Write-PSFMessage -Message "Getting Microsoft Graph Token as SP"
            $body = @{
                'tenant'        = $env:LOCAL_TenantId
                'client_id'     = $env:LOCAL_ClientId
                'scope'         = 'https://graph.microsoft.com/.default'
                'client_secret' = $env:LOCAL_ClientSecret
                'grant_type'    = 'client_credentials'
            }
            # Assemble a hashtable for splatting parameters, for readability
            # The tenant id is used in the uri of the request as well as the body
            $Params = @{

                'Uri'         = "https://login.microsoftonline.com/$env:LOCAL_TenantId/oauth2/v2.0/token"
                'Method'      = 'Post'
                'Body'        = $body
                'ContentType' = 'application/x-www-form-urlencoded'
            }

            $env:mgToken = (Invoke-RestMethod @Params).access_token #TODO Remove this variable once we move to mgGraph completely

            Write-PSFMessage -Message "Connecting to Graph API using new token"
            Connect-MgGraph -AccessToken $env:mgToken
        }
        $env:mgTokenExpiryTimeStamp = $currentTime.AddSeconds(1800) # We expire after 30 minutes to keep things fresh.
    }
    else {
        Write-PSFMessage -Message "Graph API token is not expired"
    }
    Write-PSFMessage -Message "Graph API token will expire in {0:N0} seconds" -StringValues $timeToExpire.TotalSeconds
}