# Azure Functions profile.ps1
#
# This profile.ps1 will get executed every "cold start" of your Function App.
# "cold start" occurs when:
#
# * A Function App starts up for the very first time
# * A Function App starts up after being de-allocated due to inactivity
#
# You can define helper functions, run commands, or specify environment variables
# NOTE: any variables defined that are not environment variables will get reset after the first execution

# Authenticate with Azure PowerShell using MSI.
# Remove this if you are not planning on using MSI or Azure PowerShell.
if ($env:MSI_SECRET) {
    Disable-AzContextAutosave -Scope Process | Out-Null
    Connect-AzAccount -Identity

    Write-PSFMessage -Message  "Getting Token as MSI"
    Write-PSFMessage -Message  "Getting Microsoft Graph Token"
    $resourceURI = "https://graph.microsoft.com"
    $tokenAuthURI = $env:IDENTITY_ENDPOINT + "?resource=$resourceURI&api-version=2019-08-01"
    $tokenResponse = Invoke-RestMethod -Method Get -Headers @{"X-IDENTITY-HEADER" = "$env:IDENTITY_HEADER" } -Uri $tokenAuthURI
    $env:mgToken = $tokenResponse.access_token
}
else {
    # THIS is for offline testing - using a Test SP
}

#region: Configure PSFramework logging to Workspace
Set-PSFConfig PSFramework.Logging.Internval 500
Set-PSFConfig PSFramework.Logging.Internval.Idle 500
Start-PSFRunspace psframework.logging

$paramSetPSFLoggingProvider = @{
    Name         = 'AzureLogAnalytics'
    InstanceName = 'SPOPermissions'
    WorkspaceId  = $env:_WorkspaceId
    SharedKey    = $env:_WorkspaceKey
    MaxLevel     = $enc:_LogAnalyticsMaxLevel
    Enabled      = $true
}
Set-PSFLoggingProvider @$paramSetPSFLoggingProvider
$paramSetPSFLoggingProvider = @{
    Name    = 'Console'
    Enabled = $true
}
Set-PSFLoggingProvider @$paramSetPSFLoggingProvider
Start-Sleep -Seconds 1
#endregion

Write-PSFMessage -Message "Profile load complete"


