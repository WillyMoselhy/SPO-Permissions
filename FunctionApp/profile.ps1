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
    Write-PSFMessage -Message "Running in Azure environment - Connecting with MSI"
    Disable-AzContextAutosave -Scope Process | Out-Null
    Connect-AzAccount -Identity
}
else {
    # THIS is for offline testing - using a Test SP
    Write-PSFMessage -Message  "Running in local environment - Connecting with Service Principal"
    $spCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $env:LOCAL_ClientId, ($env:LOCAL_ClientSecret | ConvertTo-SecureString -AsPlainText -Force)
    $null = Connect-AzAccount  -ServicePrincipal -Credential $spCreds -Tenant $env:LOCAL_TenantId -WarningAction SilentlyContinue
    Write-PSFMessage -Message "Connected to Azure using service principal"
}

Write-PSFMessage -Message  "Getting Microsoft Graph Token"
Update-SPOPermissionGraphAPIToken

#region: Configure PSFramework logging to Workspace
Write-PSFMessage -Message "Setting Log Analytics for PSFramework"

Set-PSFConfig PSFramework.Logging.Internval 500
Set-PSFConfig PSFramework.Logging.Internval.Idle 500
Start-PSFRunspace psframework.logging -NoMessage

$paramSetPSFLoggingProvider = @{
    Name         = 'AzureLogAnalytics'
    InstanceName = 'SPOPermissions'
    WorkspaceId  = $env:_WorkspaceId
    SharedKey    = $env:_WorkspaceKey
    MaxLevel     = $env:_LogAnalyticsMaxLevel
    Enabled      = $true
}
Set-PSFLoggingProvider @paramSetPSFLoggingProvider
$paramSetPSFLoggingProvider = @{
    Name    = 'Console'
    Enabled = $true
}
Set-PSFLoggingProvider @paramSetPSFLoggingProvider
Start-Sleep -Seconds 1
#endregion

Write-PSFMessage -Message "Profile load complete"
