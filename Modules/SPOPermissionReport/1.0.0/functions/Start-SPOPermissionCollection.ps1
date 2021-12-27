function Start-SPOPermissionCollection {
    [cmdletbinding()]
    Param
    (
        [Parameter(Mandatory = $false)] [String] $SiteURL,
        [Parameter(Mandatory = $false)] [String] $ReportFile,
        [Parameter(Mandatory = $false)] [switch] $Recursive,
        [Parameter(Mandatory = $false)] [switch] $ScanItemLevel,
        [Parameter(Mandatory = $false)] [switch] $IncludeInheritedPermissions
    )
    Write-Host "Hello World!"

    try {
        #Get the Web
        $web = Get-PnPWeb
        Write-Host "Getting Site Collection Administrators..."
        #Get Site Collection Administrators
        $siteAdmins = Get-PnPSiteCollectionAdmin
        $siteCollectionAdmins = ($siteAdmins | Select-Object -ExpandProperty Title) -join ","

        #Add the Data to Object
        $permissions = [PSCustomObject]@{
            Object               = "Site Collection"
            Title                = $web.Title
            URL                  = $web.URL
            HasUniquePermissions = "TRUE"
            Users                = $siteCollectionAdmins
            Type                 = "Site Collection Administrators"
            Permissions          = "Site Owner"
            GrantedThrough       = "Direct Permissions"
        }
        $permissions | ConvertTo-Csv | Out-String -Width 999
    }
    Catch {
        Write-Error "Error Generating Site Permission Report! $($_.Exception.Message)"
        throw $_
    }






    #Collect Report Data
}