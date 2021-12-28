function Start-SPOPermissionCollection {
    [cmdletbinding()]
    Param
    (
        [Parameter(Mandatory = $false)] [String] $SiteURL,
        [Parameter(Mandatory = $false)] [String] $ReportFile,
        [Parameter(Mandatory = $false)] [switch] $Recursive,
        [Parameter(Mandatory = $false)] [switch] $ScanItemLevel,
        [Parameter(Mandatory = $false)] [switch] $IncludeInheritedPermissions,
        [Parameter(Mandatory = $false)] [string] $BlobFunctionKey
    )
    Write-Host "Getting all site collections"
    $SitesCollections = Get-PnPTenantSite | Where-Object -Property Template -NotIn ("SRCHCEN#0", "SPSMSITEHOST#0", "APPCATALOG#0", "POINTPUBLISHINGHUB#0", "EDISC#0", "STS#-1")

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

        $permissions | Export-Csv -Path $ReportFile

        Get-PnPWebPermission -Web $Web -ReportFile $ReportFile -Recursive:$Recursive -ScanItemLevel:$ScanItemLevel -IncludeInheritedPermissions:$IncludeInheritedPermissions
        Write-Host "*** Site Permission Report Generated Successfully!***"

    }
    Catch {
        Write-Error "Error Generating Site Permission Report! $($_.Exception.Message)"
        throw $_
    }






    #Collect Report Data
}