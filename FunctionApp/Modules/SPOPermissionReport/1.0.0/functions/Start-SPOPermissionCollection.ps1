function Start-SPOPermissionCollection {
    [cmdletbinding()]
    Param
    (
        [Parameter(Mandatory = $true)] [String] $SiteURL,
        [Parameter(Mandatory = $true)] [String] $ReportFile,
        [Parameter(Mandatory = $false)] [switch] $Recursive,
        [Parameter(Mandatory = $false)] [switch] $ScanItemLevel,
        [Parameter(Mandatory = $false)] [switch] $IncludeInheritedPermissions,
        [Parameter(Mandatory = $true)] [string] $GraphApiToken
    )
    $script:GraphApiToken = $GraphApiToken


    try {
        #Get the Web
        $web = Get-PnPWeb

        #Get Site Collection Administrators
        Write-PSFMessage -Message "Getting Site Collection Administrators"
        $siteCollectionAdmins = Get-PnPSiteCollectionAdmin

        Write-PSFMessage -Message "Site Collection has {0} admins" -StringValues $siteCollectionAdmins.count
        $permissions = foreach ($admin in $siteCollectionAdmins) {
            Write-PSFMessage -Message "Site Collection admin {0} is a {1}" -StringValues $admin.Email, $admin.PrinciaplType
            switch ($admin.PrincipalType) {
                "User" {
                    [PSCustomObject] {
                        [PSCustomObject]@{
                            Object               = "Site Collection"
                            Title                = $web.Title
                            URL                  = $web.URL
                            HasUniquePermissions = "TRUE"
                            Users                = $admin.Title
                            Type                 = "User"
                            Permissions          = "Site Owner"
                            GrantedThrough       = "Direct Permissions"
                        }
                    }
                }
                "SecurityGroup" {
                    $groupMembers = Get-SPOmgGroupTransitiveMember -GroupEmail $admin.Email
                    [PSCustomObject]@{
                        Object               = "Site Collection"
                        Title                = $web.Title
                        URL                  = $web.URL
                        HasUniquePermissions = "TRUE"
                        Users                = $groupMembers -join ","
                        Type                 = "Security Group"
                        Permissions          = "Site Owner"
                        GrantedThrough       = $admin.Email
                    }
                }
                Default {
                    Write-PSFMessage -Level Error -Message "Site Collection Admin for {0} has an unexpected Principal Type: {1}" -StringValues $SiteURL, $admin.PrincipalType
                }
            }
        }

        # Save permissions to CSV
        $permissions | Export-Csv -Path $ReportFile

        Get-PnPWebPermission -Web $Web -ReportFile $ReportFile -Recursive:$Recursive -ScanItemLevel:$ScanItemLevel -IncludeInheritedPermissions:$IncludeInheritedPermissions
        Write-PSFMessage -Message "*** Site Permission Report Generated Successfully!***"

    }
    Catch {
        Write-Error "Error Generating Site Permission Report! $($_.Exception.Message)"
        throw $_
    }

    #Collect Report Data
}