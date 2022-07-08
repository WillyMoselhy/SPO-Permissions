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
            Write-PSFMessage -Message "Site Collection admin {0} is a {1}" -StringValues $admin.Email, $admin.PrincipalType
            switch ($admin.PrincipalType) {
                "User" {
                    [PSCustomObject]@{
                        Object               = "Site Collection"
                        Title                = $web.Title
                        URL                  = $web.URL
                        HasUniquePermissions = "TRUE"
                        Users                = $admin.Title
                        Type                 = "User"
                        Permissions          = "Site Owner"
                        SharePointGroup      = ""
                        GrantedThrough       = "Direct Permissions"
                    }
                }
                "SecurityGroup" {
                    $groupId = $admin.LoginName -replace ".*([\da-zA-Z]{8}-([\da-zA-Z]{4}-){3}[\da-zA-Z]{12}).*",'$1'
                    $mgGroup = Get-SPOmgGroupTransitiveMember -GroupId $groupId
                    if(-Not $mgGroup.DisplayName ){ # Handling non existing groups (Global Administrator / SharePoint Administrator / etc..)
                        $users = $admin.Title
                    }
                    elseif ($mgGroup.Users.Count -eq 0){ # Handling empty groups
                        $users = $mgGroup.DisplayName
                    }
                    else{
                        $users = $mgGroup.Users -join ","
                    }
                    [PSCustomObject]@{
                        Object               = "Site Collection"
                        Title                = $web.Title
                        URL                  = $web.URL
                        HasUniquePermissions = "TRUE"
                        Users                = $users
                        Type                 = 'Security Group'
                        Permissions          = "Site Owner"
                        SharePointGroup      = ""
                        GrantedThrough       = $mgGroup.DisplayName
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