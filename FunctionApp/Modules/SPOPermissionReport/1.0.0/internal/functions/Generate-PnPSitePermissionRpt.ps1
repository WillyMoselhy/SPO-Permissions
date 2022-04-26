Function Generate-PnPSitePermissionRpt {
    [cmdletbinding()]
    Param
    (
        [Parameter(Mandatory = $false)] [String] $SiteURL,
        [Parameter(Mandatory = $false)] [String] $ReportFile,
        [Parameter(Mandatory = $false)] [switch] $Recursive,
        [Parameter(Mandatory = $false)] [switch] $ScanItemLevel,
        [Parameter(Mandatory = $false)] [switch] $IncludeInheritedPermissions
    )
    Try {
        #Get the Web
        $Web = Get-PnPWeb


        Write-Host   "Getting Site Collection Administrators..."

        #Get Site Collection Administrators
        $SiteAdmins = Get-PnPSiteCollectionAdmin

        $SiteCollectionAdmins = ($SiteAdmins | Select-Object -ExpandProperty Title) -join ","
        #Add the Data to Object
        $Permissions = New-Object PSObject
        $Permissions | Add-Member NoteProperty Object("Site Collection")
        $Permissions | Add-Member NoteProperty Title($Web.Title)
        $Permissions | Add-Member NoteProperty URL($Web.URL)
        $Permissions | Add-Member NoteProperty HasUniquePermissions("TRUE")
        $Permissions | Add-Member NoteProperty Users($SiteCollectionAdmins)
        $Permissions | Add-Member NoteProperty Type("Site Collection Administrators")
        $Permissions | Add-Member NoteProperty Permissions("Site Owner")
        $Permissions | Add-Member NoteProperty GrantedThrough("Direct Permissions")

        #Export Permissions to CSV File
        $Permissions | Export-Csv $ReportFile -NoTypeInformation


        #Call the function with RootWeb to get site collection permissions
        Get-PnPWebPermission -Web $Web -ReportFile $ReportFile -Recursive:$Recursive -ScanItemLevel:$ScanItemLevel -IncludeInheritedPermissions:$IncludeInheritedPermissions


        Write-Host  "*** Site Permission Report Generated Successfully!***"
    }
    Catch {

        Write-Host  "Error Generating Site Permission Report! $($_.Exception.Message)"
        throw $_

    }
}