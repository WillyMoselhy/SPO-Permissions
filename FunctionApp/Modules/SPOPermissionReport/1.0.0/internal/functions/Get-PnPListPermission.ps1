#Function to Get Permissions of all lists from the given web
Function Get-PnPListPermission() {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]  [Microsoft.SharePoint.Client.Web]$Web,
        [Parameter(Mandatory = $true)]  [String] $ReportFile,
        [Parameter(Mandatory = $false)] [switch] $ScanItemLevel,
        [Parameter(Mandatory = $false)] [switch] $IncludeInheritedPermissions
    )

    #Get All Lists from the web
    $Lists = Get-PnPProperty -ClientObject $Web -Property Lists

    #Exclude system lists
    $ExcludedLists = @("Access Requests", "App Packages", "appdata", "appfiles", "Apps in Testing", "Cache Profiles", "Composed Looks", "Content and Structure Reports", "Content type publishing error log", "Converted Forms",
        "Device Channels", "Form Templates", "fpdatasources", "Get started with Apps for Office and SharePoint", "List Template Gallery", "Long Running Operation Status", "Maintenance Log Library", "Images", "site collection images"
        , "Master Docs", "Master Page Gallery", "MicroFeed", "NintexFormXml", "Quick Deploy Items", "Relationships List", "Reusable Content", "Reporting Metadata", "Reporting Templates", "Search Config List", "Site Assets", "Preservation Hold Library",
        "Site Pages", "Solution Gallery", "Style Library", "Suggested Content Browser Locations", "Theme Gallery", "TaxonomyHiddenList", "User Information List", "Web Part Gallery", "wfpub", "wfsvc", "Workflow History", "Workflow Tasks", "Pages")

    #Get all lists from the web
    ForEach ($List in ($Lists | Where-Object { -Not $_.Hidden -and $_.Title -notin $ExcludedLists })) {

        #Get Item Level Permissions if 'ScanItemLevel' switch present
        If ($ScanItemLevel) {
            #Get List Items Permissions
            Get-PnPListItemsPermission -ReportFile $ReportFile -List $List -IncludeInheritedPermissions:$IncludeInheritedPermissions
        }

        #Get Lists with Unique Permissions or Inherited Permissions based on 'IncludeInheritedPermissions' switch
        If ($IncludeInheritedPermissions) {
            Get-PnPPermissions -Object $List -ReportFile $ReportFile
        }
        Else {
            #Check if List has unique permissions
            $HasUniquePermissions = Get-PnPProperty -ClientObject $List -Property HasUniqueRoleAssignments
            If ($HasUniquePermissions -eq $True) {
                #Call the function to check permissions
                Get-PnPPermissions -Object $List -ReportFile $ReportFile
            }
        }
    }
}