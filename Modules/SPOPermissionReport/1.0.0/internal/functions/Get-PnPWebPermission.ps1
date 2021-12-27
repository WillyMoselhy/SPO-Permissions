#Function to Get Webs's Permissions from given URL
Function Get-PnPWebPermission {
    [cmdletbinding()]
    Param
    (
        [Parameter(Mandatory = $true)]  [Microsoft.SharePoint.Client.Web] $Web,
        [Parameter(Mandatory = $true)] [String] $ReportFile,
        [Parameter(Mandatory = $false)] [switch] $Recursive,
        [Parameter(Mandatory = $false)] [switch] $ScanItemLevel,
        [Parameter(Mandatory = $false)] [switch] $IncludeInheritedPermissions
    )
    #Call the function to Get permissions of the web
    Write-PSFMessage -Level Verbose -Message "Getting Permissions of the Web: $($Web.URL)..."
    Get-PnPPermissions -Object $Web -ReportFile $ReportFile

    #Get List Permissions
    Write-PSFMessage -Level Verbose -Message "`t Getting Permissions of Lists and Libraries..."
    Get-PnPListPermission -Web $Web -ReportFile $ReportFile -ScanItemLevel:$ScanItemLevel -IncludeInheritedPermissions:$IncludeInheritedPermissions

    #Recursively get permissions from all sub-webs based on the "Recursive" Switch
    If ($Recursive) {
        #Get Subwebs of the Web
        $Subwebs = Get-PnPProperty -ClientObject $Web -Property Webs

        #Iterate through each subsite in the current web
        Foreach ($Subweb in $web.Webs) {
            #Get Webs with Unique Permissions or Inherited Permissions based on 'IncludeInheritedPermissions' switch
            If ($IncludeInheritedPermissions) {
                Get-PnPWebPermission -Web $SubWeb -ReportFile $ReportFile -Recursive:$Recursive -ScanItemLevel:$ScanItemLevel -IncludeInheritedPermissions:$IncludeInheritedPermissions
            }
            Else {
                #Check if the Web has unique permissions
                $HasUniquePermissions = Get-PnPProperty -ClientObject $SubWeb -Property HasUniqueRoleAssignments

                #Get the Web's Permissions
                If ($HasUniquePermissions -eq $true) {
                    #Call the function recursively
                    Get-PnPWebPermission -Web $SubWeb -ReportFile $ReportFile -Recursive:$Recursive -ScanItemLevel:$ScanItemLevel -IncludeInheritedPermissions:$IncludeInheritedPermissions
                }
            }
        }
    }
}