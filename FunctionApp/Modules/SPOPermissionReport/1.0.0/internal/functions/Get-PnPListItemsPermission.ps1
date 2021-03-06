#Function to Get Permissions of All List Items of a given List

Function Get-PnPListItemsPermission {
    [cmdletbinding()]
    Param
    (
        [Parameter(Mandatory = $true)]  [Microsoft.SharePoint.Client.List] $List,
        [Parameter(Mandatory = $true)] [String] $ReportFile,
        [Parameter(Mandatory = $false)] [switch] $IncludeInheritedPermissions

    )
    Write-PSFMessage -Message "`t `t Getting Permissions of List Items in the List:$($List.Title)"

    #Get All Items from List in batches
    $ListItems = Get-PnPListItem -List $List


    #Loop through each List item
    ForEach ($ListItem in $ListItems) {
        #Get Objects with Unique Permissions or Inherited Permissions based on 'IncludeInheritedPermissions' switch
        If ($IncludeInheritedPermissions) {
            Get-PnPPermissions -Object $ListItem -ReportFile $ReportFile
        }
        Else {
            #Check if List Item has unique permissions
            $HasUniquePermissions = Get-PnPProperty -ClientObject $ListItem -Property HasUniqueRoleAssignments
            If ($HasUniquePermissions -eq $True) {
                #Call the function to generate Permission report
                Get-PnPPermissions -Object $ListItem -ReportFile $ReportFile
            }
        }
    }
}