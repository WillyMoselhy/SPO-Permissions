function Get-SPOmgGroupTransitiveMember {
    [CmdletBinding()]
    param (
        [string] $GroupId
    )

    # Checking if group members are already cached
    if($script:mgGroups[$GroupId]){
        $mgGroup = $script:mgGroups[$GroupId]
        Write-PSFMessage -Message "MG Group retrieved from cache"
    }
    else{
        $group = Get-MgGroup -GroupId $GroupId -Property displayName
        $users = Get-MgGroupTransitiveMember -GroupId $GroupId -Property userPrincipalName
        $mgGroup = [PSCustomObject]@{
            DisplayName = $group.DisplayName
            Users = $users.AdditionalProperties.userPrincipalName
        }

        # Add to cache
        $script:mgGroups[$GroupId] = $mgGroup
        Write-PSFMessage -Message "Members added to cache: {0} - {1} users" -StringValues $group.DisplayName,$users.count
    }
    #return list of group members
    $mgGroup
}