function Get-SPOmgGroupTransitiveMember {
    [CmdletBinding()]
    param (
        [string] $GroupId
    )
    # Checking if group members are already cached
    if($script:mgGroups[$GroupId]){
        $mgGroup = $script:mgGroups[$GroupId]
        Write-PSFMessage -Message "MG Group retrieved from cache {0}" -StringValues $mgGroup.DisplayName
    }
    else{
        # TODO: Add support for Service Principals (currently available in mg Beta)
        try{
            $group = Get-MgGroup -GroupId $GroupId -Property displayName -ErrorAction Stop
            $users = Get-MgGroupTransitiveMember -GroupId $GroupId -Property userPrincipalName -All
            if($users.Count -eq 0){
                write-PSFMessage -Level Warning -Message "Group {0} has no members and cached as null" -StringValues $group.DisplayName
            }
        }
        catch{
            Write-PSFMessage -Level Warning -Message "Could not resolve group with id: $GroupId. This could be an Azure AD Role."
            Write-PSFMessage -Level Warning -Message "Group $GroupId is cached as null."
        }

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