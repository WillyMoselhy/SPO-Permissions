function Get-SPOmgGroupTransitiveMember {
    [CmdletBinding()]
    param (
        [string] $GroupEmail
    )

    # Checking if group members are already cached
    if ($script:Groups.Name -contains $GroupEmail) {
        $groupMembers = ($Script:Groups | Where-Object { $_.Name -eq $GroupEmail -and $_.Type -eq 'SecurityGroup' }).Members
        Write-PSFMessage -Message "Members retrieved from cache"
    }
    else {
        #Get Group
        Write-PSFMessage -Message "Resolving group $GroupEmail"
        $group = Get-MgGroup -Search "Mail:$GroupEmail" -ConsistencyLevel eventual
        $groupMembers = (Get-MgGroupTransitiveMember -GroupId $group.Id -Property displayName).AdditionalProperties.displayName

        # Add members to cache
        $Script:Groups += [PSCustomObject]@{
            Name    = $RoleAssignment.Member.Title
            Type    = 'SecurityGroup'
            Members = $groupMembers
        }
        Write-PSFMessage -Message "Members added to cache"
    }

    #return list of group members
    $groupMembers

}