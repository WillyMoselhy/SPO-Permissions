function Get-SPOSharePointGroupMember {
    [CmdletBinding()]
    param (
        [string] $LoginName
    )


    if ($script:SharePointGroups[$LoginName]) {
        # Group found in cache
        $groupMembers = $script:SharePointGroups[$LoginName]
        Write-PSFMessage -Message "Members retrieved from cache"
    }
    else {
        # Group is not cached
        $pnpGroupMembers = Get-PnPGroupMember -Group $LoginName

        $users = ($pnpGroupMembers | Where-Object { $_.PrincipalType -eq 'User' }).UserPrincipalName
        if($pnpGroupMembers | Where-Object { $_.PrincipalType -eq 'SecurityGroup' }){
            $securityGroups = ($pnpGroupMembers | Where-Object { $_.PrincipalType -eq 'SecurityGroup' }).LoginName -replace ".*([\da-zA-Z]{8}-([\da-zA-Z]{4}-){3}[\da-zA-Z]{12}).*", '$1'
        }
        else{
            $securityGroups = $null
        }

        $groupMembers = [PSCustomObject]@{
            # For users we are interested in UserPrincipalName
            Users          = $users

            # For groups, we take the GUID, this will be used to get the members from MS Graph
            SecurityGroups =  $securityGroups
        }
        Write-PSFMessage -Message "Members added to cache. {0} Users and {1} Security Groups" -StringValues $groupMembers.Users.Count, $groupMembers.SecurityGroups.Count
        #Update cache
        $script:SharePointGroups[$LoginName] = $groupMembers
    }
    #return list of group members

    $groupMembers

}