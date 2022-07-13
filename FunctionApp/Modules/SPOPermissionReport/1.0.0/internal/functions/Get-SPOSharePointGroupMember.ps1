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

        $users = @()
        $securityGroups = @()
        foreach ($member in $pnpGroupMembers){
            switch -Wildcard ($member.LoginName) {
                "*|membership|*" {
                    $users += $member.UserPrincipalName
                }
                "*|federateddirectoryclaimprovider|*" {
                    $securityGroups += $member.LoginName -replace ".*([\da-zA-Z]{8}-([\da-zA-Z]{4}-){3}[\da-zA-Z]{12}).*", '$1'
                }
                "*|tenant|*" {
                    $users += $member.Title # TODO: Use Get-mgDirectoryRoleMember to expand AAD Roles like Global Admin
                }
                Default {
                    $users += $member.Title
                }
            }
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