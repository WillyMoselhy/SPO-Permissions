function Get-GApiGroupMember {
    [CmdletBinding()]
    param (
        [string] $GroupName
    )

    begin {
        $Headers = @{
            'Authorization'    = "Bearer $($Script:GroupMemberToken)"
            'ConsistencyLevel' = 'Eventual'
        }
    }

    process {
        #Get Group ID
        #$name = 'Ask HR Members' -replace '\sMembers$',''
        $requestUri = 'https://graph.microsoft.com/v1.0/groups?$search="displayName:{0}"' -f $GroupName
        $groupId = (Invoke-RestMethod -Uri $requestUri -Header $headers -Method GET -ErrorAction Stop).value.id

        #Get Group Members
        $groupMembersURI = 'https://graph.microsoft.com/v1.0/groups/{0}/transitiveMembers?$select=displayName' -f $groupId
        $groupMembers = Invoke-RestMethod -Uri $groupMembersURI -Header $headers -Method GET -ErrorAction Stop
        $groupMembers.value.displayName
    }

}