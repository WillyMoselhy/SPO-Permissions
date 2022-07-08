Function Get-PnPPermissions {
    [cmdletbinding()]
    Param
    (
        [Parameter(Mandatory = $true)] [Microsoft.SharePoint.Client.SecurableObject] $Object,
        [Parameter(Mandatory = $true)] [String] $ReportFile
    )

    #Determine the type of the object
    Write-PSFMessage -Message "Working on a new Object"
    Switch ($Object.TypedObject.ToString()) {
        "Microsoft.SharePoint.Client.Web" {
            $ObjectType = "Site"
            $ObjectURL = $Object.URL
            $ObjectTitle = $Object.Title
        }
        "Microsoft.SharePoint.Client.ListItem" {
            If ($Object.FileSystemObjectType -eq "Folder") {
                $ObjectType = "Folder"
                #Get the URL of the Folder
                Get-PnPProperty -ClientObject $Object -Property Folder #Get-PnPProperty edits web objects.
                $ObjectTitle = $Object.Folder.Name
                $ObjectURL = $("{0}{1}" -f $Web.Url.Replace($Web.ServerRelativeUrl, ''), $Object.Folder.ServerRelativeUrl)
            }
            Else {
                #File or List Item
                #Get the URL of the Object
                Get-PnPProperty -ClientObject $Object -Property File, ParentList
                If ($null -ne $Object.File.Name) {
                    $ObjectType = "File"
                    $ObjectTitle = $Object.File.Name
                    $ObjectURL = $("{0}{1}" -f $Web.Url.Replace($Web.ServerRelativeUrl, ''), $Object.File.ServerRelativeUrl)
                }
                else {

                    $ObjectType = "List Item"
                    $ObjectTitle = $Object["Title"]
                    #Get the URL of the List Item
                    $DefaultDisplayFormUrl = Get-PnPProperty -ClientObject $Object.ParentList -Property DefaultDisplayFormUrl
                    $ObjectURL = $("{0}{1}?ID={2}" -f $Web.Url.Replace($Web.ServerRelativeUrl, ''), $DefaultDisplayFormUrl, $Object.ID)
                }
            }
        }
        Default {
            $ObjectType = "List or Library"
            $ObjectTitle = $Object.Title
            #Get the URL of the List or Library
            $RootFolder = Get-PnPProperty -ClientObject $Object -Property RootFolder
            $ObjectURL = $("{0}{1}" -f $Web.Url.Replace($Web.ServerRelativeUrl, ''), $RootFolder.ServerRelativeUrl)
        }
    }
    Write-PSFMessage -Message "Object is a $ObjectType"
    Write-PSFMessage -Message "Getting permissions for $ObjectURL"

    #Get permissions assigned to the object
    Get-PnPProperty -ClientObject $Object -Property HasUniqueRoleAssignments, RoleAssignments

    #Check if Object has unique permissions
    $HasUniquePermissions = $Object.HasUniqueRoleAssignments

    #Loop through each permission assigned and extract details
    $PermissionCollection = @()
    Write-PSFMessage -Message "Object has $($Object.RoleAssignments.count) permissions"
    $Object.RoleAssignments | ForEach-Object { Get-PnPProperty -ClientObject $_ -Property RoleDefinitionBindings, Member }
    Foreach ($RoleAssignment in $Object.RoleAssignments) {
        #Get the Permission Levels assigned and Member
        #Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member

        #Get the Principal Type: User, SP Group, AD Group
        $PermissionType = $RoleAssignment.Member.PrincipalType

        #Get the Permission Levels assigned
        $PermissionLevels = $RoleAssignment.RoleDefinitionBindings | Select-Object -ExpandProperty Name

        #Remove Limited Access
        $PermissionLevels = ($PermissionLevels | Where-Object { $_ -ne "Limited Access" }) -join ","

        #Leave Principals with no Permissions
        If ($PermissionLevels.Length -eq 0) { Continue }

        # Login Name
        $loginName = $RoleAssignment.Member.LoginName

        # get user or group members
        switch ($PermissionType) {
            "SharePointGroup" {
                # Here we have 2 cases (SharePoint groups cannot be nested!)
                # user is member of SP group
                # user is member of a group that is member of SP group
                #Get Group Members

                Write-PSFMessage -Message "Getting Members of SharePoint group: $($loginName)"
                $groupMembers = Get-SPOSharePointGroupMember -LoginName $loginName

                # Permissions for users in SharePoint Group
                $permissionCollection += [PSCustomObject]@{
                    Object               = $ObjectType
                    Title                = $ObjectTitle
                    URL                  = $ObjectURL
                    HasUniquePermissions = $HasUniquePermissions
                    Users                = $groupMembers.Users -join ","
                    Type                 = $PermissionType
                    Permissions          = $PermissionLevels
                    SharePointGroup      = $RoleAssignment.Member.LoginName
                    GrantedThrough       = ""
                }

                # Permission for Security Groups in SharePoint Group
                foreach ($secGroup in $groupMembers.SecurityGroups) {
                    Write-PSFMessage -Message "Getting Members of Security Group: $($secGroup) for SharePoint Group: $loginName"
                    $mgGroup = Get-SPOmgGroupTransitiveMember -GroupId $secGroup
                    if(-Not $mgGroup.DisplayName ){ # Handling non existing groups (Global Administrator / SharePoint Administrator / etc..)
                        $users = $admin.Title
                    }
                    elseif ($mgGroup.Users.Count -eq 0){ # Handling empty groups
                        $users = $mgGroup.DisplayName
                    }
                    else{
                        $users = $mgGroup.Users -join ","
                    }
                    $permissionCollection += [PSCustomObject]@{
                        Object               = $ObjectType
                        Title                = $ObjectTitle
                        URL                  = $ObjectURL
                        HasUniquePermissions = $HasUniquePermissions
                        Users                = $users
                        Type                 = 'Security Group'
                        Permissions          = $PermissionLevels
                        SharePointGroup      = $RoleAssignment.Member.LoginName
                        GrantedThrough       = $mgGroup.DisplayName
                    }
                }
            }
            "SecurityGroup" {
                Write-PSFMessage -Message "Getting Members of Security Group: $($loginName)"
                $mgGroup = Get-SPOmgGroupTransitiveMember -GroupId $secGroup
                if(-Not $mgGroup.DisplayName ){ # Handling non existing groups (Global Administrator / SharePoint Administrator / etc..)
                    $users = $admin.Title
                }
                elseif ($mgGroup.Users.Count -eq 0){ # Handling empty groups
                    $users = $mgGroup.DisplayName
                }
                else{
                    $users = $mgGroup.Users -join ","
                }
                $permissionCollection += [PSCustomObject]@{
                    Object               = $ObjectType
                    Title                = $ObjectTitle
                    URL                  = $ObjectURL
                    HasUniquePermissions = $HasUniquePermissions
                    Users                = $users
                    Type                 = 'Security Group'
                    Permissions          = $PermissionLevels
                    SharePointGroup      = ""
                    GrantedThrough       = $mgGroup.DisplayName
                }
            }
            Default {
                #Add the Data to Object (most probably a user)
                Write-PSFMessage -Message "Adding permissions for $($RoleAssignment.Member.UserPrincipalName)"
                $permissionCollection += [PSCustomObject]@{
                    Object               = $ObjectType
                    Title                = $ObjectTitle
                    URL                  = $ObjectURL
                    HasUniquePermissions = $HasUniquePermissions
                    Users                = $RoleAssignment.Member.UserPrincipalName
                    Type                 = $PermissionType
                    Permissions          = $PermissionLevels
                    SharePointGroup      = ""
                    GrantedThrough       = 'Direct Permissions'
                }
            }
        }

    }
    #Export Permissions to CSV File
    Write-PSFMessage -Message "Appending Report: $ReportFile"
    $PermissionCollection | Export-Csv $ReportFile -NoTypeInformation -Append
}