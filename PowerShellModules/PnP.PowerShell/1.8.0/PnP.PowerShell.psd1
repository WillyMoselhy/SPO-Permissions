@{
	NestedModules =  if ($PSEdition -eq 'Core')
	{
		'Core/PnP.PowerShell.dll'
	}
	else
	{
		'Framework/PnP.PowerShell.dll'
	}
	ModuleVersion = '1.8.0'
	Description = 'Microsoft 365 Patterns and Practices PowerShell Cmdlets'
	GUID = '0b0430ce-d799-4f3b-a565-f0dca1f31e17'
	Author = 'Microsoft 365 Patterns and Practices'
	CompanyName = 'Microsoft 365 Patterns and Practices'
	CompatiblePSEditions = @("Core","Desktop")
	PowerShellVersion = '5.1'
	DotNetFrameworkVersion = '4.6.1'
	ProcessorArchitecture = 'None'
	FunctionsToExport = '*'  
	CmdletsToExport = @("Add-PnPClientSidePage","Add-PnPClientSidePageSection","Add-PnPClientSideText","Add-PnPClientSideWebPart","Export-PnPClientSidePage","Export-PnPClientSidePageMapping","Get-PnPAADUser","Get-PnPAvailableClientSideComponents","Get-PnPClientSideComponent","Get-PnPClientSidePage","Get-PnPFlowEnvironment","Get-PnPMicrosoft365GroupMembers","Get-PnPMicrosoft365GroupOwners","Get-PnPSubWebs","Invoke-PnPSearchQuery","Move-PnPClientSideComponent","Remove-PnPClientSideComponent","Remove-PnPClientSidePage","Save-PnPClientSidePageConversionLog","Set-PnPClientSidePage","Set-PnPClientSideText","Set-PnPClientSideWebPart","Add-PnPAlert","Add-PnPApp","Add-PnPApplicationCustomizer","Add-PnPAzureADGroupMember","Add-PnPAzureADGroupOwner","Add-PnPContentType","Add-PnPContentTypesFromContentTypeHub","Add-PnPContentTypeToDocumentSet","Add-PnPContentTypeToList","Add-PnPCustomAction","Add-PnPDataRowsToSiteTemplate","Add-PnPDocumentSet","Add-PnPEventReceiver","Add-PnPField","Add-PnPFieldFromXml","Add-PnPFieldToContentType","Add-PnPFile","Add-PnPFileToSiteTemplate","Add-PnPFolder","Add-PnPGroupMember","Add-PnPHtmlPublishingPageLayout","Add-PnPHubSiteAssociation","Add-PnPHubToHubAssociation","Add-PnPIndexedProperty","Add-PnPJavaScriptBlock","Add-PnPJavaScriptLink","Add-PnPListFoldersToSiteTemplate","Add-PnPListItem","Add-PnPMasterPage","Add-PnPMicrosoft365GroupMember","Add-PnPMicrosoft365GroupOwner","Add-PnPMicrosoft365GroupToSite","Add-PnPNavigationNode","Add-PnPOrgAssetsLibrary","Add-PnPOrgNewsSite","Add-PnPPage","Add-PnPPageSection","Add-PnPPageTextPart","Add-PnPPageWebPart","Add-PnPPlannerBucket","Add-PnPPlannerTask","Add-PnPPublishingImageRendition","Add-PnPPublishingPage","Add-PnPPublishingPageLayout","Add-PnPRoleDefinition","Add-PnPSiteClassification","Add-PnPSiteCollectionAdmin","Add-PnPSiteCollectionAppCatalog","Add-PnPSiteDesign","Add-PnPSiteDesignFromWeb","Add-PnPSiteDesignTask","Add-PnPSiteScript","Add-PnPSiteScriptPackage","Add-PnPSiteTemplate","Add-PnPStoredCredential","Add-PnPTaxonomyField","Add-PnPTeamsChannel","Add-PnPTeamsTab","Add-PnPTeamsTeam","Add-PnPTeamsUser","Add-PnPTenantCdnOrigin","Add-PnPTenantSequence","Add-PnPTenantSequenceSite","Add-PnPTenantSequenceSubSite","Add-PnPTenantTheme","Add-PnPTermToTerm","Add-PnPView","Add-PnPWebhookSubscription","Add-PnPWebPartToWebPartPage","Add-PnPWebPartToWikiPage","Add-PnPWikiPage","Approve-PnPTenantServicePrincipalPermissionRequest","Clear-PnPAzureADGroupMember","Clear-PnPAzureADGroupOwner","Clear-PnPDefaultColumnValues","Clear-PnPListItemAsRecord","Clear-PnPMicrosoft365GroupMember","Clear-PnPMicrosoft365GroupOwner","Clear-PnPRecycleBinItem","Clear-PnPTenantAppCatalogUrl","Clear-PnPTenantRecycleBinItem","Connect-PnPOnline","Convert-PnPFolderToSiteTemplate","Convert-PnPSiteTemplate","Convert-PnPSiteTemplateToMarkdown","ConvertTo-PnPPage","Copy-ItemProxy","Copy-PnPFile","Deny-PnPTenantServicePrincipalPermissionRequest","Disable-PnPFeature","Disable-PnPFlow","Disable-PnPSharingForNonOwnersOfSite","Disable-PnPSiteClassification","Disable-PnPTenantServicePrincipal","Disconnect-PnPOnline","Enable-PnPCommSite","Enable-PnPFeature","Enable-PnPFlow","Enable-PnPSiteClassification","Enable-PnPTenantServicePrincipal","Export-PnPFlow","Export-PnPListToSiteTemplate","Export-PnPPage","Export-PnPPageMapping","Export-PnPTaxonomy","Export-PnPTermGroupToXml","Export-PnPUserInfo","Export-PnPUserProfile","Find-PnPFile","Get-PnPAccessToken","Get-PnPAlert","Get-PnPApp","Get-PnPAppAuthAccessToken","Get-PnPAppErrors","Get-PnPAppInfo","Get-PnPApplicationCustomizer","Get-PnPAuditing","Get-PnPAuthenticationRealm","Get-PnPAvailableLanguage","Get-PnPAvailablePageComponents","Get-PnPAzureADApp","Get-PnPAzureADAppPermission","Get-PnPAzureADAppSitePermission","Get-PnPAzureADGroup","Get-PnPAzureADGroupMember","Get-PnPAzureADGroupOwner","Get-PnPAzureADUser","Get-PnPAzureCertificate","Get-PnPBrowserIdleSignout","Get-PnPBuiltInDesignPackageVisibility","Get-PnPChangeLog","Get-PnPConnection","Get-PnPContentType","Get-PnPContentTypePublishingHubUrl","Get-PnPContext","Get-PnPCustomAction","Get-PnPDefaultColumnValues","Get-PnPDeletedMicrosoft365Group","Get-PnPDiagnostics","Get-PnPDisableSpacesActivation","Get-PnPDocumentSetTemplate","Get-PnPEventReceiver","Get-PnPException","Get-PnPExternalUser","Get-PnPFeature","Get-PnPField","Get-PnPFile","Get-PnPFileVersion","Get-PnPFlow","Get-PnPFolder","Get-PnPFolderItem","Get-PnPFooter","Get-PnPGraphAccessToken","Get-PnPGraphSubscription","Get-PnPGroup","Get-PnPGroupMember","Get-PnPGroupPermissions","Get-PnPHideDefaultThemes","Get-PnPHomePage","Get-PnPHomeSite","Get-PnPHubSite","Get-PnPHubSiteChild","Get-PnPIndexedPropertyKeys","Get-PnPInPlaceRecordsManagement","Get-PnPIsSiteAliasAvailable","Get-PnPJavaScriptLink","Get-PnPKnowledgeHubSite","Get-PnPLabel","Get-PnPList","Get-PnPListInformationRightsManagement","Get-PnPListItem","Get-PnPListPermissions","Get-PnPListRecordDeclaration","Get-PnPMasterPage","Get-PnPMessageCenterAnnouncement","Get-PnPMicrosoft365Group","Get-PnPMicrosoft365GroupMember","Get-PnPMicrosoft365GroupOwner","Get-PnPNavigationNode","Get-PnPOffice365CurrentServiceStatus","Get-PnPOffice365HistoricalServiceStatus","Get-PnPOffice365ServiceMessage","Get-PnPOffice365Services","Get-PnPOrgAssetsLibrary","Get-PnPOrgNewsSite","Get-PnPPage","Get-PnPPageComponent","Get-PnPPlannerBucket","Get-PnPPlannerPlan","Get-PnPPlannerTask","Get-PnPPowerPlatformEnvironment","Get-PnPPowerShellTelemetryEnabled","Get-PnPProperty","Get-PnPPropertyBag","Get-PnPPublishingImageRendition","Get-PnPRecycleBinItem","Get-PnPRequestAccessEmails","Get-PnPRoleDefinition","Get-PnPSearchConfiguration","Get-PnPSearchCrawlLog","Get-PnPSearchSettings","Get-PnPServiceCurrentHealth","Get-PnPServiceHealthIssue","Get-PnPServiceUpdateMessage","Get-PnPSharingForNonOwnersOfSite","Get-PnPSite","Get-PnPSiteClassification","Get-PnPSiteClosure","Get-PnPSiteCollectionAdmin","Get-PnPSiteCollectionAppCatalogs","Get-PnPSiteCollectionTermStore","Get-PnPSiteDesign","Get-PnPSiteDesignRights","Get-PnPSiteDesignRun","Get-PnPSiteDesignRunStatus","Get-PnPSiteDesignTask","Get-PnPSiteGroup","Get-PnPSitePolicy","Get-PnPSiteScript","Get-PnPSiteScriptFromList","Get-PnPSiteScriptFromWeb","Get-PnPSiteSearchQueryResults","Get-PnPSiteTemplate","Get-PnPSiteUserInvitations","Get-PnPStorageEntity","Get-PnPStoredCredential","Get-PnPStructuralNavigationCacheSiteState","Get-PnPStructuralNavigationCacheWebState","Get-PnPSubscribeSharePointNewsDigest","Get-PnPSubWeb","Get-PnPSyntexModel","Get-PnPSyntexModelPublication","Get-PnPTaxonomyItem","Get-PnPTaxonomySession","Get-PnPTeamsApp","Get-PnPTeamsChannel","Get-PnPTeamsChannelMessage","Get-PnPTeamsTab","Get-PnPTeamsTeam","Get-PnPTeamsUser","Get-PnPTemporarilyDisableAppBar","Get-PnPTenant","Get-PnPTenantAppCatalogUrl","Get-PnPTenantCdnEnabled","Get-PnPTenantCdnOrigin","Get-PnPTenantCdnPolicies","Get-PnPTenantDeletedSite","Get-PnPTenantId","Get-PnPTenantRecycleBinItem","Get-PnPTenantSequence","Get-PnPTenantSequenceSite","Get-PnPTenantServicePrincipal","Get-PnPTenantServicePrincipalPermissionGrants","Get-PnPTenantServicePrincipalPermissionRequests","Get-PnPTenantSite","Get-PnPTenantSyncClientRestriction","Get-PnPTenantTemplate","Get-PnPTenantTheme","Get-PnPTerm","Get-PnPTermGroup","Get-PnPTermLabel","Get-PnPTermSet","Get-PnPTheme","Get-PnPTimeZoneId","Get-PnPUnifiedAuditLog","Get-PnPUPABulkImportStatus","Get-PnPUser","Get-PnPUserOneDriveQuota","Get-PnPUserProfileProperty","Get-PnPView","Get-PnPWeb","Get-PnPWebhookSubscriptions","Get-PnPWebPart","Get-PnPWebPartProperty","Get-PnPWebPartXml","Get-PnPWebTemplates","Get-PnPWikiPageContent","Grant-PnPAzureADAppSitePermission","Grant-PnPHubSiteRights","Grant-PnPSiteDesignRights","Grant-PnPTenantServicePrincipalPermission","Import-PnPTaxonomy","Import-PnPTermGroupFromXml","Import-PnPTermSet","Install-PnPApp","Invoke-PnPBatch","Invoke-PnPQuery","Invoke-PnPSiteDesign","Invoke-PnPSiteSwap","Invoke-PnPSiteTemplate","Invoke-PnPSPRestMethod","Invoke-PnPTenantTemplate","Invoke-PnPTransformation","Invoke-PnPWebAction","Measure-PnPList","Measure-PnPWeb","Move-ItemProxy","Move-PnPFile","Move-PnPFolder","Move-PnPListItemToRecycleBin","Move-PnPPageComponent","Move-PnPRecycleBinItem","New-PnPAzureADGroup","New-PnPAzureCertificate","New-PnPBatch","New-PnPExtensibilityHandlerObject","New-PnPGraphSubscription","New-PnPGroup","New-PnPList","New-PnPMicrosoft365Group","New-PnPPersonalSite","New-PnPPlannerPlan","New-PnPSdnProvider","New-PnPSite","New-PnPSiteCollectionTermStore","New-PnPSiteGroup","New-PnPSiteTemplate","New-PnPSiteTemplateFromFolder","New-PnPTeamsApp","New-PnPTeamsTeam","New-PnPTenantSequence","New-PnPTenantSequenceCommunicationSite","New-PnPTenantSequenceTeamNoGroupSite","New-PnPTenantSequenceTeamNoGroupSubSite","New-PnPTenantSequenceTeamSite","New-PnPTenantSite","New-PnPTenantTemplate","New-PnPTerm","New-PnPTermGroup","New-PnPTermLabel","New-PnPTermSet","New-PnPUPABulkImportJob","New-PnPUser","New-PnPWeb","Publish-PnPApp","Publish-PnPCompanyApp","Publish-PnPSyntexModel","Read-PnPSiteTemplate","Read-PnPTenantTemplate","Receive-PnPCopyMoveJobStatus","Register-PnPAppCatalogSite","Register-PnPAzureADApp","Register-PnPHubSite","Register-PnPManagementShellAccess","Remove-PnPAlert","Remove-PnPApp","Remove-PnPApplicationCustomizer","Remove-PnPAzureADApp","Remove-PnPAzureADGroup","Remove-PnPAzureADGroupMember","Remove-PnPAzureADGroupOwner","Remove-PnPContentType","Remove-PnPContentTypeFromDocumentSet","Remove-PnPContentTypeFromList","Remove-PnPCustomAction","Remove-PnPDeletedMicrosoft365Group","Remove-PnPEventReceiver","Remove-PnPExternalUser","Remove-PnPField","Remove-PnPFieldFromContentType","Remove-PnPFile","Remove-PnPFileFromSiteTemplate","Remove-PnPFileVersion","Remove-PnPFlow","Remove-PnPFolder","Remove-PnPGraphSubscription","Remove-PnPGroup","Remove-PnPGroupMember","Remove-PnPHomeSite","Remove-PnPHubSiteAssociation","Remove-PnPHubToHubAssociation","Remove-PnPIndexedProperty","Remove-PnPJavaScriptLink","Remove-PnPKnowledgeHubSite","Remove-PnPList","Remove-PnPListItem","Remove-PnPMicrosoft365Group","Remove-PnPMicrosoft365GroupMember","Remove-PnPMicrosoft365GroupOwner","Remove-PnPNavigationNode","Remove-PnPOrgAssetsLibrary","Remove-PnPOrgNewsSite","Remove-PnPPage","Remove-PnPPageComponent","Remove-PnPPlannerBucket","Remove-PnPPlannerPlan","Remove-PnPPlannerTask","Remove-PnPPropertyBagValue","Remove-PnPPublishingImageRendition","Remove-PnPRoleDefinition","Remove-PnPSdnProvider","Remove-PnPSearchConfiguration","Remove-PnPSiteClassification","Remove-PnPSiteCollectionAdmin","Remove-PnPSiteCollectionAppCatalog","Remove-PnPSiteCollectionTermStore","Remove-PnPSiteDesign","Remove-PnPSiteDesignTask","Remove-PnPSiteGroup","Remove-PnPSiteScript","Remove-PnPSiteUserInvitations","Remove-PnPStorageEntity","Remove-PnPStoredCredential","Remove-PnPTaxonomyItem","Remove-PnPTeamsApp","Remove-PnPTeamsChannel","Remove-PnPTeamsTab","Remove-PnPTeamsTeam","Remove-PnPTeamsUser","Remove-PnPTenantCdnOrigin","Remove-PnPTenantDeletedSite","Remove-PnPTenantSite","Remove-PnPTenantSyncClientRestriction","Remove-PnPTenantTheme","Remove-PnPTerm","Remove-PnPTermGroup","Remove-PnPTermLabel","Remove-PnPUser","Remove-PnPUserInfo","Remove-PnPUserProfile","Remove-PnPView","Remove-PnPWeb","Remove-PnPWebhookSubscription","Remove-PnPWebPart","Remove-PnPWikiPage","Rename-PnPFile","Rename-PnPFolder","Repair-PnPSite","Request-PnPAccessToken","Request-PnPPersonalSite","Request-PnPReIndexList","Request-PnPReIndexWeb","Request-PnPSyntexClassifyAndExtract","Reset-PnPFileVersion","Reset-PnPLabel","Reset-PnPMicrosoft365GroupExpiration","Reset-PnPUserOneDriveQuotaToDefault","Resolve-PnPFolder","Restore-PnPDeletedMicrosoft365Group","Restore-PnPFileVersion","Restore-PnPRecycleBinItem","Restore-PnPTenantRecycleBinItem","Restore-PnPTenantSite","Revoke-PnPAzureADAppSitePermission","Revoke-PnPHubSiteRights","Revoke-PnPSiteDesignRights","Revoke-PnPTenantServicePrincipalPermission","Revoke-PnPUserSession","Save-PnPPageConversionLog","Save-PnPSiteTemplate","Save-PnPTenantTemplate","Send-PnPMail","Set-PnPApplicationCustomizer","Set-PnPAppSideLoading","Set-PnPAuditing","Set-PnPAvailablePageLayouts","Set-PnPAzureADAppSitePermission","Set-PnPAzureADGroup","Set-PnPBrowserIdleSignout","Set-PnPBuiltInDesignPackageVisibility","Set-PnPContext","Set-PnPDefaultColumnValues","Set-PnPDefaultContentTypeToList","Set-PnPDefaultPageLayout","Set-PnPDisableSpacesActivation","Set-PnPDocumentSetField","Set-PnPField","Set-PnPFileCheckedIn","Set-PnPFileCheckedOut","Set-PnPFolderPermission","Set-PnPFooter","Set-PnPGraphSubscription","Set-PnPGroup","Set-PnPGroupPermissions","Set-PnPHideDefaultThemes","Set-PnPHomePage","Set-PnPHomeSite","Set-PnPHubSite","Set-PnPIndexedProperties","Set-PnPInPlaceRecordsManagement","Set-PnPKnowledgeHubSite","Set-PnPLabel","Set-PnPList","Set-PnPListInformationRightsManagement","Set-PnPListItem","Set-PnPListItemAsRecord","Set-PnPListItemPermission","Set-PnPListPermission","Set-PnPListRecordDeclaration","Set-PnPMasterPage","Set-PnPMicrosoft365Group","Set-PnPMinimalDownloadStrategy","Set-PnPPage","Set-PnPPageTextPart","Set-PnPPageWebPart","Set-PnPPlannerBucket","Set-PnPPlannerPlan","Set-PnPPlannerTask","Set-PnPPropertyBagValue","Set-PnPRequestAccessEmails","Set-PnPRoleDefinition","Set-PnPSearchConfiguration","Set-PnPSearchSettings","Set-PnPSite","Set-PnPSiteClosure","Set-PnPSiteDesign","Set-PnPSiteGroup","Set-PnPSitePolicy","Set-PnPSiteScript","Set-PnPSiteScriptPackage","Set-PnPSiteTemplateMetadata","Set-PnPStorageEntity","Set-PnPStructuralNavigationCacheSiteState","Set-PnPStructuralNavigationCacheWebState","Set-PnPSubscribeSharePointNewsDigest","Set-PnPTaxonomyFieldValue","Set-PnPTeamifyPromptHidden","Set-PnPTeamsChannel","Set-PnPTeamsTab","Set-PnPTeamsTeam","Set-PnPTeamsTeamArchivedState","Set-PnPTeamsTeamPicture","Set-PnPTemporarilyDisableAppBar","Set-PnPTenant","Set-PnPTenantAppCatalogUrl","Set-PnPTenantCdnEnabled","Set-PnPTenantCdnPolicy","Set-PnPTenantSite","Set-PnPTenantSyncClientRestriction","Set-PnPTerm","Set-PnPTermGroup","Set-PnPTermSet","Set-PnPTheme","Set-PnPTraceLog","Set-PnPUserOneDriveQuota","Set-PnPUserProfileProperty","Set-PnPView","Set-PnPWeb","Set-PnPWebhookSubscription","Set-PnPWebPartProperty","Set-PnPWebPermission","Set-PnPWebTheme","Set-PnPWikiPageContent","Submit-PnPSearchQuery","Submit-PnPTeamsChannelMessage","Sync-PnPAppToTeams","Sync-PnPSharePointUserProfilesFromAzureActiveDirectory","Test-PnPListItemIsRecord","Test-PnPMicrosoft365GroupAliasIsUsed","Test-PnPSite","Test-PnPTenantTemplate","Uninstall-PnPApp","Unpublish-PnPApp","Unpublish-PnPSyntexModel","Unregister-PnPHubSite","Update-PnPApp","Update-PnPSiteClassification","Update-PnPSiteDesignFromWeb","Update-PnPTeamsApp","Update-PnPUserType")
	VariablesToExport = '*'
	AliasesToExport = '*'
	FormatsToProcess = 'PnP.PowerShell.Format.ps1xml' 
	PrivateData = @{
		PSData = @{
			Tags = 'SharePoint','PnP','Teams','Planner'
			ProjectUri = 'https://aka.ms/sppnp'
			IconUri = 'https://raw.githubusercontent.com/pnp/media/40e7cd8952a9347ea44e5572bb0e49622a102a12/parker/ms/300w/parker-ms-300.png'
		}
	}
}

# SIG # Begin signature block
# MIIjgwYJKoZIhvcNAQcCoIIjdDCCI3ACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAUiUp4Tx6DAW5X
# PVxUz1z5HSibZLttALqY2rTBx69f8qCCDYEwggX/MIID56ADAgECAhMzAAAB32vw
# LpKnSrTQAAAAAAHfMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjAxMjE1MjEzMTQ1WhcNMjExMjAyMjEzMTQ1WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQC2uxlZEACjqfHkuFyoCwfL25ofI9DZWKt4wEj3JBQ48GPt1UsDv834CcoUUPMn
# s/6CtPoaQ4Thy/kbOOg/zJAnrJeiMQqRe2Lsdb/NSI2gXXX9lad1/yPUDOXo4GNw
# PjXq1JZi+HZV91bUr6ZjzePj1g+bepsqd/HC1XScj0fT3aAxLRykJSzExEBmU9eS
# yuOwUuq+CriudQtWGMdJU650v/KmzfM46Y6lo/MCnnpvz3zEL7PMdUdwqj/nYhGG
# 3UVILxX7tAdMbz7LN+6WOIpT1A41rwaoOVnv+8Ua94HwhjZmu1S73yeV7RZZNxoh
# EegJi9YYssXa7UZUUkCCA+KnAgMBAAGjggF+MIIBejAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUOPbML8IdkNGtCfMmVPtvI6VZ8+Mw
# UAYDVR0RBEkwR6RFMEMxKTAnBgNVBAsTIE1pY3Jvc29mdCBPcGVyYXRpb25zIFB1
# ZXJ0byBSaWNvMRYwFAYDVQQFEw0yMzAwMTIrNDYzMDA5MB8GA1UdIwQYMBaAFEhu
# ZOVQBdOCqhc3NyK1bajKdQKVMFQGA1UdHwRNMEswSaBHoEWGQ2h0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY0NvZFNpZ1BDQTIwMTFfMjAxMS0w
# Ny0wOC5jcmwwYQYIKwYBBQUHAQEEVTBTMFEGCCsGAQUFBzAChkVodHRwOi8vd3d3
# Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01pY0NvZFNpZ1BDQTIwMTFfMjAx
# MS0wNy0wOC5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAgEAnnqH
# tDyYUFaVAkvAK0eqq6nhoL95SZQu3RnpZ7tdQ89QR3++7A+4hrr7V4xxmkB5BObS
# 0YK+MALE02atjwWgPdpYQ68WdLGroJZHkbZdgERG+7tETFl3aKF4KpoSaGOskZXp
# TPnCaMo2PXoAMVMGpsQEQswimZq3IQ3nRQfBlJ0PoMMcN/+Pks8ZTL1BoPYsJpok
# t6cql59q6CypZYIwgyJ892HpttybHKg1ZtQLUlSXccRMlugPgEcNZJagPEgPYni4
# b11snjRAgf0dyQ0zI9aLXqTxWUU5pCIFiPT0b2wsxzRqCtyGqpkGM8P9GazO8eao
# mVItCYBcJSByBx/pS0cSYwBBHAZxJODUqxSXoSGDvmTfqUJXntnWkL4okok1FiCD
# Z4jpyXOQunb6egIXvkgQ7jb2uO26Ow0m8RwleDvhOMrnHsupiOPbozKroSa6paFt
# VSh89abUSooR8QdZciemmoFhcWkEwFg4spzvYNP4nIs193261WyTaRMZoceGun7G
# CT2Rl653uUj+F+g94c63AhzSq4khdL4HlFIP2ePv29smfUnHtGq6yYFDLnT0q/Y+
# Di3jwloF8EWkkHRtSuXlFUbTmwr/lDDgbpZiKhLS7CBTDj32I0L5i532+uHczw82
# oZDmYmYmIUSMbZOgS65h797rj5JJ6OkeEUJoAVwwggd6MIIFYqADAgECAgphDpDS
# AAAAAAADMA0GCSqGSIb3DQEBCwUAMIGIMQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0
# ZSBBdXRob3JpdHkgMjAxMTAeFw0xMTA3MDgyMDU5MDlaFw0yNjA3MDgyMTA5MDla
# MH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMT
# H01pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMTEwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQCr8PpyEBwurdhuqoIQTTS68rZYIZ9CGypr6VpQqrgG
# OBoESbp/wwwe3TdrxhLYC/A4wpkGsMg51QEUMULTiQ15ZId+lGAkbK+eSZzpaF7S
# 35tTsgosw6/ZqSuuegmv15ZZymAaBelmdugyUiYSL+erCFDPs0S3XdjELgN1q2jz
# y23zOlyhFvRGuuA4ZKxuZDV4pqBjDy3TQJP4494HDdVceaVJKecNvqATd76UPe/7
# 4ytaEB9NViiienLgEjq3SV7Y7e1DkYPZe7J7hhvZPrGMXeiJT4Qa8qEvWeSQOy2u
# M1jFtz7+MtOzAz2xsq+SOH7SnYAs9U5WkSE1JcM5bmR/U7qcD60ZI4TL9LoDho33
# X/DQUr+MlIe8wCF0JV8YKLbMJyg4JZg5SjbPfLGSrhwjp6lm7GEfauEoSZ1fiOIl
# XdMhSz5SxLVXPyQD8NF6Wy/VI+NwXQ9RRnez+ADhvKwCgl/bwBWzvRvUVUvnOaEP
# 6SNJvBi4RHxF5MHDcnrgcuck379GmcXvwhxX24ON7E1JMKerjt/sW5+v/N2wZuLB
# l4F77dbtS+dJKacTKKanfWeA5opieF+yL4TXV5xcv3coKPHtbcMojyyPQDdPweGF
# RInECUzF1KVDL3SV9274eCBYLBNdYJWaPk8zhNqwiBfenk70lrC8RqBsmNLg1oiM
# CwIDAQABo4IB7TCCAekwEAYJKwYBBAGCNxUBBAMCAQAwHQYDVR0OBBYEFEhuZOVQ
# BdOCqhc3NyK1bajKdQKVMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMAsGA1Ud
# DwQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFHItOgIxkEO5FAVO
# 4eqnxzHRI4k0MFoGA1UdHwRTMFEwT6BNoEuGSWh0dHA6Ly9jcmwubWljcm9zb2Z0
# LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0NlckF1dDIwMTFfMjAxMV8wM18y
# Mi5jcmwwXgYIKwYBBQUHAQEEUjBQME4GCCsGAQUFBzAChkJodHRwOi8vd3d3Lm1p
# Y3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dDIwMTFfMjAxMV8wM18y
# Mi5jcnQwgZ8GA1UdIASBlzCBlDCBkQYJKwYBBAGCNy4DMIGDMD8GCCsGAQUFBwIB
# FjNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2RvY3MvcHJpbWFyeWNw
# cy5odG0wQAYIKwYBBQUHAgIwNB4yIB0ATABlAGcAYQBsAF8AcABvAGwAaQBjAHkA
# XwBzAHQAYQB0AGUAbQBlAG4AdAAuIB0wDQYJKoZIhvcNAQELBQADggIBAGfyhqWY
# 4FR5Gi7T2HRnIpsLlhHhY5KZQpZ90nkMkMFlXy4sPvjDctFtg/6+P+gKyju/R6mj
# 82nbY78iNaWXXWWEkH2LRlBV2AySfNIaSxzzPEKLUtCw/WvjPgcuKZvmPRul1LUd
# d5Q54ulkyUQ9eHoj8xN9ppB0g430yyYCRirCihC7pKkFDJvtaPpoLpWgKj8qa1hJ
# Yx8JaW5amJbkg/TAj/NGK978O9C9Ne9uJa7lryft0N3zDq+ZKJeYTQ49C/IIidYf
# wzIY4vDFLc5bnrRJOQrGCsLGra7lstnbFYhRRVg4MnEnGn+x9Cf43iw6IGmYslmJ
# aG5vp7d0w0AFBqYBKig+gj8TTWYLwLNN9eGPfxxvFX1Fp3blQCplo8NdUmKGwx1j
# NpeG39rz+PIWoZon4c2ll9DuXWNB41sHnIc+BncG0QaxdR8UvmFhtfDcxhsEvt9B
# xw4o7t5lL+yX9qFcltgA1qFGvVnzl6UJS0gQmYAf0AApxbGbpT9Fdx41xtKiop96
# eiL6SJUfq/tHI4D1nvi/a7dLl+LrdXga7Oo3mXkYS//WsyNodeav+vyL6wuA6mk7
# r/ww7QRMjt/fdW1jkT3RnVZOT7+AVyKheBEyIXrvQQqxP/uozKRdwaGIm1dxVk5I
# RcBCyZt2WwqASGv9eZ/BvW1taslScxMNelDNMYIVWDCCFVQCAQEwgZUwfjELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9z
# b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMQITMwAAAd9r8C6Sp0q00AAAAAAB3zAN
# BglghkgBZQMEAgEFAKCBrjAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgor
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgKfSakHr/
# uf5Ji6gRoyioBltETg8Qy5IiGWF56R/oxf0wQgYKKwYBBAGCNwIBDDE0MDKgFIAS
# AE0AaQBjAHIAbwBzAG8AZgB0oRqAGGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbTAN
# BgkqhkiG9w0BAQEFAASCAQAi0KI8FqxsJrDqrHHgkoBbkysTY4uyK5N+JGi2FyP5
# qZGo4IU7ACBh25hEkQkQUECw/3Y2I0ALPQr9halCpjITHrusRVz7lxZwRsUVlwNm
# SAD5OGhWN5VPc2TicsS1DmBnDftF/1lAtXSxrau24fIxUXZvM0qthOW0gMKpQK/F
# xzQ+Fl2Pfr8cpB7jRBDsmFibJbI6S14/PvxuKiV1xHEyFW8lnfOoftktXc+hutI8
# +RkbQUIdnrpOW2TFZVLJxphI7VdvZFoLwXUVm84WLcFlaRDKAJefieyHrtVA+UeA
# ZePTlQ50TyRfauQ4wBzgdpnBdvrkU7Aau8av4bj4hE9hoYIS4jCCEt4GCisGAQQB
# gjcDAwExghLOMIISygYJKoZIhvcNAQcCoIISuzCCErcCAQMxDzANBglghkgBZQME
# AgEFADCCAVEGCyqGSIb3DQEJEAEEoIIBQASCATwwggE4AgEBBgorBgEEAYRZCgMB
# MDEwDQYJYIZIAWUDBAIBBQAEIFFbFC+Fz92l49t5e2yp7zOIQ0WapFmDE08X0lN0
# LKsOAgZhRLqkE5EYEzIwMjExMDA4MTQxNjM5LjM2OFowBIACAfSggdCkgc0wgcox
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJTAjBgNVBAsTHE1p
# Y3Jvc29mdCBBbWVyaWNhIE9wZXJhdGlvbnMxJjAkBgNVBAsTHVRoYWxlcyBUU1Mg
# RVNOOkFFMkMtRTMyQi0xQUZDMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFt
# cCBTZXJ2aWNloIIOOTCCBPEwggPZoAMCAQICEzMAAAFIoohFVrwvgL8AAAAAAUgw
# DQYJKoZIhvcNAQELBQAwfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
# b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3Jh
# dGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwHhcN
# MjAxMTEyMTgyNTU2WhcNMjIwMjExMTgyNTU2WjCByjELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2Eg
# T3BlcmF0aW9uczEmMCQGA1UECxMdVGhhbGVzIFRTUyBFU046QUUyQy1FMzJCLTFB
# RkMxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0G
# CSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQD3/3ivFYSK0dGtcXaZ8pNLEARbraJe
# wryi/JgbaKlq7hhFIU1EkY0HMiFRm2/Wsukt62k25zvDxW16fphg5876+l1wYnCl
# ge/rFlrR2Uu1WwtFmc1xGpy4+uxobCEMeIFDGhL5DNTbbOisLrBUYbyXr7fPzxbV
# kEwJDP5FG2n0ro1qOjegIkLIjXU6qahduQxTfsPOEp8jgqMKn++fpH6fvXKlewWz
# dsfvhiZ4H4Iq1CTOn+fkxqcDwTHYkYZYgqm+1X1x7458rp69qjFeVP3GbAvJbY3b
# Flq5uyxriPcZxDZrB6f1wALXrO2/IdfVEdwTWqJIDZBJjTycQhhxS3i1AgMBAAGj
# ggEbMIIBFzAdBgNVHQ4EFgQUhzLwaZ8OBLRJH0s9E63pIcWJokcwHwYDVR0jBBgw
# FoAU1WM6XIoxkPNDe3xGG8UzaFqFbVUwVgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDov
# L2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljVGltU3RhUENB
# XzIwMTAtMDctMDEuY3JsMFoGCCsGAQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0
# cDovL3d3dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNUaW1TdGFQQ0FfMjAx
# MC0wNy0wMS5jcnQwDAYDVR0TAQH/BAIwADATBgNVHSUEDDAKBggrBgEFBQcDCDAN
# BgkqhkiG9w0BAQsFAAOCAQEAZhKWwbMnC9Qywcrlgs0qX9bhxiZGve+8JED27hOi
# yGa8R9nqzHg4+q6NKfYXfS62uMUJp2u+J7tINUTf/1ugL+K4RwsPVehDasSJJj+7
# boIxZP8AU/xQdVY7qgmQGmd4F+c5hkJJtl6NReYE908Q698qj1mDpr0Mx+4LhP/t
# TqL6HpZEURlhFOddnyLStVCFdfNI1yGHP9n0yN1KfhGEV3s7MBzpFJXwOflwgyE9
# cwQ8jjOTVpNRdCqL/P5ViCAo2dciHjd1u1i1Q4QZ6xb0+B1HdZFRELOiFwf0sh3Z
# 1xOeSFcHg0rLE+rseHz4QhvoEj7h9bD8VN7/HnCDwWpBJTCCBnEwggRZoAMCAQIC
# CmEJgSoAAAAAAAIwDQYJKoZIhvcNAQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRp
# ZmljYXRlIEF1dGhvcml0eSAyMDEwMB4XDTEwMDcwMTIxMzY1NVoXDTI1MDcwMTIx
# NDY1NVowfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQG
# A1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQCpHQ28dxGKOiDs/BOX9fp/aZRrdFQQ1aUKAIKF
# ++18aEssX8XD5WHCdrc+Zitb8BVTJwQxH0EbGpUdzgkTjnxhMFmxMEQP8WCIhFRD
# DNdNuDgIs0Ldk6zWczBXJoKjRQ3Q6vVHgc2/JGAyWGBG8lhHhjKEHnRhZ5FfgVSx
# z5NMksHEpl3RYRNuKMYa+YaAu99h/EbBJx0kZxJyGiGKr0tkiVBisV39dx898Fd1
# rL2KQk1AUdEPnAY+Z3/1ZsADlkR+79BL/W7lmsqxqPJ6Kgox8NpOBpG2iAg16Hgc
# sOmZzTznL0S6p/TcZL2kAcEgCZN4zfy8wMlEXV4WnAEFTyJNAgMBAAGjggHmMIIB
# 4jAQBgkrBgEEAYI3FQEEAwIBADAdBgNVHQ4EFgQU1WM6XIoxkPNDe3xGG8UzaFqF
# bVUwGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwCwYDVR0PBAQDAgGGMA8GA1Ud
# EwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU1fZWy4/oolxiaNE9lJBb186aGMQwVgYD
# VR0fBE8wTTBLoEmgR4ZFaHR0cDovL2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwv
# cHJvZHVjdHMvTWljUm9vQ2VyQXV0XzIwMTAtMDYtMjMuY3JsMFoGCCsGAQUFBwEB
# BE4wTDBKBggrBgEFBQcwAoY+aHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraS9j
# ZXJ0cy9NaWNSb29DZXJBdXRfMjAxMC0wNi0yMy5jcnQwgaAGA1UdIAEB/wSBlTCB
# kjCBjwYJKwYBBAGCNy4DMIGBMD0GCCsGAQUFBwIBFjFodHRwOi8vd3d3Lm1pY3Jv
# c29mdC5jb20vUEtJL2RvY3MvQ1BTL2RlZmF1bHQuaHRtMEAGCCsGAQUFBwICMDQe
# MiAdAEwAZQBnAGEAbABfAFAAbwBsAGkAYwB5AF8AUwB0AGEAdABlAG0AZQBuAHQA
# LiAdMA0GCSqGSIb3DQEBCwUAA4ICAQAH5ohRDeLG4Jg/gXEDPZ2joSFvs+umzPUx
# vs8F4qn++ldtGTCzwsVmyWrf9efweL3HqJ4l4/m87WtUVwgrUYJEEvu5U4zM9GAS
# inbMQEBBm9xcF/9c+V4XNZgkVkt070IQyK+/f8Z/8jd9Wj8c8pl5SpFSAK84Dxf1
# L3mBZdmptWvkx872ynoAb0swRCQiPM/tA6WWj1kpvLb9BOFwnzJKJ/1Vry/+tuWO
# M7tiX5rbV0Dp8c6ZZpCM/2pif93FSguRJuI57BlKcWOdeyFtw5yjojz6f32WapB4
# pm3S4Zz5Hfw42JT0xqUKloakvZ4argRCg7i1gJsiOCC1JeVk7Pf0v35jWSUPei45
# V3aicaoGig+JFrphpxHLmtgOR5qAxdDNp9DvfYPw4TtxCd9ddJgiCGHasFAeb73x
# 4QDf5zEHpJM692VHeOj4qEir995yfmFrb3epgcunCaw5u+zGy9iCtHLNHfS4hQEe
# gPsbiSpUObJb2sgNVZl6h3M7COaYLeqN4DMuEin1wC9UJyH3yKxO2ii4sanblrKn
# QqLJzxlBTeCG+SqaoxFmMNO7dDJL32N79ZmKLxvHIa9Zta7cRDyXUHHXodLFVeNp
# 3lfB0d4wwP3M5k37Db9dT+mdHhk4L7zPWAUu7w2gUDXa7wknHNWzfjUeCLraNtvT
# X4/edIhJEqGCAsswggI0AgEBMIH4oYHQpIHNMIHKMQswCQYDVQQGEwJVUzETMBEG
# A1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWlj
# cm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1lcmljYSBP
# cGVyYXRpb25zMSYwJAYDVQQLEx1UaGFsZXMgVFNTIEVTTjpBRTJDLUUzMkItMUFG
# QzElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZaIjCgEBMAcG
# BSsOAwIaAxUAhyuClrocWf4SIcRafAEX1Rhs6zmggYMwgYCkfjB8MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQg
# VGltZS1TdGFtcCBQQ0EgMjAxMDANBgkqhkiG9w0BAQUFAAIFAOUKP0swIhgPMjAy
# MTEwMDgxMTUzNDdaGA8yMDIxMTAwOTExNTM0N1owdDA6BgorBgEEAYRZCgQBMSww
# KjAKAgUA5Qo/SwIBADAHAgEAAgIEUTAHAgEAAgIRNzAKAgUA5QuQywIBADA2Bgor
# BgEEAYRZCgQCMSgwJjAMBgorBgEEAYRZCgMCoAowCAIBAAIDB6EgoQowCAIBAAID
# AYagMA0GCSqGSIb3DQEBBQUAA4GBAIdb0LEHWS8VCptNZ+gRgPHobKNGOcOcwjN4
# CQZ/L57GY33FjYwCmuF9ILCs6H5EaBeCai2XN+DdqTqqBtD/hmvb8GSu+rQMn3+P
# OXG3wZyAaV/k4nyohQABVcCrx0aCopxKtRU7s7JHrk3Kso7mLdGygqBYhcEsf0eX
# 8xVzQmUZMYIDDTCCAwkCAQEwgZMwfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldh
# c2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
# b3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIw
# MTACEzMAAAFIoohFVrwvgL8AAAAAAUgwDQYJYIZIAWUDBAIBBQCgggFKMBoGCSqG
# SIb3DQEJAzENBgsqhkiG9w0BCRABBDAvBgkqhkiG9w0BCQQxIgQgzy+xk8xkOuLb
# Z6f0G5go9RwZMS4gV+n1v1cjcnyRsWUwgfoGCyqGSIb3DQEJEAIvMYHqMIHnMIHk
# MIG9BCCpkBrqjHmhvyYf5tTcTvD5Y4a+V79TwVV6T1aAwdto2DCBmDCBgKR+MHwx
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1p
# Y3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAABSKKIRVa8L4C/AAAAAAFI
# MCIEIDjq69tzoevfGDGs37NaNDCDku0uvem5pz3NPoGVAYMQMA0GCSqGSIb3DQEB
# CwUABIIBAJjaPowY31PW21VRorcg2N7Jx77niy89FWpwoXSHxs9z2hF9OWaAlIrE
# mlavB0g2NdFLo89qWGObAaYRRp1/yCXTkb6CdqYmCy20bD3TOH+7TA7N/VJdq52z
# DM1fFJCScnK1kIQn6w1DgZF5qdSU83DBWNQDkY7DpdYpEMYyrV7Cat1ru3wxlrOi
# aOs6K83ZgFFVZJka2mzWJLtvYLdqr0lVEzx9/yFRIUmT+H0b8+jNcKqPLc0cdts/
# /3AuSF1XMGLeB2YpwIlQeGwCM+Bp/Otik7ZJ37FVHiXo8Sk7kHdpBl7VmAupOgg2
# b/aORHIRboVVNyBg1q8Zjx3S82cQP94=
# SIG # End signature block
