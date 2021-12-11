@{
	# Script module or binary module file associated with this manifest
	RootModule = 'SPOPermissionReport.psm1'

	# Version number of this module.
	ModuleVersion = '1.0.0'

	# ID used to uniquely identify this module
	GUID = '29ff35c8-e66d-45d9-9053-3343189afbba'

	# Author of this module
	Author = 'wmoselhy'

	# Company or vendor of this module
	CompanyName = 'MyCompany'

	# Copyright statement for this module
	Copyright = 'Copyright (c) 2021 wmoselhy'

	# Description of the functionality provided by this module
	Description = 'PowerShell Module to collect SharePoint Online Permission Details'

	# Minimum version of the Windows PowerShell engine required by this module
	PowerShellVersion = '5.0'

	# Modules that must be imported into the global environment prior to importing
	# this module
	#RequiredModules = @(
	#	'PSFramework'
	#)

	# Assemblies that must be loaded prior to importing this module
	# RequiredAssemblies = @('bin\SPOPermissionReport.dll')

	# Type files (.ps1xml) to be loaded when importing this module
	# TypesToProcess = @('xml\SPOPermissionReport.Types.ps1xml')

	# Format files (.ps1xml) to be loaded when importing this module
	# FormatsToProcess = @('xml\SPOPermissionReport.Format.ps1xml')

	# Functions to export from this module
	FunctionsToExport = @(
		'Start-SPOPermissionCollection'
	)

	# Cmdlets to export from this module
	CmdletsToExport = ''

	# Variables to export from this module
	VariablesToExport = ''

	# Aliases to export from this module
	AliasesToExport = ''

	# List of all modules packaged with this module
	ModuleList = @()

	# List of all files packaged with this module
	FileList = @()

	# Private data to pass to the module specified in ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
	PrivateData = @{

		#Support for PowerShellGet galleries.
		PSData = @{

			# Tags applied to this module. These help with module discovery in online galleries.
			# Tags = @()

			# A URL to the license for this module.
			# LicenseUri = ''

			# A URL to the main website for this project.
			# ProjectUri = ''

			# A URL to an icon representing this module.
			# IconUri = ''

			# ReleaseNotes of this module
			# ReleaseNotes = ''

		} # End of PSData hashtable

	} # End of PrivateData hashtable
}