#
# Module manifest for module 'PSBookmark'
#
# Generated by: Ravikanth Chaganti
#
# Generated on: 5/5/2016
#

@{

# Script module or binary module file associated with this manifest.
RootModule = 'PSBookmark.psm1'

# Version number of this module.
ModuleVersion = '1.0.1'

# ID used to uniquely identify this module
GUID = '298a7c5c-d670-4093-bae2-0d6c2935a182'

# Author of this module
Author = 'Ravikanth Chaganti'

# Company or vendor of this module
CompanyName = 'PowerShell Magazine'

# Copyright statement for this module
Copyright = 'All Rights Reserved'

# Description of the functionality provided by this module
Description = 'Adds the capability to create location bookmarks for easier access.'

# Minimum version of the Windows PowerShell engine required by this module
# PowerShellVersion = ''

# Name of the Windows PowerShell host required by this module
# PowerShellHostName = ''

# Minimum version of the Windows PowerShell host required by this module
# PowerShellHostVersion = ''

# Minimum version of Microsoft .NET Framework required by this module
# DotNetFrameworkVersion = ''

# Minimum version of the common language runtime (CLR) required by this module
# CLRVersion = ''

# Processor architecture (None, X86, Amd64) required by this module
# ProcessorArchitecture = ''

# Modules that must be imported into the global environment prior to importing this module
# RequiredModules = @()

# Assemblies that must be loaded prior to importing this module
# RequiredAssemblies = @()

# Script files (.ps1) that are run in the caller's environment prior to importing this module.
# ScriptsToProcess = @()

# Type files (.ps1xml) to be loaded when importing this module
# TypesToProcess = @()

# Format files (.ps1xml) to be loaded when importing this module
# FormatsToProcess = @()

# Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
# NestedModules = @()

# Functions to export from this module
FunctionsToExport = 'Save-LocationBookmark','Set-LocationBookmarkAsPWD','Get-LocationBookmark','Remove-LocationBookmark'

# Cmdlets to export from this module
CmdletsToExport = '*'

# Variables to export from this module
VariablesToExport = '*'

# Aliases to export from this module
AliasesToExport = 'goto','save','glb','rlb'

# DSC resources to export from this module
# DscResourcesToExport = @()

# List of all modules packaged with this module
# ModuleList = @()

# List of all files packaged with this module
# FileList = @()

# Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
PrivateData = @{

    PSData = @{

        # Tags applied to this module. These help with module discovery in online galleries.
        # Tags = @()

        # A URL to the license for this module.
        # LicenseUri = ''

        # A URL to the main website for this project.
        # ProjectUri = ''

        # A URL to an icon representing this module.
        IconUri = 'http://www.powershellmagazine.com/wp-content/uploads/2012/01/logo-250x250-200x200.png'

        # ReleaseNotes of this module
        # ReleaseNotes = ''

    } # End of PSData hashtable

} # End of PrivateData hashtable

# HelpInfo URI of this module
# HelpInfoURI = ''

# Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
# DefaultCommandPrefix = ''

}

