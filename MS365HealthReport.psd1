#
# Module manifest for module 'MS365HealthReport'
#
# Generated by: tcastillotej
#
# Generated on: 1/25/2021
#

@{

    # Script module or binary module file associated with this manifest.
    RootModule        = '.\MS365HealthReport.psm1'

    # Version number of this module.
    ModuleVersion     = '2.1.6'

    # Supported PSEditions
    # CompatiblePSEditions = @()

    # ID used to uniquely identify this module
    GUID              = '4c20a01b-7720-4884-82de-2d726176ff71'

    # Author of this module
    Author            = 'June Castillote'

    # Company or vendor of this module
    CompanyName       = 'June Castillote'

    # Copyright statement for this module
    Copyright         = '(c) 2021 june.castillote@gmail.com. All rights reserved.'

    # Description of the functionality provided by this module
    Description = 'Retrieve the Microsoft 365 Service Health status and send the email report using Microsoft Graph API.'

    # Minimum version of the Windows PowerShell engine required by this module
    PowerShellVersion = '5.1'

    # Name of the Windows PowerShell host required by this module
    # PowerShellHostName = ''

    # Minimum version of the Windows PowerShell host required by this module
    # PowerShellHostVersion = ''

    # Minimum version of Microsoft .NET Framework required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
    # DotNetFrameworkVersion = ''

    # Minimum version of the common language runtime (CLR) required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
    # CLRVersion = ''

    # Processor architecture (None, X86, Amd64) required by this module
    # ProcessorArchitecture = ''

    # Modules that must be imported into the global environment prior to importing this module
    RequiredModules   = @('MSAL.PS')

    # Assemblies that must be loaded prior to importing this module
    # RequiredAssemblies = @()

    # Script files (.ps1) that are run in the caller's environment prior to importing this module.
    # ScriptsToProcess = @('.\source\private\precheck.ps1')

    # Type files (.ps1xml) to be loaded when importing this module
    # TypesToProcess = @()

    # Format files (.ps1xml) to be loaded when importing this module
    # FormatsToProcess = @()

    # Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
    # NestedModules = @()

    # Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
    FunctionsToExport = @(
        'New-MS365IncidentReport',
        'Get-MS365Messages',
        'Get-MS365HealthOverview',
        'Get-MS365CurrentStatus',
        'New-ConsolidatedCard'
    )

    # Cmdlets to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no cmdlets to export.
    CmdletsToExport   = @()

    # Variables to export from this module
    VariablesToExport = '*'

    # Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no aliases to export.
    AliasesToExport   = @()

    # DSC resources to export from this module
    # DscResourcesToExport = @()

    # List of all modules packaged with this module
    # ModuleList = @()

    # List of all files packaged with this module
    # FileList = @()

    # Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
    PrivateData       = @{

        PSData = @{

            # Tags applied to this module. These help with module discovery in online galleries.
            Tags = @('Graph','API','Office365','Health','Report','Service','Communication')

            # A URL to the license for this module.
            LicenseUri = 'https://github.com/junecastillote/MS365HealthReport/blob/main/LICENSE'

            # A URL to the main website for this project.
            ProjectUri = 'https://github.com/junecastillote/MS365HealthReport'

            # A URL to an icon representing this module.
            # IconUri = ''

            # ReleaseNotes of this module
            # ReleaseNotes = ''

        } # End of PSData hashtable

    } # End of PrivateData hashtable

    # HelpInfo URI of this module
    # HelpInfoURI = ''

    # Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
    # DefaultCommandPrefix = ''

}

