@{
    RootModule           = 'IntuneHousekeeper.psm1'
    ModuleVersion        = '0.9.1'
    GUID                 = 'b1dd7092-e846-4047-a833-3a9412845814'
    Author               = 'Konstantinos Bentis'
    Copyright            = '(c) 2026 Konstantinos Bentis. MIT licensed.'

    Description          = 'Read-only inventory of unassigned, mis-assigned and leftover Windows Intune objects, and optionally owner-scoped Entra ID assignment groups. Produces an Excel decision tracker. Issues GET requests only and never modifies a tenant.'

    # PowerShell 7 only. Windows PowerShell 5.1 fails at Connect-MgGraph on any machine
    # carrying several Microsoft.Graph module versions, because .NET Framework cannot
    # isolate the shared Microsoft.Identity.Client assembly.
    PowerShellVersion    = '7.0'
    CompatiblePSEditions = @('Core')

    RequiredModules      = @(
        @{ ModuleName = 'Microsoft.Graph.Authentication'; ModuleVersion = '2.17.0' }
        @{ ModuleName = 'ImportExcel';                    ModuleVersion = '7.8.0'  }
    )

    FunctionsToExport    = @(
        'Export-IntuneHousekeeperReport'
        'Get-IntuneHousekeeperConfig'
        'Set-IntuneHousekeeperConfig'
    )
    CmdletsToExport      = @()
    VariablesToExport    = @()
    AliasesToExport      = @()

    PrivateData = @{
        PSData = @{
            Tags         = @('Intune', 'MicrosoftGraph', 'Entra', 'Windows', 'Endpoint', 'Reporting', 'Excel', 'ReadOnly')
            LicenseUri   = 'https://github.com/kbentis/IntuneHousekeeper/blob/main/LICENSE'
            ProjectUri   = 'https://github.com/kbentis/IntuneHousekeeper'
            ReleaseNotes = 'Pre-release. See CHANGELOG.md in the project repository.'
        }
    }
}
