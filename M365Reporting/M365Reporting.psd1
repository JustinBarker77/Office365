@{
    RootModule = 'M365Reporting.psm1'
    ModuleVersion = '1.0.2.1'
    GUID = '6f505567-b194-4672-bd16-386cba3e539f'
    Author = 'Robin Dadswell & Luke Allinson & Justin Barker & Mark Lofthouse'
    CompanyName = 'Unknown'
    Copyright = '(c) 2022 Robin Dadswell, Luke Allinson, Justin Barker & Mark Lofthouse. All rights reserved.'
    Description = 'Provides tools to gather and report on some useful statistics within an M365 Tenant. This includes OneDrive Usage, Mailbox Sizes and Licence utilisation (including breakdown of assigned components)'
    RequiredModules =  @(
                            @{ ModuleName = 'ImportExcel'; ModuleVersion = '7.1.2'; },
                            @{ ModuleName = 'Microsoft.Graph.Authentication'; ModuleVersion = '1.5.0'; },
                            @{ ModuleName = 'Microsoft.Graph.Identity.DirectoryManagement'; ModuleVersion = '1.5.0'; },
                            @{ ModuleName = 'Microsoft.Graph.Groups'; ModuleVersion = '1.5.0'; },
                            @{ ModuleName = 'Microsoft.Graph.Users'; ModuleVersion = '1.5.0'; },
                            @{ ModuleName = 'ExchangeOnlineManagement'; ModuleVersion = '2.0.5'; },
                            @{ ModuleName = 'PnP.PowerShell'; ModuleVersion = '1.9.0'; MaximumVersion = '1.12.0'; }
                        )
    FunctionsToExport = @(
                            'Export-M365RMailboxSizes',
                            'Export-M365RMSOLUserLicenceBreakdown',
                            'Export-M365ROneDriveUsageReport',
                            'Export-M365RUserLicenceBreakdown',
                            'Get-M365RMailboxSizes',
                            'Get-M365ROneDriveUsageReport'
                        )
    CmdletsToExport = ''
    VariablesToExport = ''
    AliasesToExport = @(
                            'Export-M365RMSOLUserLicenseBreakdown',
                            'Export-M365RUserLicenseBreakdown'
                        )
    PrivateData = @{
        PSData = @{
            Tags = 'Office365', 'Microsoft365', 'Reporting', 'Licence', 'OneDrive'
            LicenseUri = 'https://github.com/JustinBarker77/Office365/blob/master/LICENSE'
            ProjectUri = 'https://github.com/JustinBarker77/Office365'
        }
    }
}
