# Office365
Scripts stored within this repo are to aid in the management/reporting of Microsoft 365 generally through the creation of reports that are not natively within M365.

All scripts within this repo are offered as-is without warranty or support.

## Get-MSOLUserLicence-FullBreakdown.ps1
This PowerShell script is compatible with Windows PowerShell 5 & 5.1 but will not work on newer versions of PowerShell. This script pulls back all License SKU's in a Microsoft 365 Tenancy and then reports on a per SKU basis as to which components are assigned to which users in a color coded Excel Workbook in a Friendly manner which where possible (Excel Tab naming limits being the main limiting factor) translates the name to the name within the M365 admin portal. This script is retained for those cases where an enterprise application cannot be added to Azure AD (See [the issue logged here](https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/684) for more information.)

The original base scripts used to initially create from are referenced in the script itself. Co-authors for creating the updated script include [Mark Lofthouse](https://github.com/MarkLofthouse) and [Robin Dadswell](https://github.com/RobinDadswell).

This requires the following module(s) installed:
* ImportExcel (Minimum Version 7.1.2)
* MSOnline (Minimum Version 1.1.183.57)

This can be installed with the following command: 
```powershell
Install-Module ImportExcel, MSOnline
```

## Get-M365UserLicence-FullBreakdown.ps1
This PowerShell script is compatible with Windows PowerShell 5 and all newer versions of PowerShell. This script pulls back all License SKU's in a Microsoft 365 Tenancy and then reports on a per SKU basis as to which components are assigned to which users in a color coded Excel Workbook in a Friendly manner which where possible (Excel Tab naming limits being the main limiting factor) translates the name to the name within the M365 admin portal. 

This requires the following module(s) installed: 
* ImportExcel (Minimum Version 7.1.2)
* Microsoft.Graph.Authentication (Minimum Version 1.5.0)
* Microsoft.Graph.Identity.DirectoryManagement (Minimum Version 1.5.0)
* Microsoft.Graph.Groups (Minimum Version 1.5.0)
* Microsoft.Graph.Users (Minimum Version 1.5.0)

These can be installed with the following command: 

```powershell
Install-Module ImportExcel, Microsoft.Graph.Authentication, Microsoft.Graph.Identity.DirectoryManagement, Microsoft.Graph.Groups, Microsoft.Graph.Users -AllowClobber
```

Special thanks to [Doug Finke](https://github.com/dfinke) for the work on [ImportExcel](https://github.com/dfinke/ImportExcel).

### OS Support
This script has been tested on the following platforms (all with PowerShell 7.1.3 unless otherwise stated):
* Windows 10
* Windows 10 (Windows PowerShell 5.0)
* Ubuntu 20.04

### Notes
TODO: Write Pester tests and write pipeline for this script.
## Export-TeamPrivateChannels.ps1
Export all Teams Private Channels to CSV.

This requires the following module(s) to be installed:
* MicrosoftTeams

This can be installed with the following command:
```powershell
Install-Module MicrosoftTeams
```

## Get-MSOLRoleGroupBreakdown.ps1
This exports a spreadsheet of all role group members using the MSOnline module. This PowerShell script is compatible with Windows PowerShell 5 but will not work on newer versions of PowerShell.

This version of the script requires Microsoft Excel to be installed on the machine as it interacts with the COM Objects presented by this being installed and is not cross platform due to this. 

This requires the following module(s) to be installed:
* MSOnline

This can be installed with the following command:
```powershell
Install-Module MSOnline
```
### Notes
TODO: Rewrite to use ImportExcel and Microsoft.Graph modules
## Contributors & Developers
These scripts exist with thanks to all the people who contribute - please feel free to PR any changes/improvements. 
