#Requires -Modules ImportExcel, Microsoft.Graph.Authentication, Microsoft.Graph.Identity.DirectoryManagement, Microsoft.Graph.Groups, Microsoft.Graph.Users

<#
    .SYNOPSIS
        Name: Get-M365UserLicence-FullBreakdown.ps1
        The purpose of this script is is to export licensing details to excel

    .DESCRIPTION
        This script will log in to Microsoft 365 and then create a license report by SKU, with each component level status for each user, where 1 or more is assigned. This then conditionally formats the output to colours and autofilter.

    .NOTES
        Version 2.00
        Updated: 20210604    V2.00    Migrated to the Microsoft Graph PowerShell SDK Module and enabled cross-platform support, also added more SKUs
        Updated: 20210520    V1.48    1 tab = 4 spaces
        Updated: 20210520    V1.47    Added more components, renamed some components and added more SKUs
        Updated: 20210514    V1.46    Added more components, renamed some components and updated/added more SKUs
        Updated: 20210506    V1.45    Formatted Script to remove whitespace etc.
        Updated: 20210506    V1.44    Added Windows Update for Business Deployment Service component
        Updated: 20210323    V1.43    Added more Components
        Updated: 20210323    V1.42    Added more SKUs (F3, Conf PPM, E5 without Conf)
        Updated: 20210302    V1.41    Fixed missing New-Object's
        Updated: 20210223    V1.40    performance improvements for Group Based Licensing - no longer gets all groups; only gets the group once the GUID is found as an assigning group
        Updated: 20210222    V1.39    Added some EDU Root Level SKUs
        Updated: 20210222    V1.38    Moved Autofit and Autofilter to fix autofit on GBL column
        Updated: 20210208    V1.37    No longer out-files for everyline and performance improved
        Updated: 20201216    V1.36    Added components for Power Automate User with RPA Plan
        Updated: 20201216    V1.35    Added more SKUs (Multi-Geo, Communications Credits, M365 F1, Power Automate User with RPA Plan & Dynamics 365 Remote Assist)
        Updated: 20201028    V1.34    Added additional licence components (E5 Suite, PowerApps per IW, Win10 VDAE5)
        Updated: 20201021    V1.33    Resolved GBL issues
        Updated: 20201013    V1.32    Redid group based licensing to improve performance.
        Updated: 20201013    V1.31    Added User Enabled column
        Updated: 20200929    V1.30    Added RMS_Basic
        Updated: 20200929    V1.29    Added components for E5 Compliance
        Updated: 20200929    V1.28    Added code for group assigned and direct assigned licensing
        Updated: 20200820    V1.27    Added additional Office 365 E1 components
        Updated: 20200812    V1.26    Added Links to Licensing Sheets on All Licenses Page and move All Licenses Page to be first worksheet
        Updated: 20200730    V1.25    Added AIP P2 and Project for Office (E3 + E5)
        Updated: 20200720    V1.24    Added Virtual User component
        Updated: 20200718    V1.23    Added AAD Basic friendly component name
        Updated: 20200706    V1.22    Updated SKU error and added additional friendly names
        Updated: 20200626    V1.21    Updated F1 to F3 as per Microsoft's update
        Updated: 20200625    V1.20    Added Telephony Virtual User
        Updated: 20200603    V1.19    Added Switch for no name translation
        Updated: 20200603    V1.18    Added Telephony SKU's
        Updated: 20200501    V1.17    Script readability changes
        Updated: 20200430    V1.16    Made script more readable for Product type within component breakdown
        Updated: 20200422    V1.15    Formats to Segoe UI 9pt. Removed unnecessary True output.
        Updated: 20200408    V1.14    Added Teams Exploratory SKU
        Updated: 20200204    V1.13    Added more SKU's and Components
        Updated: 20191015    V1.12    Tidied up old comments
        Updated: 20190916    V1.11    Added more components and SKU's
        Updated: 20190830    V1.10    Added more components. Updated / renamed refreshed licences
        Updated: 20190627    V1.09    Added more Components
        Updated: 20190614    V1.08    Added more SKU's and Components
        Updated: 20190602    V1.07    Parameters, Comment based help, creates folder and deletes folder for csv's, require statements

        Release Date: 20190530
        Release notes from original:
            1.0 - Initital Release
            1.1 - Added Switch for additional licence components
            1.2 - Added PowerApps Plan 2 Trial  for additional licence components
            1.3 - Added Freeze Panes to Excel output
            1.4 - Added AX7 User Trial, Project Online Professional, Visio Online Plan 2, Office 365 E1, Whiteboard SKUs
            1.5 - Added Microsoft Search, Premium Encryption and Teams Commercial Trial, RMS Ad Hoc SKUs
            1.6 - Added Microsoft 365 E3 and F1 SKU & performs actions on cell by cell basis for colouring
        Authors: Mark Lofthouse, Justin Barker & Robin Dadswell

    .EXAMPLE
    PS> ./Get-M365UserLicense-FullBreakdown.ps1 -CompanyName "Contoso" -OutputPath c:\temp

    This example shows how to run the script interactively for the Company Contoso and gives friendly names for SKUs

    .EXAMPLE
    ./Get-M365UserLicense-FullBreakdown.ps1 -CompanyName "Contos" -OutputPath c:\temp -NoNameTranslation

    This example shows how to run the script interactively for the Company Contoso and does not give friendly names for SKUs
#>
[CmdletBinding()]
param (
    [Parameter(
        Mandatory,
        HelpMessage = 'Name of the Company you are running this against. This will form part of the output file name')
    ]
    [string]$CompanyName,
    [Parameter(
        Mandatory,
        HelpMessage = 'The location you would like the final excel file to reside'
    )][ValidateScript( {
            if (!(Test-Path -Path $_))
            {
                throw "The folder $_ does not exist"
            }
            else
            {
                return $true
            }
        })]
    [System.IO.DirectoryInfo]$OutputPath,
    [Parameter(
        HelpMessage = "This stops translation into Friendly Names of SKU's and Components"
    )][switch]$NoNameTranslation
)

#Enables Information Stream
$initialInformationPreference = $InformationPreference
$InformationPreference = 'Continue'

#Following Function Switches Complicated SKU Names with Friendly Names
function LicenceTranslate
{
    param
    (
        [parameter (Mandatory = $true, Position = 1)][string]$SKU,
        [parameter (Mandatory = $true, Position = 2)][ValidateSet('Component', 'Root')]$LicenceLevel
    )
    if ($LicenceLevel -eq 'Component')
    {
        if (-not (Get-Variable -Name ComponentTranslateCache -Scope Script -ErrorAction SilentlyContinue))
        {
            $file = 'ComponentLicenses.json'
            $Script:ComponentTranslateCache = Get-Content -Path ($PSScriptRoot + [IO.Path]::DirectorySeparatorChar + 'Translations' + [IO.Path]::DirectorySeparatorChar + 'SKUTranslations' + [IO.Path]::DirectorySeparatorChar + $file) | ConvertFrom-Json
        }
        $Translatation = $Script:ComponentTranslateCache
    }
    else
    {
        if (-not (Get-Variable -Name RootTranslateCache -Scope Script -ErrorAction SilentlyContinue))
        {
            $file = 'RootLicenses.json'
            $Script:RootTranslateCache = Get-Content -Path ($PSScriptRoot + [IO.Path]::DirectorySeparatorChar + 'Translations' + [IO.Path]::DirectorySeparatorChar + 'SKUTranslations' + [IO.Path]::DirectorySeparatorChar + $file) | ConvertFrom-Json
        }
        $Translatation = $Script:RootTranslateCache
    }

    [string]$translateString = $Translatation.$SKU
    if ($translateString)
    {
        Write-Output $translateString
    }
    else
    {
        Write-Output $SKU
    }
}

#Helper function for tidier select of Groups for Group Based Licensing
Function Invoke-GroupGuidConversion
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [String[]]
        $GroupGuid,
        [Parameter(Mandatory)]
        [hashtable]
        $LicenseGroups
    )
    $output = New-Object System.Collections.Generic.List[System.Object]
    foreach ($guid in $GroupGuid)
    {
        $temp = [PSCustomObject]@{
            DisplayName = $LicenseGroups[$guid]
        }
        $output.Add($temp)
        Remove-Variable temp
    }
    Write-Output $output
}

$date = Get-Date -Format yyyyMMdd
$OutputPath = Get-Item $OutputPath
if (!$OutputPath.FullName.EndsWith([IO.Path]::DirectorySeparatorChar))
{
    $excelfilepath = $OutputPath.FullName + [IO.Path]::DirectorySeparatorChar
}
else
{
    $excelfilepath = $OutputPath.FullName
}
$XLOutput = $excelfilepath + "$CompanyName - $date.xlsx" ## Output file name
if (Test-Path $XLOutput)
{
    Write-Error -Message "The file $XLOutput already exists, please delete this and rerun the script"
    return
}
Write-Information 'Checking Connection to Office 365'
$test365 = Get-MgOrganization -ErrorAction silentlycontinue
$preConnected = $test365

if ($null -eq $test365)
{
    do
    {
        Connect-MgGraph -Scopes 'Directory.Read.All'
        $test365 = Get-MgOrganization -ErrorAction silentlycontinue
    } while ($null -eq $test365)
}
else
{
    Write-Information ('You were already signed into the tenant ' + $test365.DisplayName + ' - if this was not intentional then please cancel this script, run Disconnect-Graph and try again')
    Start-Sleep -Seconds 5
}

$allowedScopes = 'Directory.ReadWrite.All', 'Directory.Read.All'
foreach ($scope in (Get-MgContext).Scopes)
{
    if ($allowedScopes -contains $scope)
    {
        $acceptableScopes = $true
        break
    }
}

if (!$acceptableScopes)
{
    Write-Error 'The scope of the current connection is not suitable for the permissions required within this script - please run Disconnect-Graph and try again'
    Start-Sleep -Seconds 10
    return
}
Write-Information ('Connected to Microsoft Graph for the tenant with a display name of ' + $test365.DisplayName)
# Get a list of all licences that exist within the tenant
#TODO: Test with a large number of different SKUs as Get-MGSubscribedSku doesn't support custom page sizes

$licenseType = Get-MgSubscribedSku
<#
    Replace the above with the below if only a single SKU is required
    $licenseType = Get-MgSubscribedSku -All | Where-Object {$_.AccountSkuID -like "*Power*"}
#>
# Get all licences for a summary view
if ($NoNameTranslation)
{
    $licenseType | Select-Object @{Name = 'AccountLicenseSKU'; Expression = { $($_.SkuPartNumber) } }, @{ Name = 'TotalLicenses'; Expression = { $_.PrepaidUnits.Enabled + $_.PrepaidUnits.Warning } }, @{ Name = 'LicensesInWarning'; Expression = { $_.PrepaidUnits.Warning } }, ConsumedUnits, AppliesTo | Sort-Object 'AccountLicenseSKU' | Export-Excel -Path $XLOutput -WorksheetName 'AllLicenses' -FreezeTopRowFirstColumn -AutoSize
}
else
{
    $licenseType | Select-Object @{Name = 'AccountLicenseSKU(Friendly)'; Expression = { (LicenceTranslate -SKU $_.SkuPartNumber -LicenceLevel Root) } }, @{ Name = 'TotalLicenses'; Expression = { $_.PrepaidUnits.Enabled + $_.PrepaidUnits.Warning } }, @{ Name = 'LicensesInWarning'; Expression = { $_.PrepaidUnits.Warning } }, ConsumedUnits, AppliesTo | Sort-Object 'AccountLicenseSKU(Friendly)' | Export-Excel -Path $XLOutput -WorksheetName 'AllLicenses' -FreezeTopRowFirstColumn -AutoSize
}
$licenseType = $licenseType | Where-Object { $_.ConsumedUnits -ge 1 -and $_.AppliesTo -eq 'User' }
#get all users with licence
Write-Information 'Retrieving all licensed users - this may take a while.'
$allLicensedUsers = Get-MgUser -All -Property Id, DisplayName, UserPrincipalName, AccountEnabled, LicenseAssignmentStates, AssignedLicenses | Where-Object { $null -ne $_.AssignedLicenses }
$licensedGroups = @{}
# Loop through all licence types found in the tenant
foreach ($license in $licenseType)
{
    Write-Information ('Gathering users with the following subscription: ' + $license.SkuPartNumber)
    # Gather users for this particular AccountSku from pre-existing array of users
    $users = $allLicensedUsers | Where-Object { $_.AssignedLicenses.SkuId -contains $license.SkuId }
    if ($NoNameTranslation)
    {
        $rootLicence = ($($license.SkuPartNumber))
    }
    else
    {
        $rootLicence = (LicenceTranslate -SKU $($license.SkuPartNumber) -LicenceLevel Root)
    }
    $licensedUsers = New-Object System.Collections.Generic.List[System.Object]
    # Loop through all users and write them to the CSV file
    foreach ($user in $users)
    {
        Write-Verbose ('Processing ' + $user.displayname)
        $thislicense = Get-MgUserLicenseDetail -UserId $user.Id -All | Where-Object { $_.SkuId -eq $license.SkuId }
        $userHashTable = @{
            DisplayName       = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            AccountEnabled    = $user.AccountEnabled
            AccountSKU        = $rootLicence
        }
        $licenseAssignmentStates = $user.LicenseAssignmentStates | Where-Object { $_.SkuId -eq $license.SkuId }
        if ($null -eq $licenseAssignmentStates.AssignedByGroup )
        {
            $userHashTable['DirectAssigned'] = $true
            $userHashTable['GroupsAssigning'] = $false
        }
        else
        {
            $groups = $licenseAssignmentStates.AssignedByGroup | Where-Object { $null -ne $_ }
            foreach ($group in $groups)
            {
                if ($null -eq $licensedGroups[$group])
                {
                    $getGroup = Get-MgGroup -GroupId $group
                    $licensedGroups[$group] = $getGroup.DisplayName
                }
            }
            #TODO: work out why line splitting isn't happening
            $groups = (Invoke-GroupGuidConversion -GroupGuid $groups -LicenseGroups $licensedGroups).DisplayName -Join "`r`n"
            if (($licenseAssignmentStates | Where-Object { $null -eq $_.AssignedByGroup }).Count -ge 1)
            {
                $userHashTable['DirectAssigned'] = $true
            }
            else
            {
                $userHashTable['DirectAssigned'] = $false
            }
            $userHashTable['GroupsAssigning'] = $groups
        }
        foreach ($row in $($thislicense.ServicePlans))
        {
            $serviceName = $(
                if ($NoNameTranslation)
                {
                    $($row.ServicePlanName)
                }
                else
                {
                    LicenceTranslate -SKU $($row.ServicePlanName) -LicenceLevel Component
                })
            $userHashTable[$serviceName] = $row.ProvisioningStatus
        }
        $licensedUsers.Add([PSCustomObject]$userHashTable) | Out-Null
    }
    $licensedUsers | Select-Object DisplayName, UserPrincipalName, AccountEnabled, AccountSKU, DirectAssigned, GroupsAssigning, * -ErrorAction SilentlyContinue | Export-Excel -Path $XLOutput -WorksheetName $RootLicence -FreezeTopRowFirstColumn -AutoSize -AutoFilter
}
Write-Information 'Formatting Excel Workbook'
$excel = Open-ExcelPackage -Path $XLOutput
foreach ($worksheet in $excel.Workbook.Worksheets)
{
    $fullRange = ($worksheet.Dimension | Select-Object Address).Address
    $worksheet.Select($fullRange)
    $worksheet.SelectedRange.Style.Font.Name = 'Segoe UI'
    $worksheet.SelectedRange.Style.Font.Size = 9
    if ($worksheet.Name -eq 'AllLicenses')
    {
        $formattingRange = "A2:A$($worksheet.Dimension.Rows)"
        $worksheet.Select($formattingRange)

        foreach ($cell in $worksheet.SelectedRange)
        {
            if ($excel.Workbook.Worksheets | Where-Object { $_.name -eq $cell.Value })
            {
                $referenceAddress = "`'$($cell.Value)`'!A1"
                $display = $($cell.Value)
                $hyperlink = New-Object -TypeName OfficeOpenXml.ExcelHyperLink -ArgumentList $referenceAddress, $display
                $cell.Hyperlink = $hyperlink
                $cell.Style.Font.Color.SetColor([System.Drawing.Color]::Blue)
                $cell.Style.Font.UnderLine = $true
            }
        }
    }
    else
    {
        $conditionalFormattingRange = $fullRange.Replace('A1', 'G2')
        Add-ConditionalFormatting -Worksheet $worksheet -RuleType ContainsText -ConditionValue 'Success' -BackgroundColor ([System.Drawing.Color]::FromArgb(204, 255, 204)) -BackgroundPattern Solid -ForegroundColor ([System.Drawing.Color]::FromArgb(0, 51, 0)) -Range $conditionalFormattingRange
        Add-ConditionalFormatting -Worksheet $worksheet -RuleType ContainsText -ConditionValue 'Pending' -BackgroundColor ([System.Drawing.Color]::FromArgb(255, 255, 153)) -BackgroundPattern Solid -ForegroundColor ([System.Drawing.Color]::FromArgb(128, 128, 0)) -Range $conditionalFormattingRange
        Add-ConditionalFormatting -Worksheet $worksheet -RuleType ContainsText -ConditionValue 'Disabled' -BackgroundColor ([System.Drawing.Color]::FromArgb(255, 153, 204)) -BackgroundPattern Solid -ForegroundColor ([System.Drawing.Color]::FromArgb(128, 0, 0)) -Range $conditionalFormattingRange
    }
    $worksheet.Select($fullRange)
    $worksheet.SelectedRange.AutoFitColumns()
    Set-Format -Address $worksheet.Column(6) -WrapText
    $worksheet.Select('A1')
    $excel.workbook.View.ActiveTab = 0
    $excel.Save()
}
$excel | Close-ExcelPackage
if (!$preConnected)
{
    Write-Information 'You were not connected prior to running this script so disconnecting'
    Disconnect-Graph
}
else
{
    Write-Information 'You were connected at the start of this script, skipping disconnection'
}
Write-Information ("Script Completed. Results available in $XLOutput")
$InformationPreference = $initialInformationPreference
