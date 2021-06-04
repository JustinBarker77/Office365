#Requires -Modules MSOnline
#Requires -Version 5

<#
    .SYNOPSIS
        Name: Get-MSOLUserLicence-FullBreakdown.ps1
        The purpose of this script is is to export licensing details to excel

    .DESCRIPTION
        This script will log in to Microsoft 365 and then create a license report by SKU, with each component level status for each user, where 1 or more is assigned. This then conditionally formats the output to colours and autofilter.

    .NOTES
        Version 1.49
        Updated: 20210607    V1.49    Moved Translates to json files
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

        References:
            https://gallery.technet.microsoft.com/scriptcenter/Export-a-Licence-b200ca2a
            https://stackoverflow.com/questions/31183106/can-powershell-generate-a-plain-excel-file-with-multiple-sheets
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
        HelpMessage = 'Credentials to connect to Office 365 if not already connected'
    )]
    [PSCredential]$Office365Credentials,
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
#Helper function for merging CSV files and conditional formatting etc.
Function Merge-CSVFile
{
    $csvFiles = Get-ChildItem ("$CSVPath\*") -Include *.csv
    $Excel = New-Object -ComObject excel.application
    $Excel.visible = $false
    $Excel.sheetsInNewWorkbook = $csvFiles.Count
    $workbooks = $excel.Workbooks.Add()
    $CSVSheet = 1
    Foreach ($CSV in $Csvfiles)
    {
        $worksheets = $workbooks.worksheets
        $CSVFullPath = $CSV.FullName
        $SheetName = ($CSV.name -split '\.')[0]
        $worksheet = $worksheets.Item($CSVSheet)
        $worksheet.Name = $SheetName
        $TxtConnector = ('TEXT;' + $CSVFullPath)
        $CellRef = $worksheet.Range('A1')
        $Connector = $worksheet.QueryTables.add($TxtConnector, $CellRef)
        $worksheet.QueryTables.item($Connector.name).TextFileOtherDelimiter = "`t"
        $worksheet.QueryTables.item($Connector.name).TextFileParseType = 1
        $worksheet.QueryTables.item($Connector.name).Refresh()
        $worksheet.QueryTables.item($Connector.name).delete()
        $CSVSheet++
    }
    $worksheets = $workbooks.worksheets
    $xlTextString = [Microsoft.Office.Interop.Excel.XlFormatConditionType]::xlTextString
    $xlContains = [Microsoft.Office.Interop.Excel.XlContainsOperator]::xlContains
    foreach ($worksheet in $worksheets)
    {
        Write-Information ('Freezing panes on ' + $Worksheet.Name)
        $worksheet.Select()
        $worksheet.application.activewindow.splitcolumn = 1
        $worksheet.application.activewindow.splitrow = 1
        $worksheet.application.activewindow.freezepanes = $true
        $rows = $worksheet.UsedRange.Rows.count
        $columns = $worksheet.UsedRange.Columns.count
        $Selection = $worksheet.Range($worksheet.Cells(2, 5), $worksheet.Cells($rows, 6))
        [void]$Selection.Cells.Replace(';', "`n", [Microsoft.Office.Interop.Excel.XlLookAt]::xlPart)
        $Selection = $worksheet.Range($worksheet.Cells(1, 1), $worksheet.Cells($rows, $columns))
        $Selection.Font.Name = 'Segoe UI'
        $Selection.Font.Size = 9
        if ($Worksheet.Name -ne 'AllLicences')
        {
            Write-Information ('Setting Conditional Formatting on ' + $Worksheet.Name)
            $Selection = $worksheet.Range($worksheet.Cells(2, 6), $worksheet.Cells($rows, $columns))
            $Selection.FormatConditions.Add($xlTextString, '', $xlContains, 'Success', 'Success', 0, 0) | Out-Null
            $Selection.FormatConditions.Item(1).Interior.ColorIndex = 35
            $Selection.FormatConditions.Item(1).Font.ColorIndex = 51
            $Selection.FormatConditions.Add($xlTextString, '', $xlContains, 'Pending', 'Pending', 0, 0) | Out-Null
            $Selection.FormatConditions.Item(2).Interior.ColorIndex = 36
            $Selection.FormatConditions.Item(2).Font.ColorIndex = 12
            $Selection.FormatConditions.Add($xlTextString, '', $xlContains, 'Disabled', 'Disabled', 0, 0) | Out-Null
            $Selection.FormatConditions.Item(3).Interior.ColorIndex = 38
            $Selection.FormatConditions.Item(3).Font.ColorIndex = 30
        }
        else
        {
            foreach ($Item in (Import-Csv $CSVPath\AllLicences.csv -Delimiter "`t"))
            {
                if ($NoNameTranslation)
                {
                    $SearchString = $Item.'AccountLicenseSKU'
                    $Selection = $worksheet.Range('A2').EntireColumn
                    $Search = $Selection.find($SearchString, [Type]::Missing, [Type]::Missing, 1)
                    $ResultCell = "A$($Search.Row)"
                    $worksheet.Hyperlinks.Add($worksheet.Range($ResultCell), '', "`'$($SearchString)`'!A1", "$($SearchString)", $worksheet.Range($ResultCell).text)
                }
                else
                {
                    $SearchString = $Item.'AccountLicenseSKU(Friendly)'
                    $Selection = $worksheet.Range('A2').EntireColumn
                    $Search = $Selection.find($SearchString, [Type]::Missing, [Type]::Missing, 1)
                    $ResultCell = "A$($Search.Row)"
                    $worksheet.Hyperlinks.Add($worksheet.Range($ResultCell), '', "`'$($SearchString)`'!A1", "$($SearchString)", $worksheet.Range($ResultCell).text)
                }
            }
            $worksheet.Move($worksheets.Item(1))
        }
        [void]$worksheet.UsedRange.Autofilter()
        $worksheet.UsedRange.EntireColumn.AutoFit()
    }
    $workbooks.Worksheets.Item('AllLicences').Select()

    $workbooks.SaveAs($XLOutput, 51)
    $workbooks.Saved = $true
    $workbooks.Close()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbooks) | Out-Null
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

$date = Get-Date -Format yyyyMMdd
$OutputPath = Get-Item $OutputPath
if ($OutputPath.FullName -notmatch '\\$')
{
    $excelfilepath = $OutputPath.FullName + '\'
}
else
{
    $excelfilepath = $OutputPath.FullName
}
$XLOutput = $excelfilepath + "$CompanyName - $date.xlsx" ## Output file name
$CSVPath = $excelfilepath + $( -join ((65..90) + (97..122) | Get-Random -Count 14 | ForEach-Object { [char]$_ }))
$CSVPath = (New-Item -Type Directory -Path $csvpath).FullName
Write-Information 'Checking Connection to Office 365'
$test365 = Get-MsolCompanyInformation -ErrorAction silentlycontinue
if ($null -eq $test365)
{
    do
    {
        if ($Office365Credentials)
        {
            Connect-MsolService -Credential $Office365Credentials
        }
        else
        {
            Connect-MsolService
        }
        $test365 = Get-MsolCompanyInformation -ErrorAction silentlycontinue
    } while ($null -eq $test365)
}
Write-Information 'Connected to Office 365'
# Get a list of all licences that exist within the tenant
$licenseType = Get-MsolAccountSku
# Replace the above with the below if only a single SKU is required
#$licenseType = Get-MsolAccountSku | Where-Object {$_.AccountSkuID -like "*Power*"}
# Get all licences for a summary view
if ($NoNameTranslation)
{
    $licenseType | Where-Object { $_.TargetClass -eq 'User' } | Select-Object @{Name = 'AccountLicenseSKU'; Expression = { $($_.SkuPartNumber) } }, ActiveUnits, ConsumedUnits | Sort-Object 'AccountLicenseSKU' | Export-Csv $CSVPath\AllLicences.csv -NoTypeInformation -Delimiter `t
}
else
{
    $licenseType | Where-Object { $_.TargetClass -eq 'User' } | Select-Object @{Name = 'AccountLicenseSKU(Friendly)'; Expression = { (LicenceTranslate -SKU $_.SkuPartNumber -LicenceLevel Root) } }, ActiveUnits, ConsumedUnits | Sort-Object 'AccountLicenseSKU(Friendly)' | Export-Csv $CSVPath\AllLicences.csv -NoTypeInformation -Delimiter `t
}
$licenseType = $licenseType | Where-Object { $_.ConsumedUnits -ge 1 }
#get all users with licence
Write-Information 'Retrieving all licensed users - this may take a while.'
$alllicensedusers = Get-MsolUser -All | Where-Object { $_.isLicensed -eq $true }
$licensedGroups = @{}
# Loop through all licence types found in the tenant
foreach ($license in $licenseType)
{
    Write-Information ('Gathering users with the following subscription: ' + $license.accountskuid)
    # Gather users for this particular AccountSku from pre-existing array of users
    $users = $alllicensedusers | Where-Object { $_.licenses.accountskuid -contains $license.accountskuid }
    if ($NoNameTranslation)
    {
        $rootLicence = ($($license.SkuPartNumber))
    }
    else
    {
        $rootLicence = (LicenceTranslate -SKU $($license.SkuPartNumber) -LicenceLevel Root)
    }
    #$logFile = $CompanyName + "-" +$rootLicence + ".csv"
    $logFile = $CSVpath + '\' + $rootLicence + '.csv'
    $licensedUsers = New-Object System.Collections.Generic.List[System.Object]
    # Loop through all users and write them to the CSV file
    foreach ($user in $users)
    {
        Write-Verbose ('Processing ' + $user.displayname)
        $thislicense = $user.licenses | Where-Object { $_.accountskuid -eq $license.accountskuid }
        if ($user.BlockCredential -eq $true)
        {
            $enabled = $false
        }
        else
        {
            $enabled = $true
        }
        $userHashTable = @{
            DisplayName       = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            AccountEnabled    = $enabled
            AccountSKU        = $rootLicence
        }
        if ($thislicense.GroupsAssigningLicense.Count -eq 0)
        {
            $userHashTable['DirectAssigned'] = $true
            $userHashTable['GroupsAssigning'] = $false
        }
        else
        {
            if ($thislicense.GroupsAssigningLicense -contains $user.ObjectID)
            {
                $groups = $thislicense.groupsassigninglicense.guid | Where-Object { $_ -notlike $user.objectid }
                if ($null -eq $groups)
                {
                    $groups = $false
                }
                else
                {
                    foreach ($group in $groups)
                    {
                        if ($null -eq $licensedGroups[$group])
                        {
                            $getGroup = Get-MsolGroup -ObjectId $group
                            $licensedGroups[$group] = $getGroup.DisplayName
                        }
                    }
                    $groups = (Invoke-GroupGuidConversion -GroupGuid $groups -LicenseGroups $licensedGroups).DisplayName -Join ';'
                }
                $userHashTable['DirectAssigned'] = $true
                $userHashTable['GroupsAssigning'] = $groups
            }
            else
            {
                $groups = $thislicense.groupsassigninglicense.guid
                if ($null -eq $groups)
                {
                    $groups = $false
                }
                else
                {
                    foreach ($group in $groups)
                    {
                        if ($null -eq $licensedGroups[$group])
                        {
                            $getGroup = Get-MsolGroup -ObjectId $group
                            $licensedGroups[$group] = $getGroup.DisplayName
                        }
                    }
                    $groups = (Invoke-GroupGuidConversion -GroupGuid $groups -LicenseGroups $licensedGroups).DisplayName -Join ';'
                }
                $userHashTable['DirectAssigned'] = $false
                $userHashTable['GroupsAssigning'] = $groups
            }
        }
        foreach ($row in $($thislicense.ServiceStatus))
        {
            $serviceName = $(
                if ($NoNameTranslation)
                {
                    $($row.ServicePlan.ServiceName)
                }
                else
                {
                    LicenceTranslate -SKU $($row.ServicePlan.ServiceName) -LicenceLevel Component
                }
            )
            $userHashTable[$serviceName] = ($thislicense.ServiceStatus | Where-Object { $_.ServicePlan.ServiceName -eq $row.ServicePlan.ServiceName }).ProvisioningStatus
        }
        $licensedUsers.Add([PSCustomObject]$userHashTable) | Out-Null
    }
    $licensedUsers | Select-Object DisplayName, UserPrincipalName, AccountEnabled, AccountSKU, DirectAssigned, GroupsAssigning, * -ErrorAction SilentlyContinue | Export-Csv -Path $logFile -Delimiter "`t" -Encoding UTF8 -NoClobber -NoTypeInformation
}
Write-Information ('Merging CSV Files')
Merge-CSVFile -CSVPath $CSVPath -XLOutput $XLOutput | Out-Null
Write-Information ('Tidying up - deleting CSV Files')
Remove-Item $CSVPath -Recurse -Confirm:$false -Force
Write-Information ('CSV Files Deleted')
Write-Information ("Script Completed.  Results available in $XLOutput")
$InformationPreference = $initialInformationPreference
