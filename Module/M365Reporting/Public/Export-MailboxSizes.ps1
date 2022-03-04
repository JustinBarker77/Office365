#Requires -Version 5 -Modules ExchangeOnlineManagement, ImportExcel
function Export-MailboxSizes
{
    <#
        .SYNOPSIS
            Name: Export-EXOMailboxSizes.ps1
            This gathers mailbox size information including primary and archive size and item count and exports to a XLSX file.
            NOTE: If you are not connected to Exchange Online Management prior to the running of this command then the created connection will not be preserved

        .DESCRIPTION
            This script connects to EXO and then outputs Mailbox statistics to a CSV file.

        .NOTES
            Version: 0.12
            Updated: 03-03-2022 v0.12   Replaced Export-Csv with Export-Excel (from ImportExcel Module) and added Summary page
            Updated: 01-03-2022 v0.11   Updated to an export command that calls the get command
            Updated: 01-03-2022 v0.10   Included a paramter to use an input CSV file
            Updated: 06-01-2022 v0.9    Changed output file date to match order of ISO8601 standard
            Updated: 10-11-2021 v0.8    Added parameter sets to prevent use of mutually exclusive parameters
                                        Disabled write-progress if the verbose parameter is used
            Updated: 10-11-2021 v0.7    Updated to include inactive mailboxes and improved error handling
            Updated: 08-11-2021 v0.6    Fixed an issue where archive stats are not included in output if the first mailbox does not have an archive
                                        Updated filename ordering
            Updated: 19-10-2021 v0.5    Updated to use Generic List instead of ArrayList
            Updated: 18-10-2021 v0.4    Updated formatting
            Updated: 15-10-2021 v0.3    Refactored for new parameters, error handling and verbose messaging
            Updated: 14-10-2021 v0.2    Rewritten to improve speed, remove superflous information
            Updated: <unknown>  v0.1    Initial draft

            Authors: Luke Allinson (github:LukeAllinson)
                    Robin Dadswell (github:RobinDadswell)

        .PARAMETER OutputPath
            Full path to the folder where the output will be saved.
            Can be used without the parameter name in the first position only.

        .PARAMETER InactiveMailboxOnly
            Only gathers information about inactive mailboxes (active mailboxes are not included in results).

        .PARAMETER IncludeInactiveMailboxes
            Include inactive mailboxes in results; these are not included by default.

        .PARAMETER RecipientTypeDetails
            Provide one or more RecipientTypeDetails values to return only mailboxes of those types in the results. Seperate multiple values by commas.
            Valid values are: EquipmentMailbox, GroupMailbox, RoomMailbox, SchedulingMailbox, SharedMailbox, TeamMailbox, UserMailbox.

        .PARAMETER MailboxFilter
            Provide a filter to reduce the size of the Get-EXOMailbox query; this must follow oPath syntax standards.
            For example:
            'EmailAddresses -like "*bruce*"'
            'DisplayName -like "*wayne*"'
            'CustomAttribute1 -eq "InScope"'

        .PARAMETER Filter
            Alias of MailboxFilter parameter.

        .PARAMETER InputCSV
            Full path and filename to an input CSV to specify which mailboxes will be included in the report.
            The CSV must contain a 'UserPrincipalName' or 'PrimarySmtpAddress' or 'EmailAddress' column/header.
            If multiple are found, 'UserPrincipalName' is preferred if found, otherwise 'PrimarySmtpAddress'; 'EmailAddress' is included to cater for exports from non-Exchange (e.g. HR) systems or manually created files.
            Note: All mailboxes are still retrieved and then compared to the CSV to ensure all requested information is captured.
            Note2: Progress is shown as overall progress of all mailboxes plus progress of CSV contents.

        .EXAMPLE
            .\Export-EXOMailboxSizes.ps1 C:\Scripts\
            Exports size information for all mailbox types

        .EXAMPLE
            .\Export-EXOMailboxSizes.ps1 -RecipientTypeDetails RoomMailbox,EquipmentMailbox -OutputPath C:\Scripts\
            Exports size information only for Room and Equipment mailboxes

        .EXAMPLE
            .\Export-EXOMailboxSizes.ps1 C:\Scripts\ -MailboxFilter 'Department -eq "R&D"'
            Exports size information for all mailboxes from the R&D department
    #>

    [CmdletBinding(DefaultParameterSetName = 'DefaultParameters')]
    [OutputType([string])]
    param
    (
        [Parameter(
            Mandatory,
            Position = 0,
            ParameterSetName = 'DefaultParameters'
        )]
        [Parameter(
            Mandatory,
            Position = 0,
            ParameterSetName = 'InactiveOnly'
        )]
        [Parameter(
            Mandatory,
            Position = 0,
            ParameterSetName = 'IncludeInactive'
        )]
        [Parameter(
            Mandatory,
            Position = 0,
            ParameterSetName = 'InputCSV'
        )]
        [ValidateNotNullOrEmpty()]
        [ValidateScript(
            {
                if (!(Test-Path -Path $_))
                {
                    throw "The folder $_ does not exist"
                }
                else
                {
                    return $true
                }
            }
        )]
        [IO.DirectoryInfo]
        $OutputPath,
        [Parameter(
            ParameterSetName = 'InactiveOnly'
        )]
        [switch]
        $InactiveMailboxOnly,
        [Parameter(
            ParameterSetName = 'IncludeInactive'
        )]
        [switch]
        $IncludeInactiveMailboxes,
        [Parameter(
            ParameterSetName = 'DefaultParameters'
        )]
        [Parameter(
            ParameterSetName = 'InactiveOnly'
        )]
        [Parameter(
            ParameterSetName = 'IncludeInactive'
        )]
        [ValidateSet(
            'EquipmentMailbox',
            'GroupMailbox',
            'RoomMailbox',
            'SchedulingMailbox',
            'SharedMailbox',
            'TeamMailbox',
            'UserMailbox'
        )]
        [string[]]
        $RecipientTypeDetails,
        [Parameter(
            ParameterSetName = 'DefaultParameters'
        )]
        [Parameter(
            ParameterSetName = 'InactiveOnly'
        )]
        [Parameter(
            ParameterSetName = 'IncludeInactive'
        )]
        [Alias('Filter')]
        [string]
        $MailboxFilter,
        [Parameter(
            ParameterSetName = 'InputCSV'
        )]
        [ValidateNotNullOrEmpty()]
        [ValidateScript(
            {
                if (!(Test-Path -Path $_))
                {
                    throw "The file $_ does not exist"
                }
                else
                {
                    return $true
                }
            }
        )]
        [IO.FileInfo]
        $InputCSV
    )

    $commandHashTable = @{}

    if ($IncludeInactiveMailboxes)
    {
        $commandHashTable['IncludeInactiveMailbox'] = $true
    }

    if ($InactiveMailboxOnly)
    {
        $commandHashTable['InactiveMailboxOnly'] = $true
    }

    if ($RecipientTypeDetails)
    {
        $commandHashTable['RecipientTypeDetails'] = $RecipientTypeDetails
    }

    if ($MailboxFilter)
    {
        $commandHashTable['Filter'] = $MailboxFilter
    }

    if ($InputCSV)
    {
        $commandHashTable['InputCSV'] = $InputCSV
    }

    if ($DeviceAuthentication)
    {
        $commandHashTable['DeviceAuthentication'] = $true
    }

    $timeStamp = Get-Date -Format yyyyMMdd-HHmm
    $outputFile = $OutputPath.FullName.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar + 'EXOMailboxSizes' + '_' + $timeStamp + '.xlsx'

    $output = Get-MailboxSizes @commandHashTable

    if ($output.Count -ge 1)
    {
        $output | Export-Excel -Path $outputFile -WorksheetName 'MailboxStats' -FreezeTopRow -AutoSize -AutoFilter -TableName 'MailboxStats'

        ### Add summary sheet and apply formatting
        $summaryPage = New-Object System.Collections.Generic.List[System.Object]
        $mailboxTypes = ('UserMailbox', 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox')
        $t = 3
        foreach ($mailboxType in $mailboxTypes)
        {
            $summaryStats = [ordered]@{
                'MailboxType'             = $mailboxType
                'MailboxCount'            = "=COUNTIF(MailboxStats[RecipientTypeDetails],A$t)"
                'TotalSize(MB)'           = "=SUMIF(MailboxStats[RecipientTypeDetails],A$t,MailboxStats[TotalItemSize(MB)])"
                'AverageSize(MB)'         = "=IF(C$t<>0,(AVERAGEIF(MailboxStats[RecipientTypeDetails],A$t,MailboxStats[TotalItemSize(MB)])),0)"
                'TotalItemCount'          = "=SUMIF(MailboxStats[RecipientTypeDetails],A$t,MailboxStats[ItemCount])"
                'AverageItemCount'        = "=IF(E$t<>0,(AVERAGEIF(MailboxStats[RecipientTypeDetails],A$t,MailboxStats[TotalItemSize(MB)])),0)"
                'ArchiveCount'            = "=COUNTIFS(MailboxStats[RecipientTypeDetails],A$t,MailboxStats[ArchiveStatus],""Active"")"
                'ArchiveTotalSize(MB)'    = "=SUMIF(MailboxStats[RecipientTypeDetails],A$t,MailboxStats[Archive_TotalItemSize(MB)])"
                'ArchiveAverageSize(MB)'  = "=IF(H$t<>0,(AVERAGEIF(MailboxStats[RecipientTypeDetails],A$t,MailboxStats[Archive_TotalItemSize(MB)])),0)"
                'ArchiveTotalItemCount'   = "=SUMIF(MailboxStats[RecipientTypeDetails],A$t,MailboxStats[Archive_ItemCount])"
                'ArchiveAverageItemCount' = "=IF(J$t<>0,(AVERAGEIF(MailboxStats[RecipientTypeDetails],A$t,MailboxStats[Archive_ItemCount])),0)"
            }
            $summaryPage.Add([PSCustomObject]$summaryStats) | Out-Null
            $t++
        }
        $summaryStats = [ordered]@{
            'MailboxType'             = 'Total'
            'MailboxCount'            = '=SUM(B3:B6)'
            'TotalSize(MB)'           = '=SUM(C3:C6)'
            'AverageSize(MB)'         = '=SUM(D3:D6)'
            'TotalItemCount'          = '=SUM(E3:E6)'
            'AverageItemCount'        = '=SUM(F3:F6)'
            'ArchiveCount'            = '=SUM(G3:G6)'
            'ArchiveTotalSize(MB)'    = '=SUM(H3:H6)'
            'ArchiveAverageSize(MB)'  = '=SUM(I3:I6)'
            'ArchiveTotalItemCount'   = '=SUM(J3:J6)'
            'ArchiveAverageItemCount' = '=SUM(K3:K6)'
        }
        $summaryPage.Add([PSCustomObject]$summaryStats) | Out-Null
        $summaryPage | Export-Excel -Path $outputFile -WorksheetName 'Summary' -TableName 'Summary' -AutoSize -StartRow 2 -MoveToStart
        $output | Sort-Object 'TotalItemSize(MB)' -Descending | Select-Object UserPrincipalName, DisplayName, RecipientTypeDetails, 'TotalItemSize(MB)', ItemCount -First 10 | Export-Excel -Path $outputFile -WorksheetName 'Summary' -TableName 'Top10BySize' -StartRow 10
        $output | Sort-Object ItemCount -Descending | Select-Object UserPrincipalName, DisplayName, RecipientTypeDetails, 'TotalItemSize(MB)', ItemCount -First 10 | Export-Excel -Path $outputFile -WorksheetName 'Summary' -TableName 'Top10ByItemCount' -StartRow 23
        $output | Sort-Object 'Archive_TotalItemSize(MB)' -Descending | Select-Object UserPrincipalName, DisplayName, RecipientTypeDetails, 'Archive_TotalItemSize(MB)', Archive_ItemCount -First 10 | Export-Excel -Path $outputFile -WorksheetName 'Summary' -TableName 'Top10ByArchiveSize' -StartRow 36
        $output | Sort-Object Archive_ItemCount -Descending | Select-Object UserPrincipalName, DisplayName, RecipientTypeDetails, 'Archive_TotalItemSize(MB)', Archive_ItemCount -First 10 | Export-Excel -Path $outputFile -WorksheetName 'Summary' -TableName 'Top10ByArchiveItemCount' -StartRow 49
        $excelPkg = Open-ExcelPackage -Path $outputFile
        $summarySheet = $excelPkg.Workbook.Worksheets['Summary']
        $summarySheet.Cells[1, 1].Value = 'Summary'
        $summarySheet.Cells[9, 1].Value = 'Top10 Mailboxes By Size'
        $summarySheet.Cells[22, 1].Value = 'Top10 Mailboxes By ItemCount'
        $summarySheet.Cells[35, 1].Value = 'Top10 Archives By Size'
        $summarySheet.Cells[48, 1].Value = 'Top10 Archives By ItemCount'
        $summarySheet.Select('A1')
        $summarySheet.SelectedRange.Style.Font.Size = 16
        $summarySheet.SelectedRange.Style.Font.Bold = $true
        $summarySheetTitleCells = ('A9', 'A22', 'A35', 'A48')
        ForEach ($cell in $summarySheetTitleCells)
        {
            $summarySheet.Select($cell)
            $summarySheet.SelectedRange.Style.Font.Size = 13
            $summarySheet.SelectedRange.Style.Font.Bold = $true
        }
        $summarySheetBoldRanges = ('A7:K7', 'D11:D20', 'E24:E33', 'D37:D46', 'E50:E59')
        ForEach ($range in $summarySheetBoldRanges)
        {
            $summarySheet.Select($range)
            $summarySheet.SelectedRange.Style.Font.Bold = $true
        }
        $summarySheetAverageRanges = ('D3:D7', 'F3:F7', 'I3:I7', 'K3:K7')
        ForEach ($range in $summarySheetAverageRanges)
        {
            $summarySheet.Select($range)
            $summarySheet.SelectedRange.Style.Numberformat.Format = '0.00'
        }
        $fullRange = ($summarySheet.Dimension | Select-Object Address).Address
        $summarySheet.Select($fullRange)
        $summarySheet.SelectedRange.AutoFitColumns()
        $summarySheet.Select('A1')
        Close-ExcelPackage $excelPkg
        return "Mailbox size data has been exported to $outputfile"
    }
    else
    {
        return 'No results returned; no data exported'
    }
}
