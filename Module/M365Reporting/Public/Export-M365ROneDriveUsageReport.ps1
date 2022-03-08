#requires -Version 5.1 -Modules ImportExcel
function Export-M365ROneDriveUsageReport
{
    <#
        .SYNOPSIS
            Generates a basic usage report for OneDrive for Business sites

        .DESCRIPTION
            Generates a basic usage report for OneDrive for Business sites
            The report will contain the following information:
            - Owner (UPN)
            - CurrentUsage (GB)
            - PercentageOfQuotaUsed
            - Quota (GB)
            - QuotaWarning (GB)
            - QuotaType
            - LastModified
            - URL
            - Status

        .PARAMETER TenantName
            The Tenant name, like contoso if the tenant is contoso.onmicrosoft.com
            vanity names, e.g. contoso.com, are NOT supported!

        .PARAMETER OutputPath
            The folder path that you would like the xlsx output to be placed

        .PARAMETER NoStatisticsReport
            This stops the statistics report summary page from being created

        .NOTES
            Quick and dirty implementation to generate a simple XLSX report file
            This script was sourced https://hochwald.net/post/powershell-usage-reporting-onedrive-business/ and has been updated to reflect our requirements
    #>
    [CmdletBinding(ConfirmImpact = 'None')]
    param
    (
        [Parameter(
            Mandatory)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript(
            {
                if ($_ -like '*.*')
                {
                    throw 'This should be the tenant name e.g. Contoso from contoso.onmicrosoft.com'
                }
                else
                {
                    return $true
                }
            }
        )]
        [Alias('Tenant', 'M365Name', 'M365TenantName')]
        [string]
        $TenantName,
        [Parameter()]
        [Alias('Device', 'DeviceAuthentication')]
        [switch]
        $DeviceLogin,
        [Parameter(
            Mandatory
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
            })]
        [IO.DirectoryInfo]
        $OutputPath,
        [switch]$NoStatisticsReport
    )

    $paramOneDriveUsageReport = @{
        TenantName = $TenantName
    }

    if ($DeviceLogin)
    {
        $paramOneDriveUsageReport['DeviceLogin'] = $true
    }

    $report = Get-M365ROneDriveUsageReport -TenantName $TenantName

    $report
    $outputFile = $OutputPath.FullName.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar + (Get-Date -Format yyyyMMdd_HHmmss) + '-' + $TenantName + '-' + 'OneDriveUsageReport.xlsx'

    if ($report.Count -ge 1)
    {
        # Export the CSV Report
        try
        {
            $paramExportExcel = @{
                Path                    = $outputFile
                WorksheetName           = 'OneDriveUsageReport'
                ErrorAction             = 'Stop'
                FreezeTopRowFirstColumn = $true
                AutoSize                = $true
                AutoFilter              = $true
            }
            $report | Sort-Object -Property CurrentUsage -Descending | Export-Excel @paramExportExcel
            if (-not $NoStatisticsReport)
            {
                $paramExportStatistics = @{
                    Path                    = $outputFile
                    WorksheetName           = 'OneDriveUsageStatistics'
                    ErrorAction             = 'Stop'
                    FreezeTopRowFirstColumn = $true
                    AutoSize                = $true
                    AutoFilter              = $true
                    Numberformat            = 'Percentage'
                    MoveToStart             = $true
                }

                $reportCount = $report.Count
                $reportStatistics = New-Object System.Collections.Generic.List[System.Object]
                $usageHashTable = @{
                    Statistic             = 'Less than 1GB Utilisation'
                    'Percentage of Users' = [Math]::Round((($report.CurrentUsage | Where-Object { $_ -lt 1 }).Count / $reportCount), 4)
                }

                $reportStatistics.Add([PSCustomObject]$usageHashTable) | Out-Null

                foreach ($numberOfDays in 30, 60, 90)
                {
                    $usageHashTable = @{
                        Statistic             = "Content not modified within the last $numberOfDays days"
                        'Percentage of Users' = [Math]::Round((($report.LastModifiedDate | Where-Object { (Get-Date $_) -lt (Get-Date).AddDays(-$numberOfDays) }).Count / $reportCount), 4)
                    }
                    $reportStatistics.Add([PSCustomObject]$usageHashTable) | Out-Null
                }

                $reportStatistics | Select-Object Statistic, 'Percentage of Users' | Export-Excel @paramExportStatistics
            }

        }
        catch
        {
            #region ErrorHandler
            # get error record
            [Management.Automation.ErrorRecord]$e = $_

            # retrieve information about runtime error
            $info = [PSCustomObject]@{
                Exception = $e.Exception.Message
                Reason    = $e.CategoryInfo.Reason
                Target    = $e.CategoryInfo.TargetName
                Script    = $e.InvocationInfo.ScriptName
                Line      = $e.InvocationInfo.ScriptLineNumber
                Column    = $e.InvocationInfo.OffsetInLine
            }

            # output information. Post-process collected info, and log info (optional)
            $info | Out-String | Write-Verbose

            $paramWriteError = @{
                Message      = $e.Exception.Message
                ErrorAction  = 'Stop'
                Exception    = $e.Exception
                TargetObject = $e.CategoryInfo.TargetName
            }
            Write-Error @paramWriteError
            #endregion ErrorHandler
        }
        Write-Output "Please find the report $outputFile"
    }
    else
    {
        Write-Output "There were no OneDrive's to report on"
    }
}
