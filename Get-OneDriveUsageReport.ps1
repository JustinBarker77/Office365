#requires -Version 5.1 -Modules Microsoft.Online.SharePoint.PowerShell,ImportExcel
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

    .NOTES
        Quick and dirty implementation to generate a simple XLSX report file
        This script was sourced https://hochwald.net/post/powershell-usage-reporting-onedrive-business/ and has been updated to reflect our requirements
#>
[CmdletBinding(ConfirmImpact = 'None')]
param
(
    [Parameter(ValueFromPipeline,
        ValueFromPipelineByPropertyName,
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
    [Parameter(
        ValueFromPipeline,
        ValueFromPipelineByPropertyName,
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
    [switch]$StatisticsReport
)

begin
{
    # Garbage Collection
    [GC]::Collect()

    try
    {
        $paramImportModule = @{
            Name                = 'Microsoft.Online.SharePoint.PowerShell'
            DisableNameChecking = $true
            NoClobber           = $true
            Force               = $true
            ErrorAction         = 'SilentlyContinue'
            WarningAction       = 'Stop'
        }
        Import-Module @paramImportModule | Out-Null
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

        # Only here to catch a global ErrorAction overwrite
        exit 1
        #endregion ErrorHandler
    }

    # Create the Connection URI
    $AdminURL = ('https://' + $TenantName + '-admin.sharepoint.com')

    # Connect to SharePoint Online
    $paramConnectSPOService = @{
        Url         = $AdminURL
        Region      = 'Default'
        ErrorAction = 'Stop'
    }
    $null = (Connect-SPOService @paramConnectSPOService)

    # Create new object
    $Report = New-Object System.Collections.Generic.List[System.Object]
}

process
{
    $paramGetSPOSite = @{
        IncludePersonalSite = $true
        Limit               = 'all'
        Filter              = "Url -like '-my.sharepoint.com/personal/'"
        ErrorAction         = 'SilentlyContinue'
    }
    $Users = (Get-SPOSite @paramGetSPOSite)

    foreach ($user in $Users)
    {
        try
        {
            # Cleanup
            $StatsReport = $null

            # Get the dedicated Info for the user
            $paramGetSPOSite = @{
                Identity    = $user.Url
                ErrorAction = 'Stop'
            }

            # Create the Reporting object
            $StatsReport = [PSCustomObject]@{
                Owner                 = $user.Owner
                CurrentUsage          = '{0:F3}' -f ($user.StorageUsageCurrent / 1024) -as [decimal]
                PercentageOfQuotaUsed = ($user.StorageUsageCurrent / $user.StorageQuota) -as [decimal]
                Quota                 = '{0:F0}' -f ($user.StorageQuota / 1024) -as [int]
                QuotaWarning          = '{0:F0}' -f ($user.StorageQuotaWarningLevel / 1024) -as [int]
                QuotaType             = $user.StorageQuotaType
                LastModifiedDate      = $user.LastContentModifiedDate.ToLongDateString()
                URL                   = $user.URL
                Status                = $user.Status
            }

            # Append the report
            $Report.Add($StatsReport)
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

            Write-Warning -Message $info.Exception
            #endregion ErrorHandler
        }
    }
}

end
{
    # Create a Timestamp (check if this is OK for you)
    $TimeStamp = (Get-Date -Format yyyyMMdd_HHmmss)

    # Export the CSV Report
    try
    {
        $paramExportExcel = @{
            Path                    = ($OutputPath.FullName.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar + $TimeStamp + '-' + $TenantName + '-' + 'OneDriveUsageReport.xlsx')
            WorksheetName           = 'OneDriveUsageReport'
            ErrorAction             = 'Stop'
            FreezeTopRowFirstColumn = $true
            AutoSize                = $true
            AutoFilter              = $true
        }
        $Report | Sort-Object -Property CurrentUsage -Descending | Export-Excel @paramExportExcel
        if ($StatisticsReport)
        {
            $paramExportStatistics = @{
                Path                    = ($OutputPath.FullName.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar + $TimeStamp + '-' + $TenantName + '-' + 'OneDriveUsageReport.xlsx')
                WorksheetName           = 'OneDriveUsageStatistics'
                ErrorAction             = 'Stop'
                FreezeTopRowFirstColumn = $true
                AutoSize                = $true
                AutoFilter              = $true
            }
            $reportCount = $Report.Count
            $reportStatistics = New-Object System.Collections.Generic.List[System.Object]
            $usageHashTable = @{
                Statistic        = 'Less than 1GB Utilisation'
                'OneDrive Usage' = [Math]::Round((($Report.CurrentUsage | Where-Object { $_ -lt 1 }).Count / $reportCount), 4)
            }
            $reportStatistics.Add([PSCustomObject]$usageHashTable) | Out-Null
            foreach ($numberOfDays in 30, 60, 90)
            {
                $usageHashTable = @{
                    Statistic        = "Content not modified within the last $numberOfDays days"
                    'OneDrive Usage' = [Math]::Round((($report.LastModifiedDate | Where-Object { (Get-Date $_) -lt (Get-Date).AddDays(-$numberOfDays) }).Count / $reportCount), 4)

                }
                $reportStatistics.Add([PSCustomObject]$usageHashTable) | Out-Null
            }

            $reportStatistics | Export-Excel @paramExportStatistics
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
    finally
    {
        # Cleanup
        $Report = $null

        # Disconnect from SharePoint Online
        (Disconnect-SPOService -ErrorAction SilentlyContinue) | Out-Null

        # Garbage Collection
        [GC]::Collect()
    }
}


