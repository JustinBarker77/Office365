#requires -Version 5.1 -Modules PnP.PowerShell
function Get-OneDriveUsageReport
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

        .NOTES
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
        $TenantName
    )

    process
    {
        # Garbage Collection
        [GC]::Collect()


        # Create the Connection URI
        $AdminURL = ('https://' + $TenantName + '-admin.sharepoint.com')

        #TODO: Better cross platform authentication support (e.g. for headless environments etc.)
        # Connect to SharePoint Online
        $paramConnectPnPPowerShell = @{
            URL = $AdminURL
            ErrorAction    = 'Stop'
            Interactive    = $true
        }
        try
        {
            Get-PnPTenant -ErrorAction Stop | Out-Null
            $preConnected = $true
            Write-Verbose 'Previously connected to SPO Service, no extra authentication to occur'
        }
        catch
        {
            Write-Verbose 'Not Connected to SPO Service, connecting to SPO Service'
            Connect-PnPOnline @paramConnectPnPPowerShell | Out-Null
        }
        try
        {
            Get-PnPTenant -ErrorAction Stop | Out-Null
        }
        catch
        {
            return
        }


        # Create new object
        $Report = New-Object System.Collections.Generic.List[System.Object]

        $paramGetO4BSites = @{
            IncludeOneDriveSites = $true
            Filter              = "Url -like '-my.sharepoint.com/personal/'"
            ErrorAction         = 'SilentlyContinue'
        }
        $sites = (Get-PnPTenantSite @paramGetO4BSites)

        foreach ($site in $sites)
        {
            try
            {
                # Cleanup
                $StatsReport = $null

                # Create the Reporting object
                $StatsReport = [PSCustomObject]@{
                    Owner                 = $site.Owner
                    CurrentUsage          = '{0:F3}' -f ($site.StorageUsageCurrent / 1024) -as [decimal]
                    PercentageOfQuotaUsed = ($site.StorageUsageCurrent / $site.StorageQuota) -as [decimal]
                    Quota                 = '{0:F0}' -f ($site.StorageQuota / 1024) -as [int]
                    QuotaWarning          = '{0:F0}' -f ($site.StorageQuotaWarningLevel / 1024) -as [int]
                    QuotaType             = $site.StorageQuotaType
                    LastModifiedDate      = $site.LastContentModifiedDate.ToLongDateString()
                    URL                   = $site.URL
                    Status                = $site.Status
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
        # Disconnect from SharePoint Online
        try
        {
            if (-not $preConnected)
            {
                Write-Verbose 'Not previously connected to SPO Service, disconnecting from SPO'
                (Disconnect-PnPOnline -ErrorAction SilentlyContinue) | Out-Null
            }
            else
            {
                Write-Verbose 'Previously connected to SPO, will remain connected to SPO'
            }
        }
        catch
        {
            [GC]::Collect()
        }

        # Garbage Collection
        [GC]::Collect()
        return $Report
    }
}
