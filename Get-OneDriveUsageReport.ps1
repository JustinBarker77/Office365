#requires -Version 5.1 -Modules Microsoft.Online.SharePoint.PowerShell,ImportExcel
<#
    .SYNOPSIS
        Generates a basic usage report for OneDrive for Business sites

    .DESCRIPTION
        Generates a basic usage report for OneDrive for Business sites
        The report will contain the following information:
        - Owner (UPN)
        - CurrentUsage (GB)
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
        This script was sourced https://hochwald.net/post/powershell-usage-reporting-onedrive-business/ and has been updated to reflect our requirments
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
    $OutputPath
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
            $Stats = $null
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
        ($Report | Sort-Object -Property CurrentUsage -Descending | Export-Excel @paramExportExcel)
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
        $Excel = $nul

        # Disconnect from SharePoint Online
        (Disconnect-SPOService -ErrorAction SilentlyContinue) | Out-Null

        # Garbage Collection
        [GC]::Collect()
    }
}

#region LICENSE
<#
      BSD 3-Clause License
      Copyright (c) 2021,  enabling Technology
      All rights reserved.
      Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
      1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
      2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
      3. Neither the name of the copyright holder nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.
      THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
#>
#endregion LICENSE

#region DISCLAIMER
<#
      DISCLAIMER:
      - Use at your own risk, etc.
      - This is open-source software, if you find an issue try to fix it yourself. There is no support and/or warranty in any kind
      - This is a third-party Software
      - The developer of this Software is NOT sponsored by or affiliated with Microsoft Corp (MSFT) or any of its subsidiaries in any way
      - The Software is not supported by Microsoft Corp (MSFT)
      - By using the Software, you agree to the License, Terms, and any Conditions declared and described above
      - If you disagree with any of the terms, and any conditions declared: Just delete it and build your own solution
#>
#endregion DISCLAIMER
