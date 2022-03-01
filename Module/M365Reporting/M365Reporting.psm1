if ($PSEdition -eq 'Core' -and $IsWindows)
{
    Import-WinModule 'MSOnline', 'Microsoft.Online.SharePoint.PowerShell' -DisableNameChecking
}

foreach ($directory in @('Private', 'Public'))
{
    if ($PSEdition -eq 'Core' -and -not $IsWindows)
    {
        Get-ChildItem -Path "$PSScriptRoot\$directory\*.ps1" -Exclude 'Get-OneDriveUsageReport.ps1', 'Export-OneDriveUsageReport.ps1' | ForEach-Object { . $_.FullName }
    }
    else
    {
        Get-ChildItem -Path "$PSScriptRoot\$directory\*.ps1" | ForEach-Object { . $_.FullName }
    }
}
