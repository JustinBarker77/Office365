if ($PSEdition -eq 'Core' -and $IsWindows)
{
    Import-Module 'MSOnline', 'Microsoft.Online.SharePoint.PowerShell' -DisableNameChecking -UseWindowsPowerShell
}

foreach ($directory in @('Private', 'Public'))
{
    if ($PSEdition -eq 'Core' -and -not $IsWindows)
    {
        Get-ChildItem -Path $(Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath $directory) -ChildPath '*.ps1') -Exclude 'Export-MSOLUserLicenceBreakdown.ps1','Get-OneDriveUsageReport.ps1', 'Export-OneDriveUsageReport.ps1' | ForEach-Object { . $_.FullName }
    }
    else
    {
        Get-ChildItem -Path $(Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath $directory) -ChildPath '*.ps1') | ForEach-Object { . $_.FullName }
    }
}
