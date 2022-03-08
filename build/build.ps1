$ErrorActionPreference = 'Stop'
$VerbosePreference = 'SilentlyContinue'
#Install-PackageProvider -Name NuGet -force | Out-Null
#Set-PSRepository -Name PSGallery -InstallationPolicy Trusted

[version]$CloudVersion = (Find-Module M365Reporting).ModuleVersion
[version]$LocalVersion = (Import-PowerShellDataFile '.\M365Reporting\M365Reporting.psd1').ModuleVersion

if ($LocalVersion -gt $CloudVersion)
{
    [version]$NewVersion = "{0}.{1}.{2}" -f $LocalVersion.Major, $LocalVersion.Minor, ($LocalVersion.Build + 1)
}
else
{
    [version]$NewVersion = "{0}.{1}.{2}" -f $CloudVersion.Major, $CloudVersion.Minor, ($CloudVersion.Build + 1)
}

"Cloud Version is $($CloudVersion)"
"Local Version is $($LocalVersion)"
"New Version is $($NewVersion)"

#Update-ModuleManifest -Path .\M365Reporting\M365Reporting.psd1 -ModuleVersion $NewVersion -ReleaseNotes $($env:COMMIT_MESSAGE)
