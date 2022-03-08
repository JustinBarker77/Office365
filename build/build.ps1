$ErrorActionPreference = 'Stop'
$VerbosePreference = 'SilentlyContinue'
#Install-PackageProvider -Name NuGet -force | Out-Null
#Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
#Gets Cloud Version in case of major change
[version]$CloudVersion = (Find-Module M365Reporting).Version
#Gets Manifest Version in case of major change
[version]$LocalVersion = (Import-PowerShellDataFile '.\M365Reporting\M365Reporting.psd1').ModuleVersion

#Checks if Major Version Change to Module Manifest
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

#Updates Module Manifest ready for publish
Update-ModuleManifest -Path .\M365Reporting\M365Reporting.psd1 -ModuleVersion $NewVersion -ReleaseNotes $($env:COMMIT_MESSAGE)
