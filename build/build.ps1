$ErrorActionPreference = 'Stop'
$VerbosePreference = 'SilentlyContinue'

#Setup PS Repository for easy install
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
#Gets Cloud Version in case of major change
[version]$CloudVersion = (Find-Module M365Reporting).Version
#Gets Manifest Version in case of major change
$manifest = Import-PowerShellDataFile '.\M365Reporting\M365Reporting.psd1'
[version]$LocalVersion = $manifest.ModuleVersion

#Checks if Major Version Change to Module Manifest
if ($LocalVersion -gt $CloudVersion)
{
    [version]$NewVersion = "{0}.{1}.{2}" -f $LocalVersion.Major, $LocalVersion.Minor, ($LocalVersion.Build)
}
else
{
    [version]$NewVersion = "{0}.{1}.{2}" -f $CloudVersion.Major, $CloudVersion.Minor, ($CloudVersion.Build + 1)
}

#Installs Requires Modules
foreach ($module in $manifest.RequiredModules.ModuleName)
{
    Install-Module -Name $module -Confirm:$false -Force -AllowClobber
}

#Updates Module Manifest ready for publish
Update-ModuleManifest -Path .\M365Reporting\M365Reporting.psd1 -ModuleVersion $NewVersion -ReleaseNotes $($env:COMMIT_MESSAGE)
