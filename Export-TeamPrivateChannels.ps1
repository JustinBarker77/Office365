#Requires -Module MicrosoftTeams
<#
TODO: Help information
#>

[CmdletBinding()]
param (
    [Parameter(
        Mandatory,
        HelpMessage = 'The location you would like the export csv to reside'
    )][ValidateScript( {
            if (!(Test-Path -Path $_))
            {
                throw "The folder $_ does not exist"
            }
            else
            {
                return $true
            }
        })]
    [System.IO.DirectoryInfo]$OutputPath
)
$date = Get-Date -Format yyyyMMdd_hh_mm
if (!$OutputPath.FullName.EndsWith([System.IO.Path]::DirectorySeparatorChar))
{
    $outputFolder = $OutputPath.FullName + [System.IO.Path]::DirectorySeparatorChar
}
else
{
    $outputFolder = $OutputPath.FullName
}

$filepath = $outputFolder + "$date - Teams_PrivateChannels.csv"

if (Test-Path $filepath -ErrorAction SilentlyContinue)
{
    Write-Error -Message "$filepath already exists - please delete the file and try again"
    return
}

$teams = Get-Team
$total = 0
$privatechannels = @()
$count = 0
foreach ($team in $teams)
{
    $count ++
    Write-Progress -Activity 'Enumerating Private Teams' -Status "Checking Team $count of $($teams.count)" -PercentComplete (($count / ($teams.count)) * 100)
    $private = Get-TeamChannel -GroupId $team.groupid -MembershipType Private
    $privateCount = $private.count
    if ($privateCount -gt 0)
    {
        $total = $total + $privateCount
        $privatechannels += $private
    }
}
$total

$privatechannels | Export-Csv $filepath -NoClobber -NoTypeInformation
