$users = New-Object System.Collections.Generic.List[System.Object]
$outputPath = "OUTPUTPATH"

$headers = @{"Authorization" = "Bearer TOKEN"}

$csv = Import-Csv -path "INPUTPATH" -encoding utf8

foreach ($user in $csv)
{
    $SourceUPN = $user.SourceUPN
    $uri = 'https://graph.microsoft.com/beta/users/' + $SourceUPN
    $uri
    try {
        $data = Invoke-RestMethod -Method get -Uri $uri -Headers $headers
        $users.Add($data)
    }
    catch {
        "User $sourceUPN was not found"
    }
}

$users | Select-Object userprincipalname,onPremisesSamAccountName,onpremisesDomainName,onPremisesDistinguishedName | export-csv -Encoding UTF8 -Path $outputPath -NoClobber -NoTypeInformation
