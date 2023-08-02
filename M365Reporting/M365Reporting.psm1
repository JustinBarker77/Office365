foreach ($directory in @('Private', 'Public'))
{
    Get-ChildItem -Path $(Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath $directory) -ChildPath '*.ps1') | ForEach-Object { . $_.FullName }
}
