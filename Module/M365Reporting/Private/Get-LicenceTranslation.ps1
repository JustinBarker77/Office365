#TODO: Update this to be capable of multiple language SKU translations
function Get-LicenceTranslation
{
    param
    (
        [parameter (Mandatory = $true, Position = 0)][string]$SKU,
        [parameter (Mandatory = $true, Position = 1)][ValidateSet('Component', 'Root')]$LicenceLevel
    )
    if ($LicenceLevel -eq 'Component')
    {
        if (-not (Get-Variable -Name ComponentTranslateCache -Scope Script -ErrorAction SilentlyContinue))
        {
            $file = 'ComponentLicenses.json'
            $Script:ComponentTranslateCache = Get-Content -Path ($PSScriptRoot + [IO.Path]::DirectorySeparatorChar + 'Translations' + [IO.Path]::DirectorySeparatorChar + 'SKUTranslations' + [IO.Path]::DirectorySeparatorChar + $file) | ConvertFrom-Json
        }
        $Translatation = $Script:ComponentTranslateCache
    }
    else
    {
        if (-not (Get-Variable -Name RootTranslateCache -Scope Script -ErrorAction SilentlyContinue))
        {
            $file = 'RootLicenses.json'
            $Script:RootTranslateCache = Get-Content -Path ($PSScriptRoot + [IO.Path]::DirectorySeparatorChar + 'Translations' + [IO.Path]::DirectorySeparatorChar + 'SKUTranslations' + [IO.Path]::DirectorySeparatorChar + $file) | ConvertFrom-Json
        }
        $Translatation = $Script:RootTranslateCache
    }

    [string]$translateString = $Translatation.$SKU
    if ($translateString)
    {
        Write-Output $translateString
    }
    else
    {
        Write-Output $SKU
    }
}
