#TODO: Update this to be capable of multiple language SKU translations
function Get-LicenceTranslation
{
    param
    (
        [parameter (Mandatory = $true, Position = 0)][string]$SKU,
        [parameter (Mandatory = $true, Position = 1)][ValidateSet('Component', 'Root')]$LicenceLevel
    )
    if (!$Script:translationFile)
    {
        $Script:translationFile = Invoke-RestMethod -Method Get -Uri "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv" | ConvertFrom-Csv
    }

    if ($LicenceLevel -eq 'Component')
    {
        [string]$translateString = $Script:translationFile.where({$_.Service_Plan_Name -eq $SKU})[0].Service_Plans_Included_Friendly_Names
    }
    else
    {
        [string]$translateString = $Script:translationFile.where({$_.String_Id -eq $SKU})[0].Product_Display_Name
    }

    if ($translateString)
    {
        Write-Output $translateString
    }
    else
    {
        Write-Output $SKU
    }
}
