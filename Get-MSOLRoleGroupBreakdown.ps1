#Requires -Version 5
<#
	.SYNOPSIS
		Name: Get-MSOLRoleGroupBreakdown.ps1
		The purpose of this script is is to export MSOL RoleGroups to excel

	.DESCRIPTION
		This script will log in to Office 365 and then create a RoleGroup report in excel, applying formating and an autofilter.

	.NOTES
		Version: 0.2
        	Updated: 14-08-2020	v0.2	Updated references
		Updated: 27-07-2020	v0.1	Initial draft

		Authors: Luke Allinson, Justin Barker

        TODO: Rewrite to use ImportExcel Module and Microsoft.Graph Module
		References:
			https://stackoverflow.com/questions/31183106/can-powershell-generate-a-plain-excel-file-with-multiple-sheets
			https://learn-powershell.net/2015/10/02/quick-hits-adding-a-hyperlink-to-excel-spreadsheet/
#>

[CmdletBinding()]
param (
	[Parameter(
		Mandatory,
		HelpMessage="Name of the Company you are running this against. This will form part of the output file name:")
		]
		[string]$CompanyName,
	[Parameter(
		Mandatory,
		HelpMessage = "The location you would like the final excel file to reside (please specify the full path):"
		)][ValidateScript({
			If (!(Test-Path -Path $_)) {
				Throw "The folder $_ does not exist"
			} Else {
				Return $true
			}
		})]
	[System.IO.DirectoryInfo]$OutputPath,
	[Parameter(
		HelpMessage = "Credentials to connect to Office 365 if not already connected:"
	)]
	[PSCredential]$Office365Credentials
)

Function Merge-CSVFiles {
	## Gather all CSV files in temporary folder
    $CsvFiles = Get-ChildItem ("$CSVPath\*") -Include *.csv
    $CsvFiles = $CsvFiles | Sort-Object Name
	## Create new Excel object
    $Excel = New-Object -ComObject excel.application
	$Excel.Visible = $False
	$Excel.SheetsInNewWorkbook = $CsvFiles.Count
	$Workbooks = $Excel.Workbooks.Add()
    $Worksheets = $Workbooks.worksheets
    $CSVSheet = 1
    $SummaryInfo = @()
	## Process each CSV file
    Foreach ($CSV in $CsvFiles) {
		$CSVFullPath = $CSV.FullName
		## Generate truncated sheet names (31 character limit)
        $SheetName = ($CSV.name -split "\.")[0]
        If ($SheetName.Length -gt 30) {
    		$TruncSheetName = ($SheetName.Split(" ") | ForEach-Object {$_.Substring(0,5)}) -join ""
        } Else {
            $TruncSheetName = $SheetName.Replace(" ","")
        }
        ### Select the worksheet and apply the truncated name
        $Worksheet = $Worksheets.Item($CSVSheet)
		$Worksheet.Name = $TruncSheetName
        ## Add title to sheet and apply formatting
        $Worksheet.Range("A1").Value = $SheetName
        $Worksheet.Range("A1").Font.Name = "Segoe UI"
        $Worksheet.Range("A1").Font.Size = "14"
        $Worksheet.Range("A1").Font.FontStyle = "Bold"
        $Worksheet.Range("A1").Font.ColorIndex = 55
        $TxtConnector = ("TEXT;" + $CSVFullPath)
        ## Import CSV file
        $CellRef = $worksheet.Range("A2")
		$Connector = $worksheet.QueryTables.Add($TxtConnector,$CellRef)
		$Worksheet.QueryTables.Item($Connector.Name).TextFileOtherDelimiter = ","
		$Worksheet.QueryTables.Item($Connector.Name).TextFileParseType  = 1
		$Worksheet.QueryTables.Item($Connector.Name).Refresh()
		$Worksheet.QueryTables.Item($Connector.Name).Delete()
        ## Capture sheet information for link processing later
        $SheetObj = New-Object PSObject
        $SheetObj | Add-Member NoteProperty -Name "Index" -Value $CSVSheet
        $SheetObj | Add-Member NoteProperty -Name "SheetName" -Value $SheetName
        $SheetObj | Add-Member NoteProperty -Name "TruncSheetName" -Value $TruncSheetName
        $SummaryInfo += $SheetObj
        $CSVSheet++
	}
    Write-Host "Applying formatting to Worksheets..." -ForegroundColor Magenta
    ForEach ($Worksheet in $Worksheets) {
        ## Add links to summary page
        If ($Worksheet.Name -eq "Summary") {
            Write-Host "Adding summary links to other worksheets..." -ForegroundColor Magenta
            ForEach ($Item in $SummaryInfo) {
                If ($Item.SheetName -ne "Summary") {
                    #Write-Host "$($Item.Index) --- $($Item.SheetName) --- $($Item.TruncSheetName)" -ForegroundColor Cyan
                    $SearchString = $Item.Sheetname
                    $Selection = $worksheet.Range("B3").EntireColumn
                    $Search = $Selection.find($SearchString,[Type]::Missing,[Type]::Missing,1)
                    $ResultCell = "B$($Search.Row)"
                    $worksheet.Hyperlinks.Add($worksheet.Range($ResultCell),"","$($Item.TruncSheetName)!A1","$($Item.SheetName)",$worksheet.Range($ResultCell).text)
                }
            }
            ## Move the summary sheet to the front of the workbook
            $worksheet.Move($Worksheets.Item(1))
        }
        ## Freeze the header row
		$Worksheet.Select()
		$Worksheet.Application.Activewindow.SplitRow = 2
		$Worksheet.Application.Activewindow.FreezePanes = $True
		## Format the header Row
		$Rows = $Worksheet.UsedRange.Rows.Count
		$Columns = $Worksheet.UsedRange.Columns.Count
        $Selection = $Worksheet.Range($Worksheet.Cells(2,1), $Worksheet.Cells(2,$Columns))
		$Selection.Font.Name = "Segoe UI"
		$Selection.Font.Size = 9
		$Selection.Font.FontStyle = "Bold"
        $Selection.Interior.ColorIndex = 48
		## Format data
        $Selection = $Worksheet.Range($Worksheet.Cells(3,1), $Worksheet.Cells($Rows,$Columns))
		$Selection.Font.Name = "Segoe UI"
		$Selection.Font.Size = 9
        ## Add filter and autofit
        $Selection = $Worksheet.Range($Worksheet.Cells(2,1), $Worksheet.Cells($Rows,$Columns))
        $Selection.AutoFilter()
        $Selection.EntireColumn.AutoFit()
    }
    ## Focus the Summary sheet and tidy up
	$Workbooks.Worksheets.Item("Summary").Select()
	$workbooks.SaveAs($XLOutput,51)
	$workbooks.Saved = $true
	$workbooks.Close()
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbooks) | Out-Null
	$excel.Quit()
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
	[System.GC]::Collect()
	[System.GC]::WaitForPendingFinalizers()
    Remove-Variable Excel
}

## Set constants
$Date = Get-Date -Format ddMMyyyy-HHmm
$Output = @()
$OutputPath = Get-Item $OutputPath
If ($OutputPath.FullName -notmatch '\\$') {
	$ExcelFilePath = $OutputPath.FullName + "\"
} Else {
	$ExcelFilePath = $OutputPath.FullName
}
## Output file name
$XLOutput= $ExcelFilePath + "$CompanyName - MSOL Roles - $Date.xlsx"
$CSVPath = $ExcelFilePath + $(-join ((65..90) + (97..122) | Get-Random -Count 14 | ForEach-Object {[char]$_}))
$CSVPath = (New-Item -Type Directory -Path $CSVPath).FullName
$LogFile1 = $CSVpath + "\Summary.csv"
## Check Office365 connectivity
Write-Host "Checking Connection to Office 365" -ForegroundColor Cyan
$Test365 = Get-MsolCompanyInformation -ErrorAction SilentlyContinue
If ($Null -eq $Test365) {
	Do {
		If ($Office365Credentials) {
		Connect-MsolService -Credential $Office365Credentials
		} Else {
			Connect-MsolService
		}
		$Test365 = Get-MsolCompanyInformation -ErrorAction silentlycontinue
	} while ($Null -eq $Test365)
}
Write-Host "Connected to Office 365..." -ForegroundColor Green
## Get all MSOL roles
$MSOLRole = Get-MsolRole
## Generate CSVs for each Role (with members)
Write-Host "Getting Role Members..." -ForegroundColor Cyan
ForEach ($Role in $MSOLRole) {
    $RoleName = $Role.Name
    $RoleName
    $RoleMembers = Get-MsolRoleMember -RoleObjectId $Role.ObjectId
    $RoleMembersCount = $RoleMembers.Count
    $LogFile2 = $CSVpath + "\" + $RoleName + ".csv"
    If ($RoleMembersCount -ne "") {
        $UserOutput = @()
        ForEach ($RoleMember in $RoleMembers) {
            If ($RoleMember.RoleMemberType -eq "User") {
                $MSOLUser = Get-MsolUser -UserPrincipalName $RoleMember.EmailAddress
                $LastDirSyncTime = $MSOLUser.LastDirSyncTime
            } Else {
                $LastDirSyncTime = "N/A"
            }
            $UserObj = New-Object PSObject
            $UserObj | Add-Member NoteProperty -Name "RoleMemberType" -Value $RoleMember.RoleMemberType
            $UserObj | Add-Member NoteProperty -Name "EmailAddress" -Value $RoleMember.EmailAddress
            $UserObj | Add-Member NoteProperty -Name "DisplayName" -Value $RoleMember.DisplayName
            $UserObj | Add-Member NoteProperty -Name "isLicensed" -Value $RoleMember.isLicensed
            $UserObj | Add-Member NoteProperty -Name "LastDirSyncTime" -Value $LastDirSyncTime
            $UserOutput += $UserObj
        }
        $UserOutput | Export-Csv -NoClobber -NoTypeInformation -Encoding UTF8 $LogFile2
    }
    $RoleObj = New-Object PSObject
    $RoleObj | Add-Member NoteProperty -Name "ObjectId" -Value $Role.ObjectId
    $RoleObj | Add-Member NoteProperty -Name "Name" -Value $Role.Name
    $RoleObj | Add-Member NoteProperty -Name "MemberCount" -Value $RoleMembersCount
    $Output += $RoleObj
}
$Output = $Output | Sort-Object Name
$Output | Export-Csv -NoClobber -NoTypeInformation -Encoding UTF8 $LogFile1
## Merge the CSV files into a single Excel workbook and tidy up
Write-Host "Merging CSV Files..." -ForegroundColor Cyan
Merge-CSVFiles -CSVPath $CSVPath -XLOutput $XLOutput | Out-Null
Write-Host "Tidying up CSV Files..." -ForegroundColor Cyan
Remove-Item $CSVPath -Recurse -Confirm:$false -Force
Write-Host "CSV Files Deleted" -ForegroundColor Green
Write-Host "Script Completed.  Results available in $XLOutput" -ForegroundColor Green
