#Requires -Modules @{ ModuleName ="Pester"; ModuleVersion="5.2.0" }
if (!(Get-Module PSScriptAnalyzer -ListAvailable))
{
    Install-Module -Name PSScriptAnalyzer -Force
}

$folder = (Get-Item $PSScriptRoot).Parent.FullName
describe 'repo-level tests' {
    It "passes all PSScriptAnalyzer rules" {
        Invoke-ScriptAnalyzer -Path $folder -Recurse | should -BeNullOrEmpty
    }
}
