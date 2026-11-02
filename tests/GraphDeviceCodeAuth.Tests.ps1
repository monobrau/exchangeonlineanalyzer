#Requires -Version 5.1
#Requires -Modules Pester

$modulePath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Modules\GraphOnline.psm1'
Import-Module $modulePath -Force

Describe 'Graph device-code auth helpers' {
    It 'detects browser/WAM failure messages that should fall back to device code' {
        Test-MgGraphBrowserAuthFailure 'InteractiveBrowserCredential authentication failed' | Should -Be $true
        Test-MgGraphBrowserAuthFailure 'AADSTS500014 unrelated business error' | Should -Be $false
        Test-MgGraphBrowserAuthFailure 'no DISPLAY available for browser' | Should -Be $true
    }

    It 'resolves a PowerShell executable without requiring .exe suffix' {
        $exe = Get-EOAPowerShellExecutable
        $exe | Should -Not -BeNullOrEmpty
        Test-Path -LiteralPath $exe | Should -Be $true
    }
}
