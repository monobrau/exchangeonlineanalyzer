#Requires -Version 5.1
#Requires -Modules Pester

$modulePath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Modules\Remediation.psm1'
Import-Module $modulePath -Force

Describe 'Remediation Graph error helpers' {
    It 'treats Request_ResourceNotFound as already gone' {
        $text = @'
Response status code does not indicate success: 404 (Not Found).
{
  "error": {
    "code": "Request_ResourceNotFound",
    "message": "Resource 'f1744981-85dc-46de-b8de-e4bc9ee53a4b' does not exist or one of its queried reference-property objects are not present."
  }
}
'@
        Test-RemediationGraphAlreadyGone $text | Should -Be $true
    }

    It 'does not treat a permission error as already gone' {
        $text = 'Graph DELETE failed (403): Authorization_RequestDenied'
        Test-RemediationGraphAlreadyGone $text | Should -Be $false
    }

    It 'detects Intune not licensed on the tenant' {
        $text = @'
Response status code does not indicate success: 400 (Bad Request).
{ "error": { "code": "BadRequest", "message": "Request not applicable to target tenant." } }
'@
        Test-RemediationIntuneUnavailable $text | Should -Be $true
        Test-RemediationGraphAlreadyGone $text | Should -Be $false
    }
}

Describe 'Remediation date conversion' {
    It 'unwraps EXO REST DateTime/Kind objects' {
        $wrapped = [pscustomobject]@{ DateTime = '2026-09-11T19:21:47'; Kind = 'Utc' }
        Convert-RemediationDateToText $wrapped | Should -Be '2026-09-11T19:21:47'
    }

    It 'formats UTC DateTime as ISO' {
        $dt = [datetime]::new(2026, 1, 15, 12, 0, 0, [datetimekind]::Utc)
        Convert-RemediationDateToText $dt | Should -Be '2026-01-15T12:00:00Z'
    }

    It 'reads WhenChangedUTC when WhenChanged is a blank wrapper' {
        $rule = [pscustomobject]@{
            WhenChanged    = [pscustomobject]@{ DateTime = ''; Kind = 'Utc' }
            WhenChangedUTC = '2026-09-04T08:00:00Z'
        }
        Get-RemediationObjectDateText -Object $rule -PropertyNames @('WhenChanged', 'WhenChangedUTC') | Should -Be '2026-09-04T08:00:00Z'
    }

    It 'reads Graph startDateTime from a credential object' {
        $cred = [pscustomobject]@{ startDateTime = '2026-03-01T00:00:00Z'; endDateTime = '2028-03-01T00:00:00Z' }
        Get-RemediationObjectDateText -Object $cred -PropertyNames @('startDateTime', 'StartDateTime') | Should -Be '2026-03-01T00:00:00Z'
    }
}
