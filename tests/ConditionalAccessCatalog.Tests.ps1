#Requires -Version 5.1
#Requires -Modules Pester

$modulePath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Modules\ConditionalAccessCatalog.psm1'
Import-Module $modulePath -Force

Describe 'Conditional Access catalog summaries' {
    It 'summarizes All users plus exclusions' {
        $users = [pscustomobject]@{
            includeUsers  = @('All')
            excludeUsers  = @('u1', 'u2')
            includeGroups = @('g1')
        }
        $s = Get-CaUserScopeSummary -Users $users
        $s | Should -Match 'Include: All users'
        $s | Should -Match 'Exclude users: 2'
        $s | Should -Match 'Include groups: 1'
    }

    It 'counts specific include users and omits empty roles' {
        $users = [pscustomobject]@{
            IncludeUsers = @('u1', 'u2', 'u3')
        }
        Get-CaUserScopeSummary -Users $users | Should -Be 'Include users: 3'
    }

    It 'returns none when user conditions are empty' {
        Get-CaUserScopeSummary -Users $null | Should -Be '(none)'
        Get-CaUserScopeSummary -Users ([pscustomobject]@{}) | Should -Be '(none)'
    }

    It 'maps policy location include and exclude ids' {
        $policy = [pscustomobject]@{
            id          = 'p1'
            displayName = 'Block countries'
            state       = 'enabled'
            conditions  = [pscustomobject]@{
                users     = [pscustomobject]@{ includeUsers = @('All') }
                locations = [pscustomobject]@{
                    includeLocations = @('All')
                    excludeLocations = @('loc-office')
                }
            }
        }
        $s = ConvertTo-CaPolicySummary -Policy $policy
        $s.id | Should -Be 'p1'
        $s.state | Should -Be 'enabled'
        $s.userScope | Should -Match 'All users'
        $s.includeLocationIds | Should -Be @('All')
        $s.excludeLocationIds | Should -Be @('loc-office')
        $s.exclusions | Should -Be '—'
    }

    It 'labels country and IP named locations' {
        $country = [pscustomobject]@{
            id                 = 'loc1'
            displayName        = 'Blocked countries'
            '@odata.type'      = '#microsoft.graph.countryNamedLocation'
            countriesAndRegions = @('RU', 'CN')
        }
        $c = ConvertTo-CaLocationSummary -Location $country
        $c.type | Should -Be 'Country blocklist'
        $c.countryMode | Should -Be 'blocklist'
        $c.countries | Should -Be @('RU', 'CN')
        $c.allowedCountries | Should -Not -Contain 'RU'
        $c.allowedCountries | Should -Not -Contain 'CN'
        $c.allowedCountries | Should -Contain 'US'
        $c.detail | Should -Match '^Allowed:'
        $c.detail | Should -Not -Match '\bRU\b'

        $ip = [pscustomobject]@{
            id            = 'loc2'
            displayName   = 'Office egress'
            '@odata.type' = '#microsoft.graph.ipNamedLocation'
            ipRanges      = @(@{ cidrAddress = '203.0.113.0/24' }, @{ cidrAddress = '198.51.100.0/24' })
        }
        $i = ConvertTo-CaLocationSummary -Location $ip
        $i.type | Should -Be 'IP range'
        $i.detail | Should -Be '203.0.113.0/24, 198.51.100.0/24'
        $i.ipRanges | Should -Be @('203.0.113.0/24', '198.51.100.0/24')
    }

    It 'extracts CIDRs from Graph ipRanges objects' {
        $cidrs = Get-CaIpRangeCidrs @(
            [pscustomobject]@{ cidrAddress = '10.0.0.0/8' },
            [pscustomobject]@{ '@odata.type' = '#microsoft.graph.iPv4CidrRange'; CidrAddress = '192.168.1.0/24' }
        )
        $cidrs | Should -Be @('10.0.0.0/8', '192.168.1.0/24')
    }

    It 'keeps trusted country allowlists as listed codes' {
        $trusted = [pscustomobject]@{
            id                  = 'loc3'
            displayName         = 'Trusted countries'
            '@odata.type'       = '#microsoft.graph.countryNamedLocation'
            countriesAndRegions = @('US', 'CA', 'GB')
        }
        $t = ConvertTo-CaLocationSummary -Location $trusted
        $t.type | Should -Be 'Country'
        $t.countryMode | Should -Be 'listed'
        $t.detail | Should -Be 'US, CA, GB'
        @($t.allowedCountries).Count | Should -Be 0
    }

    It 'inverts an unnamed country list as a blocklist' {
        $loc = [pscustomobject]@{
            id                  = 'loc4'
            displayName         = 'Named location 1'
            countriesAndRegions = @('RU', 'CN')
        }
        $s = ConvertTo-CaLocationSummary -Location $loc
        $s.countryMode | Should -Be 'blocklist'
        $s.allowedCountries | Should -Contain 'US'
        $s.allowedCountries | Should -Not -Contain 'RU'
        $s.allowedCountries | Should -Not -Contain 'CN'
    }

    It 'summarizes user and group exclusions' {
        $users = [pscustomobject]@{
            includeUsers  = @('All')
            excludeUsers  = @('u1')
            excludeGroups = @('g1', 'g2')
        }
        $e = Get-CaExclusionSummary -Users $users
        $e.excludeUserCount | Should -Be 1
        $e.excludeGroupCount | Should -Be 2
        $e.text | Should -Match 'Users: 1'
        $e.text | Should -Match 'Groups: 2'
    }

    It 'rejects unknown worker commands and missing policy id' {
        Invoke-CaCatalogWorkerCommand -Command 'CA_NOPE' | Should -Be 'CA_FAILED:Unknown CA command'
        Invoke-CaCatalogWorkerCommand -Command 'CA_GET_POLICY' | Should -Be 'CA_FAILED:Missing POLICY_ID'
    }

    It 'maps built-in role template ids to display names' {
        Get-CaWellKnownRoleName -Id '62e90394-69f5-4237-9190-012177145e10' | Should -Be 'Global Administrator'
        $list = ConvertTo-CaNamedList -Ids @('62e90394-69f5-4237-9190-012177145e10', 'All') -Kind role
        @($list).Count | Should -Be 2
        $list[0].label | Should -Be 'Global Administrator'
        $list[1].label | Should -Be 'All roles'
    }

    It 'keeps All users readable without Graph' {
        $list = ConvertTo-CaNamedList -Ids @('All', 'GuestsOrExternalUsers') -Kind user
        $list[0].label | Should -Be 'All users'
        $list[1].label | Should -Be 'Guests or external users'
    }

    It 'merges exclude users without duplicates' {
        $m = Merge-CaPolicyUserIds -Current @('u1') -Ids @('u1', 'u2') -Op add -Scope exclude
        @($m) | Should -Be @('u1', 'u2')
    }

    It 'replaces All when adding include users' {
        $m = Merge-CaPolicyUserIds -Current @('All') -Ids @('u1') -Op add -Scope include
        @($m) | Should -Be @('u1')
    }

    It 'rejects emptying include without fallback' {
        { Merge-CaPolicyUserIds -Current @('u1') -Ids @('u1') -Op remove -Scope include } | Should -Throw
    }

    It 'falls back to All users when allowed' {
        $m = Merge-CaPolicyUserIds -Current @('u1') -Ids @('u1') -Op remove -Scope include -AllowAllFallback
        @($m) | Should -Be @('All')
    }

    It 'parses JSON write payloads and rejects a short user search' {
        $p = Get-CaCatalogJsonPayload -Command 'CA_PATCH_USERS|{"policyId":"p1","scope":"exclude","op":"add"}'
        $p.policyId | Should -Be 'p1'
        $p.scope | Should -Be 'exclude'
        { Search-CaCatalogUsers -Query 'a' } | Should -Throw
        Invoke-CaCatalogWorkerCommand -Command 'CA_PATCH_USERS' | Should -Be 'CA_FAILED:Missing JSON payload'
    }
}
