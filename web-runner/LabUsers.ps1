# People come from VScanMagic. Name, email, and Manage member identifier stay there.
# This file only reads that roster. Selection is the lab_user cookie.

function Get-LabPeopleUrl {
    if (-not [string]::IsNullOrWhiteSpace([string]$env:LAB_PEOPLE_URL)) {
        return [string]$env:LAB_PEOPLE_URL
    }
    return 'http://127.0.0.1:8815/api/people'
}

function Get-VScanMagicRosterUrl {
    if (-not [string]::IsNullOrWhiteSpace([string]$env:VSCANMAGIC_ROSTER_URL)) {
        return [string]$env:VSCANMAGIC_ROSTER_URL
    }
    return 'http://10.23.20.106:8080/api/users/roster'
}

function Get-LabUsers {
    $now = Get-Date
    if ($script:LabRoster -and ($now - $script:LabRosterAt).TotalSeconds -lt 15) {
        return @($script:LabRoster)
    }
    $resp = $null
    try {
        $resp = Invoke-RestMethod -Uri (Get-LabPeopleUrl) -TimeoutSec 8
    } catch {
        try {
            $resp = Invoke-RestMethod -Uri (Get-VScanMagicRosterUrl) -TimeoutSec 8
        } catch {
            throw "Could not load people. $($_.Exception.Message)"
        }
    }
    $rows = @()
    if ($resp.users) { $rows = @($resp.users) }
    $script:LabRoster = @($rows | ForEach-Object {
        $row = [ordered]@{
            id               = [string]$_.id
            name             = [string]$_.name
            email            = [string]$_.email
            memberIdentifier = [string]$_.memberIdentifier
        }
        if ($_.values) {
            foreach ($prop in $_.values.PSObject.Properties) {
                $row[$prop.Name] = [string]$prop.Value
            }
        }
        [pscustomobject]$row
    })
    $script:LabRosterAt = $now
    return @($script:LabRoster)
}

function Get-LabUser {
    param([string]$Id)
    if ([string]::IsNullOrWhiteSpace($Id)) { return $null }
    $want = $Id.Trim()
    return @(Get-LabUsers) | Where-Object { $_.id -eq $want } | Select-Object -First 1
}

function Save-LabUser {
    param(
        [string]$Name,
        [string]$Email = '',
        [string]$MemberIdentifier = '',
        [string]$Id = ''
    )
    throw 'Add and edit people in VScanMagic.'
}

function Remove-LabUser {
    param([string]$Id)
    throw 'Add and edit people in VScanMagic.'
}
