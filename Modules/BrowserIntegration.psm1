function Get-FirefoxProfiles {
    try {
        $profilesIni = Join-Path $env:APPDATA 'Mozilla\Firefox\profiles.ini'
        if (-not (Test-Path $profilesIni)) { return @() }
        $content = Get-Content $profilesIni -ErrorAction Stop
        $profiles = @()
        $current = @{}
        foreach ($line in $content) {
            if ($line -match '^\[Profile') { if ($current.Count) { $profiles += [pscustomobject]$current }; $current = @{}; continue }
            if ($line -match '^Name=(.*)$') { $current.Name = $Matches[1] }
            elseif ($line -match '^Path=(.*)$') { $current.Path = $Matches[1] }
            elseif ($line -match '^Default=(.*)$') { $current.Default = ($Matches[1] -eq '1') }
        }
        if ($current.Count) { $profiles += [pscustomobject]$current }
        return $profiles
    } catch { return @() }
}

function Get-FirefoxContainers {
    param([Parameter(Mandatory=$true)][string]$ProfilePath)
    try {
        $containersPath = Join-Path $ProfilePath 'containers.json'
        if (-not (Test-Path $containersPath)) { return @() }
        $json = Get-Content $containersPath -Raw | ConvertFrom-Json -ErrorAction Stop
        if ($json.identities) { return $json.identities }
        return @()
    } catch { return @() }
}

function Get-TenantIdentity {
    $result = [ordered]@{ TenantDisplayName=$null; Domains=@(); PrimaryDomain=$null }
    try {
        $ctx = Get-MgContext -ErrorAction Stop
        if ($ctx -and $ctx.TenantId) {
            try {
                $org = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization' -ErrorAction Stop
                if ($org.value -and $org.value[0]) {
                    $o = $org.value[0]
                    $result.TenantDisplayName = $o.displayName
                    if ($o.verifiedDomains) { $result.Domains = @($o.verifiedDomains | ForEach-Object { $_.name }) }
                    if ($o.verifiedDomains) { $result.PrimaryDomain = ($o.verifiedDomains | Where-Object { $_.isDefault -eq $true } | Select-Object -First 1 -ExpandProperty name) }
                }
            } catch {}
        }
    } catch {}
    if (-not $result.PrimaryDomain) {
        try { $Accepted = Get-AcceptedDomain -ErrorAction SilentlyContinue; if ($Accepted) { $result.PrimaryDomain = ($Accepted | Where-Object Default -eq $true | Select-Object -First 1 -Expand DomainName) } } catch {}
    }
    return [pscustomobject]$result
}

function Measure-StringSimilarity {
    param([string]$A,[string]$B)
    if ([string]::IsNullOrWhiteSpace($A) -or [string]::IsNullOrWhiteSpace($B)) { return 0.0 }
    $a = $A.ToLower(); $b = $B.ToLower()
    # Simple Jaccard of character bigrams
    $n=2
    $sa = for($i=0;$i -le $a.Length-$n;$i++){ $a.Substring($i,$n) }
    $sb = for($i=0;$i -le $b.Length-$n;$i++){ $b.Substring($i,$n) }
    $inter = (@($sa) + @($sb)) | Group-Object | Where-Object { $_.Count -gt 1 } | Measure-Object | Select-Object -ExpandProperty Count
    if (-not $inter) { $inter = 0 }
    $union = ([System.Collections.Generic.HashSet[string]]::new(($sa + $sb))).Count
    if ($union -eq 0) { return 0.0 }
    return [double]($inter/$union)
}

function Select-BestContainer {
    param(
        [Parameter(Mandatory=$true)]$Containers,
        [Parameter(Mandatory=$true)]$TenantIdentity
    )
    if (-not $Containers -or $Containers.Count -eq 0) { return $null }
    $candidates = @()
    $names = @()
    if ($TenantIdentity.TenantDisplayName) { $names += $TenantIdentity.TenantDisplayName }
    if ($TenantIdentity.PrimaryDomain) { $names += $TenantIdentity.PrimaryDomain }
    if ($TenantIdentity.Domains) { $names += $TenantIdentity.Domains }
    foreach ($cont in $Containers) {
        $best=0.0
        foreach ($n in $names) { $s = Measure-StringSimilarity -A $cont.name -B $n; if ($s -gt $best) { $best=$s } }
        $candidates += [pscustomobject]@{ Container=$cont; Score=$best }
    }
    $pick = $candidates | Sort-Object Score -Descending | Select-Object -First 1
    return $pick.Container
}

function Open-FirefoxUrlInProfile {
    param([Parameter(Mandatory=$true)][string]$ProfileName,[Parameter(Mandatory=$true)][string]$Url)
    $exe = 'firefox.exe'
    Start-Process -FilePath $exe -ArgumentList @('-P', $ProfileName, '-new-tab', $Url) -WindowStyle Normal | Out-Null
}

function Open-FirefoxUrlInContainer {
    param(
        [Parameter(Mandatory=$true)][string]$ProfileName,
        [Parameter(Mandatory=$true)][string]$ContainerName,
        [Parameter(Mandatory=$true)][string]$Url
    )
    # Attempt container-aware open via extension protocol (requires appropriate helper extension installed)
    $encoded = [System.Web.HttpUtility]::UrlEncode($Url)
    $extUrl = "ext+container:name=$ContainerName&url=$encoded"
    try { Open-FirefoxUrlInProfile -ProfileName $ProfileName -Url $extUrl } catch { Open-FirefoxUrlInProfile -ProfileName $ProfileName -Url $Url }
}

function Open-EntraDeepLink {
    param(
        [Parameter(Mandatory=$true)][string]$ProfileName,
        [Parameter(Mandatory=$false)][string]$ContainerName,
        [Parameter(Mandatory=$true)][ValidateSet('SignIns','RestrictedEntities','ConditionalAccess')][string]$Target
    )
    switch ($Target) {
        'SignIns'            { $url = 'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/SignInsMenuBlade/~/SignIns' }
        'RestrictedEntities' { $url = 'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/SecurityRestrictedEntitiesBlade/~/Overview' }
        'ConditionalAccess'  { $url = 'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/ConditionalAccessBlade/~/Policies' }
    }
    if ($ContainerName) { Open-FirefoxUrlInContainer -ProfileName $ProfileName -ContainerName $ContainerName -Url $url }
    else { Open-FirefoxUrlInProfile -ProfileName $ProfileName -Url $url }
}

Export-ModuleMember -Function Get-FirefoxProfiles,Get-FirefoxContainers,Get-TenantIdentity,Select-BestContainer,Open-FirefoxUrlInProfile,Open-FirefoxUrlInContainer,Open-EntraDeepLink


