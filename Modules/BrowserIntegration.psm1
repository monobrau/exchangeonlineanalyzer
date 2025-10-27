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

function Get-FirefoxProfilePathByName {
    param([Parameter(Mandatory=$true)][string]$ProfileName)
    $profiles = Get-FirefoxProfiles
    $prof = $profiles | Where-Object { $_.Name -eq $ProfileName } | Select-Object -First 1
    if (-not $prof) { return $null }
    $ppath = if ($prof.Path -like '*:*') { $prof.Path } else { Join-Path (Join-Path $env:APPDATA 'Mozilla\Firefox') $prof.Path }
    return $ppath
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

function Normalize-Name {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return '' }
    $s = $Name.ToLower()
    $s = ($s -replace '[^a-z0-9 ]',' ')
    $s = ($s -replace '\s+',' ').Trim()
    # Remove common suffixes
    $tokens = $s.Split(' ')
    $stop = @('inc','incorporated','llc','corp','corporation','company','co','group','ltd','limited','the')
    $filtered = $tokens | Where-Object { $_ -and ($stop -notcontains $_) }
    return ($filtered -join ' ').Trim()
}

function Compute-NameScore {
    param([string]$Candidate,[string]$Target)
    $c = Normalize-Name $Candidate
    $t = Normalize-Name $Target
    if (-not $c -or -not $t) { return 0.0 }
    # Base score: Jaccard bigram
    $score = Measure-StringSimilarity -A $c -B $t
    # Prefix boost
    $pref = 0.0
    $len = [Math]::Min($c.Length,$t.Length)
    $i=0; while($i -lt $len -and $c[$i] -eq $t[$i]){ $i++ }
    if ($len -gt 0) { $pref = $i / $len }
    if ($pref -gt $score) { $score = $pref }
    # Token overlap
    $ctoks = ($c -split ' ' | Where-Object { $_ })
    $ttoks = ($t -split ' ' | Where-Object { $_ })
    if ($ctoks.Count -gt 0 -and $ttoks.Count -gt 0) {
        $common = @($ctoks | Where-Object { $ttoks -contains $_ }).Count
        $tokScore = $common / [double]([Math]::Max($ctoks.Count,$ttoks.Count))
        if ($tokScore -gt $score) { $score = $tokScore }
        # Prefix token overlap (handles truncated last token: 'interio' vs 'interior')
        $prefMatches = 0
        foreach ($ct in $ctoks) {
            foreach ($tt in $ttoks) {
                if ([string]::IsNullOrEmpty($ct) -or [string]::IsNullOrEmpty($tt)) { continue }
                if ($tt.StartsWith($ct) -or $ct.StartsWith($tt)) { $prefMatches++; break }
            }
        }
        $prefTokScore = $prefMatches / [double]([Math]::Max($ctoks.Count,$ttoks.Count))
        if ($prefTokScore -gt $score) { $score = $prefTokScore }
    }
    # Substring/prefix heuristic to favor truncated names (e.g., 'creative business interio' vs 'creative business interiors')
    if ($c.Length -gt 0 -and $t.Length -gt 0) {
        $isPref = $t.StartsWith($c) -or $c.StartsWith($t)
        $isSub  = (-not $isPref) -and ($t.Contains($c) -or $c.Contains($t))
        if ($isPref -or $isSub) {
            $shared = [Math]::Min($c.Length,$t.Length)
            $baseRatio = $shared / [double]([Math]::Max($c.Length,$t.Length))
            $boost = [Math]::Min(1.0, $baseRatio + 0.1)
            if ($boost -gt $score) { $score = $boost }
        }
    }
    return [double]$score
}

function Get-LevenshteinDistance {
    param([string]$a,[string]$b)
    if ($a -eq $b) { return 0 }
    if ([string]::IsNullOrEmpty($a)) { return $b.Length }
    if ([string]::IsNullOrEmpty($b)) { return $a.Length }
    $m = $a.Length; $n = $b.Length
    $d = New-Object 'int[,]' ($m+1),($n+1)
    for($i=0;$i -le $m;$i++){ $d[$i,0]=$i }
    for($j=0;$j -le $n;$j++){ $d[0,$j]=$j }
    for($i=1;$i -le $m;$i++){
        for($j=1;$j -le $n;$j++){
            $cost = ([int]($a[$i-1] -ne $b[$j-1]))
            $d[$i,$j] = [Math]::Min([Math]::Min($d[$i-1,$j]+1,$d[$i,$j-1]+1), $d[$i-1,$j-1]+$cost)
        }
    }
    return $d[$m,$n]
}

function Get-BestContainerName {
    param(
        [Parameter(Mandatory=$true)][string[]]$ContainerNames,
        [Parameter(Mandatory=$true)]$TenantIdentity
    )
    $targets = @()
    if ($TenantIdentity.TenantDisplayName) { $targets += $TenantIdentity.TenantDisplayName }
    if ($TenantIdentity.PrimaryDomain) { $targets += $TenantIdentity.PrimaryDomain }
    if ($TenantIdentity.Domains) { $targets += $TenantIdentity.Domains }
    try { if ($TenantIdentity.PrimaryDomain) { $host = ($TenantIdentity.PrimaryDomain -split '\.')[0]; if ($host) { $targets += $host } } } catch {}
    $more = @(); foreach ($n in $targets) { if ($n -and $n.ToString().EndsWith('s')) { $more += $n.Substring(0,$n.Length-1) } }
    $targets += $more

    $bestName = $null; $bestScore = 0.0
    foreach ($name in $ContainerNames) {
        $localBest = 0.0
        foreach ($t in $targets) {
            $score = Compute-NameScore -Candidate $name -Target $t
            # Levenshtein normalized score boost
            $lev = Get-LevenshteinDistance -a (Normalize-Name $name) -b (Normalize-Name $t)
            $maxLen = [Math]::Max((Normalize-Name $name).Length,(Normalize-Name $t).Length)
            if ($maxLen -gt 0) { $levScore = 1.0 - ($lev / [double]$maxLen) } else { $levScore = 0.0 }
            if ($levScore -gt $score) { $score = $levScore }
            if ($score -gt $localBest) { $localBest = $score }
        }
        if ($localBest -gt $bestScore) { $bestScore = $localBest; $bestName = $name }
    }
    return [pscustomobject]@{ Name=$bestName; Score=[double]$bestScore }
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
    # Add primary domain host (leftmost label) as a candidate and singularized forms
    try {
        if ($TenantIdentity.PrimaryDomain) {
            $host = ($TenantIdentity.PrimaryDomain -split '\.')[0]
            if ($host) { $names += $host }
        }
    } catch {}
    $more = @()
    foreach ($n in $names) {
        if ($n -and $n.ToString().EndsWith('s')) { $more += $n.Substring(0, $n.Length-1) }
    }
    $names += $more
    foreach ($cont in $Containers) {
        $best=0.0
        foreach ($n in $names) { $s = Compute-NameScore -Candidate $cont.name -Target $n; if ($s -gt $best) { $best=$s } }
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
    # Attempt container-aware open via extension protocol (requires helper extension installed)
    $encodedUrl = [System.Web.HttpUtility]::UrlEncode($Url)
    $encodedName = [System.Web.HttpUtility]::UrlEncode($ContainerName)

    # Try to resolve cookieStoreId (cid) for a name with special characters
    $cid = $null
    try {
        $ppath = Get-FirefoxProfilePathByName -ProfileName $ProfileName
        if ($ppath) {
            $containers = Get-FirefoxContainers -ProfilePath $ppath
            $match = $containers | Where-Object { $_.name -and ($_.name -eq $ContainerName -or $_.name.ToLower() -eq $ContainerName.ToLower()) } | Select-Object -First 1
            if ($match -and $match.cookieStoreId) { $cid = $match.cookieStoreId }
        }
    } catch {}

    # Candidate invocations (encoded first, then raw), plus cid if available
    $candidates = @(
        "ext+container:name=$encodedName&url=$encodedUrl",
        "ext+container:container=$encodedName&url=$encodedUrl"
    )
    if ($cid) { $candidates += "ext+container:cid=$cid&url=$encodedUrl" }
    $candidates += @(
        "ext+container:name=$ContainerName&url=$encodedUrl",
        "ext+container:container=$ContainerName&url=$encodedUrl"
    )

    foreach ($c in $candidates) {
        try { Start-Process 'firefox.exe' -ArgumentList $c -WindowStyle Normal | Out-Null; return } catch {}
        try { Start-Process 'firefox.exe' -ArgumentList @('-P', $ProfileName, '-url', $c) -WindowStyle Normal | Out-Null; return } catch {}
        try { Start-Process 'firefox.exe' -ArgumentList @('-P', $ProfileName, '-new-tab', $c) -WindowStyle Normal | Out-Null; return } catch {}
    }

    # 4) Final fallback: open plain URL in profile
    Open-FirefoxUrlInProfile -ProfileName $ProfileName -Url $Url
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


