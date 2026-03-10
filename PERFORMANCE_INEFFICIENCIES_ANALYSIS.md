# Performance Inefficiencies Analysis

## Overview
This document identifies inefficiencies in processing large files/lists and opportunities to replace loops with more efficient operations.

---

## 1. ARRAY CONCATENATION INEFFICIENCIES ⚡ HIGH IMPACT

### Problem
Using `+=` operator on arrays creates a new array each time, copying all existing elements. This is O(n) for each operation.

### Found Instances

#### ExportUtils.psm1
- **Line 117**: `$mfaPolicies += $p` (in loop)
- **Line 136**: `$userPage += $u` (in loop)
- **Line 172**: `$userGroups += $m.Id` (in loop)
- **Line 173**: `$userRoles += $m.Id` (in loop)
- **Line 203**: `$users += [pscustomobject]@{...}` (in loop)
- **Line 244**: `$users += $u` (in loop)
- **Line 262**: `$groups += $name` (in loop)
- **Line 411-420**: Multiple `$collectorNames += '...'` operations
- **Line 797**: `$deviceIdsFromSignIns += $signIn.DeviceId` (in loop)
- **Line 1066**: `$deviceIdsFromSignIns += $signIn.DeviceId` (in loop)
- **Line 1887**: `$usersToExport += $mfaUser` (in loop)
- **Line 1893**: `$usersToExport += [pscustomobject]@{...}` (in loop)
- **Line 2091**: `$selectedUserList += $upn` (in loop)
- **Line 2184**: `$uniqueResults += $item` (in loop)
- **Line 2260**: `$mailboxes += $mbx` (in loop)
- **Line 2330**: `$result += [pscustomobject]@{...}` (in loop)
- **Line 2520**: `$userIds += $mgUser.Id` (in loop)
- **Line 2598**: `$targets += "..."` (in loop)
- **Line 2608**: `$modProps += "..."` (in loop)
- **Line 2615**: `$details += "..."` (in loop)
- **Line 2705**: `$pageCount += $results.Count` (in loop)
- **Line 2728**: `$pageCount += $results.Count` (in loop)
- **Line 2828**: `$userIds += $user.Id` (in loop)
- **Line 2893-2894**: `$deviceParts += ...` (in loop)
- **Line 2994-3038**: Multiple `$caPolicyNames += ...`, `$caPolicyResults += ...`, `$caPolicyDetails += ...` (in nested loops)
- **Line 3056-3058**: `$authMethods += ...` (in loop)
- **Line 3082-3084**: `$locParts += ...` (in loop)
- **Line 3176-3180**: `$userUpns += ...` (in loop)
- **Line 3223**: `$flattenedDevices += $deviceRecord` (in loop)
- **Line 3273**: `$mailboxes += $mbx` (in loop)
- **Line 3356**: `$mailboxes += $mbx` (in loop)
- **Line 4245**: `$nonEmptyCsvJsonFiles += $file` (in loop)
- **Line 4290**: `$zipPaths += $zipPath` (in loop)
- **Line 4298**: `$fileGroups += ,$group` (in loop)
- **Line 4318**: `$zipPaths += $zipPath` (in loop)
- **Line 4530-4532**: `$licenses += ...` (in loop)
- **Line 4571**: `$users += $user` (in loop)
- **Line 4599-4601**: `$licenses += ...` (in loop)
- **Line 4609**: `$report += [PSCustomObject]@{...}` (in loop)
- **Line 488**: `$allLogs += $enhancedLog` (in EntraInvestigator.psm1, line 488)

#### ExchangeOnlineAnalyzer.ps1
- **Line 255**: `$candidateUpns += $upnVal` (in loop)
- **Line 288**: `$autoKeywords += $host` (in loop)
- **Line 292-294**: `$allKw += ...` (multiple operations)
- **Line 618-692**: Multiple `$note += "..."` string concatenations (should use StringBuilder)
- **Line 792**: `$selectedAccounts += [PSCustomObject]@{...}` (in loop)
- **Line 816-1203**: Multiple `$report += "..."` string concatenations (should use StringBuilder)
- **Line 1044-1046**: `$nonUSSignIns += $log`, `$usSignIns += $log` (in loop)
- **Line 4135**: `$allFoundUsers += $users` (in loop)
- **Line 4751**: `$allFoundMailboxes += $mailboxes` (in loop)
- **Line 4776**: `$script:allLoadedMailboxUPNs += $mbx.UserPrincipalName` (in loop)
- **Line 7845**: `$allFoundUsers += $users` (in loop)

#### BulkTenantExporter.ps1
- **Line 95**: `$usersAlt1 += Get-MgUser ...` (in loop)
- **Line 100**: `$usersAlt2 += Get-MgUser ...` (in loop)
- **Line 158**: `$allFoundUsers += $users` (in loop)
- **Line 1196**: `$allFoundUsers += $users` (in runspace script)
- **Line 1336**: `$usersAlt1 += Get-MgUser ...` (in runspace script)
- **Line 1340**: `$usersAlt2 += Get-MgUser ...` (in runspace script)
- **Line 1369**: `$allFoundUsers += $users` (in runspace script)
- **Line 3601**: `$authenticatedClients += $clientNum` (in loop)

### Solution
Replace with `ArrayList` or `List<T>`:

```powershell
# Instead of:
$users = @()
foreach ($user in $allUsers) {
    $users += $user  # Creates new array each time!
}

# Use:
$users = [System.Collections.ArrayList]::new()
foreach ($user in $allUsers) {
    [void]$users.Add($user)  # O(1) operation
}
# Or PowerShell 5.1+:
$users = New-Object System.Collections.ArrayList
```

### Impact
- **For 1000 items**: Current = 500,000 operations, Optimized = 1,000 operations
- **Performance gain**: ~500x faster for large collections

---

## 2. LINEAR SEARCH INEFFICIENCIES ⚡ HIGH IMPACT

### Problem
Using `-contains` or `-in` on arrays performs O(n) linear search. For lookups in loops, this becomes O(n²).

### Found Instances

#### BulkTenantExporter.ps1 (from OPTIMIZATION_ANALYSIS.md)
- **Lines 6085, 6101, 6111, 6210, 6234**: `$completedTenants -contains $i` (O(n) search in loop)

#### ExchangeOnlineAnalyzer.ps1
- **Line 1044-1046**: Filtering sign-ins by location (could use hashtable lookup)
- **Line 4140**: `Sort-Object UserPrincipalName -Unique` after building array (could deduplicate during build)

#### ExportUtils.psm1
- **Line 2184**: Building `$uniqueResults` by checking duplicates (could use hashtable)
- Multiple instances where arrays are searched for membership

### Solution
Use hashtables for O(1) lookups:

```powershell
# Instead of:
$completedTenants = @()
if ($completedTenants -contains $i) { continue }  # O(n)
$completedTenants += $i  # O(n)

# Use:
$completedTenants = @{}
if ($completedTenants.ContainsKey($i)) { continue }  # O(1)
$completedTenants[$i] = $true  # O(1)
```

### Impact
- **For 10 tenants, 100 iterations**: Current = 5,000 comparisons, Optimized = 1,000 hash lookups
- **Performance gain**: ~5x faster

---

## 3. STRING CONCATENATION INEFFICIENCIES ⚡ MEDIUM IMPACT

### Problem
String concatenation with `+=` creates new string objects each time, copying all existing characters.

### Found Instances

#### ExchangeOnlineAnalyzer.ps1
- **Lines 618-692**: Building `$note` with 30+ `+=` operations
- **Lines 816-1203**: Building `$report` with 100+ `+=` operations
- **Line 2598**: `$targets += "..."` (in loop)
- **Line 2608**: `$modProps += "..."` (in loop)
- **Line 2615**: `$details += "..."` (in loop)

#### ExportUtils.psm1
- **Line 2598**: `$targets += "..."` (in loop)
- **Line 2608**: `$modProps += "..."` (in loop)
- **Line 2615**: `$details += "..."` (in loop)
- **Line 2994-3038**: Building policy strings with `+=` in nested loops

### Solution
Use `StringBuilder` for large string operations:

```powershell
# Instead of:
$report = ""
$report += "Section 1`n"
$report += "Section 2`n"
# ... 100+ operations

# Use:
$sb = [System.Text.StringBuilder]::new()
[void]$sb.AppendLine("Section 1")
[void]$sb.AppendLine("Section 2")
$report = $sb.ToString()
```

### Impact
- **For 100 concatenations**: Current = 5,050 character copies, Optimized = 100 appends
- **Performance gain**: ~50x faster for large strings

---

## 4. FILE I/O INEFFICIENCIES ⚡ MEDIUM IMPACT

### Problem
Repeated file operations (Test-Path, Get-Content) without caching.

### Found Instances

#### BulkTenantExporter.ps1
- **Lines 784, 836, 900, 914, 950, 962, 973, 985, 996**: Multiple `Test-Path` calls for same paths
- **Lines 1969, 1971**: `Test-Path` and `Get-Content` called every 2 seconds in monitoring loop
- **Lines 2728, 2731, 2882, 2884, 2893, 2894, 2932, 2935**: Repeated `Test-Path` and `Get-Content` calls
- **Line 1971**: `Get-Content $statusFilePath -Tail 5` reads entire file every iteration

#### ExportUtils.psm1
- Status file reading in monitoring loops (if any)

### Solution
Cache file existence and track read positions:

```powershell
# Instead of:
while ($running) {
    if (Test-Path $statusFile) {  # Every iteration!
        $content = Get-Content $statusFile -Tail 5
    }
    Start-Sleep -Seconds 2
}

# Use:
$lastReadPosition = 0
$fileInfo = [System.IO.FileInfo]::new($statusFile)
while ($running) {
    if ($fileInfo.Exists -and $fileInfo.LastWriteTime -gt $lastReadTime) {
        $stream = [System.IO.File]::OpenRead($statusFile)
        $stream.Position = $lastReadPosition
        $newContent = $stream.ReadToEnd()
        $lastReadPosition = $stream.Position
        $stream.Close()
        $lastReadTime = $fileInfo.LastWriteTime
    }
    Start-Sleep -Seconds 2
}
```

### Impact
- **For monitoring loop**: Reduces file system calls by ~95%
- **Performance gain**: Eliminates unnecessary I/O overhead

---

## 5. INEFFICIENT FILTERING/SORTING ⚡ MEDIUM IMPACT

### Problem
Filtering large collections multiple times or sorting unnecessarily.

### Found Instances

#### ExchangeOnlineAnalyzer.ps1
- **Line 4087**: `$users = @($users1) + @($users2) + @($users3) + @($users4) | Sort-Object UserPrincipalName -Unique`
  - Creates 4 arrays, concatenates, then sorts
  - Could use HashSet during collection
- **Line 4126**: Same pattern repeated
- **Line 4140**: `$uniqueUsers = $allFoundUsers | Sort-Object UserPrincipalName -Unique`
  - Sorting entire collection when could deduplicate during build
- **Line 4759**: `$uniqueMailboxes = $allFoundMailboxes | Sort-Object UserPrincipalName -Unique`
  - Same issue

#### ExportUtils.psm1
- **Line 2184**: Building unique results by checking array membership (O(n²))
- Multiple instances of `Sort-Object -Unique` after building arrays

### Solution
Use HashSet for deduplication during collection:

```powershell
# Instead of:
$allFoundUsers = @()
foreach ($term in $searchTerms) {
    $users = Get-MgUser -Filter "..."
    $allFoundUsers += $users  # May have duplicates
}
$uniqueUsers = $allFoundUsers | Sort-Object UserPrincipalName -Unique  # Expensive!

# Use:
$uniqueUsers = [System.Collections.Generic.HashSet[string]]::new()
foreach ($term in $searchTerms) {
    $users = Get-MgUser -Filter "..."
    foreach ($user in $users) {
        [void]$uniqueUsers.Add($user.UserPrincipalName)  # Automatic deduplication
    }
}
# No sorting needed if order doesn't matter, or sort only once at end
```

### Impact
- **For 1000 users, 10 search terms**: Current = 10,000 items sorted, Optimized = 1,000 items in HashSet
- **Performance gain**: ~10x faster, eliminates duplicates during collection

---

## 6. NESTED LOOPS / O(n²) OPERATIONS ⚡ HIGH IMPACT

### Problem
Nested loops create O(n²) or worse complexity.

### Found Instances

#### ExportUtils.psm1
- **Lines 2994-3038**: Nested loops building CA policy strings
  ```powershell
  foreach ($signIn in $signIns) {
      foreach ($policy in $policies) {
          $caPolicyNames += $policyName  # In nested loop!
      }
  }
  ```
- **Line 2184**: Checking array membership in loop (O(n²))
- **Line 1850**: Creating lookup dictionary - this is GOOD, but could be optimized further

#### ExchangeOnlineAnalyzer.ps1
- **Lines 4073-4140**: Searching for users with multiple fallback methods (could be optimized)
- **Lines 4742-4759**: Searching mailboxes with fallback (could be optimized)

### Solution
Pre-compute lookups and use hashtables:

```powershell
# Instead of:
foreach ($signIn in $signIns) {
    foreach ($policy in $policies) {
        if ($signIn.AppliedPolicies -contains $policy.Id) {
            $caPolicyNames += $policy.Name
        }
    }
}

# Use:
# Pre-build lookup
$policyLookup = @{}
foreach ($policy in $policies) {
    $policyLookup[$policy.Id] = $policy.Name
}

# Single loop
foreach ($signIn in $signIns) {
    foreach ($appliedPolicyId in $signIn.AppliedPolicies) {
        if ($policyLookup.ContainsKey($appliedPolicyId)) {
            $caPolicyNames.Add($policyLookup[$appliedPolicyId])
        }
    }
}
```

### Impact
- **For 1000 sign-ins, 50 policies**: Current = 50,000 iterations, Optimized = 1,000 + 50 = 1,050 operations
- **Performance gain**: ~47x faster

---

## 7. REPEATED API CALLS ⚡ MEDIUM IMPACT

### Problem
Making the same API calls multiple times or not caching results.

### Found Instances

#### ExchangeOnlineAnalyzer.ps1
- **Lines 4080-4086**: Making 4 separate `Get-MgUser` calls with different case variations
  - Could use single call with better filter or client-side filtering
- **Lines 4110-4118**: Fetching ALL users for client-side filtering (could cache)

#### BulkTenantExporter.ps1
- **Lines 95, 100**: Multiple `Get-MgUser` calls in fallback logic
- **Lines 1336, 1340**: Same pattern in runspace scripts

### Solution
Cache API results and use single optimized call:

```powershell
# Instead of:
$users1 = Get-MgUser -Filter "startsWith(DisplayName,'$term')" ...
$users2 = Get-MgUser -Filter "startsWith(DisplayName,'$termLower')" ...
$users3 = Get-MgUser -Filter "startsWith(DisplayName,'$termUpper')" ...
$users4 = Get-MgUser -Filter "startsWith(DisplayName,'$termTitle')" ...

# Use:
# Single call, then client-side case-insensitive filtering
$allUsers = Get-MgUser -Filter "startsWith(DisplayName,'$term') or startsWith(UserPrincipalName,'$term')" -All
$users = $allUsers | Where-Object { 
    $_.DisplayName -ilike "*$term*" -or $_.UserPrincipalName -ilike "*$term*"
}
```

### Impact
- **For 4 search terms**: Current = 16 API calls, Optimized = 4 API calls
- **Performance gain**: 4x fewer API calls, faster execution

---

## 8. PROCESS MONITORING INEFFICIENCIES ⚡ MEDIUM IMPACT

### Problem
Repeatedly calling `Get-Process` when process objects already have status.

### Found Instances

#### BulkTenantExporter.ps1 (from OPTIMIZATION_ANALYSIS.md)
- **Line 6105**: `Get-Process -Id $tenantProc.Process.Id` called every 2 seconds
- **Lines 1914, 1992, 2834**: Multiple `Get-Process` calls

### Solution
Use process object properties directly:

```powershell
# Instead of:
while ($running) {
    $proc = Get-Process -Id $process.Id -ErrorAction SilentlyContinue
    if (-not $proc) { break }
    Start-Sleep -Seconds 2
}

# Use:
while ($running) {
    try {
        if ($process.HasExited) { break }
        # Process still running, no need to call Get-Process
    } catch {
        # Process already exited
        break
    }
    Start-Sleep -Seconds 2
}
```

### Impact
- **For 10 tenants, 2-hour monitoring**: Current = 36,000 Get-Process calls, Optimized = ~10 calls
- **Performance gain**: ~99.97% reduction in system calls

---

## SUMMARY OF RECOMMENDATIONS

### High Priority (Quick Wins)
1. ✅ Replace array `+=` with `ArrayList` (50+ instances)
2. ✅ Replace `-contains` lookups with hashtables (5+ instances)
3. ✅ Optimize nested loops with pre-computed lookups (3+ instances)
4. ✅ Use `StringBuilder` for large string building (2 major instances)

### Medium Priority (Good ROI)
5. ⚠️ Cache file I/O operations (20+ instances)
6. ⚠️ Use HashSet for deduplication during collection (5+ instances)
7. ⚠️ Reduce repeated API calls (10+ instances)
8. ⚠️ Optimize process monitoring (4+ instances)

### Low Priority (Nice to Have)
9. 📝 Cache DateTime operations
10. 📝 Throttle UI updates
11. 📝 Extract magic numbers to constants

---

## ESTIMATED PERFORMANCE IMPROVEMENTS

### Current Performance (Estimated)
- **Array operations**: ~50ms per 1000-item loop
- **String building**: ~100ms per 100 concatenations
- **File I/O**: ~10ms per Test-Path + Get-Content
- **API calls**: ~200ms per Graph API call

### After High-Priority Optimizations
- **Array operations**: ~1ms per 1000-item loop (**50x faster**)
- **String building**: ~2ms per 100 concatenations (**50x faster**)
- **File I/O**: ~0.5ms per cached check (**20x faster**)
- **API calls**: Same, but 4x fewer calls (**4x overall improvement**)

### Overall Impact
For a typical bulk export with 10 tenants:
- **Current**: ~5-10 minutes total
- **Optimized**: ~1-2 minutes total
- **Improvement**: **5x faster**

---

## IMPLEMENTATION PRIORITY

1. **Week 1**: High-priority array/string optimizations (biggest impact, easiest to implement)
2. **Week 2**: Hashtable lookups and nested loop optimizations
3. **Week 3**: File I/O caching and API call optimization
4. **Week 4**: Process monitoring and remaining optimizations

---

## NOTES

- Most optimizations are straightforward refactorings
- Test thoroughly after each optimization
- Monitor performance improvements with timing measurements
- Some optimizations may require PowerShell 5.1+ (ArrayList, HashSet)
- Consider backward compatibility if supporting older PowerShell versions
