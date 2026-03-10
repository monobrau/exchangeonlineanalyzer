# Code Review: Unused Code and Duplicate Code Analysis

## Summary
This document identifies unused code, duplicate code patterns, and recommendations for cleanup.

---

## 1. UNUSED MODULES

### 1.1 `Modules/ExportUtils_new.psm1` ❌ **UNUSED**
- **Status**: Not imported by any script
- **Size**: ~540 lines
- **Reason**: Appears to be an alternate/simplified version of `ExportUtils.psm1`
- **Recommendation**: **DELETE** if not needed, or move to `Archive/` if kept for reference

### 1.2 `Modules/RestrictedSender.psm1` ❌ **UNUSED**
- **Status**: Only imported by `Archive/365analyzerv7.ps1` (which is confirmed unused)
- **Size**: 2,103 lines
- **Current Usage**: Not used by `ExchangeOnlineAnalyzer.ps1` or `BulkTenantExporter.ps1`
- **Recommendation**: **DELETE** - Archive file is not used, so this module is also unused

---

## 2. UNUSED EXPORTED FUNCTIONS

### 2.1 `Modules/MailboxAnalysis.psm1` - Missing Function Definitions
**Exported but NOT defined:**
- `Get-InternalDomains` ❌
- `Analyze-MailboxRules` ❌

**Status**: These functions are exported but never defined in the module. They may be:
- Defined elsewhere (not found in codebase)
- Legacy exports that should be removed
- Functions that were renamed/refactored

**Recommendation**: Remove from `Export-ModuleMember` if not used, or implement if needed.

### 2.2 `Modules/ExchangeOnline.psm1` - Missing Function Definitions
**Exported but NOT defined:**
- `Connect-ExchangeOnlineAnalyzer` ❌
- `Disconnect-ExchangeOnlineAnalyzer` ❌

**Status**: These functions are exported but not defined in the module. They are likely defined in `ExchangeOnlineAnalyzer.ps1` main script.

**Recommendation**: 
- Remove from `Export-ModuleMember` if they're only used within the main script
- Or move the function definitions into the module if they should be reusable

---

## 3. DUPLICATE CODE PATTERNS

### 3.1 `Safe-ImportModule` Function - **DUPLICATED 3 TIMES**
**Locations:**
1. `ExchangeOnlineAnalyzer.ps1` (line 51)
2. `BulkTenantExporter.ps1` (line 35)
3. `Archive/365analyzerv7.ps1` (line 50)

**Differences:**
- `ExchangeOnlineAnalyzer.ps1`: Uses `[System.Windows.Forms.MessageBox]` for errors
- `BulkTenantExporter.ps1`: Uses `Write-Error` and `exit 1`, includes success message
- `Archive/365analyzerv7.ps1`: Similar to ExchangeOnlineAnalyzer version

**Recommendation**: 
- **Create shared module** `Modules/ModuleLoader.psm1` with `Safe-ImportModule`
- Or add to `GraphOnline.psm1` or `Logging.psm1` as a utility function
- Update all three scripts to use the shared function

### 3.2 `Analyze-MailboxRulesEnhanced` Function - **DUPLICATED**
**Locations:**
1. `ExchangeOnlineAnalyzer.ps1` (line 458)
2. `Archive/365analyzerv7.ps1` (line 412)

**Status**: Same function exists in both active and archive versions.

**Recommendation**: 
- If archive is deprecated, this is fine
- If function should be shared, move to `Modules/MailboxAnalysis.psm1`

### 3.3 Repeated Module Import Patterns
**Pattern**: Multiple scripts import `Settings.psm1` repeatedly throughout code:
- `ExchangeOnlineAnalyzer.ps1`: Imports `Settings.psm1` **15+ times** in different event handlers
- `BulkTenantExporter.ps1`: Imports `Settings.psm1` **6+ times** in different handlers

**Recommendation**: 
- Import once at the top of the script (already done for some modules)
- Remove redundant imports from event handlers if module is already loaded globally

### 3.4 Duplicate Error Handling Patterns
**Pattern**: Similar try-catch blocks for module imports repeated throughout:
```powershell
try { Import-Module "$PSScriptRoot\Modules\Settings.psm1" -Force -ErrorAction SilentlyContinue } catch {}
```

**Occurrences**: Found 20+ instances across both main scripts.

**Recommendation**: 
- Create helper function: `Import-ModuleSafely -ModulePath $path`
- Or ensure modules are loaded globally at startup

---

## 4. UNUSED TEST/ARCHIVE FILES

### 4.1 Test Scripts (Already Cleaned Up)
- ✅ `test_ticket_simple.ps1` - Removed from git tracking
- ✅ `test_ticket_integration.ps1` - Removed from git tracking
- ✅ `test_memberberry_execution.ps1` - Removed from git tracking
- ✅ `test_memberberry_integration.ps1` - Removed from git tracking

### 4.2 Archive Files
- `Archive/365analyzerv7.ps1` - Old version (v8.0) ❌ **CONFIRMED UNUSED**
  - **Status**: Not used by any active scripts
  - **Size**: 5,307 lines
  - **Recommendation**: **DELETE** - No longer needed

### 4.3 Test Scripts in Scripts Folder
- `Scripts/Test-8.1b.ps1` - Test script for Graph/Browser modules
  - **Status**: May be useful for testing
  - **Recommendation**: Keep if actively used, otherwise move to test directory or remove

---

## 5. CODE DUPLICATION IN MODULES

### 5.1 `ExportUtils.psm1` - Self-Import Pattern
**Pattern**: Module imports itself in multiple functions:
```powershell
$ExportUtilsPath = Join-Path $PSScriptRoot 'ExportUtils.psm1'
Import-Module $ExportUtilsPath -Force -ErrorAction Stop
```

**Occurrences**: Lines 389, 685, 749 (in runspace scripts)

**Reason**: Used in parallel execution contexts (runspaces) where module needs to be reloaded.

**Recommendation**: 
- This is intentional for parallel execution - **KEEP**
- Consider documenting why this pattern is used

### 5.2 Graph Module Import Patterns
**Pattern**: Similar Microsoft.Graph module import patterns repeated:
```powershell
Import-Module Microsoft.Graph.Xxx -ErrorAction SilentlyContinue | Out-Null
```

**Occurrences**: Found in `ExportUtils.psm1`, `EntraInvestigator.psm1`, `SecurityAnalysis.psm1`

**Recommendation**: 
- Already centralized in `GraphOnline.psm1` with `Import-GraphModulesOnDemand`
- Some direct imports may be needed for specific cmdlets - **REVIEW** if all can use centralized function

---

## 6. RECOMMENDATIONS SUMMARY

### High Priority
1. ✅ **DELETE** `Modules/ExportUtils_new.psm1` (unused alternate version)
2. ✅ **DELETE** `Modules/RestrictedSender.psm1` (unused, only used by unused archive file)
3. ✅ **DELETE** `Archive/365analyzerv7.ps1` (confirmed unused)
4. ✅ **REMOVE** unused exports from `MailboxAnalysis.psm1` (`Get-InternalDomains`, `Analyze-MailboxRules`)
5. ✅ **REMOVE** unused exports from `ExchangeOnline.psm1` (`Connect-ExchangeOnlineAnalyzer`, `Disconnect-ExchangeOnlineAnalyzer`) OR move function definitions into module
6. ✅ **CREATE** shared `Safe-ImportModule` function to eliminate duplication

### Medium Priority
7. ⚠️ **REDUCE** redundant `Settings.psm1` imports in event handlers (import once globally)
8. ⚠️ **CONSOLIDATE** error handling patterns for module imports

### Low Priority
8. 📝 **DOCUMENT** why `ExportUtils.psm1` imports itself (for parallel execution)
9. 📝 **REVIEW** if all Graph module imports can use centralized `Import-GraphModulesOnDemand`

---

## 7. ESTIMATED IMPACT

### Code Reduction
- **ExportUtils_new.psm1**: ~540 lines
- **RestrictedSender.psm1**: 2,103 lines
- **Archive/365analyzerv7.ps1**: 5,307 lines
- **Duplicate Safe-ImportModule**: ~30 lines × 2 = 60 lines (if consolidated)
- **Unused exports cleanup**: ~5 lines
- **Total potential reduction**: ~8,015 lines

### Maintenance Benefits
- Single source of truth for module loading
- Reduced risk of inconsistencies between duplicate functions
- Cleaner module exports (only export what's actually used)
- Easier to maintain and test

---

## 8. NEXT STEPS

1. Review this analysis with the team
2. Confirm which archive/test files can be removed
3. Implement high-priority recommendations
4. Test after changes to ensure no functionality is broken
5. Update documentation if module structure changes
