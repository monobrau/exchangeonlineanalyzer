# Inbox Rule Keyword Enhancement Analysis

## Current Keywords
```powershell
$BaseSuspiciousKeywords = @("invoice", "payment", "password", "confidential", "urgent", "bank", "account", "auto forward", "external", "hidden")
```

## Analysis

### Current Coverage
- **Financial**: invoice, payment, bank, account
- **Security**: password, confidential
- **Urgency**: urgent
- **Forwarding**: auto forward, external
- **Visibility**: hidden

### Gaps Identified

1. **Missing common phishing/BEC keywords**
2. **Missing data exfiltration indicators**
3. **Missing social engineering terms**
4. **Missing impersonation indicators**
5. **Missing common attack pattern keywords**

---

## Recommended Enhanced Keyword List

### Category 1: Financial Fraud & BEC (Business Email Compromise)
**Current:** invoice, payment, bank, account
**Add:**
- "wire"
- "transfer"
- "refund"
- "payroll"
- "vendor"
- "supplier"
- "invoice"
- "billing"
- "receipt"
- "purchase order"
- "po number"
- "ach"
- "swift"
- "routing"
- "tax"
- "irs"
- "w2"
- "w-2"
- "1099"

### Category 2: Credential Theft & Account Takeover
**Current:** password
**Add:**
- "credential"
- "login"
- "sign in"
- "sign-in"
- "authentication"
- "verify"
- "verification"
- "security"
- "mfa"
- "2fa"
- "two factor"
- "reset"
- "expired"
- "suspended"
- "locked"
- "unlock"
- "activate"
- "reactivate"

### Category 3: Data Exfiltration & Forwarding
**Current:** auto forward, external
**Add:**
- "forward"
- "redirect"
- "copy"
- "bcc"
- "blind copy"
- "send copy"
- "archive"
- "backup"
- "export"
- "sync"
- "mirror"
- "duplicate"
- "replicate"

### Category 4: Social Engineering & Urgency
**Current:** urgent, confidential
**Add:**
- "asap"
- "immediately"
- "critical"
- "important"
- "action required"
- "action needed"
- "response required"
- "verify now"
- "confirm"
- "review"
- "approve"
- "authorize"
- "sign"
- "execute"
- "deadline"
- "overdue"
- "past due"
- "final notice"
- "last chance"

### Category 5: Impersonation & Authority Abuse
**Add:**
- "ceo"
- "cfo"
- "executive"
- "president"
- "director"
- "manager"
- "boss"
- "supervisor"
- "hr"
- "human resources"
- "it support"
- "help desk"
- "admin"
- "administrator"
- "service desk"

### Category 6: Malicious Actions & Evasion
**Current:** hidden
**Add:**
- "delete"
- "remove"
- "move"
- "archive"
- "mark as read"
- "mark read"
- "junk"
- "spam"
- "trash"
- "permanent"
- "cleanup"
- "organize"
- "sort"
- "filter"
- "silent"
- "quiet"
- "stealth"
- "invisible"

### Category 7: Common Attack Patterns
**Add:**
- "phishing"
- "malware"
- "ransomware"
- "breach"
- "compromise"
- "hack"
- "attack"
- "alert"
- "warning"
- "notification"
- "update"
- "upgrade"
- "patch"
- "fix"
- "vulnerability"

### Category 8: Compliance & Legal (Often Used in Phishing)
**Add:**
- "legal"
- "compliance"
- "audit"
- "lawsuit"
- "subpoena"
- "court"
- "attorney"
- "lawyer"
- "settlement"
- "dispute"
- "claim"

### Category 9: Technology & System Terms (Often Used in Phishing)
**Add:**
- "microsoft"
- "office365"
- "office 365"
- "azure"
- "sharepoint"
- "onedrive"
- "teams"
- "outlook"
- "exchange"
- "update"
- "upgrade"
- "maintenance"
- "migration"
- "sync"

### Category 10: Common Phishing Domains & Services
**Add:**
- "gmail"
- "yahoo"
- "hotmail"
- "outlook.com"
- "icloud"
- "protonmail"
- "tutanota"
- "mail.com"
- "yandex"
- "qq.com"

---

## Recommended Enhanced List (Prioritized)

### High Priority Additions (Most Common in Attacks)
```powershell
$BaseSuspiciousKeywords = @(
    # Financial/BEC (existing + additions)
    "invoice", "payment", "password", "confidential", "urgent", "bank", "account", 
    "auto forward", "external", "hidden",
    
    # High-priority additions
    "wire", "transfer", "refund", "payroll", "vendor", "supplier",
    "credential", "login", "verify", "verification", "reset", "expired",
    "forward", "redirect", "bcc", "archive", "export",
    "asap", "critical", "action required", "response required",
    "ceo", "cfo", "executive", "hr", "it support", "admin",
    "delete", "remove", "move", "mark as read", "junk", "spam",
    "phishing", "alert", "warning", "update", "upgrade",
    "legal", "compliance", "audit", "lawsuit",
    "microsoft", "office365", "azure", "sharepoint"
)
```

### Medium Priority Additions
```powershell
# Add to the list above:
"billing", "receipt", "purchase order", "ach", "swift",
"sign in", "sign-in", "mfa", "2fa", "suspended", "locked",
"copy", "sync", "mirror", "backup",
"immediately", "important", "confirm", "approve", "deadline",
"director", "manager", "boss", "administrator",
"trash", "permanent", "cleanup", "filter",
"malware", "breach", "compromise",
"attorney", "court", "settlement",
"onedrive", "teams", "outlook", "exchange", "migration"
```

### Lower Priority (But Still Valuable)
```powershell
# Add if needed:
"tax", "irs", "w2", "w-2", "1099",
"authentication", "two factor", "unlock", "activate",
"blind copy", "send copy", "duplicate", "replicate",
"asap", "review", "authorize", "sign", "execute", "overdue", "final notice",
"president", "supervisor", "help desk", "service desk",
"sort", "silent", "quiet", "stealth", "invisible",
"ransomware", "hack", "attack", "patch", "fix", "vulnerability",
"subpoena", "dispute", "claim",
"gmail", "yahoo", "hotmail", "icloud"
```

---

## Implementation Recommendation

### Option 1: Comprehensive List (Recommended)
Replace the current list with a comprehensive set covering all high and medium priority keywords:

```powershell
$BaseSuspiciousKeywords = @(
    # Financial & BEC
    "invoice", "payment", "bank", "account", "wire", "transfer", "refund", 
    "payroll", "vendor", "supplier", "billing", "receipt", "purchase order",
    "ach", "swift", "routing", "tax", "irs", "w2", "w-2", "1099",
    
    # Credentials & Authentication
    "password", "credential", "login", "sign in", "sign-in", "authentication",
    "verify", "verification", "reset", "expired", "suspended", "locked",
    "unlock", "activate", "reactivate", "mfa", "2fa", "two factor",
    
    # Data Exfiltration & Forwarding
    "auto forward", "external", "forward", "redirect", "copy", "bcc",
    "blind copy", "send copy", "archive", "backup", "export", "sync",
    "mirror", "duplicate", "replicate",
    
    # Social Engineering & Urgency
    "urgent", "confidential", "asap", "immediately", "critical", "important",
    "action required", "action needed", "response required", "verify now",
    "confirm", "review", "approve", "authorize", "sign", "execute",
    "deadline", "overdue", "past due", "final notice", "last chance",
    
    # Impersonation
    "ceo", "cfo", "executive", "president", "director", "manager", "boss",
    "supervisor", "hr", "human resources", "it support", "help desk",
    "service desk", "admin", "administrator",
    
    # Malicious Actions
    "hidden", "delete", "remove", "move", "mark as read", "mark read",
    "junk", "spam", "trash", "permanent", "cleanup", "organize", "sort",
    "filter", "silent", "quiet", "stealth", "invisible",
    
    # Attack Patterns
    "phishing", "malware", "ransomware", "breach", "compromise", "hack",
    "attack", "alert", "warning", "notification", "update", "upgrade",
    "patch", "fix", "vulnerability",
    
    # Compliance & Legal
    "legal", "compliance", "audit", "lawsuit", "subpoena", "court",
    "attorney", "lawyer", "settlement", "dispute", "claim",
    
    # Technology Terms
    "microsoft", "office365", "office 365", "azure", "sharepoint", "onedrive",
    "teams", "outlook", "exchange", "migration", "maintenance"
)
```

### Option 2: Tiered Approach
Create separate arrays for different severity levels:

```powershell
# High severity - most suspicious
$HighSeverityKeywords = @("password", "credential", "wire", "transfer", "forward", 
    "redirect", "external", "hidden", "delete", "ceo", "cfo", "phishing")

# Medium severity - suspicious but may be legitimate
$MediumSeverityKeywords = @("invoice", "payment", "urgent", "verify", "reset",
    "archive", "asap", "critical", "admin", "update")

# Combined for analysis
$BaseSuspiciousKeywords = $HighSeverityKeywords + $MediumSeverityKeywords
```

### Option 3: Categorized with Comments
Keep keywords organized by category for easier maintenance:

```powershell
$BaseSuspiciousKeywords = @(
    # Financial & BEC
    "invoice", "payment", "bank", "account", "wire", "transfer", "refund", 
    "payroll", "vendor", "supplier", "billing", "receipt", "purchase order",
    
    # Credentials & Authentication  
    "password", "credential", "login", "sign in", "verify", "reset", "expired",
    
    # Data Exfiltration
    "auto forward", "external", "forward", "redirect", "copy", "bcc", "archive",
    
    # Social Engineering
    "urgent", "confidential", "asap", "critical", "action required",
    
    # Impersonation
    "ceo", "cfo", "executive", "hr", "it support", "admin",
    
    # Malicious Actions
    "hidden", "delete", "remove", "move", "mark as read", "junk", "spam",
    
    # Attack Patterns
    "phishing", "malware", "alert", "update", "upgrade",
    
    # Compliance
    "legal", "compliance", "audit", "lawsuit"
)
```

---

## Additional Recommendations

### 1. Case-Insensitive Matching
✅ Already implemented correctly with `-match [regex]::Escape($kw)`

### 2. Partial Word Matching
✅ Already implemented - matches substrings within rule names

### 3. Consider Adding Pattern Matching
Consider detecting:
- Rules with only numbers (e.g., "12345")
- Rules with only special characters
- Rules with suspicious Unicode characters
- Rules with very long names (>100 chars)
- Rules with suspicious character combinations

### 4. Context-Aware Detection
Consider flagging rules that combine:
- Forwarding action + suspicious keyword
- Delete action + suspicious keyword  
- Hidden + forwarding
- Multiple suspicious keywords in one rule name

### 5. Regular Updates
Recommend reviewing and updating keyword list quarterly based on:
- Current threat intelligence
- Incident post-mortems
- Industry threat reports
- User feedback

---

## Testing Recommendations

1. Test with legitimate rule names to avoid false positives:
   - "Invoice Processing" (legitimate)
   - "Payment Reminders" (legitimate)
   - "Archive Old Emails" (legitimate)

2. Test with malicious rule names:
   - "Forward to external"
   - "Delete urgent"
   - "CEO wire transfer"
   - "Password reset verify"

3. Monitor false positive rate and adjust accordingly

---

## Summary

**Current Keywords:** 10
**Recommended High Priority:** ~40-50 keywords
**Recommended Comprehensive:** ~80-100 keywords

**Recommendation:** Start with Option 1 (Comprehensive List) covering high and medium priority keywords. This provides the best balance of detection coverage while maintaining reasonable performance.










