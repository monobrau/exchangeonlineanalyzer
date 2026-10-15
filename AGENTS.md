# Workspace notes

- **Client auth Graph app dropdown:** Commit **`4cc23b0`** is the reference version where the client-auth Graph app dropdown correctly shows **client name and tenant number**.
- **Web API (bulk jobs):** Removed from this repo. An archived snapshot of the former `web/` tree (FastAPI console, deploy docs, `WebExoLinuxRunner.ps1`, etc.) lives outside the repo at **`C:\git\archives\exchangeonlineanalyzer-web-archive.tgz`** with **`C:\git\archives\README-exchangeonlineanalyzer-web.txt`**. Restore with `tar -xzvf …` if you need that code again.

## Cloud Agent (Linux)

Cloud Agent VMs are Linux. `pwsh` 7 is on `PATH`. The environment install places these modules in the CurrentUser scope: PSScriptAnalyzer, ExchangeOnlineManagement, Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Users.Actions, Microsoft.Graph.Identity.SignIns, Microsoft.Graph.Reports, and Microsoft.Graph.Mail.

`ExchangeOnlineAnalyzer.ps1` loads `System.Windows.Forms` at startup, so the GUI does not open here. Excel COM and Windows Credential Manager are Windows-only. Exercise the analysis modules instead:

```powershell
Import-Module .\Modules\MailboxAnalysis.psm1, .\Modules\ReportAnalysis.psm1, .\Modules\Settings.psm1, .\Modules\ExchangeOnline.psm1
Get-AutoDetectedDomains -MailboxUPNs @('user@contoso.com')
Get-InboxRuleFindings -Rules @([pscustomobject]@{ MailboxOwner='user@contoso.com'; Name='.'; ForwardTo='a@micr0soft.com'; IsHidden='True'; SubjectContains='.' })
```

Parse a script with `[System.Management.Automation.Language.Parser]::ParseFile` before treating an edit as valid. `Invoke-ScriptAnalyzer -Severity ParseError,Error` is available.
