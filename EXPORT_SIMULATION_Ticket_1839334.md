# Export Simulation: Ticket #1234567 (All Options Checked)

**Simulated for:** Service Ticket #1234567 - Barracuda XDR SentinelOne New Threat Mitigated  
**Company:** Contoso Accounting LLP  
**Assumption:** All report checkboxes selected in the GUI, ticket pasted, Memberberry enabled.

---

## 1. Extracted Data from Ticket

| Field | Value |
|-------|-------|
| **Ticket #** | 1234567 |
| **Company** | Contoso Accounting LLP |
| **Contact** | Jane Smith |
| **Contact Email** | jane.smith@contoso.com |
| **Site** | Contoso Accounting LLP |
| **Domain (from ticket)** | contoso.com |
| **User in alert** | jdoe (from CEF: sourceUserName=jdoe) |
| **Hostname** | WORKSTATION-01 |

**Emails extracted from ticket (Extract Emails):** `jane.smith@contoso.com` (SOC@BARRACUDA.COM is external, not tenant)

---

## 2. Memberberry-Slim Package Output

**create-slim-package.ps1** runs with:
- `TicketContent` = cleaned ticket text
- `TicketNumbers` = @('1234567')
- `CompanyName` = "Contoso Accounting LLP" (from extract-company.ps1 → Company: field)

**Files produced:**

| File | Description |
|------|-------------|
| `Ticket-1234567.txt` | Cleaned ticket with header "TICKET INFORMATION - Ticket #1234567" |
| `ClientExceptions-Ticket-1234567-Contoso-Accounting-LLP.txt` | Client-specific exceptions (if matched in exceptions.json) or `ClientExceptions-Contoso-Accounting-LLP-Default.txt` if no match |
| `ClientExceptions-Ticket-1234567-Contoso-Accounting-LLP.zip` | Zip of the exceptions file |
| `GlobalExceptions.txt` | Global exceptions from exceptions.json |
| `Settings.txt` | technician_name, timezone, export_time, signature_format |
| `always_include.md` | If present in memberberry folder |

**Client match logic:**  
- Exact match: "Contoso Accounting LLP" in exceptions.json  
- Partial match: Only if multi-word (avoids "Todd" → "The Todd Organization")  
- If no match: Uses "Client matched from ticket but no exceptions configured yet. Use GLOBAL EXCEPTIONS for default rules."

---

## 3. Report CSVs (All Options Checked)

**Ticket suffix:** `_Ticket_1234567`

| Report | Filename | Scope |
|--------|----------|-------|
| Message Trace | `MessageTrace_Ticket_1234567.csv` | Last 10 days; filtered by SelectedUsers if user filter enabled |
| Inbox Rules | `InboxRules_Ticket_1234567.csv` | Per SelectedUsers |
| Transport Rules | `TransportRules_Ticket_1234567.csv` | Tenant-wide |
| Mail Flow Connectors | `MailFlowConnectors_Ticket_1234567.csv` | Tenant-wide |
| Graph Audit Logs | `GraphAuditLogs_Ticket_1234567.csv` | Tenant-wide |
| Unified Audit Logs | `UnifiedAuditLogs_Ticket_1234567.csv` | Per SelectedUsers, 1 query per user |
| Conditional Access Policies | `ConditionalAccessPolicies_Ticket_1234567.csv` | Tenant-wide |
| App Registrations | `AppRegistrations_Ticket_1234567.csv` | Tenant-wide |
| Sign-in Logs | `SignInLogs_Ticket_1234567.csv` | Per SelectedUsers (requires Azure AD Premium; 7 days default) |
| Intune Devices | `IntuneDevices_Ticket_1234567.csv` | Tenant-wide (requires DeviceManagementManagedDevices.Read.All) |
| SharePoint Activity | `SharePointActivity_Ticket_1234567.csv` | Per SelectedUsers (requires E5/Reports.Read.All) |
| OneDrive Activity | `OneDriveActivity_Ticket_1234567.csv` | Per SelectedUsers (requires E5/Reports.Read.All) |
| Teams Activity | `TeamsActivity_Ticket_1234567.csv` | Per SelectedUsers (requires E5/Reports.Read.All) |
| SharePoint Sharing | `SharePointSharing_Ticket_1234567.csv` | Per SelectedUsers |
| Security Alerts | `SecurityAlerts_Ticket_1234567.csv` | Tenant-wide (requires E5/SecurityAlert.Read.All) |
| Security Incidents | `SecurityIncidents_Ticket_1234567.csv` | Tenant-wide (requires E5/SecurityIncident.Read.All) |
| Mailbox Forwarding | (in UserSecurityPosture) | Per SelectedUsers |
| MFA Coverage | (in UserSecurityPosture) | Per SelectedUsers |

**UserSecurityPosture_Ticket_1234567.csv** – Combined view: MFA status, mailbox forwarding, delegation, security groups.

**Findings_Ticket_1234567.csv** – Rule-based analysis findings (if automated analysis enabled).  
**_Automated_Summary_Ticket_1234567.txt** – Summary of findings.

---

## 4. User Scope

**If "Filter to users in ticket" is OFF:**  
- All tenant users for user-scoped reports (Message Trace, Inbox Rules, UAL, Sign-in Logs, SharePoint/OneDrive/Teams Activity, Mailbox Forwarding, MFA Coverage).

**If "Filter to users in ticket" is ON and Extract Emails was run:**  
- `jane.smith@contoso.com` (only tenant email in ticket).  
- `jdoe` would need to be added manually (e.g. via User Search) – it appears in the CEF log but not as an email.

---

## 5. Output Folder Structure

```
{Documents}\ExchangeOnlineAnalyzer\SecurityInvestigation\{TenantName}\{yyyyMMdd_HHmmss}\
├── Ticket-1234567.txt
├── ClientExceptions-Ticket-1234567-Contoso-Accounting-LLP.txt
├── ClientExceptions-Ticket-1234567-Contoso-Accounting-LLP.zip
├── GlobalExceptions.txt
├── Settings.txt
├── always_include.md (if exists)
├── MessageTrace_Ticket_1234567.csv
├── InboxRules_Ticket_1234567.csv
├── TransportRules_Ticket_1234567.csv
├── MailFlowConnectors_Ticket_1234567.csv
├── GraphAuditLogs_Ticket_1234567.csv
├── UnifiedAuditLogs_Ticket_1234567.csv
├── ConditionalAccessPolicies_Ticket_1234567.csv
├── AppRegistrations_Ticket_1234567.csv
├── SignInLogs_Ticket_1234567.csv
├── IntuneDevices_Ticket_1234567.csv
├── SharePointActivity_Ticket_1234567.csv
├── OneDriveActivity_Ticket_1234567.csv
├── TeamsActivity_Ticket_1234567.csv
├── SharePointSharing_Ticket_1234567.csv
├── SecurityAlerts_Ticket_1234567.csv
├── SecurityIncidents_Ticket_1234567.csv
├── UserSecurityPosture_Ticket_1234567.csv
├── Findings_Ticket_1234567.csv (if automated analysis)
├── _Automated_Summary_Ticket_1234567.txt (if automated analysis)
└── SecurityInvestigation_Ticket_1234567_{timestamp}.zip
```

---

## 6. Zip Contents

**SecurityInvestigation_Ticket_1234567_{timestamp}.zip** includes:
- All CSV/JSON report files
- Memberberry-Slim files: Ticket-1234567.txt, ClientExceptions-*.txt, GlobalExceptions.txt, Settings.txt, always_include.md
- ClientExceptions zip (standalone deliverable)

---

## 7. Notes

- **Company extraction:** Uses "Company:" field → "Contoso Accounting LLP". No Todd false match (multi-word company).
- **Ticket number:** Extracted via `Extract-TicketNumbers` (Pattern 1: "Service Ticket #1234567").
- **Permissions:** Some reports need E5, Azure AD Premium, or specific Graph/Exchange roles; missing permissions produce error/info files instead of data.
- **Days back:** Message Trace and UAL use 10 days; Sign-in Logs use 7 days by default.
