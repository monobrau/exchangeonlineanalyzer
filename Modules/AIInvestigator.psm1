# AIInvestigator.psm1
# AI-driven threat investigation engine with streaming and function calling

# Import required modules
$script:requiredModules = @(
    'SecurityAnalysis',
    'EntraInvestigator',
    'MailboxAnalysis',
    'MemberberryIntegration'
)

foreach ($mod in $script:requiredModules) {
    $modPath = Join-Path $PSScriptRoot "$mod.psm1"
    if (Test-Path $modPath) {
        Import-Module $modPath -Force -ErrorAction SilentlyContinue
    }
}

# Module state
$script:InvestigationCache = @{}
$script:InvestigationStats = @{
    ToolCalls = 0
    CacheHits = 0
    TotalDuration = 0
}
$script:DefaultClaudeModel = 'claude-sonnet-5-5'

function ConvertTo-StableJson {
    param($Value, [int]$Depth = 20)
    if ($Depth -lt 0) { return '"…"' }
    if ($null -eq $Value) { return 'null' }
    if ($Value -is [bool]) { if ($Value) { return 'true' } else { return 'false' } }
    if ($Value -is [datetime]) { return ('"' + $Value.ToUniversalTime().ToString('o') + '"') }
    if ($Value -is [string]) {
        $escaped = $Value.Replace('\', '\\').Replace('"', '\"').Replace("`r", '\r').Replace("`n", '\n').Replace("`t", '\t')
        return '"' + $escaped + '"'
    }
    if ($Value -is [byte] -or $Value -is [int16] -or $Value -is [int] -or $Value -is [int64] -or $Value -is [decimal] -or $Value -is [double] -or $Value -is [single]) {
        return ([string]([System.Convert]::ToString($Value, [cultureinfo]::InvariantCulture)))
    }
    if ($Value -is [System.Collections.IDictionary]) {
        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($key in @($Value.Keys)) {
            $parts.Add(((ConvertTo-StableJson ([string]$key)) + ':' + (ConvertTo-StableJson $Value[$key] ($Depth - 1))))
        }
        return '{' + ($parts -join ',') + '}'
    }
    if ($Value -is [System.Collections.IEnumerable]) {
        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($item in $Value) { $parts.Add((ConvertTo-StableJson $item ($Depth - 1))) }
        return '[' + ($parts -join ',') + ']'
    }
    $hash = @{}
    $wrote = $false
    foreach ($prop in $Value.PSObject.Properties) {
        if ($prop.MemberType -eq 'NoteProperty' -or $prop.MemberType -eq 'Property') {
            $hash[$prop.Name] = $prop.Value
            $wrote = $true
        }
    }
    if ($wrote) { return ConvertTo-StableJson $hash ($Depth - 1) }
    return (ConvertTo-StableJson ([string]$Value))
}

function ConvertTo-HashtableDeep {
    param($InputObject)
    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [string] -or $InputObject -is [ValueType]) { return $InputObject }
    if ($InputObject -is [System.Collections.IDictionary]) {
        $hash = @{}
        foreach ($key in @($InputObject.Keys)) { $hash[$key] = ConvertTo-HashtableDeep $InputObject[$key] }
        return $hash
    }
    if ($InputObject -is [System.Collections.IEnumerable]) {
        $items = New-Object System.Collections.Generic.List[object]
        foreach ($item in $InputObject) { $items.Add((ConvertTo-HashtableDeep $item)) }
        return ,$items.ToArray()
    }
    $hash = @{}
    $wrote = $false
    foreach ($prop in $InputObject.PSObject.Properties) {
        if ($prop.MemberType -eq 'NoteProperty' -or $prop.MemberType -eq 'Property') {
            $hash[$prop.Name] = ConvertTo-HashtableDeep $prop.Value
            $wrote = $true
        }
    }
    if ($wrote) { return $hash }
    return $InputObject
}

function New-ClaudeStreamAssembler {
    return @{
        Blocks = @{}
        Order = New-Object System.Collections.Generic.List[int]
        StopReason = $null
        Usage = @{}
        Error = $null
        LastTextDelta = ''
    }
}

function Add-ClaudeStreamEvent {
    param($State, $Event)
    if ($null -eq $Event) { return }
    switch ([string]$Event.type) {
        'content_block_start' {
            $index = [int]$Event.index
            $block = $Event.content_block
            $State.Blocks[$index] = @{
                type = [string]$block.type
                text = [string]$block.text
                id = [string]$block.id
                name = [string]$block.name
                partial_json = ''
                input = $null
            }
            if (-not $State.Order.Contains($index)) { [void]$State.Order.Add($index) }
        }
        'content_block_delta' {
            $index = [int]$Event.index
            $entry = $State.Blocks[$index]
            if (-not $entry) { return }
            $deltaType = [string]$Event.delta.type
            if ($deltaType -eq 'text_delta') {
                $piece = [string]$Event.delta.text
                $entry.text += $piece
                $State.LastTextDelta = $piece
            } elseif ($deltaType -eq 'input_json_delta') {
                $entry.partial_json += [string]$Event.delta.partial_json
            }
        }
        'content_block_stop' {
            $index = [int]$Event.index
            $entry = $State.Blocks[$index]
            if ($entry -and $entry.type -eq 'tool_use') {
                if ([string]::IsNullOrWhiteSpace($entry.partial_json)) {
                    $entry.input = @{}
                } else {
                    try {
                        $entry.input = ConvertTo-HashtableDeep ($entry.partial_json | ConvertFrom-Json)
                        if ($entry.input -isnot [hashtable]) { $entry.input = @{} }
                    } catch {
                        $entry.input = @{ Error = $_.Exception.Message }
                    }
                }
            }
        }
        'message_delta' {
            if ($Event.delta.stop_reason) { $State.StopReason = [string]$Event.delta.stop_reason }
            if ($Event.usage) { $State.Usage = $Event.usage }
        }
        'error' {
            $State.Error = [string]$Event.error.message
        }
    }
}

function Get-ClaudeStreamContent {
    param($State)
    $content = New-Object System.Collections.Generic.List[object]
    foreach ($index in @($State.Order)) {
        $entry = $State.Blocks[$index]
        if ($entry.type -eq 'text') {
            $content.Add(@{ type = 'text'; text = [string]$entry.text })
        } elseif ($entry.type -eq 'tool_use') {
            $content.Add(@{ type = 'tool_use'; id = [string]$entry.id; name = [string]$entry.name; input = $entry.input })
        }
    }
    # A bare array is unwrapped by PowerShell. Keep the blocks on a property.
    return [pscustomobject]@{ Items = $content.ToArray() }
}

function Get-InvestigationReplyText {
    param([string]$Text)
    $result = @{
        Verdict = 'Unknown'
        Confidence = 'Unknown'
        Summary = ''
        ClientReply = $Text
    }
    if ([string]::IsNullOrWhiteSpace($Text)) { return $result }
    if ($Text -match '(?m)^VERDICT:\s*(.+)$') { $result.Verdict = $Matches[1].Trim() }
    if ($Text -match '(?m)^CONFIDENCE:\s*(.+)$') { $result.Confidence = $Matches[1].Trim() }
    if ($Text -match '(?m)^SUMMARY:\s*(.+)$') { $result.Summary = $Matches[1].Trim() }
    if ($Text -match '(?s)CLIENT REPLY:\s*(.+)$') { $result.ClientReply = $Matches[1].Trim() }
    return $result
}

function Get-ToolInt {
    param($Parameters, [string]$Name, [int]$Default, [int]$Minimum = 1, [int]$Maximum = 30)
    if ($null -eq $Parameters -or -not $Parameters.ContainsKey($Name) -or $null -eq $Parameters[$Name] -or "$($Parameters[$Name])" -eq '') {
        return $Default
    }
    $number = 0
    if (-not [int]::TryParse([string]$Parameters[$Name], [ref]$number)) { return $Default }
    if ($number -lt $Minimum) { return $Minimum }
    if ($number -gt $Maximum) { return $Maximum }
    return $number
}

<#
.SYNOPSIS
    Defines available tools for AI investigation
.DESCRIPTION
    Maps AI tool names to PowerShell functions with parameter schemas
#>
function Get-AIToolDefinitions {
    return @(
        @{
            name = "get_signin_logs"
            description = "Retrieves Entra ID sign-in logs for specified users over a time period. Returns authentication attempts, locations, devices, IP addresses, conditional access results, and MFA status."
            parameters = @{
                type = "object"
                properties = @{
                    user_principals = @{
                        type = "array"
                        items = @{ type = "string" }
                        description = "User principal names (UPNs) to query, e.g. ['user@domain.com']"
                    }
                    days_back = @{
                        type = "integer"
                        description = "Number of days to look back (max 30)"
                        default = 7
                    }
                }
                required = @("user_principals")
            }
        },
        @{
            name = "get_inbox_rules"
            description = "Gets all inbox rules for specified mailboxes. Detects forwarding, redirecting, deletion, and suspicious patterns."
            parameters = @{
                type = "object"
                properties = @{
                    mailbox_upns = @{
                        type = "array"
                        items = @{ type = "string" }
                        description = "Mailbox UPNs to check"
                    }
                }
                required = @("mailbox_upns")
            }
        },
        @{
            name = "get_user_details"
            description = "Gets comprehensive Entra ID user details including roles, groups, MFA status, and account status."
            parameters = @{
                type = "object"
                properties = @{
                    user_principal = @{
                        type = "string"
                        description = "User principal name"
                    }
                }
                required = @("user_principal")
            }
        },
        @{
            name = "get_audit_logs"
            description = "Retrieves Entra ID audit logs for user actions (mailbox changes, admin operations, permission grants)."
            parameters = @{
                type = "object"
                properties = @{
                    user_principal = @{
                        type = "string"
                        description = "User principal name to query"
                    }
                    days_back = @{
                        type = "integer"
                        description = "Days to look back (max 30)"
                        default = 7
                    }
                }
                required = @("user_principal")
            }
        },
        @{
            name = "get_mailbox_permissions"
            description = "Checks for suspicious mailbox permission changes (FullAccess, SendAs, SendOnBehalf grants)."
            parameters = @{
                type = "object"
                properties = @{
                    days_back = @{
                        type = "integer"
                        description = "Days to look back"
                        default = 7
                    }
                    target_mailboxes = @{
                        type = "array"
                        items = @{ type = "string" }
                        description = "Specific mailboxes to check (optional, checks all if empty)"
                    }
                }
            }
        },
        @{
            name = "get_oauth_permissions"
            description = "Identifies OAuth app consent grants, especially suspicious or user-level consents to external apps."
            parameters = @{
                type = "object"
                properties = @{
                    days_back = @{
                        type = "integer"
                        description = "Days to look back"
                        default = 30
                    }
                }
            }
        },
        @{
            name = "get_device_compliance"
            description = "Gets Intune device compliance status for devices used by specified users."
            parameters = @{
                type = "object"
                properties = @{
                    user_principals = @{
                        type = "array"
                        items = @{ type = "string" }
                        description = "User UPNs to check devices for"
                    }
                }
                required = @("user_principals")
            }
        },
        @{
            name = "get_mfa_status"
            description = "Gets MFA methods and whether Security Defaults or Conditional Access require MFA for one user."
            parameters = @{
                type = "object"
                properties = @{
                    user_principal = @{
                        type = "string"
                        description = "User principal name"
                    }
                }
                required = @("user_principal")
            }
        },
        @{
            name = "search_message_trace"
            description = "Searches Exchange message trace logs for sent/received emails matching criteria."
            parameters = @{
                type = "object"
                properties = @{
                    sender = @{
                        type = "string"
                        description = "Sender email address (optional)"
                    }
                    recipient = @{
                        type = "string"
                        description = "Recipient email address (optional)"
                    }
                    days_back = @{
                        type = "integer"
                        description = "Days to search (max 10). Provide a sender or a recipient."
                        default = 7
                    }
                }
            }
        }
    )
}

<#
.SYNOPSIS
    Executes an AI tool call by mapping to PowerShell function
.DESCRIPTION
    Handles caching, parameter mapping, error handling
#>
function Invoke-AITool {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$ToolName,
        
        [Parameter(Mandatory=$false)]
        [object]$Parameters = @{},
        
        [Parameter(Mandatory=$false)]
        [switch]$UseCache = $true
    )
    
    $script:InvestigationStats.ToolCalls++
    $startTime = Get-Date
    
    try {
        if ($Parameters -isnot [hashtable]) {
            $Parameters = ConvertTo-HashtableDeep $Parameters
        }
        if ($Parameters -isnot [hashtable]) { $Parameters = @{} }

        $cacheKey = "$ToolName|$(ConvertTo-StableJson $Parameters)"
        
        if ($UseCache -and $script:InvestigationCache.ContainsKey($cacheKey)) {
            $script:InvestigationStats.CacheHits++
            Write-Host "  [Cache Hit] $ToolName" -ForegroundColor DarkGray
            return $script:InvestigationCache[$cacheKey]
        }
        
        Write-Host "  [Executing] $ToolName..." -ForegroundColor Cyan

        $graphTools = @('get_signin_logs','get_user_details','get_audit_logs','get_oauth_permissions','get_device_compliance','get_mfa_status')
        $exchangeTools = @('get_inbox_rules','get_mailbox_permissions','search_message_trace')
        if ($graphTools -contains $ToolName) {
            $graphContext = $null
            if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) {
                $graphContext = Get-MgContext -ErrorAction SilentlyContinue
            }
            if (-not $graphContext) {
                return @{ Error = 'Not connected to Microsoft Graph. Run Connect-MgGraph before investigating.' }
            }
        }
        if ($exchangeTools -contains $ToolName -and -not (Get-Command Get-InboxRule -ErrorAction SilentlyContinue) -and $ToolName -eq 'get_inbox_rules') {
            return @{ Error = 'Exchange Online is not loaded. Run Connect-ExchangeOnline before investigating.' }
        }
        
        $result = switch ($ToolName) {
            "get_signin_logs" {
                $upns = @($Parameters['user_principals'])
                $days = Get-ToolInt -Parameters $Parameters -Name 'days_back' -Default 7 -Maximum 30
                Get-EntraSignInLogs -UserPrincipalNames $upns -Days $days
            }
            
            "get_inbox_rules" {
                $upns = @($Parameters['mailbox_upns'])
                $allRules = @()
                foreach ($upn in $upns) {
                    if ([string]::IsNullOrWhiteSpace($upn)) { continue }
                    try {
                        $rules = @(Get-InboxRule -Mailbox $upn -ErrorAction Stop)
                        foreach ($rule in $rules) {
                            $allRules += [pscustomobject]@{
                                Mailbox = $upn
                                Name = $rule.Name
                                Enabled = $rule.Enabled
                                Priority = $rule.Priority
                                From = ($rule.From -join '; ')
                                SubjectContainsWords = ($rule.SubjectContainsWords -join '; ')
                                ForwardTo = ($rule.ForwardTo -join '; ')
                                RedirectTo = ($rule.RedirectTo -join '; ')
                                ForwardAsAttachmentTo = ($rule.ForwardAsAttachmentTo -join '; ')
                                DeleteMessage = $rule.DeleteMessage
                                MoveToFolder = $rule.MoveToFolder
                                StopProcessingRules = $rule.StopProcessingRules
                                Description = $rule.Description
                            }
                        }
                    } catch {
                        $allRules += @{ Mailbox = $upn; Error = $_.Exception.Message }
                    }
                }
                $allRules
            }
            
            "get_user_details" {
                Get-EntraUserDetailsAndRoles -UserPrincipalName $Parameters['user_principal']
            }
            
            "get_audit_logs" {
                $days = Get-ToolInt -Parameters $Parameters -Name 'days_back' -Default 7 -Maximum 30
                Get-EntraUserAuditLogs -UserPrincipalName $Parameters['user_principal'] -Days $days
            }
            
            "get_mailbox_permissions" {
                $days = Get-ToolInt -Parameters $Parameters -Name 'days_back' -Default 7 -Maximum 30
                $targets = @($Parameters['target_mailboxes'] | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                if (-not (Get-Command Get-MailboxPermissionChanges -ErrorAction SilentlyContinue)) {
                    @{ Error = "Mailbox permission tracking not available" }
                } elseif ($targets.Count -gt 0) {
                    Get-MailboxPermissionChanges -DaysBack $days -SelectedUsers $targets
                } else {
                    Get-MailboxPermissionChanges -DaysBack $days
                }
            }
            
            "get_oauth_permissions" {
                $days = Get-ToolInt -Parameters $Parameters -Name 'days_back' -Default 30 -Maximum 90
                if (Get-Command Get-OAuthPermissionGrants -ErrorAction SilentlyContinue) {
                    Get-OAuthPermissionGrants -DaysBack $days
                } else {
                    @{ Error = "OAuth permission tracking not available" }
                }
            }
            
            "get_device_compliance" {
                $upns = @($Parameters['user_principals'])
                if (Get-Command Get-IntuneDeviceRecords -ErrorAction SilentlyContinue) {
                    Get-IntuneDeviceRecords -UserPrincipalNames $upns
                } else {
                    @{ Error = "Device compliance tracking not available (requires Intune)" }
                }
            }

            "get_mfa_status" {
                Get-EntraUserMfaStatus -UserPrincipalName $Parameters['user_principal']
            }
            
            "search_message_trace" {
                $sender = [string]$Parameters['sender']
                $recipient = [string]$Parameters['recipient']
                $days = Get-ToolInt -Parameters $Parameters -Name 'days_back' -Default 7 -Maximum 10
                if ([string]::IsNullOrWhiteSpace($sender) -and [string]::IsNullOrWhiteSpace($recipient)) {
                    @{ Error = 'Message trace needs a sender or a recipient.' }
                } else {
                    $traceParams = @{
                        StartDate = (Get-Date).AddDays(-$days)
                        EndDate = Get-Date
                        PageSize = 100
                    }
                    if ($sender) { $traceParams['SenderAddress'] = $sender }
                    if ($recipient) { $traceParams['RecipientAddress'] = $recipient }
                    $pages = @()
                    for ($page = 1; $page -le 5; $page++) {
                        $traceParams['Page'] = $page
                        $batch = @(Get-MessageTrace @traceParams -ErrorAction Stop)
                        if ($batch.Count -eq 0) { break }
                        $pages += $batch | Select-Object Received, SenderAddress, RecipientAddress, Subject, Status, FromIP, ToIP, Size
                        if ($batch.Count -lt 100) { break }
                    }
                    $pages
                }
            }
            
            default {
                @{ Error = "Unknown tool: $ToolName" }
            }
        }
        
        $isError = ($result -is [System.Collections.IDictionary] -and $result.Contains('Error'))
        if ($UseCache -and -not $isError) {
            $script:InvestigationCache[$cacheKey] = $result
        }
        
        $duration = ((Get-Date) - $startTime).TotalSeconds
        Write-Host "  [Complete] $ToolName ($([math]::Round($duration, 2))s)" -ForegroundColor Green
        
        return $result
        
    } catch {
        Write-Error "Tool execution failed: $ToolName - $_"
        return @{ Error = $_.Exception.Message }
    }
}

<#
.SYNOPSIS
    Converts PowerShell object to JSON summary for AI consumption
.DESCRIPTION
    Reduces large objects to concise summaries to save tokens
#>
function ConvertTo-AIFriendlyJson {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [object]$Data,
        
        [Parameter(Mandatory=$false)]
        [int]$MaxItems = 50,
        
        [Parameter(Mandatory=$false)]
        [int]$MaxDepth = 3
    )
    
    try {
        # Handle arrays
        if ($Data -is [Array]) {
            if ($Data.Count -eq 0) {
                return "[]"
            }
            
            # Limit array size
            $limitedData = if ($Data.Count -gt $MaxItems) {
                $Data | Select-Object -First $MaxItems
            } else {
                $Data
            }
            
            # Convert to JSON
            $json = $limitedData | ConvertTo-Json -Depth $MaxDepth -Compress
            
            if ($Data.Count -gt $MaxItems) {
                $json += " [... $($Data.Count - $MaxItems) more items truncated]"
            }
            
            return $json
        }
        
        # Handle single objects
        return $Data | ConvertTo-Json -Depth $MaxDepth -Compress
        
    } catch {
        return "{ `"error`": `"Failed to serialize data`" }"
    }
}

<#
.SYNOPSIS
    Calls Claude API with streaming support
.DESCRIPTION
    Handles function calling, streaming responses, and error recovery
#>
function Invoke-ClaudeInvestigation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$ApiKey,
        
        [Parameter(Mandatory=$true)]
        [string]$SystemPrompt,
        
        [Parameter(Mandatory=$true)]
        [array]$Messages,
        
        [Parameter(Mandatory=$false)]
        [array]$Tools = @(),
        
        [Parameter(Mandatory=$false)]
        [scriptblock]$OnStreamChunk = $null,
        
        [Parameter(Mandatory=$false)]
        [int]$MaxTokens = 4096,
        
        [Parameter(Mandatory=$false)]
        [string]$Model = "claude-sonnet-5-5"
    )
    
    try {
        $apiUrl = "https://api.anthropic.com/v1/messages"
        
        # Build request body
        $requestBody = @{
            model = $Model
            max_tokens = $MaxTokens
            system = $SystemPrompt
            messages = $Messages
        }
        
        if ($Tools.Count -gt 0) {
            $requestBody['tools'] = $Tools
        }
        
        # Enable streaming if callback provided
        if ($OnStreamChunk) {
            $requestBody['stream'] = $true
        }
        
        $headers = @{
            'anthropic-version' = '2023-06-01'
            'x-api-key' = $ApiKey
            'content-type' = 'application/json'
        }
        
        $jsonBody = ConvertTo-StableJson $requestBody
        
        if ($OnStreamChunk) {
            # Streaming mode
            return Invoke-ClaudeStreaming -Url $apiUrl -Headers $headers -Body $jsonBody -OnChunk $OnStreamChunk
        } else {
            # Non-streaming mode
            $response = Invoke-RestMethod -Uri $apiUrl -Method Post -Headers $headers -Body $jsonBody -ErrorAction Stop
            return $response
        }
        
    } catch {
        Write-Error "Claude API error: $_"
        return $null
    }
}

<#
.SYNOPSIS
    Handles streaming Claude API response
#>
function Invoke-ClaudeStreaming {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Url,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$Headers,
        
        [Parameter(Mandatory=$true)]
        [string]$Body,
        
        [Parameter(Mandatory=$true)]
        [scriptblock]$OnChunk
    )
    
    try {
        # Use WebRequest for streaming (Invoke-RestMethod doesn't support SSE well)
        $request = [System.Net.WebRequest]::Create($Url)
        $request.Method = "POST"
        $request.ContentType = "application/json"
        
        foreach ($key in $Headers.Keys) {
            if ($key -eq 'content-type') { continue }
            $request.Headers.Add($key, $Headers[$key])
        }
        
        # Write body
        $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($Body)
        $requestStream = $request.GetRequestStream()
        $requestStream.Write($bodyBytes, 0, $bodyBytes.Length)
        $requestStream.Close()
        
        # Get response stream
        $response = $request.GetResponse()
        $stream = $response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($stream)
        
        $assembler = New-ClaudeStreamAssembler
        while (-not $reader.EndOfStream) {
            $line = $reader.ReadLine()
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            if ($line.StartsWith('event:')) { continue }
            if (-not $line.StartsWith('data:')) { continue }

            $data = $line.Substring(5).Trim()
            if ($data -eq '[DONE]') { break }
            $event = $data | ConvertFrom-Json
            Add-ClaudeStreamEvent -State $assembler -Event $event
            if ($assembler.LastTextDelta -and $OnChunk -and $event.type -eq 'content_block_delta' -and $event.delta.type -eq 'text_delta') {
                & $OnChunk $assembler.LastTextDelta
                $assembler.LastTextDelta = ''
            }
            if ($assembler.Error) { throw $assembler.Error }
        }

        $reader.Close()
        $stream.Close()
        $response.Close()

        return @{
            content = @(Get-ClaudeStreamContent -State $assembler).Items
            stop_reason = $assembler.StopReason
            usage = $assembler.Usage
        }
        
    } catch {
        Write-Error "Streaming error: $_"
        return $null
    }
}

<#
.SYNOPSIS
    Starts an AI-driven threat investigation
.DESCRIPTION
    Main entry point for investigation workflow with streaming and function calling
#>
function Start-AIInvestigation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$TicketContent,
        
        [Parameter(Mandatory=$false)]
        [string]$CompanyName = "",
        
        [Parameter(Mandatory=$false)]
        [string]$MemberberryInstructions = "",
        
        [Parameter(Mandatory=$false)]
        [string]$ApiKey = "",
        
        [Parameter(Mandatory=$false)]
        [string]$Model = "",
        
        [Parameter(Mandatory=$false)]
        [scriptblock]$OnStatusUpdate = $null,
        
        [Parameter(Mandatory=$false)]
        [scriptblock]$OnStreamChunk = $null,
        
        [Parameter(Mandatory=$false)]
        [int]$MaxIterations = 10
    )
    
    # Reset stats
    $script:InvestigationStats = @{
        ToolCalls = 0
        CacheHits = 0
        TotalDuration = 0
        StartTime = Get-Date
    }
    
    # Clear cache for new investigation
    $script:InvestigationCache = @{}
    
    try {
        if ([string]::IsNullOrWhiteSpace($Model)) { $Model = $script:DefaultClaudeModel }

        if ([string]::IsNullOrWhiteSpace($ApiKey)) {
            throw "API key required for AI investigation"
        }
        
        # Status update helper
        $updateStatus = {
            param($message)
            if ($OnStatusUpdate) {
                & $OnStatusUpdate $message
            }
            Write-Host $message -ForegroundColor Cyan
        }
        
        & $updateStatus "🔍 Starting AI investigation..."
        
        # Build system prompt
        $systemPrompt = @"
You are a cybersecurity analyst investigating a security alert. Your goal is to determine if this is:
1. TRUE POSITIVE - Actual malicious activity requiring action
2. FALSE POSITIVE - Legitimate activity that triggered an alert
3. AUTHORIZED ACTIVITY - Intentional activity by authorized personnel

WORKFLOW:
1. Analyze the ticket content to understand the alert
2. Request data you need via available tools (sign-in logs, inbox rules, audit logs, etc.)
3. Gather evidence iteratively - request more data if initial findings are inconclusive
4. Make a determination with confidence level
5. Draft a client-facing reply explaining findings

TOOL USE GUIDELINES:
- Start with broader queries (e.g., 7 days of sign-in logs)
- Request more specific data if you find suspicious patterns
- You can call tools multiple times with different parameters
- Typical investigation needs 3-7 tool calls

REPLY FORMAT:
When ready to draft the reply, output it in this structure:

VERDICT: [True Positive / False Positive / Authorized Activity]
CONFIDENCE: [High / Medium / Low]
SUMMARY: [One-line summary of findings]

CLIENT REPLY:
[Professional explanation in plain language for the client]

$MemberberryInstructions
"@
        
        # Get tool definitions
        $tools = Get-AIToolDefinitions
        
        # Convert tools to Claude format
        $claudeTools = $tools | ForEach-Object {
            @{
                name = $_.name
                description = $_.description
                input_schema = $_.parameters
            }
        }
        
        # Initial message
        $messages = @(
            @{
                role = "user"
                content = "Please investigate this security alert:`n`n$TicketContent"
            }
        )
        
        $iteration = 0
        $investigation = @{
            Verdict = "Unknown"
            Confidence = "Unknown"
            Summary = ""
            ClientReply = ""
            Evidence = @()
            ToolCalls = @()
            Complete = $false
        }
        
        # Investigation loop
        while ($iteration -lt $MaxIterations -and -not $investigation.Complete) {
            $iteration++
            & $updateStatus "📊 Investigation iteration $iteration/$MaxIterations..."
            
            # Call Claude
            $response = Invoke-ClaudeInvestigation `
                -ApiKey $ApiKey `
                -SystemPrompt $systemPrompt `
                -Messages $messages `
                -Tools $claudeTools `
                -OnStreamChunk $OnStreamChunk `
                -Model $Model `
                -MaxTokens 4096
            
            if (-not $response) {
                throw "Failed to get AI response"
            }
            if (((Get-Date) - $script:InvestigationStats.StartTime).TotalMinutes -ge 5) {
                throw "Investigation stopped after 5 minutes."
            }
            
            $hasToolUse = $false
            $textContent = ""
            $toolResults = New-Object System.Collections.Generic.List[object]
            $assistantContent = New-Object System.Collections.Generic.List[object]
            
            foreach ($block in @($response.content)) {
                $blockType = [string]$block.type
                if ($blockType -eq "text") {
                    $textContent += [string]$block.text
                    $assistantContent.Add(@{ type = 'text'; text = [string]$block.text })
                } elseif ($blockType -eq "tool_use") {
                    $hasToolUse = $true
                    $toolName = [string]$block.name
                    $toolParams = ConvertTo-HashtableDeep $block.input
                    if ($toolParams -isnot [hashtable]) { $toolParams = @{} }
                    $toolId = [string]$block.id
                    
                    & $updateStatus "  Calling tool: $toolName"
                    
                    # Graph and Exchange sessions cannot be shared across runspaces, so tools in one turn run back to back in this session. The cache skips repeats.
                    $toolResult = Invoke-AITool -ToolName $toolName -Parameters $toolParams
                    
                    $investigation.ToolCalls += @{
                        Tool = $toolName
                        Parameters = $toolParams
                        Result = $toolResult
                    }
                    
                    $resultJson = ConvertTo-AIFriendlyJson -Data $toolResult
                    $assistantContent.Add(@{ type = 'tool_use'; id = $toolId; name = $toolName; input = $toolParams })
                    $toolResults.Add(@{
                        type = "tool_result"
                        tool_use_id = $toolId
                        content = $resultJson
                    })
                }
            }
            
            $messages += @{
                role = "assistant"
                content = $assistantContent.ToArray()
            }
            
            if ($hasToolUse) {
                $messages += @{
                    role = "user"
                    content = $toolResults.ToArray()
                }
            } else {
                $investigation.Complete = $true
                $parsed = Get-InvestigationReplyText -Text $textContent
                $investigation.Verdict = $parsed.Verdict
                $investigation.Confidence = $parsed.Confidence
                $investigation.Summary = $parsed.Summary
                $investigation.ClientReply = $parsed.ClientReply
                
                & $updateStatus "Investigation complete: $($investigation.Verdict) ($($investigation.Confidence) confidence)"
            }
        }
        
        if (-not $investigation.Complete) {
            Write-Warning "Investigation reached maximum iterations without completing"
            $investigation.Complete = $true
            $investigation.Verdict = "Inconclusive"
            $investigation.Confidence = "Low"
            $investigation.Summary = "Investigation exceeded maximum iterations"
        }
        
        # Add stats
        $script:InvestigationStats.TotalDuration = ((Get-Date) - $script:InvestigationStats.StartTime).TotalSeconds
        $investigation.Stats = $script:InvestigationStats
        
        return $investigation
        
    } catch {
        Write-Error "Investigation failed: $_"
        return @{
            Verdict = "Error"
            Confidence = "N/A"
            Summary = $_.Exception.Message
            ClientReply = "An error occurred during investigation. Manual review required."
            Evidence = @()
            ToolCalls = @()
            Complete = $true
            Error = $_.Exception.Message
        }
    }
}

<#
.SYNOPSIS
    Clears investigation cache
#>
function Update-AIClientDraft {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$ApiKey,
        [Parameter(Mandatory=$true)][string]$CurrentDraft,
        [Parameter(Mandatory=$true)][string]$Clarification,
        [Parameter(Mandatory=$false)][string]$MemberberryInstructions = "",
        [Parameter(Mandatory=$false)][string]$Model = "",
        [Parameter(Mandatory=$false)][scriptblock]$OnStreamChunk = $null
    )
    if ([string]::IsNullOrWhiteSpace($Model)) { $Model = $script:DefaultClaudeModel }
    $systemPrompt = @"
You revise a client-facing security ticket reply. Return only the revised reply.
Keep every claim grounded in the current draft. Apply the owner's clarification.
Do not add a verdict header.

$MemberberryInstructions
"@
    if ($systemPrompt.Length -gt 8000) { $systemPrompt = $systemPrompt.Substring(0, 8000) }
    $messages = @(
        @{
            role = 'user'
            content = "Current draft:`n$CurrentDraft`n`nOwner clarification:`n$Clarification`n`nWrite the revised reply only."
        }
    )
    $response = Invoke-ClaudeInvestigation -ApiKey $ApiKey -SystemPrompt $systemPrompt -Messages $messages -Model $Model -OnStreamChunk $OnStreamChunk -MaxTokens 2000
    if (-not $response) { throw 'Failed to revise the draft.' }
    $text = ''
    foreach ($block in @($response.content)) {
        if ([string]$block.type -eq 'text') { $text += [string]$block.text }
    }
    return $text.Trim()
}

function Clear-InvestigationCache {
    $script:InvestigationCache = @{}
    $script:InvestigationStats = @{
        ToolCalls = 0
        CacheHits = 0
        TotalDuration = 0
    }
    Write-Host "Investigation cache cleared" -ForegroundColor Yellow
}

# Export module members
Export-ModuleMember -Function @(
    'Get-AIToolDefinitions',
    'Invoke-AITool',
    'Start-AIInvestigation',
    'Update-AIClientDraft',
    'Clear-InvestigationCache',
    'ConvertTo-StableJson',
    'ConvertTo-HashtableDeep',
    'New-ClaudeStreamAssembler',
    'Add-ClaudeStreamEvent',
    'Get-ClaudeStreamContent',
    'Get-InvestigationReplyText'
)
