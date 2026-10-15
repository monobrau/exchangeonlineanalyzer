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
                        description = "Days to search (max 10)"
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
        [hashtable]$Parameters = @{},
        
        [Parameter(Mandatory=$false)]
        [switch]$UseCache = $true
    )
    
    $script:InvestigationStats.ToolCalls++
    $startTime = Get-Date
    
    try {
        # Generate cache key
        $cacheKey = "$ToolName|$($Parameters | ConvertTo-Json -Compress)"
        
        # Check cache
        if ($UseCache -and $script:InvestigationCache.ContainsKey($cacheKey)) {
            $script:InvestigationStats.CacheHits++
            Write-Host "  [Cache Hit] $ToolName" -ForegroundColor DarkGray
            return $script:InvestigationCache[$cacheKey]
        }
        
        Write-Host "  [Executing] $ToolName..." -ForegroundColor Cyan
        
        # Map tool to PowerShell function
        $result = switch ($ToolName) {
            "get_signin_logs" {
                $upns = $Parameters['user_principals']
                $days = if ($Parameters['days_back']) { $Parameters['days_back'] } else { 7 }
                Get-EntraSignInLogs -UserPrincipalNames $upns -Days $days
            }
            
            "get_inbox_rules" {
                $upns = $Parameters['mailbox_upns']
                $allRules = @()
                foreach ($upn in $upns) {
                    try {
                        $rules = Get-InboxRule -Mailbox $upn -ErrorAction SilentlyContinue
                        if ($rules) {
                            $allRules += $rules
                        }
                    } catch {
                        Write-Warning "Could not get inbox rules for $upn : $_"
                    }
                }
                $allRules
            }
            
            "get_user_details" {
                $upn = $Parameters['user_principal']
                Get-EntraUserDetailsAndRoles -UserPrincipalName $upn
            }
            
            "get_audit_logs" {
                $upn = $Parameters['user_principal']
                $days = if ($Parameters['days_back']) { $Parameters['days_back'] } else { 7 }
                Get-EntraUserAuditLogs -UserPrincipalName $upn -Days $days
            }
            
            "get_mailbox_permissions" {
                $days = if ($Parameters['days_back']) { $Parameters['days_back'] } else { 7 }
                $targets = $Parameters['target_mailboxes']
                if (Get-Command Get-MailboxPermissionChanges -ErrorAction SilentlyContinue) {
                    if ($targets) {
                        Get-MailboxPermissionChanges -DaysBack $days -TargetMailboxes $targets
                    } else {
                        Get-MailboxPermissionChanges -DaysBack $days
                    }
                } else {
                    @{ Error = "Mailbox permission tracking not available" }
                }
            }
            
            "get_oauth_permissions" {
                $days = if ($Parameters['days_back']) { $Parameters['days_back'] } else { 30 }
                if (Get-Command Get-OAuthPermissionGrants -ErrorAction SilentlyContinue) {
                    Get-OAuthPermissionGrants -DaysBack $days
                } else {
                    @{ Error = "OAuth permission tracking not available" }
                }
            }
            
            "get_device_compliance" {
                $upns = $Parameters['user_principals']
                if (Get-Command Get-IntuneDeviceRecords -ErrorAction SilentlyContinue) {
                    Get-IntuneDeviceRecords -UserPrincipalNames $upns
                } else {
                    @{ Error = "Device compliance tracking not available (requires Intune)" }
                }
            }
            
            "search_message_trace" {
                $sender = $Parameters['sender']
                $recipient = $Parameters['recipient']
                $days = if ($Parameters['days_back']) { $Parameters['days_back'] } else { 7 }
                
                $traceParams = @{}
                if ($sender) { $traceParams['SenderAddress'] = $sender }
                if ($recipient) { $traceParams['RecipientAddress'] = $recipient }
                $traceParams['StartDate'] = (Get-Date).AddDays(-$days)
                $traceParams['EndDate'] = Get-Date
                
                try {
                    Get-MessageTrace @traceParams -ErrorAction SilentlyContinue
                } catch {
                    @{ Error = "Message trace failed: $_" }
                }
            }
            
            default {
                @{ Error = "Unknown tool: $ToolName" }
            }
        }
        
        # Cache result
        if ($UseCache) {
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
        [string]$Model = "claude-3-5-sonnet-20241022"
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
        
        $jsonBody = $requestBody | ConvertTo-Json -Depth 10 -Compress
        
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
        
        $fullResponse = @{
            content = @()
            stop_reason = $null
            usage = @{}
        }
        
        # Read SSE stream
        while (-not $reader.EndOfStream) {
            $line = $reader.ReadLine()
            
            if ($line.StartsWith("data: ")) {
                $data = $line.Substring(6)
                
                if ($data -eq "[DONE]") {
                    break
                }
                
                try {
                    $event = $data | ConvertFrom-Json
                    
                    # Handle different event types
                    if ($event.type -eq "content_block_delta") {
                        if ($event.delta.type -eq "text_delta") {
                            $text = $event.delta.text
                            $fullResponse.content += @{ type = "text"; text = $text }
                            
                            # Call chunk callback
                            if ($OnChunk) {
                                & $OnChunk $text
                            }
                        }
                    } elseif ($event.type -eq "message_stop") {
                        $fullResponse.stop_reason = $event.message.stop_reason
                        break
                    } elseif ($event.type -eq "message_delta") {
                        $fullResponse.usage = $event.usage
                    }
                } catch {
                    # Ignore JSON parse errors in stream
                }
            }
        }
        
        $reader.Close()
        $stream.Close()
        $response.Close()
        
        return $fullResponse
        
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
                -MaxTokens 4096
            
            if (-not $response) {
                throw "Failed to get AI response"
            }
            
            # Process response
            $hasToolUse = $false
            $textContent = ""
            $toolResults = @()
            
            foreach ($block in $response.content) {
                if ($block.type -eq "text") {
                    $textContent += $block.text
                } elseif ($block.type -eq "tool_use") {
                    $hasToolUse = $true
                    $toolName = $block.name
                    $toolParams = $block.input
                    $toolId = $block.id
                    
                    & $updateStatus "  🔧 Calling tool: $toolName"
                    
                    # Execute tool
                    $toolResult = Invoke-AITool -ToolName $toolName -Parameters $toolParams
                    
                    # Record tool call
                    $investigation.ToolCalls += @{
                        Tool = $toolName
                        Parameters = $toolParams
                        Result = $toolResult
                    }
                    
                    # Convert result to AI-friendly format
                    $resultJson = ConvertTo-AIFriendlyJson -Data $toolResult
                    
                    $toolResults += @{
                        type = "tool_result"
                        tool_use_id = $toolId
                        content = $resultJson
                    }
                }
            }
            
            # Add assistant response to messages
            $messages += @{
                role = "assistant"
                content = $response.content
            }
            
            if ($hasToolUse) {
                # Add tool results and continue loop
                $messages += @{
                    role = "user"
                    content = $toolResults
                }
            } else {
                # No tool use - AI has made final determination
                $investigation.Complete = $true
                
                # Parse verdict from response
                if ($textContent -match "VERDICT:\s*(.+)") {
                    $investigation.Verdict = $Matches[1].Trim()
                }
                if ($textContent -match "CONFIDENCE:\s*(.+)") {
                    $investigation.Confidence = $Matches[1].Trim()
                }
                if ($textContent -match "SUMMARY:\s*(.+)") {
                    $investigation.Summary = $Matches[1].Trim()
                }
                if ($textContent -match "CLIENT REPLY:\s*(.+)") {
                    $investigation.ClientReply = $Matches[1].Trim()
                } else {
                    # Fallback: use entire text as reply
                    $investigation.ClientReply = $textContent
                }
                
                & $updateStatus "✅ Investigation complete: $($investigation.Verdict) ($($investigation.Confidence) confidence)"
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
    'Clear-InvestigationCache'
)
