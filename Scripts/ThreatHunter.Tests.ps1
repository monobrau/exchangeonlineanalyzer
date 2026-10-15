# Behavioral tests for the threat hunter helpers. No network calls.
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $root 'Modules/AIInvestigator.psm1') -Force
Import-Module (Join-Path $root 'Modules/ConnectWiseIntegration.psm1') -Force

$failed = 0
function Assert-True {
    param($Condition, [string]$Name)
    if ($Condition) { Write-Host "PASS  $Name" -ForegroundColor Green }
    else { Write-Host "FAIL  $Name" -ForegroundColor Red; $script:failed++ }
}

$patch = @(
    @{ op = 'replace'; path = '/status/id'; value = 42 }
)
$patchJson = ConvertTo-ConnectWiseJson $patch
Assert-True ($patchJson.StartsWith('[')) "ConnectWise patch stays a JSON array: $patchJson"
Assert-True ($patchJson -match '"/status/id"') 'ConnectWise patch path is a JSON pointer'

$tool = Get-AIToolDefinitions | Where-Object { $_.name -eq 'get_signin_logs' } | Select-Object -First 1
$toolJson = ConvertTo-StableJson @{ name = $tool.name; input_schema = $tool.parameters }
Assert-True ($toolJson -match '"required":\["user_principals"\]') "Tool schema keeps a one-item required array: $toolJson"

$state = New-ClaudeStreamAssembler
$events = @(
    '{"type":"content_block_start","index":0,"content_block":{"type":"text","text":""}}',
    '{"type":"content_block_delta","index":0,"delta":{"type":"text_delta","text":"Checking sign-ins."}}',
    '{"type":"content_block_stop","index":0}',
    '{"type":"content_block_start","index":1,"content_block":{"type":"tool_use","id":"toolu_1","name":"get_signin_logs","input":{}}}',
    '{"type":"content_block_delta","index":1,"delta":{"type":"input_json_delta","partial_json":"{\"user_principals\":[\"a@b.com\"],\"days_back\":14}"}}',
    '{"type":"content_block_stop","index":1}',
    '{"type":"message_delta","delta":{"stop_reason":"tool_use","stop_sequence":null},"usage":{"output_tokens":12}}',
    '{"type":"message_stop"}'
)
foreach ($line in $events) {
    Add-ClaudeStreamEvent -State $state -Event ($line | ConvertFrom-Json)
}
$content = @((Get-ClaudeStreamContent -State $state).Items)
Assert-True ($content.Count -eq 2) "Stream assembled $($content.Count) blocks"
Assert-True ($content[1].name -eq 'get_signin_logs') 'Stream kept the tool name'
Assert-True ($content[1].input -is [hashtable]) 'Tool input is a hashtable'
Assert-True (@($content[1].input.user_principals) -contains 'a@b.com') 'Tool input kept the user principal'
Assert-True ([int]$content[1].input.days_back -eq 14) 'Tool input kept days_back'
Assert-True ($state.StopReason -eq 'tool_use') 'Stop reason comes from message_delta'

$reply = @'
VERDICT: False Positive
CONFIDENCE: High
SUMMARY: The sign-in matches the user's usual device.

CLIENT REPLY:
Hello,

This alert is a false positive. The sign-in came from the user's known device.
'@
$parsed = Get-InvestigationReplyText -Text $reply
Assert-True ($parsed.Verdict -eq 'False Positive') 'Verdict line parsed'
Assert-True ($parsed.ClientReply -match 'known device') 'Client reply keeps later lines'
Assert-True ($parsed.ClientReply -notmatch 'VERDICT:') 'Client reply omits the verdict header'

$nested = ConvertTo-HashtableDeep ((ConvertTo-StableJson $content[1].input) | ConvertFrom-Json)
Assert-True ($nested -is [hashtable]) 'Round-trip tool input is still a hashtable'
Assert-True ((Invoke-AITool -ToolName 'get_signin_logs' -Parameters $nested -UseCache:$false).Error -match 'Not connected') 'Disconnected Graph returns an error instead of throwing'

if ($failed -gt 0) { throw "$failed test(s) failed" }
Write-Host "All threat hunter tests passed." -ForegroundColor Green
