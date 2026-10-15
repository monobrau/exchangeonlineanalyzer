# ConnectWiseIntegration.psm1
# REST API integration for ConnectWise Manage (PSA)

# Module-level variables for connection state
$script:CWSession = $null

<#
.SYNOPSIS
    Connects to ConnectWise Manage REST API
.DESCRIPTION
    Authenticates with ConnectWise using company ID and API keys (public + private).
    Stores connection state in module scope and validates connectivity.
.PARAMETER CompanyID
    ConnectWise company identifier (e.g., "yourcompany")
.PARAMETER PublicKey
    API public key (ClientID)
.PARAMETER PrivateKey
    API private key (SecureString)
.PARAMETER BaseUrl
    ConnectWise API base URL (default: https://api-na.myconnectwise.net)
.EXAMPLE
    $privKey = Read-Host "Private Key" -AsSecureString
    Connect-ConnectWise -CompanyID "acme" -PublicKey "ABC123" -PrivateKey $privKey
#>
function Connect-ConnectWise {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$CompanyID,
        
        [Parameter(Mandatory=$true)]
        [string]$PublicKey,
        
        [Parameter(Mandatory=$true)]
        [securestring]$PrivateKey,
        
        [Parameter(Mandatory=$false)]
        [string]$BaseUrl = "https://api-na.myconnectwise.net"
    )
    
    try {
        # Convert SecureString to plain text for API (REST requires it)
        $privateKeyPlain = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($PrivateKey)
        )
        
        # Build Basic Auth header: Base64(CompanyID+PublicKey:PrivateKey)
        $authString = "$CompanyID+$PublicKey`:$privateKeyPlain"
        $authBytes = [System.Text.Encoding]::UTF8.GetBytes($authString)
        $authBase64 = [System.Convert]::ToBase64String($authBytes)
        
        # Clear plaintext private key from memory
        $privateKeyPlain = $null
        [System.GC]::Collect()
        
        # Build headers
        $headers = @{
            'Authorization' = "Basic $authBase64"
            'clientId' = [guid]::NewGuid().ToString()  # Required by CW API
            'Content-Type' = 'application/json'
            'Accept' = 'application/json'
        }
        
        # Test connection with /system/info endpoint
        $testUrl = "$BaseUrl/v4_6_release/apis/3.0/system/info"
        $response = Invoke-RestMethod -Uri $testUrl -Method Get -Headers $headers -ErrorAction Stop
        
        # Store session
        $script:CWSession = @{
            CompanyID = $CompanyID
            BaseUrl = $BaseUrl
            Headers = $headers
            Connected = $true
            ConnectedAt = Get-Date
            Version = $response.version
        }
        
        Write-Host "✓ Connected to ConnectWise: $CompanyID (v$($response.version))" -ForegroundColor Green
        return $true
        
    } catch {
        Write-Error "Failed to connect to ConnectWise: $($_.Exception.Message)"
        $script:CWSession = $null
        return $false
    }
}

<#
.SYNOPSIS
    Disconnects from ConnectWise and clears session
#>
function Disconnect-ConnectWise {
    $script:CWSession = $null
    Write-Host "Disconnected from ConnectWise" -ForegroundColor Yellow
}

<#
.SYNOPSIS
    Gets current ConnectWise connection status
#>
function Get-ConnectWiseStatus {
    if ($script:CWSession -and $script:CWSession.Connected) {
        return @{
            Connected = $true
            CompanyID = $script:CWSession.CompanyID
            ConnectedAt = $script:CWSession.ConnectedAt
            Version = $script:CWSession.Version
        }
    } else {
        return @{ Connected = $false }
    }
}

<#
.SYNOPSIS
    Makes a REST API call to ConnectWise
.DESCRIPTION
    Internal helper for making authenticated API calls
.PARAMETER Endpoint
    API endpoint path (e.g., "service/tickets/12345")
.PARAMETER Method
    HTTP method (GET, POST, PATCH, PUT, DELETE)
.PARAMETER Body
    Request body (will be converted to JSON)
#>
function Invoke-CWRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Endpoint,
        
        [Parameter(Mandatory=$false)]
        [Microsoft.PowerShell.Commands.WebRequestMethod]$Method = 'Get',
        
        [Parameter(Mandatory=$false)]
        [object]$Body = $null
    )
    
    if (-not $script:CWSession -or -not $script:CWSession.Connected) {
        throw "Not connected to ConnectWise. Call Connect-ConnectWise first."
    }
    
    try {
        # Build full URL
        $url = "$($script:CWSession.BaseUrl)/v4_6_release/apis/3.0/$Endpoint"
        
        # Prepare parameters
        $params = @{
            Uri = $url
            Method = $Method
            Headers = $script:CWSession.Headers
            ErrorAction = 'Stop'
        }
        
        if ($Body) {
            $params['Body'] = ($Body | ConvertTo-Json -Depth 10 -Compress)
        }
        
        # Make request
        $response = Invoke-RestMethod @params
        return $response
        
    } catch {
        Write-Error "ConnectWise API error: $($_.Exception.Message)"
        throw
    }
}

<#
.SYNOPSIS
    Gets a ConnectWise ticket by ID
.PARAMETER TicketNumber
    Ticket number (integer)
.EXAMPLE
    $ticket = Get-CWTicket -TicketNumber 1838914
#>
function Get-CWTicket {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [int]$TicketNumber
    )
    
    try {
        $ticket = Invoke-CWRequest -Endpoint "service/tickets/$TicketNumber" -Method Get
        return $ticket
    } catch {
        Write-Error "Failed to get ticket $TicketNumber : $_"
        return $null
    }
}

<#
.SYNOPSIS
    Creates a new ConnectWise ticket
.PARAMETER Summary
    Ticket summary/subject (required)
.PARAMETER Board
    Board name or ID (optional, uses default if not specified)
.PARAMETER Company
    Company name or ID (required)
.PARAMETER ContactName
    Contact name (optional)
.PARAMETER Description
    Initial ticket description (optional)
.PARAMETER Priority
    Priority name or ID (optional)
.PARAMETER Type
    Ticket type name or ID (optional)
.PARAMETER SubType
    Ticket subtype name or ID (optional)
.PARAMETER AssignTo
    Member identifier to assign ticket to (optional)
.PARAMETER Status
    Initial status (optional, defaults to board default)
.EXAMPLE
    $ticket = New-CWTicket -Summary "Security Alert Investigation" -Company "Acme Corp" -Description $ticketText -AssignTo "jdoe"
#>
function New-CWTicket {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Summary,
        
        [Parameter(Mandatory=$false)]
        [string]$Board = "",
        
        [Parameter(Mandatory=$true)]
        [string]$Company,
        
        [Parameter(Mandatory=$false)]
        [string]$ContactName = "",
        
        [Parameter(Mandatory=$false)]
        [string]$Description = "",
        
        [Parameter(Mandatory=$false)]
        [string]$Priority = "",
        
        [Parameter(Mandatory=$false)]
        [string]$Type = "",
        
        [Parameter(Mandatory=$false)]
        [string]$SubType = "",
        
        [Parameter(Mandatory=$false)]
        [string]$AssignTo = "",
        
        [Parameter(Mandatory=$false)]
        [string]$Status = ""
    )
    
    try {
        # Build ticket object
        $ticketBody = @{
            summary = $Summary
            company = @{ identifier = $Company }
        }
        
        if ($Board) { $ticketBody['board'] = @{ name = $Board } }
        if ($ContactName) { $ticketBody['contact'] = @{ name = $ContactName } }
        if ($Description) { $ticketBody['initialDescription'] = $Description }
        if ($Priority) { $ticketBody['priority'] = @{ name = $Priority } }
        if ($Type) { $ticketBody['type'] = @{ name = $Type } }
        if ($SubType) { $ticketBody['subType'] = @{ name = $SubType } }
        if ($AssignTo) { $ticketBody['owner'] = @{ identifier = $AssignTo } }
        if ($Status) { $ticketBody['status'] = @{ name = $Status } }
        
        $ticket = Invoke-CWRequest -Endpoint "service/tickets" -Method Post -Body $ticketBody
        Write-Host "✓ Created ticket #$($ticket.id): $($ticket.summary)" -ForegroundColor Green
        return $ticket
        
    } catch {
        Write-Error "Failed to create ticket: $_"
        return $null
    }
}

<#
.SYNOPSIS
    Updates a ConnectWise ticket status
.PARAMETER TicketNumber
    Ticket number to update
.PARAMETER StatusName
    Status name (e.g., "In Progress", "Waiting on Client", "Review")
.EXAMPLE
    Set-CWTicketStatus -TicketNumber 1838914 -StatusName "In Progress"
#>
function Set-CWTicketStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [int]$TicketNumber,
        
        [Parameter(Mandatory=$true)]
        [string]$StatusName
    )
    
    try {
        # PATCH operation to update status
        $patchBody = @(
            @{
                op = "replace"
                path = "status/name"
                value = $StatusName
            }
        )
        
        $result = Invoke-CWRequest -Endpoint "service/tickets/$TicketNumber" -Method Patch -Body $patchBody
        Write-Host "✓ Updated ticket #$TicketNumber status to: $StatusName" -ForegroundColor Green
        return $result
        
    } catch {
        Write-Error "Failed to update ticket status: $_"
        return $null
    }
}

<#
.SYNOPSIS
    Assigns a ticket to a member
.PARAMETER TicketNumber
    Ticket number to update
.PARAMETER MemberIdentifier
    Member identifier (username)
.EXAMPLE
    Set-CWTicketOwner -TicketNumber 1838914 -MemberIdentifier "jdoe"
#>
function Set-CWTicketOwner {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [int]$TicketNumber,
        
        [Parameter(Mandatory=$true)]
        [string]$MemberIdentifier
    )
    
    try {
        $patchBody = @(
            @{
                op = "replace"
                path = "owner/identifier"
                value = $MemberIdentifier
            }
        )
        
        $result = Invoke-CWRequest -Endpoint "service/tickets/$TicketNumber" -Method Patch -Body $patchBody
        Write-Host "✓ Assigned ticket #$TicketNumber to: $MemberIdentifier" -ForegroundColor Green
        return $result
        
    } catch {
        Write-Error "Failed to assign ticket: $_"
        return $null
    }
}

<#
.SYNOPSIS
    Adds a note (internal or detail) to a ticket
.PARAMETER TicketNumber
    Ticket number
.PARAMETER Text
    Note text (HTML supported)
.PARAMETER InternalNote
    If true, creates internal note. If false, creates detail note (visible to client)
.PARAMETER ProcessNotifications
    If true, sends email notifications (default: false for automated notes)
.EXAMPLE
    Add-CWTicketNote -TicketNumber 1838914 -Text "Investigation complete. False positive." -InternalNote $false
#>
function Add-CWTicketNote {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [int]$TicketNumber,
        
        [Parameter(Mandatory=$true)]
        [string]$Text,
        
        [Parameter(Mandatory=$false)]
        [bool]$InternalNote = $false,
        
        [Parameter(Mandatory=$false)]
        [bool]$ProcessNotifications = $false
    )
    
    try {
        $noteBody = @{
            text = $Text
            detailDescriptionFlag = -not $InternalNote  # detailDescriptionFlag = true means visible to client
            internalAnalysisFlag = $InternalNote
            processNotifications = $ProcessNotifications
        }
        
        $result = Invoke-CWRequest -Endpoint "service/tickets/$TicketNumber/notes" -Method Post -Body $noteBody
        $noteType = if ($InternalNote) { "internal note" } else { "detail note" }
        Write-Host "✓ Added $noteType to ticket #$TicketNumber" -ForegroundColor Green
        return $result
        
    } catch {
        Write-Error "Failed to add note to ticket: $_"
        return $null
    }
}

<#
.SYNOPSIS
    Gets available statuses for a board
.PARAMETER BoardName
    Board name (e.g., "Service Board")
.EXAMPLE
    $statuses = Get-CWBoardStatuses -BoardName "Service Board"
#>
function Get-CWBoardStatuses {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$BoardName
    )
    
    try {
        # Get board first
        $boards = Invoke-CWRequest -Endpoint "service/boards?conditions=name='$BoardName'" -Method Get
        if (-not $boards -or $boards.Count -eq 0) {
            Write-Error "Board not found: $BoardName"
            return @()
        }
        
        $boardId = $boards[0].id
        
        # Get statuses for board
        $statuses = Invoke-CWRequest -Endpoint "service/boards/$boardId/statuses" -Method Get
        return $statuses
        
    } catch {
        Write-Error "Failed to get board statuses: $_"
        return @()
    }
}

<#
.SYNOPSIS
    Saves ConnectWise credentials to Windows Credential Manager
.PARAMETER CompanyID
    ConnectWise company ID
.PARAMETER PublicKey
    API public key
.PARAMETER PrivateKey
    API private key (SecureString)
.PARAMETER BaseUrl
    API base URL (optional)
#>
function Save-CWCredential {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$CompanyID,
        
        [Parameter(Mandatory=$true)]
        [string]$PublicKey,
        
        [Parameter(Mandatory=$true)]
        [securestring]$PrivateKey,
        
        [Parameter(Mandatory=$false)]
        [string]$BaseUrl = "https://api-na.myconnectwise.net"
    )
    
    try {
        # Use CredentialManager module or native CredWrite
        # Target name: "ConnectWise_<CompanyID>"
        $targetName = "ConnectWise_$CompanyID"
        
        # Store as JSON: { "CompanyID": "...", "PublicKey": "...", "PrivateKey": "...", "BaseUrl": "..." }
        $privateKeyPlain = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($PrivateKey)
        )
        
        $credJson = @{
            CompanyID = $CompanyID
            PublicKey = $PublicKey
            PrivateKey = $privateKeyPlain
            BaseUrl = $BaseUrl
        } | ConvertTo-Json -Compress
        
        # Clear plaintext
        $privateKeyPlain = $null
        [System.GC]::Collect()
        
        # Try CredentialManager module first
        if (Get-Command New-StoredCredential -ErrorAction SilentlyContinue) {
            # Use CredentialManager module
            $existingCred = Get-StoredCredential -Target $targetName -ErrorAction SilentlyContinue
            if ($existingCred) {
                Remove-StoredCredential -Target $targetName -ErrorAction SilentlyContinue
            }
            
            New-StoredCredential -Target $targetName -UserName $CompanyID -Password $credJson -Type Generic -Persist LocalMachine | Out-Null
            Write-Host "✓ Saved ConnectWise credentials to Windows Credential Manager" -ForegroundColor Green
            return $true
        } else {
            Write-Error "CredentialManager module not available. Install with: Install-Module CredentialManager"
            Write-Host "Alternatively, manually configure CW credentials in Settings." -ForegroundColor Yellow
            return $false
        }
        
    } catch {
        Write-Error "Failed to save ConnectWise credentials: $_"
        return $false
    }
}

<#
.SYNOPSIS
    Loads ConnectWise credentials from Windows Credential Manager
.PARAMETER CompanyID
    ConnectWise company ID
.OUTPUTS
    Hashtable with CompanyID, PublicKey, PrivateKey (SecureString), BaseUrl
#>
function Get-CWCredential {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$CompanyID
    )
    
    try {
        $targetName = "ConnectWise_$CompanyID"
        
        if (Get-Command Get-StoredCredential -ErrorAction SilentlyContinue) {
            $cred = Get-StoredCredential -Target $targetName -ErrorAction SilentlyContinue
            if (-not $cred) {
                Write-Warning "ConnectWise credentials not found in Credential Manager for: $CompanyID"
                return $null
            }
            
            # Parse JSON from password field
            $credJson = $cred.GetNetworkCredential().Password | ConvertFrom-Json
            
            # Convert PrivateKey back to SecureString
            $privateKeySecure = ConvertTo-SecureString -String $credJson.PrivateKey -AsPlainText -Force
            
            return @{
                CompanyID = $credJson.CompanyID
                PublicKey = $credJson.PublicKey
                PrivateKey = $privateKeySecure
                BaseUrl = $credJson.BaseUrl
            }
        } else {
            Write-Error "CredentialManager module not available. Install with: Install-Module CredentialManager"
            return $null
        }
        
    } catch {
        Write-Error "Failed to load ConnectWise credentials: $_"
        return $null
    }
}

# Export module members
Export-ModuleMember -Function @(
    'Connect-ConnectWise',
    'Disconnect-ConnectWise',
    'Get-ConnectWiseStatus',
    'Get-CWTicket',
    'New-CWTicket',
    'Set-CWTicketStatus',
    'Set-CWTicketOwner',
    'Add-CWTicketNote',
    'Get-CWBoardStatuses',
    'Save-CWCredential',
    'Get-CWCredential'
)
