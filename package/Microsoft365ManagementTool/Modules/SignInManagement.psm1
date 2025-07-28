function Set-UserSignInBlockedState {
    param(
        [Parameter(Mandatory=$true)]
        [array]$UserPrincipalNames,
        [Parameter(Mandatory=$true)]
        [bool]$Blocked,
        [Parameter(Mandatory=$false)]
        [System.Windows.Forms.ToolStripStatusLabel]$StatusLabel,
        [Parameter(Mandatory=$false)]
        [System.Windows.Forms.ProgressBar]$ProgressBar,
        [Parameter(Mandatory=$false)]
        [System.Windows.Forms.Form]$MainForm
    )
    
    # Check if we're connected to Microsoft Graph
    try {
        $context = Get-MgContext -ErrorAction Stop
        if (-not $context) {
            throw "Not connected to Microsoft Graph. Please connect first."
        }
    } catch {
        Write-Error "Microsoft Graph connection required. Error: $($_.Exception.Message)"
        if ($StatusLabel) { $StatusLabel.Text = "Error: Microsoft Graph connection required" }
        return
    }
    
    $successCount = 0
    $failCount = 0
    
    foreach ($upn in $UserPrincipalNames) {
        try {
            if ($StatusLabel) { $StatusLabel.Text = "Processing: $upn" }
            
            # Use Microsoft Graph API to update user account enabled status
            $params = @{
                AccountEnabled = -not $Blocked
            }
            
            Update-MgUser -UserId $upn -BodyParameter $params -ErrorAction Stop
            
            $successCount++
            Write-Host "Successfully $(if($Blocked){'blocked'}else{'unblocked'}) sign-in for: $upn" -ForegroundColor Green
            
        } catch {
            $failCount++
            Write-Error "Failed to $(if($Blocked){'block'}else{'unblock'}) sign-in for $upn`: $($_.Exception.Message)"
        }
    }
    
    $message = "Completed: $successCount successful, $failCount failed"
    if ($StatusLabel) { $StatusLabel.Text = $message }
    Write-Host $message -ForegroundColor $(if($failCount -eq 0){'Green'}else{'Yellow'})
}
Export-ModuleMember -Function Set-UserSignInBlockedState 