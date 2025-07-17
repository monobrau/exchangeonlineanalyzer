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
    # ... (copy the full function from the main script)
}
Export-ModuleMember -Function Set-UserSignInBlockedState 