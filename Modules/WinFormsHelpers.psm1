<#
.SYNOPSIS
    Small WinForms helpers for Exchange Online Analyzer and related tools.
.DESCRIPTION
    Use these when building or editing dialogs to reduce repetitive New-Object / property boilerplate.
#>

function Add-ToolTip {
    <#
    .SYNOPSIS
        Attaches a ToolTip to a WinForms control (same behavior as legacy in-script helper).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.Control]$Control,
        [Parameter(Mandatory = $true)]
        [string]$Text
    )
    $tooltip = New-Object System.Windows.Forms.ToolTip
    $tooltip.AutoPopDelay = 5000
    $tooltip.InitialDelay = 1000
    $tooltip.ReshowDelay = 500
    $tooltip.ShowAlways = $true
    $tooltip.SetToolTip($Control, $Text)
}

function New-EOALabel {
    <#
    .SYNOPSIS
        Creates a Label with optional font, bounds, and autosize.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Text,
        [System.Drawing.Point]$Location,
        [System.Drawing.Size]$Size,
        [System.Drawing.Font]$Font,
        [switch]$AutoSize
    )
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Text
    if ($PSBoundParameters.ContainsKey('Location')) { $lbl.Location = $Location }
    if ($PSBoundParameters.ContainsKey('Size')) { $lbl.Size = $Size }
    if ($Font) { $lbl.Font = $Font }
    if ($AutoSize) { $lbl.AutoSize = $true }
    return $lbl
}

function New-EOAButton {
    <#
    .SYNOPSIS
        Creates a Button with text, location, and size.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Text,
        [Parameter(Mandatory = $true)][System.Drawing.Point]$Location,
        [Parameter(Mandatory = $true)][System.Drawing.Size]$Size,
        [System.Drawing.Font]$Font
    )
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $Text
    $btn.Location = $Location
    $btn.Size = $Size
    if ($Font) { $btn.Font = $Font }
    return $btn
}

function New-EOATextBox {
    <#
    .SYNOPSIS
        Creates a TextBox with optional multiline and scrollbars.
    #>
    param(
        [Parameter(Mandatory = $true)][System.Drawing.Point]$Location,
        [Parameter(Mandatory = $true)][System.Drawing.Size]$Size,
        [switch]$Multiline,
        [System.Windows.Forms.ScrollBars]$ScrollBars = 'None'
    )
    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Location = $Location
    $tb.Size = $Size
    if ($Multiline) {
        $tb.Multiline = $true
        $tb.ScrollBars = $ScrollBars
    }
    return $tb
}

function New-EOAPanel {
    <#
    .SYNOPSIS
        Creates a Panel with Dock and optional padding.
    #>
    param(
        [System.Windows.Forms.DockStyle]$Dock = 'None',
        [System.Windows.Forms.Padding]$Padding
    )
    $p = New-Object System.Windows.Forms.Panel
    if ($Dock -ne 'None') { $p.Dock = $Dock }
    if ($PSBoundParameters.ContainsKey('Padding')) { $p.Padding = $Padding }
    return $p
}

Export-ModuleMember -Function Add-ToolTip, New-EOALabel, New-EOAButton, New-EOATextBox, New-EOAPanel
