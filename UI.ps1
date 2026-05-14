# ============================================================
# Intune Discovered Apps - UI Launcher
# ============================================================
# Provides a graphical interface to run Get-DiscoveredAppsAllDevices.ps1
# with parameter selection

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ============================================================
# THEME & COLORS (Patch My PC Design System)
# ============================================================
$colors = @{
    Primary        = [System.Drawing.Color]::FromArgb(27, 188, 155)      # Sea-Foam Green #1BBC9B
    PrimaryHover   = [System.Drawing.Color]::FromArgb(17, 164, 134)      # Sea-Foam Green #11A486
    Secondary      = [System.Drawing.Color]::FromArgb(4, 144, 218)       # PMPC Blue #0490DA
    Background     = [System.Drawing.Color]::FromArgb(21, 21, 33)        # Base 900 #151521
    BackgroundAlt  = [System.Drawing.Color]::FromArgb(30, 30, 45)        # Base 800 #1E1E2D
    BackgroundAlt2 = [System.Drawing.Color]::FromArgb(42, 42, 60)        # Base 700 #2A2A3C
    TextPrimary    = [System.Drawing.Color]::White
    TextSecondary  = [System.Drawing.Color]::FromArgb(133, 141, 151)     # Neutral 400 #858D97
    Border         = [System.Drawing.Color]::FromArgb(171, 173, 207)     # Tertiary #ABADCF
    Success        = [System.Drawing.Color]::FromArgb(80, 205, 137)      # Mint Accent #50CD89
    Warning        = [System.Drawing.Color]::FromArgb(255, 199, 0)       # Golden Accent #FFC700
    Error          = [System.Drawing.Color]::FromArgb(241, 65, 108)      # Crimson Accent #F1416C
}

# ============================================================
# MAIN FORM
# ============================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Intune Discovered Apps - Configuration"
$form.Size = New-Object System.Drawing.Size(550, 550)
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.BackColor = $colors.Background
$form.ForeColor = $colors.TextPrimary
$form.Font = New-Object System.Drawing.Font("Poppins", 10, [System.Drawing.FontStyle]::Regular)

# ============================================================
# TITLE PANEL
# ============================================================
$titlePanel = New-Object System.Windows.Forms.Panel
$titlePanel.Location = New-Object System.Drawing.Point(0, 0)
$titlePanel.Size = New-Object System.Drawing.Size(500, 60)
$titlePanel.BackColor = $colors.BackgroundAlt
$titlePanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Script Configuration"
$titleLabel.Font = New-Object System.Drawing.Font("Poppins", 16, [System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor = $colors.TextPrimary
$titleLabel.Location = New-Object System.Drawing.Point(20, 15)
$titleLabel.AutoSize = $true
$titlePanel.Controls.Add($titleLabel)

$form.Controls.Add($titlePanel)

# ============================================================
# MAIN CONTENT PANEL
# ============================================================
$contentPanel = New-Object System.Windows.Forms.Panel
$contentPanel.Location = New-Object System.Drawing.Point(0, 60)
$contentPanel.Size = New-Object System.Drawing.Size(500, 340)
$contentPanel.BackColor = $colors.Background
$contentPanel.AutoScroll = $false

# ============================================================
# DISPLAY NAME INPUT SECTION
# ============================================================
$displayNameLabel = New-Object System.Windows.Forms.Label
$displayNameLabel.Text = "Filter by App Name (Optional)"
$displayNameLabel.Location = New-Object System.Drawing.Point(20, 20)
$displayNameLabel.Size = New-Object System.Drawing.Size(460, 20)
$displayNameLabel.ForeColor = $colors.TextPrimary
$displayNameLabel.Font = New-Object System.Drawing.Font("Poppins", 10, [System.Drawing.FontStyle]::Bold)
$contentPanel.Controls.Add($displayNameLabel)

$displayNameHint = New-Object System.Windows.Forms.Label
$displayNameHint.Text = "Use * for wildcards (e.g., 'Java*' or '*Chrome*')"
$displayNameHint.Location = New-Object System.Drawing.Point(20, 40)
$displayNameHint.Size = New-Object System.Drawing.Size(460, 16)
$displayNameHint.ForeColor = $colors.TextSecondary
$displayNameHint.Font = New-Object System.Drawing.Font("Poppins", 8, [System.Drawing.FontStyle]::Italic)
$contentPanel.Controls.Add($displayNameHint)

$displayNameInput = New-Object System.Windows.Forms.TextBox
$displayNameInput.Location = New-Object System.Drawing.Point(20, 60)
$displayNameInput.Size = New-Object System.Drawing.Size(460, 32)
$displayNameInput.BackColor = $colors.BackgroundAlt2
$displayNameInput.ForeColor = $colors.TextPrimary
$displayNameInput.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$displayNameInput.Font = New-Object System.Drawing.Font("Poppins", 10)
$displayNameInput.Padding = New-Object System.Windows.Forms.Padding(8)
$contentPanel.Controls.Add($displayNameInput)

# ============================================================
# SWITCHES SECTION
# ============================================================
$switchLabel = New-Object System.Windows.Forms.Label
$switchLabel.Text = "Export Options"
$switchLabel.Location = New-Object System.Drawing.Point(20, 110)
$switchLabel.Size = New-Object System.Drawing.Size(460, 20)
$switchLabel.ForeColor = $colors.TextPrimary
$switchLabel.Font = New-Object System.Drawing.Font("Poppins", 10, [System.Drawing.FontStyle]::Bold)
$contentPanel.Controls.Add($switchLabel)

# CSV Switch
$csvCheckbox = New-Object System.Windows.Forms.CheckBox
$csvCheckbox.Location = New-Object System.Drawing.Point(20, 140)
$csvCheckbox.Size = New-Object System.Drawing.Size(460, 24)
$csvCheckbox.Text = "Export as CSV (Default: Excel)"
$csvCheckbox.ForeColor = $colors.TextPrimary
$csvCheckbox.BackColor = $colors.Background
$csvCheckbox.Font = New-Object System.Drawing.Font("Poppins", 10)
$csvCheckbox.Appearance = [System.Windows.Forms.Appearance]::Button
$csvCheckbox.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$csvCheckbox.FlatAppearance.BorderColor = $colors.Border
$csvCheckbox.FlatAppearance.CheckedBackColor = $colors.Primary
$csvCheckbox.FlatAppearance.MouseDownBackColor = $colors.PrimaryHover
$contentPanel.Controls.Add($csvCheckbox)

# BuildGroup Switch
$buildGroupCheckbox = New-Object System.Windows.Forms.CheckBox
$buildGroupCheckbox.Location = New-Object System.Drawing.Point(20, 170)
$buildGroupCheckbox.Size = New-Object System.Drawing.Size(460, 24)
$buildGroupCheckbox.Text = "Build Dynamic Device Group"
$buildGroupCheckbox.ForeColor = $colors.TextPrimary
$buildGroupCheckbox.BackColor = $colors.Background
$buildGroupCheckbox.Font = New-Object System.Drawing.Font("Poppins", 10)
$buildGroupCheckbox.Appearance = [System.Windows.Forms.Appearance]::Button
$buildGroupCheckbox.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buildGroupCheckbox.FlatAppearance.BorderColor = $colors.Border
$buildGroupCheckbox.FlatAppearance.CheckedBackColor = $colors.Primary
$buildGroupCheckbox.FlatAppearance.MouseDownBackColor = $colors.PrimaryHover
$contentPanel.Controls.Add($buildGroupCheckbox)

# ============================================================
# INFO PANEL
# ============================================================
$infoPanel = New-Object System.Windows.Forms.Panel
$infoPanel.Location = New-Object System.Drawing.Point(20, 210)
$infoPanel.Size = New-Object System.Drawing.Size(460, 80)
$infoPanel.BackColor = $colors.BackgroundAlt2
$infoPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

$infoTitle = New-Object System.Windows.Forms.Label
$infoTitle.Text = "Info"
$infoTitle.Location = New-Object System.Drawing.Point(10, 5)
$infoTitle.Size = New-Object System.Drawing.Size(440, 16)
$infoTitle.ForeColor = $colors.TextSecondary
$infoTitle.Font = New-Object System.Drawing.Font("Poppins", 9, [System.Drawing.FontStyle]::Bold)
$infoPanel.Controls.Add($infoTitle)

$infoText = New-Object System.Windows.Forms.Label
$infoText.Text = "The script will retrieve all discovered applications from Intune and generate a comprehensive report. This may take several minutes for large environments."
$infoText.Location = New-Object System.Drawing.Point(10, 25)
$infoText.Size = New-Object System.Drawing.Size(440, 50)
$infoText.ForeColor = $colors.TextSecondary
$infoText.Font = New-Object System.Drawing.Font("Poppins", 9)
$infoText.AutoSize = $false
$infoPanel.Controls.Add($infoText)

$contentPanel.Controls.Add($infoPanel)

$form.Controls.Add($contentPanel)

# ============================================================
# FOOTER PANEL (Buttons)
# ============================================================
$footerPanel = New-Object System.Windows.Forms.Panel
$footerPanel.Location = New-Object System.Drawing.Point(0, 400)
$footerPanel.Size = New-Object System.Drawing.Size(500, 50)
$footerPanel.BackColor = $colors.BackgroundAlt
$footerPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None

# Cancel Button
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(20, 10)
$cancelButton.Size = New-Object System.Drawing.Size(200, 30)
$cancelButton.Text = "Cancel"
$cancelButton.BackColor = $colors.BackgroundAlt2
$cancelButton.ForeColor = $colors.TextPrimary
$cancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$cancelButton.FlatAppearance.BorderColor = $colors.Border
$cancelButton.Font = New-Object System.Drawing.Font("Poppins", 10, [System.Drawing.FontStyle]::Bold)
$cancelButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$cancelButton.Add_Click({
    $form.Close()
})
$footerPanel.Controls.Add($cancelButton)

# Run Button
$runButton = New-Object System.Windows.Forms.Button
$runButton.Location = New-Object System.Drawing.Point(280, 10)
$runButton.Size = New-Object System.Drawing.Size(200, 30)
$runButton.Text = "Run Script"
$runButton.BackColor = $colors.Primary
$runButton.ForeColor = $colors.TextPrimary
$runButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$runButton.FlatAppearance.BorderColor = $colors.Primary
$runButton.Font = New-Object System.Drawing.Font("Poppins", 10, [System.Drawing.FontStyle]::Bold)
$runButton.Cursor = [System.Windows.Forms.Cursors]::Hand

$runButton.Add_MouseEnter({
    $runButton.BackColor = $colors.PrimaryHover
})

$runButton.Add_MouseLeave({
    $runButton.BackColor = $colors.Primary
})

$runButton.Add_Click({
    # Validation: BuildGroup requires DisplayName
    if ($buildGroupCheckbox.Checked -and [string]::IsNullOrWhiteSpace($displayNameInput.Text)) {
        [System.Windows.Forms.MessageBox]::Show(
            "App Name is required when 'Build Dynamic Device Group' is selected.",
            "Validation Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }
    
    # Build command arguments
    $scriptPath = Split-Path -Parent $PSCommandPath
    $scriptFile = Join-Path $scriptPath "Get-DiscoveredAppsAllDevices.ps1"
    
    $args = @()
    
    if ($displayNameInput.Text) {
        $args += "-DisplayName '$($displayNameInput.Text)'"
    }
    
    if ($csvCheckbox.Checked) {
        $args += "-CSV"
    }
    
    if ($buildGroupCheckbox.Checked) {
        $args += "-BuildGroup"
    }
    
    # Build command string
    $cmdString = "& '$scriptFile'"
    if ($args.Count -gt 0) {
        $cmdString += " " + ($args -join " ")
    }
    
    Write-Host "Executing: $cmdString" -ForegroundColor Cyan
    
    # Execute script in new PowerShell window
    Start-Process PowerShell -ArgumentList "-NoExit", "-Command", $cmdString
    
    $form.Close()
})

$footerPanel.Controls.Add($runButton)

$form.Controls.Add($footerPanel)

# ============================================================
# SHOW FORM
# ============================================================
[void]$form.ShowDialog()
