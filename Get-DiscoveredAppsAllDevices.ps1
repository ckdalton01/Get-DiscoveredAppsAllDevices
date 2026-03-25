# ============================================================
# PARAMETERS
# ============================================================
param(
    [Parameter(Mandatory = $false)]
    [string]$DisplayName,
    
    [Parameter(Mandatory = $false)]
    [switch]$CSV
)

# ============================================================
# MODULE CHECKS
# ============================================================
$requiredModules = @('Microsoft.Graph', 'ImportExcel')
$missingModules = @()

Write-Host "Checking for required modules..." -ForegroundColor Cyan

foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        $missingModules += $module
        Write-Host "  [MISSING] $module" -ForegroundColor Red
    } else {
        Write-Host "  [OK] $module" -ForegroundColor Green
    }
}

if ($missingModules.Count -gt 0) {
    Write-Host "`nERROR: Missing required modules!" -ForegroundColor Red
    Write-Host "Please install the following modules before running this script:" -ForegroundColor Yellow
    foreach ($module in $missingModules) {
        Write-Host "  Install-Module $module -Scope CurrentUser" -ForegroundColor White
    }
    Write-Host "`nScript execution halted." -ForegroundColor Red
    exit 1
}

Write-Host "All required modules are installed.`n" -ForegroundColor Green

# ============================================================
# GRAPH AUTH
# ============================================================
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan

try {
    Connect-MgGraph -Scopes `
        "DeviceManagementApps.Read.All",
        "DeviceManagementManagedDevices.Read.All" `
        -ErrorAction Stop
    
    # Verify authentication was successful
    $context = Get-MgContext
    
    if (-not $context) {
        throw "Authentication verification failed - no context available"
    }
    
    Write-Host "Successfully authenticated to Microsoft Graph" -ForegroundColor Green
    Write-Host "  Account: $($context.Account)" -ForegroundColor Gray
    Write-Host "  Scopes: $($context.Scopes -join ', ')`n" -ForegroundColor Gray
}
catch {
    Write-Host "`nERROR: Failed to authenticate to Microsoft Graph!" -ForegroundColor Red
    
    # Provide specific guidance for common errors
    if ($_.Exception.Message -like "*window handle*") {
        Write-Host "`nThis error typically occurs when running in certain environments." -ForegroundColor Yellow
        Write-Host "Try one of these solutions:" -ForegroundColor Yellow
        Write-Host "  1. Run PowerShell as Administrator" -ForegroundColor White
        Write-Host "  2. Use a different authentication method:" -ForegroundColor White
        Write-Host "     Connect-MgGraph -UseDeviceCode" -ForegroundColor Cyan
        Write-Host "  3. Run from Windows PowerShell instead of PowerShell Core" -ForegroundColor White
    }
    elseif ($_.Exception.Message -like "*canceled*") {
        Write-Host "`nAuthentication was canceled by user." -ForegroundColor Yellow
        Write-Host "Please run the script again and complete the authentication process." -ForegroundColor Yellow
    }
    else {
        Write-Host "`nError details: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "`nTroubleshooting steps:" -ForegroundColor Yellow
        Write-Host "  1. Ensure you have the required permissions (Intune Administrator or Global Administrator)" -ForegroundColor White
        Write-Host "  2. Try disconnecting first: Disconnect-MgGraph" -ForegroundColor White
        Write-Host "  3. Check your internet connection" -ForegroundColor White
        Write-Host "  4. Try using device code authentication: Connect-MgGraph -UseDeviceCode" -ForegroundColor White
    }
    
    Write-Host "`nScript execution halted." -ForegroundColor Red
    exit 1
}

# ============================================================
# PATHS & SETTINGS
# ============================================================
$basePath   = "."
$csvPath    = Join-Path $basePath "Intune_DiscoveredApps_WithDevices.csv"
$xlsxPath   = Join-Path $basePath "Intune_Software_Inventory.xlsx"

$maxRetries  = 5
$baseDelayMs = 300

# ============================================================
# CSV HEADER (RAW DATA CONTRACT)
# ============================================================
"AppName,Version,InstallCount,DeviceName,DeviceId,OS,UserPrincipal,RetrievalError" |
    Out-File -FilePath $csvPath -Encoding UTF8

Write-Host "Collecting discovered apps from Intune..." -ForegroundColor Cyan

# ============================================================
# DATA COLLECTION (STREAMED TO CSV)
# ============================================================

# Verify we're still authenticated before making API calls
if (-not (Get-MgContext)) {
    Write-Host "`nERROR: Not authenticated to Microsoft Graph!" -ForegroundColor Red
    Write-Host "Please run Connect-MgGraph and try again." -ForegroundColor Yellow
    Write-Host "Script execution halted." -ForegroundColor Red
    exit 1
}

try {
    $apps = Get-MgDeviceManagementDetectedApp -All -ErrorAction Stop
}
catch {
    Write-Host "`nERROR: Failed to retrieve discovered apps from Intune!" -ForegroundColor Red
    Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "`nPossible causes:" -ForegroundColor Yellow
    Write-Host "  - Insufficient permissions (requires DeviceManagementApps.Read.All)" -ForegroundColor White
    Write-Host "  - Network connectivity issues" -ForegroundColor White
    Write-Host "  - Authentication token expired" -ForegroundColor White
    Write-Host "`nScript execution halted." -ForegroundColor Red
    exit 1
}

# Apply DisplayName filter if specified
if ($DisplayName) {
    Write-Host "Filtering apps by DisplayName: '$DisplayName'" -ForegroundColor Yellow
    $apps = $apps | Where-Object { $_.DisplayName -like $DisplayName }
    Write-Host "Found $($apps.Count) matching app(s)" -ForegroundColor Yellow
}
else {
    Write-Host "No filter applied - processing all $($apps.Count) app(s)" -ForegroundColor Cyan
}

foreach ($app in $apps) {

    Write-Host "Processing app: $($app.DisplayName)" -ForegroundColor Yellow

    # Skip pointless expansions
    if ($app.DeviceCount -le 0) {
        "$($app.DisplayName),$($app.Version),0,,,,,No installs" |
            Out-File -FilePath $csvPath -Append -Encoding UTF8
        continue
    }

    Start-Sleep -Milliseconds $baseDelayMs

    $attempt = 0
    $devices = $null

    do {
        try {
            $attempt++

            $devices = Get-MgDeviceManagementDetectedAppManagedDevice `
                -DetectedAppId $app.Id -All -ErrorAction Stop

            break
        }
        catch {
            $retryAfter = $null

            # Safely inspect Retry-After header (may not exist)
            if ($_.Exception -and
                $_.Exception.PSObject.Properties.Name -contains "ResponseHeaders" -and
                $_.Exception.ResponseHeaders -and
                $_.Exception.ResponseHeaders["Retry-After"]) {

                $retryAfter = [int]$_.Exception.ResponseHeaders["Retry-After"]
            }

            if ($attempt -ge $maxRetries) {
                "$($app.DisplayName),$($app.Version),$($app.DeviceCount),,,,,Failed after retries" |
                    Out-File -FilePath $csvPath -Append -Encoding UTF8
                break
            }

            if ($retryAfter) {
                Start-Sleep -Seconds $retryAfter
            }
            else {
                Start-Sleep -Seconds ([math]::Pow(2, $attempt))
            }
        }
    }
    while ($attempt -lt $maxRetries)

    if (-not $devices) { continue }

    foreach ($device in $devices) {
        "$($app.DisplayName),$($app.Version),$($app.DeviceCount),$($device.DeviceName),$($device.Id),$($device.OperatingSystem),$($device.UserPrincipalName)," |
            Out-File -FilePath $csvPath -Append -Encoding UTF8
    }
}

Write-Host "CSV export complete." -ForegroundColor Green

# ============================================================
# OUTPUT LOGIC
# ============================================================
if ($CSV) {
    # CSV-only output requested
    Write-Host "`nCSV output complete: $csvPath" -ForegroundColor Green
    Write-Host "Use the -CSV switch to generate CSV output only." -ForegroundColor Gray
    exit 0
}

# Default: Generate Excel and remove CSV
Write-Host "Generating Excel report..." -ForegroundColor Cyan

$data = Import-Csv $csvPath

# ------------------------------------------------------------
# OVERVIEW TAB
# ------------------------------------------------------------
$overview = @(
    [PSCustomObject]@{ Metric = "Total Discovered Apps"; Value = ($data | Select-Object AppName -Unique).Count }
    [PSCustomObject]@{ Metric = "Total Install Records"; Value = ($data | Where-Object DeviceName).Count }
    [PSCustomObject]@{ Metric = "Apps With Errors"; Value = ($data | Where-Object RetrievalError).Count }
)

# ------------------------------------------------------------
# APPLICATIONS SUMMARY TAB
# ------------------------------------------------------------
$appSummary = $data |
    Where-Object { $_.RetrievalError -eq "" } |
    Group-Object AppName, Version |
    ForEach-Object {
        [PSCustomObject]@{
            AppName      = $_.Group[0].AppName
            Version      = $_.Group[0].Version
            InstallCount = $_.Group[0].InstallCount
        }
    } |
    Sort-Object AppName, Version

# ------------------------------------------------------------
# INSTALLATIONS TAB (ONE ROW PER APP, ALL DEVICES INLINE)
# ------------------------------------------------------------
$installations = $data |
    Where-Object { $_.DeviceName } |
    Group-Object AppName, Version |
    ForEach-Object {

        $deviceList = $_.Group | ForEach-Object {
            "$($_.DeviceName) ($($_.DeviceId))"
        }

        [PSCustomObject]@{
            AppVersion   = "$($_.Group[0].AppName) ($($_.Group[0].Version))"
            InstallCount = $_.Group[0].InstallCount
            Devices      = ($deviceList -join "`n")
        }
    } |
    Sort-Object AppVersion -Descending

# ------------------------------------------------------------
# ERRORS TAB
# ------------------------------------------------------------
$errors = $data |
    Where-Object { $_.RetrievalError } |
    Select-Object AppName, Version, RetrievalError -Unique

# ============================================================
# WRITE EXCEL WORKBOOK
# ============================================================
Remove-Item $xlsxPath -ErrorAction SilentlyContinue

$overview | Export-Excel $xlsxPath `
    -WorksheetName "Overview" `
    -TableName "Overview" `
    -TableStyle Medium2 `
    -BoldTopRow -AutoSize

$appSummary | Export-Excel $xlsxPath `
    -WorksheetName "Applications" `
    -TableName "Applications" `
    -TableStyle Medium2 `
    -BoldTopRow -AutoSize `
    -FreezeTopRow

$installations | Export-Excel $xlsxPath `
    -WorksheetName "Installations" `
    -TableName "Installations" `
    -TableStyle Medium2 `
    -BoldTopRow -AutoSize `
    -FreezeTopRow

$errors | Export-Excel $xlsxPath `
    -WorksheetName "Errors" `
    -TableName "Errors" `
    -TableStyle Medium2 `
    -BoldTopRow -AutoSize

# ------------------------------------------------------------
# FINAL FORMATTING (WRAP DEVICES COLUMN)
# ------------------------------------------------------------
$pkg = Open-ExcelPackage -Path $xlsxPath
$ws  = $pkg.Workbook.Worksheets["Installations"]

# Wrap and align devices column
$devicesColumn = 3  # Note: Column index - Devices is the 3rd column (AppVersion, InstallCount, Devices)

for ($row = 2; $row -le $ws.Dimension.End.Row; $row++) {
    $cell = $ws.Cells[$row, $devicesColumn]
    
    # Set WrapText on the CELL, not the column
    $cell.Style.WrapText = $true
    $cell.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Top
    
    # Auto-adjust row height based on line count
    $lineCount = ([regex]::Matches($cell.Text, "`n")).Count + 1
    $minHeight = $lineCount * 15

    if ($ws.Row($row).Height -lt $minHeight) {
        $ws.Row($row).Height = $minHeight
    }
}

# Set column width after setting cell properties
$ws.Column($devicesColumn).Width = 80

Close-ExcelPackage $pkg

# Clean up intermediate CSV file
Remove-Item $csvPath -ErrorAction SilentlyContinue

Write-Host "`nExcel report created: $xlsxPath" -ForegroundColor Green
Write-Host "Intermediate CSV file has been removed. Use -CSV switch if you need CSV output." -ForegroundColor Gray
