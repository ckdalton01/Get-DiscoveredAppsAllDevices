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
Write-Host "Connecting to GraphAPI.`n" -ForegroundColor Green
Connect-MgGraph -Scopes `
    "DeviceManagementApps.Read.All",
    "DeviceManagementManagedDevices.Read.All"

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
$apps = Get-MgDeviceManagementDetectedApp -All

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
# EXCEL REPORT GENERATION
# ============================================================
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

Write-Host "Excel report created: $xlsxPath" -ForegroundColor Green
