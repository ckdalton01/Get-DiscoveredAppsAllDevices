# ============================================================
# PARAMETERS
# ============================================================
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$DisplayName,

    [Parameter(Mandatory = $false)]
    [switch]$CSV,

    [Parameter(Mandatory = $false)]
    [switch]$BuildGroup
)

# ============================================================
# PATHS & SETTINGS
# ============================================================
$basePath   = "."
$csvPath    = Join-Path $basePath "Intune_DiscoveredApps_WithDevices.csv"
$xlsxPath   = Join-Path $basePath "Intune_Software_Inventory.xlsx"
$logPath    = Join-Path $basePath "Get-DiscoveredAppsAllDevices.log"

$maxRetries  = 5
$baseDelayMs = 300

$script:EnableVerboseLogging = $PSBoundParameters.ContainsKey('Verbose')
$script:DirectoryDeviceCache = @{}
$script:ManagedDeviceCache   = @{}
$script:BuildGroupFailed     = $false
$script:GroupEligibleDevices = 0
$script:GroupLookupFailures  = 0
$script:GroupTargetDeviceIds = New-Object 'System.Collections.Generic.HashSet[string]'

function Write-CMTraceLog {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [int]$Type = 1,

        [Parameter(Mandatory = $false)]
        [string]$Component = 'Get-DiscoveredAppsAllDevices'
    )

    if (-not $script:EnableVerboseLogging) {
        return
    }

    $safeMessage = $Message -replace ']', ']]'
    $timeStamp = Get-Date
    $logLine = [string]::Format(
        '<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="{4}" type="{5}" thread="{6}" file="">',
        $safeMessage,
        $timeStamp.ToString('HH:mm:ss.fff'),
        $timeStamp.ToString('MM-dd-yyyy'),
        $Component,
        ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name),
        $Type,
        $PID
    )

    $logLine | Out-File -FilePath $logPath -Append -Encoding UTF8
}

function Write-ScriptMessage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Info', 'Warning', 'Error', 'Success', 'Verbose')]
        [string]$Level = 'Info',

        [Parameter(Mandatory = $false)]
        [string]$Component = 'Get-DiscoveredAppsAllDevices'
    )

    $logType = switch ($Level) {
        'Warning' { 2 }
        'Error' { 3 }
        default { 1 }
    }

    switch ($Level) {
        'Info'    { Write-Host $Message -ForegroundColor Cyan }
        'Warning' { Write-Host $Message -ForegroundColor Yellow }
        'Error'   { Write-Host $Message -ForegroundColor Red }
        'Success' { Write-Host $Message -ForegroundColor Green }
        'Verbose' { Write-Verbose $Message }
    }

    Write-CMTraceLog -Message $Message -Type $logType -Component $Component
}

function Get-CollectionCount {
    param(
        [Parameter(Mandatory = $false)]
        [object]$Value
    )

    if ($null -eq $Value) {
        return 0
    }

    if ($Value -is [System.Array]) {
        return $Value.Count
    }

    return 1
}

function Get-GroupNameToken {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    $token = $Value -replace '[*?]', ''
    $token = $token -replace '[^A-Za-z0-9]', ''

    return $token
}

function Get-PropertyValue {
    param(
        [Parameter(Mandatory = $true)]
        [object]$InputObject,

        [Parameter(Mandatory = $true)]
        [string[]]$PropertyNames
    )

    foreach ($propertyName in $PropertyNames) {
        if ($InputObject.PSObject.Properties.Name -contains $propertyName -and $InputObject.$propertyName) {
            return $InputObject.$propertyName
        }
    }

    return $null
}

function Get-ManagedDeviceAzureDeviceId {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ManagedDevice
    )

    $azureDeviceId = Get-PropertyValue -InputObject $ManagedDevice -PropertyNames @('AzureADDeviceId', 'AzureActiveDirectoryDeviceId')
    if ($azureDeviceId) {
        return $azureDeviceId
    }

    if ($script:ManagedDeviceCache.ContainsKey($ManagedDevice.Id)) {
        return $script:ManagedDeviceCache[$ManagedDevice.Id]
    }

    try {
        $expandedManagedDevice = Get-MgDeviceManagementManagedDevice -ManagedDeviceId $ManagedDevice.Id -ErrorAction Stop
        $azureDeviceId = Get-PropertyValue -InputObject $expandedManagedDevice -PropertyNames @('AzureADDeviceId', 'AzureActiveDirectoryDeviceId')
        $script:ManagedDeviceCache[$ManagedDevice.Id] = $azureDeviceId
        return $azureDeviceId
    }
    catch {
        Write-ScriptMessage -Message "Failed to expand managed device '$($ManagedDevice.DeviceName)' for group sync: $($_.Exception.Message)" -Level Warning -Component 'GroupSync'
        $script:ManagedDeviceCache[$ManagedDevice.Id] = $null
        return $null
    }
}

function Resolve-DirectoryDeviceObjectId {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ManagedDevice
    )

    $azureDeviceId = Get-ManagedDeviceAzureDeviceId -ManagedDevice $ManagedDevice
    if (-not $azureDeviceId) {
        return $null
    }

    if ($script:DirectoryDeviceCache.ContainsKey($azureDeviceId)) {
        return $script:DirectoryDeviceCache[$azureDeviceId]
    }

    try {
        $directoryDevice = Get-MgDevice -Filter "deviceId eq '$azureDeviceId'" -ErrorAction Stop | Select-Object -First 1
        if (-not $directoryDevice) {
            Write-ScriptMessage -Message "No Entra device object was found for Azure AD device ID '$azureDeviceId'." -Level Warning -Component 'GroupSync'
            $script:DirectoryDeviceCache[$azureDeviceId] = $null
            return $null
        }

        $script:DirectoryDeviceCache[$azureDeviceId] = $directoryDevice.Id
        return $directoryDevice.Id
    }
    catch {
        Write-ScriptMessage -Message "Failed to resolve Entra device object for '$($ManagedDevice.DeviceName)': $($_.Exception.Message)" -Level Warning -Component 'GroupSync'
        $script:DirectoryDeviceCache[$azureDeviceId] = $null
        return $null
    }
}

function Get-GroupMailNickname {
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupDisplayName
    )

    $nicknameBase = ($GroupDisplayName -replace '[^A-Za-z0-9]', '').ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($nicknameBase)) {
        $nicknameBase = 'appgroup'
    }

    if ($nicknameBase.Length -gt 40) {
        $nicknameBase = $nicknameBase.Substring(0, 40)
    }

    $sha1 = [System.Security.Cryptography.SHA1]::Create()
    try {
        $hashBytes = $sha1.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($GroupDisplayName))
    }
    finally {
        $sha1.Dispose()
    }

    $hashText = ([System.BitConverter]::ToString($hashBytes)).Replace('-', '').Substring(0, 8).ToLowerInvariant()
    return "$nicknameBase$hashText"
}

function Get-AppGroup {
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupDisplayName
    )

    $escapedDisplayName = $GroupDisplayName.Replace("'", "''")
    return Get-MgGroup -Filter "displayName eq '$escapedDisplayName'" -ErrorAction Stop | Select-Object -First 1
}

function New-AppGroup {
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupDisplayName,

        [Parameter(Mandatory = $true)]
        [string]$FilterDisplayName
    )

    $mailNickname = Get-GroupMailNickname -GroupDisplayName $GroupDisplayName
    $newGroupParams = @{
        DisplayName     = $GroupDisplayName
        Description     = "Managed by Get-DiscoveredAppsAllDevices for filter '$FilterDisplayName' by CDALTON"
        MailEnabled     = $false
        MailNickname    = $mailNickname
        SecurityEnabled = $true
        ErrorAction     = 'Stop'
    }

    return New-MgGroup @newGroupParams
}

function Sync-AppGroupMembership {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Group,

        [Parameter(Mandatory = $true)]
        [string[]]$TargetDeviceIds,

        [Parameter(Mandatory = $false)]
        [bool]$AllowRemoval = $true
    )

    $targetSet = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($targetDeviceId in $TargetDeviceIds) {
        if (-not [string]::IsNullOrWhiteSpace($targetDeviceId)) {
            [void]$targetSet.Add($targetDeviceId)
        }
    }

    $existingMembers = Get-MgGroupMember -GroupId $Group.Id -All -ErrorAction Stop
    $existingDeviceIds = @(
        $existingMembers |
            Where-Object {
                $_.AdditionalProperties -and
                $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.device'
            } |
            Select-Object -ExpandProperty Id
    )

    $nonDeviceMemberCount = @(
        $existingMembers |
            Where-Object {
                -not $_.AdditionalProperties -or
                $_.AdditionalProperties['@odata.type'] -ne '#microsoft.graph.device'
            }
    ).Count

    if ($nonDeviceMemberCount -gt 0) {
        Write-ScriptMessage -Message "Group '$($Group.DisplayName)' contains $nonDeviceMemberCount non-device member(s). They were left unchanged." -Level Warning -Component 'GroupSync'
    }

    $membersToAdd = @($targetSet | Where-Object { $_ -notin $existingDeviceIds })
    $membersToRemove = @()

    if ($AllowRemoval) {
        $membersToRemove = @($existingDeviceIds | Where-Object { $_ -notin $targetSet })
    }
    else {
        Write-ScriptMessage -Message "Device lookups were incomplete. Existing device members were not removed from '$($Group.DisplayName)'." -Level Warning -Component 'GroupSync'
    }

    foreach ($directoryDeviceId in $membersToAdd) {
        try {
            New-MgGroupMemberByRef -GroupId $Group.Id -BodyParameter @{ '@odata.id' = "https://graph.microsoft.com/v1.0/devices/$directoryDeviceId" } -ErrorAction Stop | Out-Null
            Write-ScriptMessage -Message "Added device '$directoryDeviceId' to group '$($Group.DisplayName)'." -Level Verbose -Component 'GroupSync'
        }
        catch {
            Write-ScriptMessage -Message "Failed to add device '$directoryDeviceId' to group '$($Group.DisplayName)': $($_.Exception.Message)" -Level Warning -Component 'GroupSync'
            $script:BuildGroupFailed = $true
        }
    }

    foreach ($directoryDeviceId in $membersToRemove) {
        try {
            Remove-MgGroupMemberByRef -GroupId $Group.Id -DirectoryObjectId $directoryDeviceId -ErrorAction Stop
            Write-ScriptMessage -Message "Removed device '$directoryDeviceId' from group '$($Group.DisplayName)'." -Level Verbose -Component 'GroupSync'
        }
        catch {
            Write-ScriptMessage -Message "Failed to remove device '$directoryDeviceId' from group '$($Group.DisplayName)': $($_.Exception.Message)" -Level Warning -Component 'GroupSync'
            $script:BuildGroupFailed = $true
        }
    }

    Write-ScriptMessage -Message "Group sync complete for '$($Group.DisplayName)'. Added: $($membersToAdd.Count). Removed: $($membersToRemove.Count). Target devices: $($targetSet.Count)." -Level Success -Component 'GroupSync'
}

if ($script:EnableVerboseLogging) {
    Remove-Item $logPath -ErrorAction SilentlyContinue
    Write-CMTraceLog -Message "Verbose logging enabled. Log file initialized at $logPath." -Component 'Logging'
    Write-ScriptMessage -Message "Verbose logging enabled. Writing CMTrace log to $logPath" -Level Info -Component 'Logging'
}

if ($BuildGroup -and -not $DisplayName) {
    Write-ScriptMessage -Message 'ERROR: -BuildGroup requires -DisplayName so the script can name and populate the target Entra group.' -Level Error
    exit 1
}

$groupDisplayName = $null
if ($BuildGroup) {
    $groupToken = Get-GroupNameToken -Value $DisplayName
    if ([string]::IsNullOrWhiteSpace($groupToken)) {
        Write-ScriptMessage -Message "ERROR: The -DisplayName value '$DisplayName' does not contain usable text for a group name after removing wildcards." -Level Error
        exit 1
    }

    $groupDisplayName = "App.$groupToken"
    Write-ScriptMessage -Message "BuildGroup enabled. Target group name: '$groupDisplayName'" -Level Verbose -Component 'GroupSync'
}

# ============================================================
# MODULE CHECKS
# ============================================================
$requiredModules = @('Microsoft.Graph', 'ImportExcel')
$missingModules = @()

Write-ScriptMessage -Message 'Checking for required modules...' -Level Info -Component 'ModuleCheck'

foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        $missingModules += $module
        Write-ScriptMessage -Message "  [MISSING] $module" -Level Error -Component 'ModuleCheck'
    }
    else {
        Write-ScriptMessage -Message "  [OK] $module" -Level Success -Component 'ModuleCheck'
    }
}

if ($missingModules.Count -gt 0) {
    Write-ScriptMessage -Message "`nERROR: Missing required modules!" -Level Error -Component 'ModuleCheck'
    Write-ScriptMessage -Message 'Please install the following modules before running this script:' -Level Warning -Component 'ModuleCheck'
    foreach ($module in $missingModules) {
        Write-Host "  Install-Module $module -Scope CurrentUser" -ForegroundColor White
        Write-CMTraceLog -Message "Suggested install command for missing module: Install-Module $module -Scope CurrentUser" -Type 2 -Component 'ModuleCheck'
    }
    Write-ScriptMessage -Message "`nScript execution halted." -Level Error -Component 'ModuleCheck'
    exit 1
}

Write-ScriptMessage -Message 'All required modules are installed.' -Level Success -Component 'ModuleCheck'
Write-Host ''

# ============================================================
# GRAPH AUTH
# ============================================================
$graphScopes = @(
    'DeviceManagementApps.Read.All',
    'DeviceManagementManagedDevices.Read.All'
)

if ($BuildGroup) {
    $graphScopes += @(
        'Group.ReadWrite.All',
        'Device.Read.All'
    )
}

Write-ScriptMessage -Message 'Connecting to Microsoft Graph...' -Level Info -Component 'Authentication'
Write-ScriptMessage -Message "Requested Graph scopes: $($graphScopes -join ', ')" -Level Verbose -Component 'Authentication'

try {
    Connect-MgGraph -Scopes $graphScopes -ErrorAction Stop

    $context = Get-MgContext
    if (-not $context) {
        throw 'Authentication verification failed - no context available'
    }

    Write-ScriptMessage -Message 'Successfully authenticated to Microsoft Graph' -Level Success -Component 'Authentication'
    Write-Host "  Account: $($context.Account)" -ForegroundColor Gray
    Write-Host "  Scopes: $($context.Scopes -join ', ')" -ForegroundColor Gray
    Write-Host ''
    Write-CMTraceLog -Message "Authenticated account: $($context.Account)" -Component 'Authentication'
    Write-CMTraceLog -Message "Granted scopes: $($context.Scopes -join ', ')" -Component 'Authentication'
}
catch {
    Write-ScriptMessage -Message "`nERROR: Failed to authenticate to Microsoft Graph!" -Level Error -Component 'Authentication'

    if ($_.Exception.Message -like '*window handle*') {
        Write-ScriptMessage -Message 'This error typically occurs when running in certain environments.' -Level Warning -Component 'Authentication'
        Write-ScriptMessage -Message 'Try one of these solutions:' -Level Warning -Component 'Authentication'
        Write-Host '  1. Run PowerShell as Administrator' -ForegroundColor White
        Write-Host '  2. Use a different authentication method:' -ForegroundColor White
        Write-Host '     Connect-MgGraph -UseDeviceCode' -ForegroundColor Cyan
        Write-Host '  3. Run from Windows PowerShell instead of PowerShell Core' -ForegroundColor White
    }
    elseif ($_.Exception.Message -like '*canceled*') {
        Write-ScriptMessage -Message 'Authentication was canceled by user.' -Level Warning -Component 'Authentication'
        Write-ScriptMessage -Message 'Please run the script again and complete the authentication process.' -Level Warning -Component 'Authentication'
    }
    else {
        Write-ScriptMessage -Message "Error details: $($_.Exception.Message)" -Level Warning -Component 'Authentication'
        Write-ScriptMessage -Message 'Troubleshooting steps:' -Level Warning -Component 'Authentication'
        Write-Host '  1. Ensure you have the required permissions (Intune Administrator or Global Administrator)' -ForegroundColor White
        Write-Host '  2. Try disconnecting first: Disconnect-MgGraph' -ForegroundColor White
        Write-Host '  3. Check your internet connection' -ForegroundColor White
        Write-Host '  4. Try using device code authentication: Connect-MgGraph -UseDeviceCode' -ForegroundColor White
    }

    Write-ScriptMessage -Message 'Script execution halted.' -Level Error -Component 'Authentication'
    exit 1
}

# ============================================================
# CSV HEADER (RAW DATA CONTRACT)
# ============================================================
"AppName,Version,InstallCount,DeviceName,DeviceId,OS,UserPrincipal,RetrievalError" |
    Out-File -FilePath $csvPath -Encoding UTF8

Write-ScriptMessage -Message 'Collecting discovered apps from Intune...' -Level Info -Component 'DataCollection'

# ============================================================
# DATA COLLECTION (STREAMED TO CSV)
# ============================================================
if (-not (Get-MgContext)) {
    Write-ScriptMessage -Message 'ERROR: Not authenticated to Microsoft Graph!' -Level Error -Component 'DataCollection'
    Write-ScriptMessage -Message 'Please run Connect-MgGraph and try again.' -Level Warning -Component 'DataCollection'
    Write-ScriptMessage -Message 'Script execution halted.' -Level Error -Component 'DataCollection'
    exit 1
}

try {
    $apps = Get-MgDeviceManagementDetectedApp -All -ErrorAction Stop
}
catch {
    Write-ScriptMessage -Message 'ERROR: Failed to retrieve discovered apps from Intune!' -Level Error -Component 'DataCollection'
    Write-ScriptMessage -Message "Error details: $($_.Exception.Message)" -Level Warning -Component 'DataCollection'
    Write-ScriptMessage -Message 'Possible causes:' -Level Warning -Component 'DataCollection'
    Write-Host '  - Insufficient permissions (requires DeviceManagementApps.Read.All)' -ForegroundColor White
    Write-Host '  - Network connectivity issues' -ForegroundColor White
    Write-Host '  - Authentication token expired' -ForegroundColor White
    Write-ScriptMessage -Message 'Script execution halted.' -Level Error -Component 'DataCollection'
    exit 1
}

if ($DisplayName) {
    Write-ScriptMessage -Message "Filtering apps by DisplayName: '$DisplayName'" -Level Warning -Component 'DataCollection'
    $apps = $apps | Where-Object { $_.DisplayName -like $DisplayName }
}

$appCount = Get-CollectionCount -Value $apps

if ($DisplayName) {
    Write-ScriptMessage -Message "Found $appCount matching app(s)" -Level Warning -Component 'DataCollection'
}
else {
    Write-ScriptMessage -Message "No filter applied - processing all $appCount app(s)" -Level Info -Component 'DataCollection'
}

foreach ($app in $apps) {
    Write-ScriptMessage -Message "Processing app: $($app.DisplayName)" -Level Warning -Component 'DataCollection'

    if ($app.DeviceCount -le 0) {
        "$($app.DisplayName),$($app.Version),0,,,,,No installs" |
            Out-File -FilePath $csvPath -Append -Encoding UTF8
        Write-ScriptMessage -Message "Skipping app '$($app.DisplayName)' because Intune reported no installs." -Level Verbose -Component 'DataCollection'
        continue
    }

    Start-Sleep -Milliseconds $baseDelayMs

    $attempt = 0
    $devices = $null

    do {
        try {
            $attempt++
            $devices = Get-MgDeviceManagementDetectedAppManagedDevice -DetectedAppId $app.Id -All -ErrorAction Stop
            break
        }
        catch {
            $retryAfter = $null

            if ($_.Exception -and
                $_.Exception.PSObject.Properties.Name -contains 'ResponseHeaders' -and
                $_.Exception.ResponseHeaders -and
                $_.Exception.ResponseHeaders['Retry-After']) {

                $retryAfter = [int]$_.Exception.ResponseHeaders['Retry-After']
            }

            if ($attempt -ge $maxRetries) {
                "$($app.DisplayName),$($app.Version),$($app.DeviceCount),,,,,Failed after retries" |
                    Out-File -FilePath $csvPath -Append -Encoding UTF8
                Write-ScriptMessage -Message "Failed to retrieve devices for '$($app.DisplayName)' after $attempt attempt(s)." -Level Warning -Component 'DataCollection'
                break
            }

            if ($retryAfter) {
                Write-ScriptMessage -Message "Throttled while retrieving '$($app.DisplayName)'. Waiting $retryAfter second(s) before retry $($attempt + 1)." -Level Verbose -Component 'DataCollection'
                Start-Sleep -Seconds $retryAfter
            }
            else {
                $retryDelaySeconds = [math]::Pow(2, $attempt)
                Write-ScriptMessage -Message "Retrying '$($app.DisplayName)' after $retryDelaySeconds second(s). Attempt $($attempt + 1) of $maxRetries." -Level Verbose -Component 'DataCollection'
                Start-Sleep -Seconds $retryDelaySeconds
            }
        }
    }
    while ($attempt -lt $maxRetries)

    if (-not $devices) {
        continue
    }

    foreach ($device in $devices) {
        "$($app.DisplayName),$($app.Version),$($app.DeviceCount),$($device.DeviceName),$($device.Id),$($device.OperatingSystem),$($device.UserPrincipalName)," |
            Out-File -FilePath $csvPath -Append -Encoding UTF8

        if ($BuildGroup) {
            $script:GroupEligibleDevices++
            $directoryDeviceId = Resolve-DirectoryDeviceObjectId -ManagedDevice $device

            if ($directoryDeviceId) {
                [void]$script:GroupTargetDeviceIds.Add($directoryDeviceId)
            }
            else {
                $script:GroupLookupFailures++
            }
        }
    }
}

Write-ScriptMessage -Message 'CSV export complete.' -Level Success -Component 'DataCollection'

if ($BuildGroup) {
    if ($appCount -eq 0) {
        Write-ScriptMessage -Message "BuildGroup requested, but no apps matched '$DisplayName'. Group sync was skipped." -Level Warning -Component 'GroupSync'
    }
    elseif ($script:GroupTargetDeviceIds.Count -eq 0 -and $script:GroupEligibleDevices -gt 0) {
        Write-ScriptMessage -Message "BuildGroup found $($script:GroupEligibleDevices) device record(s), but none could be resolved to Entra device objects. Group sync was skipped to avoid destructive changes." -Level Error -Component 'GroupSync'
        $script:BuildGroupFailed = $true
    }
    else {
        try {
            $group = Get-AppGroup -GroupDisplayName $groupDisplayName
            if ($group) {
                Write-ScriptMessage -Message "Using existing Entra group '$groupDisplayName'." -Level Success -Component 'GroupSync'
            }
            else {
                Write-ScriptMessage -Message "Creating Entra group '$groupDisplayName'." -Level Info -Component 'GroupSync'
                $group = New-AppGroup -GroupDisplayName $groupDisplayName -FilterDisplayName $DisplayName
                Write-ScriptMessage -Message "Created Entra group '$groupDisplayName'." -Level Success -Component 'GroupSync'
            }

            $allowRemoval = ($script:GroupLookupFailures -eq 0)
            Sync-AppGroupMembership -Group $group -TargetDeviceIds @($script:GroupTargetDeviceIds) -AllowRemoval:$allowRemoval
        }
        catch {
            Write-ScriptMessage -Message "Group sync failed for '$groupDisplayName': $($_.Exception.Message)" -Level Error -Component 'GroupSync'
            $script:BuildGroupFailed = $true
        }
    }
}

# ============================================================
# OUTPUT LOGIC
# ============================================================
if ($CSV) {
    Write-ScriptMessage -Message "CSV output complete: $csvPath" -Level Success -Component 'Output'
    Write-ScriptMessage -Message 'Use the -CSV switch to generate CSV output only.' -Level Verbose -Component 'Output'

    if ($script:BuildGroupFailed) {
        exit 1
    }

    exit 0
}

Write-ScriptMessage -Message 'Generating Excel report...' -Level Info -Component 'Output'

$data = Import-Csv $csvPath

# ------------------------------------------------------------
# OVERVIEW TAB
# ------------------------------------------------------------
$overview = @(
    [PSCustomObject]@{ Metric = 'Total Discovered Apps'; Value = ($data | Select-Object AppName -Unique).Count }
    [PSCustomObject]@{ Metric = 'Total Install Records'; Value = ($data | Where-Object DeviceName).Count }
    [PSCustomObject]@{ Metric = 'Apps With Errors'; Value = ($data | Where-Object RetrievalError).Count }
)

# ------------------------------------------------------------
# APPLICATIONS SUMMARY TAB
# ------------------------------------------------------------
$appSummary = $data |
    Where-Object { $_.RetrievalError -eq '' } |
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
    -WorksheetName 'Overview' `
    -TableName 'Overview' `
    -TableStyle Medium2 `
    -BoldTopRow -AutoSize

$appSummary | Export-Excel $xlsxPath `
    -WorksheetName 'Applications' `
    -TableName 'Applications' `
    -TableStyle Medium2 `
    -BoldTopRow -AutoSize `
    -FreezeTopRow

$installations | Export-Excel $xlsxPath `
    -WorksheetName 'Installations' `
    -TableName 'Installations' `
    -TableStyle Medium2 `
    -BoldTopRow -AutoSize `
    -FreezeTopRow

$errors | Export-Excel $xlsxPath `
    -WorksheetName 'Errors' `
    -TableName 'Errors' `
    -TableStyle Medium2 `
    -BoldTopRow -AutoSize

# ------------------------------------------------------------
# FINAL FORMATTING (WRAP DEVICES COLUMN)
# ------------------------------------------------------------
$pkg = Open-ExcelPackage -Path $xlsxPath
$ws = $pkg.Workbook.Worksheets['Installations']
$devicesColumn = 3

for ($row = 2; $row -le $ws.Dimension.End.Row; $row++) {
    $cell = $ws.Cells[$row, $devicesColumn]
    $cell.Style.WrapText = $true
    $cell.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Top

    $lineCount = ([regex]::Matches($cell.Text, "`n")).Count + 1
    $minHeight = $lineCount * 15

    if ($ws.Row($row).Height -lt $minHeight) {
        $ws.Row($row).Height = $minHeight
    }
}

$ws.Column($devicesColumn).Width = 80

Close-ExcelPackage $pkg
Remove-Item $csvPath -ErrorAction SilentlyContinue

Write-ScriptMessage -Message "Excel report created: $xlsxPath" -Level Success -Component 'Output'
Write-ScriptMessage -Message 'Intermediate CSV file has been removed. Use -CSV switch if you need CSV output.' -Level Verbose -Component 'Output'

if ($script:BuildGroupFailed) {
    exit 1
}
