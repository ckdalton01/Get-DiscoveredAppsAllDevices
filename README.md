# Intune Software Inventory Report

A PowerShell script that generates a comprehensive Excel report of all software applications discovered across your Intune-managed devices.

## Overview

This script collects discovered applications from Microsoft Intune and generates a detailed Excel workbook with multiple worksheets showing application summaries, installation details, and error tracking.

## Prerequisites

### Required Modules

1. **Microsoft.Graph PowerShell SDK**
   ```powershell
   Install-Module Microsoft.Graph -Scope CurrentUser
   ```

2. **ImportExcel Module**
   ```powershell
   Install-Module ImportExcel -Scope CurrentUser
   ```

### Required Permissions

The script requires the following Microsoft Graph API permissions:

- `DeviceManagementApps.Read.All`
- `DeviceManagementManagedDevices.Read.All`

### Azure AD Requirements

- An Azure AD account with sufficient permissions to consent to the above Graph API scopes
- Typically requires one of these roles:
  - Global Administrator
  - Intune Administrator
  - Cloud Device Administrator

## Installation

1. Clone or download this script to your local machine
2. Install the required PowerShell modules (see Prerequisites)
3. Ensure you have appropriate permissions in your Azure AD tenant

## Usage

### Running the Script

**Basic usage (all apps):**
```powershell
.\Get-DiscoveredAppsAllDevices.ps1
```

**Filter by application name:**
```powershell
# Exact match
.\Get-DiscoveredAppsAllDevices.ps1 -DisplayName "Google Chrome"

# Wildcard patterns
.\Get-DiscoveredAppsAllDevices.ps1 -DisplayName "Java*"           # Apps starting with "Java"
.\Get-DiscoveredAppsAllDevices.ps1 -DisplayName "*Office*"        # Apps containing "Office"
.\Get-DiscoveredAppsAllDevices.ps1 -DisplayName "*Adobe Reader*"  # Apps containing "Adobe Reader"
```

### Parameters

**`-DisplayName` (Optional)**
- Filters the discovered apps by display name
- Supports PowerShell wildcards (`*` for any characters, `?` for single character)
- Case-insensitive matching
- Examples:
  - `"Google Chrome"` - Exact match
  - `"Java*"` - All apps starting with "Java"
  - `"*Java*"` - All apps containing "Java"
  - `"*"` - All apps (same as omitting the parameter)

**`-CSV` (Optional Switch)**
- Outputs data in CSV format only, skipping Excel generation
- By default (without this switch), only the Excel file is created
- Examples:
  ```powershell
  # Default: Create Excel file only
  .\Get-DiscoveredAppsAllDevices.ps1 -DisplayName "Chrome*"
  
  # Create CSV file only (no Excel)
  .\Get-DiscoveredAppsAllDevices.ps1 -DisplayName "Chrome*" -CSV
  ```

### What Happens

1. **Authentication**: The script will prompt you to sign in to Microsoft Graph
2. **Data Collection**: Retrieves discovered apps from Intune and their associated devices
3. **Output Generation**: 
   - **Default (no `-CSV` switch)**: Creates formatted Excel workbook, removes intermediate CSV
   - **With `-CSV` switch**: Creates CSV file only, skips Excel generation

### Output Files

**Default behavior (Excel only):**
- `Intune_Software_Inventory.xlsx` - Formatted Excel report with 4 worksheets

**With `-CSV` switch:**
- `Intune_DiscoveredApps_WithDevices.csv` - Raw data export

Files are created in the current directory where the script is executed.

## Excel Report Structure

The generated Excel workbook contains four worksheets:

### 1. Overview
High-level metrics including:
- Total discovered applications
- Total installation records
- Applications with retrieval errors

### 2. Applications
Summary of all applications with:
- Application name
- Version
- Installation count

### 3. Installations
Detailed view showing:
- Application name and version
- Installation count
- List of all devices with the application installed (with Device IDs)
- Wrapped text formatting for easy reading

### 4. Errors
Lists any applications that encountered errors during retrieval:
- Application name
- Version
- Error description

## Features

- **Flexible Output Format**: Choose between Excel (default) or CSV output using the `-CSV` switch
- **Application Filtering**: Filter apps by name with wildcard support using the `-DisplayName` parameter
- **Retry Logic**: Automatically retries failed API calls up to 5 times with exponential backoff
- **Rate Limiting**: Respects Microsoft Graph API rate limits with configurable delays
- **Error Tracking**: Captures and reports applications that couldn't be fully processed
- **Formatted Output**: Professional Excel formatting with:
  - Frozen header rows
  - Auto-sized columns
  - Wrapped text for multi-line content
  - Styled tables
  - Automatic row height adjustment

## Configuration

You can modify these variables at the top of the script:

```powershell
$basePath   = "."              # Output directory
$maxRetries  = 5               # Maximum API retry attempts
$baseDelayMs = 300             # Base delay between API calls (milliseconds)
```

## Troubleshooting

### Authentication Issues

**Error: "InteractiveBrowserCredential authentication failed: A window handle must be configured"**

This error occurs when running in certain PowerShell environments. Try these solutions:
1. Run PowerShell as Administrator
2. Use device code authentication:
   ```powershell
   # Disconnect first if needed
   Disconnect-MgGraph
   
   # Connect with device code
   Connect-MgGraph -Scopes "DeviceManagementApps.Read.All","DeviceManagementManagedDevices.Read.All" -UseDeviceCode
   
   # Then run the script
   .\Get-DiscoveredAppsAllDevices.ps1
   ```
3. Run from Windows PowerShell instead of PowerShell Core/7

**Error: "User canceled authentication"**

Simply re-run the script and complete the authentication process in the browser window.

**Error: "Authentication needed. Please call Connect-MgGraph"**

The authentication failed or was not completed. The script will automatically detect this and provide guidance.

**General authentication troubleshooting:**
- Ensure you have the required permissions in Azure AD (Intune Administrator or Global Administrator)
- Try disconnecting first: `Disconnect-MgGraph`
- Verify your account has access to Intune managed devices
- Check your internet connection

### Module Not Found Errors
```powershell
# Check installed modules
Get-Module -ListAvailable Microsoft.Graph*
Get-Module -ListAvailable ImportExcel

# Reinstall if needed
Install-Module Microsoft.Graph -Force
Install-Module ImportExcel -Force
```

### API Rate Limiting
If you encounter persistent rate limiting:
- Increase `$baseDelayMs` to 500 or 1000
- The script already implements retry logic with exponential backoff

### Excel Formatting Issues
If the Devices column doesn't wrap properly:
- Ensure you're using the latest version of the ImportExcel module
- Try opening and saving the file in Excel to refresh formatting

## Performance Considerations

- **Large Environments**: For tenants with thousands of applications and devices, the script may take 30+ minutes to complete
- **API Throttling**: The script includes delays and retry logic to avoid overwhelming the Graph API
- **Memory Usage**: Large datasets are streamed to CSV to minimize memory footprint

## License

This script is provided as-is for use in your organization.

## Support

For issues or questions:
- Check the Errors worksheet in the generated Excel file
- Review the console output for any warning messages
- Ensure all prerequisites are met and modules are up to date

## Version History

- **v1.1** - Authentication and filtering enhancements
  - Added `-DisplayName` parameter with wildcard support for filtering apps
  - Added `-CSV` switch to control output format (Excel or CSV)
  - Implemented robust authentication error handling with specific guidance
  - Added authentication verification before API calls
  - Improved error messages for common authentication issues
  - Default behavior now creates Excel only and removes intermediate CSV file
- **v1.0** - Initial release
  - Basic app discovery and Excel reporting
  - Retry logic and error handling
  - Multi-worksheet Excel output
