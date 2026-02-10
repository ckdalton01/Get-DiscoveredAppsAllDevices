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

```powershell
.\Get-DiscoveredAppsAllDevices.ps1
```

### What Happens

1. **Authentication**: The script will prompt you to sign in to Microsoft Graph
2. **Data Collection**: Retrieves all discovered apps from Intune and their associated devices
3. **CSV Export**: Creates an intermediate CSV file with raw data
4. **Excel Generation**: Produces a formatted Excel workbook with multiple worksheets

### Output Files

The script generates two files in the current directory:

- `Intune_DiscoveredApps_WithDevices.csv` - Raw data export
- `Intune_Software_Inventory.xlsx` - Formatted Excel report

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
- Ensure you have the required permissions in Azure AD
- Try running `Disconnect-MgGraph` then re-running the script
- Verify your account has access to Intune managed devices

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

- **v1.0** - Initial release
  - Basic app discovery and Excel reporting
  - Retry logic and error handling
  - Multi-worksheet Excel output
