# PowerShell Profile

A comprehensive PowerShell profile script that automatically configures your PowerShell environment with essential modules, applications, and customizations.

## Overview

This PowerShell profile is designed to streamline your PowerShell experience by automatically:

- Installing and managing essential PowerShell modules
- Checking for WinGet application updates
- Configuring Oh My Posh for an enhanced prompt
- Setting up PSReadLine with intelligent predictions
- Providing comprehensive logging and error handling
- Running background update checks for optimal performance

## Features

### 🚀 Automatic Module Management

- **ExchangeOnlineManagement**: For Microsoft 365 Exchange management
- **Microsoft.Graph**: Microsoft Graph PowerShell SDK
- **PSReadLine**: Enhanced command-line editing experience
- **MATC.TS.Exchange**: Custom Exchange module (from MATC.TS repository)

### 📦 WinGet Application Monitoring

Monitors and notifies about updates for:

- Microsoft Windows Terminal
- Git
- Microsoft PowerShell
- Visual Studio Code
- Oh My Posh

### 🎨 Visual Enhancements

- **Oh My Posh**: Custom prompt configuration with themes
- **PSReadLine**: Intelligent command prediction and history
- **Progress Indicators**: Visual feedback during initialization

### 📊 Logging & Monitoring

- Comprehensive logging to `~/.config/profile.log`
- Automatic log rotation (maintains last 1000 entries)
- Network connectivity testing
- Background update checking with timeout protection

### ⚡ Performance Optimizations

- Asynchronous update checking to prevent blocking
- Daily update check throttling
- Timeout protection for network operations
- Background job management for heavy operations

## Usage

### Installation

1. Copy `Microsoft.PowerShell_profile.ps1` to your PowerShell profile directory
2. The profile will automatically run when you start a new PowerShell session

### Configuration

The profile uses a centralized configuration function `Get-ProfileConfiguration()` that you can modify to customize:

- Module list and repositories
- Oh My Posh theme URL
- Global variables
- WinGet applications to monitor

### Commands

The profile adds several utility functions:

#### Logging Functions

```powershell
# View recent profile logs
Get-ProfileLog -Last 50

# View only error logs
Get-ProfileLog -Level ERROR
```

#### Update Functions

```powershell
# Manually check for WinGet updates
Search-WingetUpdate -Apps $config.WinGetApps

# Check for PowerShell module updates
Search-ModuleUpdate -Modules $config.PowerShellModules
```

## Requirements

- **PowerShell 5.0 or higher**
- **Windows operating system**
- **Internet connection** (for module installation and updates)
- **Administrator privileges** (may be required for some module installations)

## Global Variables

The profile automatically sets these global variables:

- `$Global:TenantUrl`: SharePoint admin URL
- `$Global:instance`: Database instance
- `$Global:database`: Database name

## Troubleshooting

### Log Files

Profile logs are stored at: `~/.config/profile.log`

Use `Get-ProfileLog` to view recent activity and errors.

### Common Issues

1. **Module Installation Failures**: Ensure you have internet connectivity and appropriate permissions
2. **Oh My Posh Not Loading**: Check network connectivity to the theme configuration URL
3. **Slow Startup**: Background update checks run asynchronously to minimize impact

### Network Connectivity

The profile includes network connectivity testing and will gracefully handle offline scenarios.

## Customization

### Adding New Modules

Modify the `PowerShellModules` array in `Get-ProfileConfiguration()`:

```powershell
PowerShellModules = @(
    @{Name = "YourModule"; PSRepository = "PSGallery" }
    # Add more modules here
)
```

### Changing Oh My Posh Theme

Update the `OhMyPoshConfig` URL in `Get-ProfileConfiguration()`:

```powershell
OhMyPoshConfig = "https://your-theme-url.json"
```

## License

This project is open source and available under standard usage terms.

## Contributing

Feel free to submit issues, fork the repository, and create pull requests for any improvements.
