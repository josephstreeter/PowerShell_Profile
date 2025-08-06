<#
.SYNOPSIS
PowerShell profile script for initializing environment, installing modules, and configuring Oh My Posh.

.DESCRIPTION
This script is executed whenever a new PowerShell session is started. It sets up the environment by loading necessary modules,
configuring the prompt, and applying user preferences.
#>

[CmdletBinding()]
param()

# Check minimum PowerShell version
if ($PSVersionTable.PSVersion.Major -lt 5)
{
    Write-Warning "This profile requires PowerShell 5.0 or higher. Current version: $($PSVersionTable.PSVersion)"
    return
}

function Write-ProfileLog()
{
    [CmdletBinding()]
    [OutputType([void])]
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [Parameter(Mandatory = $false)][ValidateSet("INFO", "WARNING", "ERROR", "DEBUG")][string]$Level = "INFO"
    )
        
    try
    {
        $logPath = Join-Path $env:USERPROFILE ".config"
        $logFile = Join-Path $logPath "profile.log"
                
        if (-not (Test-Path $logPath))
        {
            New-Item $logPath -ItemType Directory -Force | Out-Null
        }
                
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "[$timestamp] [$Level] $Message"
                
        # Keep log file size manageable (last 1000 lines)
        if (Test-Path $logFile)
        {
            $existingLines = Get-Content $logFile -ErrorAction SilentlyContinue
            if ($existingLines.Count -gt 1000)
            {
                $existingLines | Select-Object -Last 900 | Set-Content $logFile
            }
        }
                
        Add-Content $logFile $logEntry -ErrorAction SilentlyContinue
    }
    catch
    {
        # Log the error to Windows Event Log instead of failing silently
        Write-EventLog -LogName "Application" -Source "PowerShell" -EventId 1001 -EntryType Warning -Message "Profile logging failed: $($_.Exception.Message)" -ErrorAction SilentlyContinue
    }
}

function Get-ProfileLog()
{
    [CmdletBinding()]
    [OutputType([void])]
    param(
        [Parameter(Mandatory = $false)][int]$Last = 50,
        [Parameter(Mandatory = $false)][ValidateSet("INFO", "WARNING", "ERROR", "DEBUG")][string]$Level
    )
        
    try
    {
        $logFile = Join-Path $env:USERPROFILE ".config\profile.log"
        if (-not (Test-Path $logFile))
        {
            Write-Information "No profile log found at $logFile" -InformationAction Continue
            return
        }
                
        $logs = Get-Content $logFile | Select-Object -Last $Last
                
        if ($Level)
        {
            $logs = $logs | Where-Object { $_ -match "\[$Level\]" }
        }
                
        $logs | ForEach-Object {
            Write-Information $_ -InformationAction Continue
        }
    }
    catch
    {
        Write-Error "Failed to read profile log: $($_.Exception.Message)"
    }
}

function Test-NetworkConnectivity()
{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $false)][int]$TimeoutSeconds = 10
    )
        
    try
    {
        $request = [System.Net.WebRequest]::Create($Uri)
        $request.Timeout = $TimeoutSeconds * 1000
        $request.Method = "HEAD"
        $response = $request.GetResponse()
        $response.Close()
        return $true
    }
    catch
    {
        Write-ProfileLog "Network connectivity test failed for $Uri : $($_.Exception.Message)" "WARNING"
        return $false
    }
}

function Get-ProfileConfiguration()
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param()
        
    return @{
        WinGetApps        = @(
            "Microsoft.WindowsTerminal",
            "Git.Git",
            "Microsoft.PowerShell",
            "Microsoft.VisualStudioCode",
            "JanDeDobbeleer.OhMyPosh"
        )
        PowerShellModules = @(
            @{Name = "ExchangeOnlineManagement"; PSRepository = "PSGallery" }
            @{Name = "Microsoft.Graph"; PSRepository = "PSGallery" }
            @{Name = "PSReadLine"; PSRepository = "PSGallery" }
            @{Name = "MATC.TS.Exchange"; PSRepository = "MATC.TS" }
        )
        OhMyPoshConfig    = "https://raw.githubusercontent.com/josephstreeter/Oh-My-Posh/refs/heads/main/jstreeter.omp.json"
        GlobalVariables   = @{
            TenantUrl = "https://madisoncollege365-admin.sharepoint.com/"
            Instance  = "IDMDBPRD01\MIMStage"
            Database  = "StagingDirectory"
        }
    }
}

function Compare-UpdateLastRun()
{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param()
        
    try
    {
        $configPath = Join-Path $env:USERPROFILE ".config"
        $lastRunFile = Join-Path $configPath "lastrun.txt"
                
        if (-not (Test-Path $configPath))
        {
            New-Item $configPath -ItemType Directory -Force | Out-Null
        }

        if (-not (Test-Path $lastRunFile))
        {
            (Get-Date).AddDays(-1).ToString("yyyy-MM-dd") | Set-Content $lastRunFile -Force
        }

        $lastRun = Get-Content $lastRunFile -ErrorAction Stop
        $today = (Get-Date).ToString("yyyy-MM-dd")
                
        if ($lastRun -eq $today)
        {
            Write-Verbose "Skipping update check, last run was $lastRun"
            Write-ProfileLog "Skipping update check, last run was $lastRun" "INFO"
            return $true
        }
        else
        {
            Write-Verbose "Checking for updates, last run was $lastRun"
            Write-ProfileLog "Running update check, last run was $lastRun" "INFO"
            $today | Set-Content $lastRunFile
            return $false
        }
    }
    catch
    {
        Write-Warning "Failed to check last run date: $($_.Exception.Message)"
        Write-ProfileLog "Failed to check last run date: $($_.Exception.Message)" "ERROR"
        return $false  # Default to checking updates on error
    }
}

function Search-WingetUpdate()
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)][array]$Apps
    )
    try
    {
        Write-ProfileLog "Starting WinGet update check" "INFO"
        Write-Progress -Activity "Checking WinGet Updates" -Status "Initializing..." -PercentComplete 0
                
        if (-not (Get-Module microsoft.winget.client -ListAvailable))
        {
            Write-Output "Installing WinGet PowerShell module...."
            Write-Progress -Activity "Checking WinGet Updates" -Status "Installing WinGet module..." -PercentComplete 25
            Install-Module microsoft.winget.client -Repository PSGallery -Force -Scope CurrentUser
        }

        Write-Verbose "Checking for WinGet updates...."
        Write-Progress -Activity "Checking WinGet Updates" -Status "Scanning for updates..." -PercentComplete 50

        $Updates = Get-WinGetPackage | Where-Object { $Apps -contains $_.id -and $_.IsUpdateAvailable }
                
        Write-Progress -Activity "Checking WinGet Updates" -Status "Processing results..." -PercentComplete 75

        if ($Updates.count -eq 0)
        {
            Write-Output "No WinGet updates available"
            Write-ProfileLog "No WinGet updates available" "INFO"
            Write-Progress -Activity "Checking WinGet Updates" -Completed
            return
        }

        $updateList = ($Updates | ForEach-Object { $_.id }) -join ", "
        Write-Warning "Updates available for: $updateList`nRun 'winget upgrade --all' or individual updates with 'winget upgrade <package-id>'"
        Write-ProfileLog "WinGet updates available: $updateList" "INFO"
        Write-Progress -Activity "Checking WinGet Updates" -Completed
        return
    }
    catch
    {
        Write-Error -Message ("Failed to check for updates: {0}" -f $_.Exception.Message)
        Write-ProfileLog "Failed to check WinGet updates: $($_.Exception.Message)" "ERROR"
        Write-Progress -Activity "Checking WinGet Updates" -Completed
        return
    }
}

function Install-CommonModule()
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)][array]$Modules
    )
        
    Write-Verbose "Installing common PowerShell modules...."
    Write-ProfileLog "Starting module installation check" "INFO"
        
    $totalModules = $Modules.Count
    $currentModule = 0
        
    foreach ($Module in $Modules)
    {
        $currentModule++
        $percentComplete = [math]::Round(($currentModule / $totalModules) * 100)
        Write-Progress -Activity "Installing PowerShell Modules" -Status "Processing $($Module.Name)..." -PercentComplete $percentComplete
                
        try
        {
            if (-not (Get-Module $Module.Name -ListAvailable))
            {
                Write-Output "Installing module: $($Module.Name)"
                Write-ProfileLog "Installing module: $($Module.Name)" "INFO"
                Install-Module $Module.Name -Repository $Module.PSRepository -Force -Scope CurrentUser
            }
            else
            {
                Write-Verbose "Module $($Module.Name) already installed"
                Write-ProfileLog "Module $($Module.Name) already installed" "DEBUG"
            }
        }
        catch
        {
            Write-Warning "Failed to install module $($Module.Name): $($_.Exception.Message)"
            Write-ProfileLog "Failed to install module $($Module.Name): $($_.Exception.Message)" "ERROR"
        }
    }
        
    Write-Progress -Activity "Installing PowerShell Modules" -Completed
}

function Search-ModuleUpdate()
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)][array]$Modules
    )

    try
    {
        Write-Verbose "Checking for PowerShell module updates...."

        $Updates = @()
        foreach ($Module in $Modules)
        {
            # Check if module is installed first
            $installedModule = Get-InstalledModule -Name $Module.Name -ErrorAction SilentlyContinue
            if (-not $installedModule)
            {
                Write-Warning "Module $($Module.Name) is not installed"
                continue
            }
                        
            $availableModule = Find-Module -Name $Module.Name -Repository $Module.PSRepository -ErrorAction SilentlyContinue
            if (-not $availableModule)
            {
                Write-Warning "Module $($Module.Name) not found in repository $($Module.PSRepository)"
                continue
            }
                        
            if ($installedModule.Version -ne $availableModule.Version)
            {
                $Updates += $Module
            }
        }

        if ($Updates.count -eq 0)
        {
            Write-Output "All modules are up to date"
            return
        }

        $moduleNames = ($Updates | ForEach-Object { $_.Name }) -join ", "
        Write-Warning "Updates available for: $moduleNames`nRun 'Update-Module <ModuleName>' to install updates"
    }
    catch
    {
        Write-Error -Message ("Failed to check for updates: {0}" -f $_.Exception.Message)
    }
}

function Import-LocalModule()
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)][string]$ModulePath
    )
    try
    {
        if (Test-Path $ModulePath)
        {
            Write-Output ("Importing module from {0}" -f $ModulePath)
            Import-Module $ModulePath -ErrorAction Stop
        }
        else
        {
            Write-Warning -Message ("Module not found at {0}" -f $ModulePath)
        }
    }
    catch
    {
        Write-Error -Message ("Failed to import module: {0}" -f $_.Exception.Message)
    }
}

function Set-CommonVariable()
{
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $false)][hashtable]$Config = (Get-ProfileConfiguration)
    )
    try
    {
        if ($PSCmdlet.ShouldProcess("Global Variables", "Set common variables"))
        {
            Write-Information "Setting common variables...." -InformationAction Continue
            $Global:TenantUrl = $Config.GlobalVariables.TenantUrl
            $Global:instance = $Config.GlobalVariables.Instance
            $Global:database = $Config.GlobalVariables.Database
        }
    }
    catch
    {
        Write-Error -Message ("Failed to set common variables: {0}" -f $_.Exception.Message)
    }
}

function Initialize-OhMyPosh()
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)][string]$ConfigFile,
        [Parameter(Mandatory = $false)][int]$TimeoutSeconds = 30
    )

    try
    {
        if (-not(Get-WinGetPackage JanDeDobbeleer.OhMyPosh -ErrorAction SilentlyContinue))
        {
            Write-Output "Installing Oh My Posh...."
            Write-ProfileLog "Installing Oh My Posh" "INFO"
            Install-WinGetPackage -Id JanDeDobbeleer.OhMyPosh -Source winget -Force
        }
    }
    catch
    {
        Write-Warning "Failed to install Oh My Posh: $($_.Exception.Message)"
        Write-ProfileLog "Failed to install Oh My Posh: $($_.Exception.Message)" "ERROR"
        return
    }

    try
    {
        # Test network connectivity to config file with timeout
        if ($ConfigFile -match "^https?://")
        {
            Write-Verbose "Testing connectivity to Oh My Posh config file..."
            if (-not (Test-NetworkConnectivity -Uri $ConfigFile -TimeoutSeconds $TimeoutSeconds))
            {
                Write-Warning "Cannot reach Oh My Posh config file within $TimeoutSeconds seconds. Skipping initialization."
                Write-ProfileLog "Oh My Posh config file unreachable: $ConfigFile" "WARNING"
                return
            }
        }

        Write-Output "Initializing Oh My Posh...."
        Write-ProfileLog "Initializing Oh My Posh with config: $ConfigFile" "INFO"
                
        # Use a job with timeout for initialization
        $job = Start-Job -ScriptBlock {
            param($configFileParam)
            oh-my-posh --init --shell pwsh --config $configFileParam
        } -ArgumentList $ConfigFile
                
        if (Wait-Job $job -Timeout $TimeoutSeconds)
        {
            $result = Receive-Job $job
            # Instead of using Invoke-Expression, create a script block
            $scriptBlock = [ScriptBlock]::Create($result)
            & $scriptBlock
            Write-ProfileLog "Oh My Posh initialized successfully" "INFO"
        }
        else
        {
            Write-Warning "Oh My Posh initialization timed out after $TimeoutSeconds seconds"
            Write-ProfileLog "Oh My Posh initialization timeout" "WARNING"
            Stop-Job $job
        }
                
        Remove-Job $job -Force -ErrorAction SilentlyContinue
    }
    catch
    {
        Write-Warning "Failed to initialize oh-my-posh: $($_.Exception.Message)"
        Write-ProfileLog "Failed to initialize Oh My Posh: $($_.Exception.Message)" "ERROR"
    }
}

function Start-AsyncUpdateCheck()
{
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)][hashtable]$Config
    )
        
    if ($PSCmdlet.ShouldProcess("Background Update Check", "Start async update check"))
    {
        # Pass the configuration values directly as parameters to avoid Using scope issues
        $job = Start-Job -ScriptBlock {
            param($wingetApps, $powershellModules)
                    
            # Import the functions we need in the job
            function Search-WingetUpdate()
            {
                param([array]$Apps)
                try
                {
                    if (-not (Get-Module microsoft.winget.client -ListAvailable))
                    {
                        Install-Module microsoft.winget.client -Repository PSGallery -Force -Scope CurrentUser
                    }
                    $Updates = Get-WinGetPackage | Where-Object { $Apps -contains $_.id -and $_.IsUpdateAvailable }
                    if ($Updates.count -gt 0)
                    {
                        $updateList = ($Updates | ForEach-Object { $_.id }) -join ", "
                        return "WinGet updates available for: $updateList"
                    }
                    return "No WinGet updates available"
                }
                catch
                {
                    return "Failed to check WinGet updates: $($_.Exception.Message)"
                }
            }
                    
            function Search-ModuleUpdate()
            {
                param([array]$Modules)
                try
                {
                    $Updates = @()
                    foreach ($Module in $Modules)
                    {
                        $installedModule = Get-InstalledModule -Name $Module.Name -ErrorAction SilentlyContinue
                        if ($installedModule)
                        {
                            $availableModule = Find-Module -Name $Module.Name -Repository $Module.PSRepository -ErrorAction SilentlyContinue
                            if ($availableModule -and $installedModule.Version -ne $availableModule.Version)
                            {
                                $Updates += $Module
                            }
                        }
                    }
                    if ($Updates.count -gt 0)
                    {
                        $moduleNames = ($Updates | ForEach-Object { $_.Name }) -join ", "
                        return "Module updates available for: $moduleNames"
                    }
                    return "All modules are up to date"
                }
                catch
                {
                    return "Failed to check module updates: $($_.Exception.Message)"
                }
            }
                    
            $wingetResult = Search-WingetUpdate -Apps $wingetApps
            $moduleResult = Search-ModuleUpdate -Modules $powershellModules
                    
            return @{
                WinGetResult = $wingetResult
                ModuleResult = $moduleResult
            }
        } -ArgumentList $Config.WinGetApps, $Config.PowerShellModules
            
        Write-ProfileLog "Started async update check job" "INFO"
        return $job
    }
}

function Initialize-PSReadLine()
{
    [CmdletBinding()]
    param()

    try
    {
        if (-not (Get-Module PSReadLine -ListAvailable))
        {
            Write-Output "Installing PSReadLine...."
            Install-Module PSReadLine -Repository PSGallery -Force -Scope CurrentUser
        }

        Write-Output "Initializing PSReadLine...."

        $Params = @{
            BellStyle           = "None"
            PredictionSource    = "History"
            PredictionViewStyle = "ListView"
        }

        Set-PSReadLineOption @Params -ErrorAction Stop
    }
    catch
    {
        Write-Error -Message ("Failed to initialize PSReadLine: {0}" -f $_.Exception.Message)
    }
}

# End Functions ############################################################################################################

try
{
    Write-ProfileLog "Starting PowerShell profile initialization" "INFO"
    $config = Get-ProfileConfiguration
    $SkipUpdates = Compare-UpdateLastRun
        
    # Start async update check if not skipping
    $updateJob = $null
    if ($SkipUpdates -eq $false)
    {
        Write-Verbose "Starting background update check..."
        $updateJob = Start-AsyncUpdateCheck -Config $config -Confirm:$false
                
        # Install modules synchronously as they may be needed immediately
        Install-CommonModule -Modules $config.PowerShellModules
    }

    #Import-LocalModule "C:\Users\JStreeter\source\repos\IAM_PowerShell_Modules\Exchange\PS-EX-Administration.psm1"

    # Initialize PowerShell Environment
    Write-Progress -Activity "Initializing Profile" -Status "Setting variables..." -PercentComplete 20
    Set-CommonVariable -Config $config -Confirm:$false
        
    Write-Progress -Activity "Initializing Profile" -Status "Configuring Oh My Posh..." -PercentComplete 50
    Initialize-OhMyPosh -ConfigFile $config.OhMyPoshConfig -TimeoutSeconds 15
        
    Write-Progress -Activity "Initializing Profile" -Status "Configuring PSReadLine..." -PercentComplete 80
    Initialize-PSReadLine
        
    Write-Progress -Activity "Initializing Profile" -Status "Completing..." -PercentComplete 100
        
    # Check async update results if job was started
    if ($updateJob)
    {
        if (Wait-Job $updateJob -Timeout 5)
        {
            $updateResults = Receive-Job $updateJob
            if ($updateResults.WinGetResult -notmatch "No.*updates")
            {
                Write-Information $updateResults.WinGetResult -InformationAction Continue
            }
            if ($updateResults.ModuleResult -notmatch "up to date")
            {
                Write-Information $updateResults.ModuleResult -InformationAction Continue
            }
            Write-ProfileLog "Async update check completed" "INFO"
        }
        else
        {
            Write-Verbose "Update check still running in background..."
            Write-ProfileLog "Update check job still running" "INFO"
        }
        Remove-Job $updateJob -Force -ErrorAction SilentlyContinue
    }
        
    Write-Progress -Activity "Initializing Profile" -Completed
    Write-Information "PowerShell profile loaded successfully" -InformationAction Continue
    Write-ProfileLog "PowerShell profile initialization completed successfully" "INFO"
}
catch
{
    Write-Warning "Profile initialization encountered an error: $($_.Exception.Message)"
    Write-ProfileLog "Profile initialization error: $($_.Exception.Message)" "ERROR"
    Write-Progress -Activity "Initializing Profile" -Completed
}