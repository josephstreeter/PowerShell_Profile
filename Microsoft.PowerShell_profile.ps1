function Compare-UpdateLastRun()
{
	try
	{
		if (-not (Test-Path ~\.config))
		{
			New-Item ~\.config -ItemType Directory -ErrorAction Stop | Out-Null
		}
		
		if (Test-Path ~\.config\lastrun.txt)
		{
			New-Item ~\.config\lastrun.txt -ItemType File -ErrorAction Stop | Out-Null
			(Get-Date).AddDays(-1).ToString("yyyy-MM-dd") | Set-Content ~\.config\lastrun.txt
		}
		
	}
	catch
	{
		Write-Error ("Failed to create lastrun.txt: {0}" -f $_.Exception.Message)
	}


	try
	{
		$lastrun = Get-Content ~\.config\lastrun.txt
	  
		if (((Get-Date $lastrun).ToString("yyyy-MM-dd") -eq (Get-Date).ToString("yyyy-MM-dd")) -eq $true)
		{
			#Write-Output "Skipping update check, last run was $lastrun"
			Return $true
		}
		else
		{
			#Write-Output "Checking for updates, last run was $lastrun"
			(Get-Date).ToString("yyyy-MM-dd") | Set-Content ~\.config\lastrun.txt
			Return $false
		}
	}
	catch
	{
		Write-Error ("Failed to read lastrun.txt: {0}" -f $_.Exception.Message)
	}
}

function Search-WingetUpdates()
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory = $true)][array]$Apps
	)
	try
	{
		if (-not (Get-Module microsoft.winget.client -ListAvailable))
		{
			Install-Module microsoft.winget.client -Repository PSGallery -Force
		}
  
		Write-Output "Checking for WinGet updates...."
    
		$Updates = Get-WinGetPackage | Where-Object { $Apps -contains $_.id -and $_.IsUpdateAvailable }
    
		if ($Updates.count -eq 0)
		{
			Write-Output "No WinGet updates available"
			Return
		}
    
		Write-Warning -Message "Updates available for $($Updates.id -join ", ")`nRun 'winget upgrade $($Updates.id -join ", ")' to install updates"
		Return
	}
	catch
	{
		Write-Error -Message ("Failed to check for updates: {0}" -f $_.Exception.Message)
		Return
	}
}

function Install-CommonModule()
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory = $true)][array]$Modules
	)
	try
	{
		Write-Output "Installing common PowerShell modules...."

		foreach ($Module in $Modules)
		{
			if ((Get-Module $Module.name -ListAvailable).count -eq 0)
			{
				Write-Output ("Installing module {0}" -f $Module.Name)
				Install-Module $Module.Name -Repository $Module.PSRepository -Force -ErrorAction Continue
			}
		}
	}
	catch
	{
		Write-Host ("Failed to install module: {0}" -f $_.Exception.Message)
	}
}

function Search-ModuleUpdate()
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory = $true)][array]$Modules
	)

	try
	{
		Write-Output "Checking for PowerShell module updates...."
  
		$Updates = @()
		foreach ($Module in $Modules)
		{
			$InstalledVersion = (Get-InstalledModule -Name $Module.Name).version
			$AvailableVersion = (Find-Module -Name $Module.Name -Repository $Module.PSRepository).version

			if ($InstalledVersion -ne $AvailableVersion)
			{
				$Updates += $Module
			}
		}
  
		if ($Updates.count -eq 0)
		{
			Write-Output "Modules are up to date"
			Return
		}
  
		Write-Warning -Message "Updates available for $($Updates -join ", ")`nRun 'Update-Module $($Updates -join ", ")' to install updates"
	}
	catch
	{
		Write-Error -Message ("Failed to check for updates: {0}" -f $_.Exception.Message)
	}
}

function Import-LocalModule()
{
	[CmdletBinding()]
	Param
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

Function Set-CommonVariables()
{
	[CmdletBinding()]
	Param()
	try
	{
		Write-Output "Setting common variables...."
		$Global:TenantUrl = "https://madisoncollege365-admin.sharepoint.com/"
		$Global:instance = "IDMDBPRD01\MIMStage"
		$Global:database = "StagingDirectory"

	}
	catch
	{
		Write-Error -Message ("Failed to set common variables: {0}" -f $_.Exception.Message)
	}
}

function Initialize-OhMyPosh()
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory = $true)][string]$ConfigFile
	)
  
	if (Get-WinGetPackage JanDeDobbeleer.OhMyPosh)
	{
		Write-Output "Oh My Posh is already installed"
	}
	else
	{
		Install-WinGetPackage -Id JanDeDobbeleer.OhMyPosh -Source winget -Force
	}
  
	Write-Output "Initializing Oh My Posh...."
  
	oh-my-posh --init --shell pwsh --config $ConfigFile | Invoke-Expression
}

function Initialize-PSReadLine()
{
	[CmdletBinding()]
	Param()

	try
	{
		if (-not (Get-Module PSReadLine -ListAvailable))
		{
			Install-Module PSReadLine -Repository PSGallery -Force
		}
  
		Write-Output "Initializing PSReadLine...."

		$Params = @{
			BellStyle           = "None"
			PredictionSource    = "History"
			PredictionViewStyle = "Listview"
		}
  
		Set-PSReadLineOption @Params -ErrorAction Stop
	}
	catch
	{
		Write-Error -Message ("Failed to initialize PSReadLine: {0}" -f $_.Exception.Message)
	}
}

# End Functions ############################################################################################################

$SkipUpdates = Compare-UpdateLastRun

if ($SkipUpdates -eq $true)
{
	Write-Output "Skipping update check"
}
else
{
	# Manage WinGet Apps
	$Apps = "Microsoft.WindowsTerminal", "Git.Git", "Microsoft.PowerShell", "Microsoft.VisualStudioCode", "JanDeDobbeleer.OhMyPosh"
	Search-WingetUpdates -Apps $Apps
	
	# Manage PowerShell Modules
	$Modules = @(
		[PSCustomObject]@{"Name" = "ExchangeOnlineManagement"; "PSRepository" = "PSGallery" }
		[PSCustomObject]@{"Name" = "Microsoft.Graph"; "PSRepository" = "PSGallery" }
		[PSCustomObject]@{"Name" = "PSReadLine"; "PSRepository" = "PSGallery" }
	)
	
	Search-ModuleUpdate -Modules $Modules
	Install-CommonModule -Modules $Modules
}

Import-LocalModule "C:\Users\JStreeter\source\repos\IAM_PowerShell_Modules\Exchange\PS-EX-Administration.psm1"

# Initialize PowerShell Environment
Set-CommonVariables
Initialize-OhMyPosh -ConfigFile "https://raw.githubusercontent.com/josephstreeter/Oh-My-Posh/refs/heads/main/jstreeter.omp.json"
Initialize-PSReadLine
