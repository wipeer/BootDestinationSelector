# BootDestinationSelector.ps1
# Script to select which Windows installation to boot to next and restart
# Author: Slavomir Prkno
# Version: 1.4
# Description: Allows selecting a Windows installation for one-time boot with optional BitLocker suspension

#region Auto-Elevation
# Check if script is running as Administrator and self-elevate if not
$currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
$adminRole = [Security.Principal.WindowsBuiltInRole]::Administrator

# If not running as Administrator, restart script with elevation
if (-not $principal.IsInRole($adminRole)) {
    try {
        # Get the script path
        $scriptPath = $MyInvocation.MyCommand.Definition
        
        # Create a new process info object
        $processInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processInfo.FileName = "powershell.exe"
        $processInfo.Arguments = "-ExecutionPolicy Bypass -File `"$scriptPath`""
        $processInfo.Verb = "runas" # This triggers the UAC prompt
        $processInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Normal
        
        # Start the process
        [System.Diagnostics.Process]::Start($processInfo) | Out-Null
        
        # Exit the current non-elevated script
        exit
    }
    catch {
        Write-Host "Error: Failed to restart script with admin privileges." -ForegroundColor Red
        Write-Host "Please right-click on the script and select 'Run with PowerShell as Administrator'." -ForegroundColor Yellow
        Write-Host "Error details: $_" -ForegroundColor Gray
        
        # Pause so the user can read the message
        Write-Host "Press any key to exit..." -ForegroundColor Cyan
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit
    }
}
#endregion

#region Functions
<#
.SYNOPSIS
    Retrieves Windows boot entries from the system.
.DESCRIPTION
    Parses the output of bcdedit /enum to extract information about Windows boot entries.
    Only returns Windows Boot Loader entries (actual OS installations).
.OUTPUTS
    Array of hashtables containing boot entry information.
#>
function Get-BootEntries {
    # Get boot entries using bcdedit with error handling
    try {
        $bcdeditOutput = bcdedit /enum 2>&1
        
        # Check if bcdedit returned an error
        if ($LASTEXITCODE -ne 0 -or $bcdeditOutput -match "error") {
            throw "Failed to execute bcdedit /enum. Error code: $LASTEXITCODE. Output: $bcdeditOutput"
        }
        
        # Convert to string if it's not already
        $bcdeditOutput = $bcdeditOutput | Out-String
    }
    catch {
        throw "Error retrieving boot entries: $_"
    }
    
    # Process the output to extract boot entries
    $entries = @()
    $currentEntry = $null
    $lines = $bcdeditOutput -split "`r`n"
    
    foreach ($line in $lines) {
        # Skip empty lines
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        
        # Start of a new entry
        if ($line -match "^Windows Boot (Manager|Loader)$") {
            # Save previous entry if exists
            if ($null -ne $currentEntry) {
                $entries += $currentEntry
            }
            
            # Create new entry
            $currentEntry = @{
                Type = $line
                Identifier = ""
                Description = ""
                Device = ""
                Path = ""
                Properties = @{}
                IsDefault = $false
            }
        }
        elseif ($line -match "^-+$") {
            # Separator line, skip
            continue
        }
        elseif ($null -ne $currentEntry) {
            # Process entry properties with improved pattern matching
            if ($line -match "^\s*identifier\s+(.+)$") {
                $currentEntry.Identifier = $matches[1].Trim()
                # Mark default entry
                if ($currentEntry.Identifier -eq "{default}") {
                    $currentEntry.IsDefault = $true
                }
            }
            elseif ($line -match "^\s*description\s+(.+)$") {
                $currentEntry.Description = $matches[1].Trim()
            }
            elseif ($line -match "^\s*device\s+(.+)$") {
                $currentEntry.Device = $matches[1].Trim()
            }
            elseif ($line -match "^\s*path\s+(.+)$") {
                $currentEntry.Path = $matches[1].Trim()
            }
            elseif ($line -match "^\s*(\w+)\s+(.+)$") {
                # Store other properties
                $currentEntry.Properties[$matches[1].Trim()] = $matches[2].Trim()
            }
        }
    }
    
    # Add the last entry
    if ($null -ne $currentEntry) {
        $entries += $currentEntry
    }
    
    # Return only Windows Boot Loader entries (actual OS installations)
    $loaderEntries = $entries | Where-Object { $_.Type -eq "Windows Boot Loader" }
    
    # Return entries in original order (no sorting)
    return $loaderEntries
}

<#
.SYNOPSIS
    Displays an interactive menu for selecting a boot entry.
.DESCRIPTION
    Shows a formatted menu of available Windows boot entries and handles user selection.
.PARAMETER Entries
    Array of boot entry objects to display in the menu.
.OUTPUTS
    Selected boot entry or $null if cancelled.
#>
function Show-Menu {
    param (
        [Parameter(Mandatory=$true)]
        [array]$Entries
    )
    
    Clear-Host
    
    # Fixed header without using expressions that show in the output
    Write-Host "" 
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "                  Windows Boot Selector                     " -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan
    
    Write-Host "`nSelect the Windows installation to boot to for the next restart:" -ForegroundColor Yellow
    Write-Host "(System will return to default boot entry after this restart)`n" -ForegroundColor Yellow
    
    # Display entries with improved formatting
    for ($i = 0; $i -lt $Entries.Count; $i++) {
        $entry = $Entries[$i]
        $indicators = @()
        
        if ($entry.Identifier -eq "{current}") {
            $indicators += "Current"
        }
        if ($entry.IsDefault) {
            $indicators += "Default"
        }
        
        $indicatorStr = if ($indicators.Count -gt 0) { " (" + ($indicators -join ", ") + ")" } else { "" }
        
        Write-Host "[$($i + 1)] $($entry.Description)$indicatorStr" -ForegroundColor Green
        
        # Show device info with better formatting
        $deviceInfo = $entry.Device -replace "partition=", "Drive: "
        Write-Host "    $deviceInfo" -ForegroundColor Gray
        
        # Show path in a friendly format
        $pathInfo = $entry.Path -replace "\\", "\"
        Write-Host "    Path: $pathInfo" -ForegroundColor Gray
        Write-Host ""
    }
    
    Write-Host "[R] Refresh list" -ForegroundColor Blue
    Write-Host "[C] Cancel" -ForegroundColor Red
    Write-Host ""
    
    $validSelection = $false
    while (-not $validSelection) {
        $selection = Read-Host "Enter your choice"
        
        # Handle special commands
        if ($selection -eq "C" -or $selection -eq "c") {
            return $null
        }
        elseif ($selection -eq "R" -or $selection -eq "r") {
            # Refresh the entries and redisplay menu
            try {
                $refreshedEntries = Get-BootEntries
                return Show-Menu -Entries $refreshedEntries
            }
            catch {
                Write-Host "Error refreshing boot entries: $_" -ForegroundColor Red
                Start-Sleep -Seconds 2
                return Show-Menu -Entries $Entries
            }
        }
        
        # Handle numeric selection
        if ($selection -match "^\d+$") {
            $index = [int]$selection - 1
            if ($index -ge 0 -and $index -lt $Entries.Count) {
                $validSelection = $true
                return $Entries[$index]
            }
        }
        
        # Invalid selection
        Write-Host "Invalid selection. Please try again." -ForegroundColor Red
    }
}

<#
.SYNOPSIS
    Checks BitLocker status and optionally suspends for one reboot.
.DESCRIPTION
    Determines if a drive is BitLocker encrypted and offers to suspend protection
    for one reboot if it is.
.PARAMETER MountPoint
    The drive letter to check and potentially suspend BitLocker for.
.OUTPUTS
    Boolean indicating whether BitLocker was suspended.
#>
function Manage-BitLocker {
    param (
        [Parameter(Mandatory=$true)]
        [string]$MountPoint
    )
    
    # Make sure mount point has the correct format
    if (-not $MountPoint.EndsWith(":")) {
        $MountPoint = $MountPoint + ":"
    }
    
    try {
        # Try to get BitLocker status, will fail if BitLocker is not available
        $bitlockerVolume = Get-BitLockerVolume -MountPoint $MountPoint -ErrorAction SilentlyContinue
        
        # Skip if BitLocker isn't installed or if command fails
        if ($null -eq $bitlockerVolume) {
            Write-Host "BitLocker is not available on this system or drive." -ForegroundColor Yellow
            return $false
        }
        
        # Check if BitLocker is enabled on the selected drive
        if ($bitlockerVolume.VolumeStatus -eq "FullyEncrypted" -or 
            $bitlockerVolume.VolumeStatus -eq "EncryptionInProgress" -or
            $bitlockerVolume.ProtectionStatus -eq "On") {
            
            Write-Host "`nBitLocker is enabled on drive $MountPoint" -ForegroundColor Yellow
            $suspendBitlocker = Read-Host "Would you like to suspend BitLocker protection for one reboot? (Y/n)"
            
            # Default to yes if empty input
            if ([string]::IsNullOrWhiteSpace($suspendBitlocker) -or 
                $suspendBitlocker.ToLower() -eq "y") {
                
                Write-Host "Suspending BitLocker for one reboot cycle..." -ForegroundColor Yellow
                Suspend-BitLocker -MountPoint $MountPoint -RebootCount 1
                
                if ($LASTEXITCODE -eq 0) {
                    Write-Host "BitLocker suspended successfully for one reboot." -ForegroundColor Green
                    return $true
                }
                else {
                    Write-Host "Failed to suspend BitLocker. Error code: $LASTEXITCODE" -ForegroundColor Red
                    return $false
                }
            }
            else {
                Write-Host "BitLocker suspension skipped." -ForegroundColor Yellow
                return $false
            }
        }
        else {
            Write-Host "BitLocker is not enabled on drive $MountPoint." -ForegroundColor Gray
            return $false
        }
    }
    catch {
        Write-Host "Error checking BitLocker status: $_" -ForegroundColor Red
        return $false
    }
}

<#
.SYNOPSIS
    Sets a one-time boot entry and optionally restarts the system.
.DESCRIPTION
    Configures the system to boot to the specified entry on next restart.
    Optionally initiates a system restart with confirmation.
.PARAMETER Entry
    The boot entry to set for the next restart.
#>
function Set-OneTimeBoot {
    param (
        [Parameter(Mandatory=$true)]
        [hashtable]$Entry
    )
    
    Write-Host "`nSetting one-time boot sequence to $($Entry.Description)..." -ForegroundColor Yellow
    
    try {
        # Execute bcdedit command with proper error handling
        $output = bcdedit /bootsequence $Entry.Identifier /addfirst 2>&1
        
        if ($LASTEXITCODE -ne 0) {
            throw "bcdedit command failed with exit code $LASTEXITCODE. Output: $output"
        }
        
        Write-Host "Boot sequence successfully set." -ForegroundColor Green
        Write-Host "System will boot to $($Entry.Description) on next restart only." -ForegroundColor Green
        
        # Check for BitLocker on the current drive
        if ($Entry.Identifier -eq "{current}") {
            # If current OS, get its drive letter (usually C:)
            $driveToCheck = $env:SystemDrive
        }
        else {
            # For other entries, extract drive letter from device property
            $driveToCheck = $Entry.Device -replace '.*partition=', ''
            if ($driveToCheck -match "\\Device\\HarddiskVolume\d+") {
                # For device path format, just use C: as default since we can't easily map this
                $driveToCheck = "C:"
            }
        }
        
        # Check and optionally suspend BitLocker
        $bitlockerSuspended = Manage-BitLocker -MountPoint $driveToCheck
        
        # Ask for confirmation before restarting with improved default handling
        $confirmRestart = Read-Host "`nSystem will restart now and boot to $($Entry.Description). Continue? (Y/n)"
        
        # Default to yes if empty input
        if ([string]::IsNullOrWhiteSpace($confirmRestart) -or 
            $confirmRestart.ToLower() -eq "y") {
            
            # Countdown timer for restart
            $seconds = 5
            Write-Host "`nRestarting system in $seconds seconds..." -ForegroundColor Yellow
            
            while ($seconds -gt 0) {
                Write-Host "`rCountdown: $seconds..." -ForegroundColor Red -NoNewline
                Start-Sleep -Seconds 1
                $seconds--
            }
            
            Write-Host "`rSystem is restarting now. Goodbye!               " -ForegroundColor Red
            
            # Initiate restart
            shutdown /r /t 0
        }
        else {
            Write-Host "`nRestart cancelled. The boot sequence has been set for the next restart." -ForegroundColor Cyan
            Write-Host "When you're ready to restart, use the normal Windows restart procedure." -ForegroundColor Cyan
            
            if ($bitlockerSuspended) {
                Write-Host "NOTE: BitLocker has been suspended for one reboot." -ForegroundColor Yellow
            }
        }
    }
    catch {
        Write-Host "Error: $_" -ForegroundColor Red
        return $false
    }
    
    return $true
}
#endregion

#region Main Script
try {
    Write-Host "Scanning for Windows installations... Please wait." -ForegroundColor Cyan
    
    # Get boot entries
    $bootEntries = Get-BootEntries
    
    if ($bootEntries.Count -eq 0) {
        throw "No Windows boot entries found. This is unusual and may indicate a system configuration issue."
    }
    
    # Show menu and get selection
    $selectedEntry = Show-Menu -Entries $bootEntries
    
    # Handle selection
    if ($null -eq $selectedEntry) {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        exit 0
    }
    
    # Set the selected entry as the next boot target
    Set-OneTimeBoot -Entry $selectedEntry
}
catch {
    # Enhanced error handling with detailed messages
    Write-Host "`nERROR: An unexpected error occurred:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    
    # Show stack trace in debug mode
    $VerbosePreference = "Continue"
    Write-Verbose "Stack Trace: $($_.ScriptStackTrace)"
    
    # Provide potential solutions
    Write-Host "`nPossible solutions:" -ForegroundColor Yellow
    Write-Host "1. Make sure you're running PowerShell as Administrator" -ForegroundColor Yellow
    Write-Host "2. Check if bcdedit.exe is accessible (it should be part of Windows)" -ForegroundColor Yellow
    Write-Host "3. Ensure your system boot configuration is not corrupted" -ForegroundColor Yellow
    Write-Host "4. If BitLocker error: Verify BitLocker is enabled in Windows features" -ForegroundColor Yellow
    
    # Pause so error can be read
    Write-Host "`nPress any key to exit..." -ForegroundColor Cyan
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}
finally {
    # Cleanup code if needed (currently none required)
}
#endregion