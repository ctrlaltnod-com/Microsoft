<#
.SYNOPSIS
    VLC Media Player Detection Script - INTUNE OPTIMIZED

.DESCRIPTION
    This PowerShell script detects the presence of VLC Media Player software
    on the system for Microsoft Intune Win32 app deployment. Returns correct exit codes
    and STDOUT output for proper Intune detection logic.
    Supports both 32-bit and 64-bit installations.

.PARAMETER SearchPaths
    Array of path patterns where VLC executable should be located.
    Default: 64-bit and 32-bit Program Files directories

.OUTPUTS
    System.String
    Returns detection status ONLY when application is found (for Intune STDOUT requirement)

.EXAMPLE
    PS> .\Detect-VLCMediaPlayer-Intune.ps1
    VLC Media Player found - Version: 3.0.18.0 (64-bit)
    (Exit code 0 = App detected)

.EXAMPLE
    PS> .\Detect-VLCMediaPlayer-Intune.ps1
    (No output, Exit code 1 = App not detected)

.NOTES
    File Name      : Detect-VLCMediaPlayer-Intune.ps1
    Author         : ctrlaltnod.com
    Prerequisite   : PowerShell 3.0 or higher
    Creation Date  : September 16, 2025
    Purpose/Change : Intune Win32 app detection with proper STDOUT handling
    Context        : Runs in SYSTEM context via Intune Management Extension

.LINK
    https://ctrlaltnod.com

#>

# ┌─┐┌┬┐┬─┐┬  ┌─┐┬ ┌┬┐┌┐┌┌─┐┌┬┐  ┌─┐┌─┐┌┬┐
# │   │ ├┬┘│  ├─┤│  │ ││││ │ ││  │  │ ││││
# └─┘ ┴ ┴└─┴─┘┴ ┴┴─┘┴ ┘┘└┘└─┘─┴┘ ┘└─┘└─┘┴ ┴
# 
# Site: https://ctrlaltnod.com
# Slogan: Access Granted : CTRL maîtrisé, ALT-ernative validée, NOD confirmé
# Copyright © 2025 ctrlaltnod.com - All rights reserved
# 
# ============================================================================
# SCRIPT: VLC Media Player Detection - INTUNE OPTIMIZED
# PURPOSE: Win32 app detection for Microsoft Intune deployment
# AUTHOR: ctrlaltnod.com MSP Team
# VERSION: 1.1 (Intune-Ready)
# CREATED: September 16, 2025
# UPDATED: September 16, 2025
# ============================================================================

#Requires -Version 3.0

# Initialize script parameters and variables
param(
    [Parameter(Mandatory = $false)]
    [string[]]$SearchPaths = @(
        "C:\Program Files\VideoLAN\VLC\vlc.exe",
        "C:\Program Files (x86)\VideoLAN\VLC\vlc.exe"
    )
)

# Set error handling preferences for Intune context
$ErrorActionPreference = "SilentlyContinue"

# Main detection logic optimized for Intune Win32 apps
try {
    # Search for VLC executable in both 64-bit and 32-bit locations
    $VLCExe = Get-ChildItem -Path $SearchPaths -ErrorAction SilentlyContinue | Select-Object -First 1

    # Check if executable was found
    if ($null -ne $VLCExe) {
        # App found - Extract version information
        $VersionInfo = Get-Item -Path $VLCExe.FullName -ErrorAction SilentlyContinue
        $FileVersion = $VersionInfo.VersionInfo.FileVersion

        # Determine architecture based on path
        $Architecture = if ($VLCExe.FullName -match "Program Files \(x86\)") { "32-bit" } else { "64-bit" }

        # ✅ INTUNE SUCCESS: Output to STDOUT + Exit 0 = App Detected
        Write-Output "VLC Media Player found - Version: $FileVersion ($Architecture)"
        Exit 0
    }
    else {
        # ✅ INTUNE NOT FOUND: NO Output + Exit 1 = App Not Detected  
        # Important: Do NOT use Write-Output here or Intune will detect the app incorrectly
        Exit 1
    }
}
catch {
    # ✅ INTUNE ERROR: NO Output + Exit 1 = App Not Detected
    # Important: Do NOT use Write-Output for errors in Intune detection scripts
    Exit 1
}

# End of script - Intune Detection Optimized
# ============================================================================
# ctrlaltnod.com - Your MSP IT Reflex
# Intune-ready scripts for professional MSP deployments
# ============================================================================
