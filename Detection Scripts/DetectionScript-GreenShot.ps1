<#
.SYNOPSIS
    Greenshot Detection Script

.DESCRIPTION
    This PowerShell script detects the presence of Greenshot software
    on the system and returns version information. Designed for enterprise deployment
    and software inventory management in MSP environments.
    Supports both 32-bit and 64-bit installations.

.PARAMETER SearchPaths
    Array of path patterns where Greenshot executable should be located.
    Default: 64-bit and 32-bit Program Files directories

.OUTPUTS
    System.String
    Returns detection status and version information, or error messages.

.EXAMPLE
    PS> .\Detect-Greenshot.ps1
    Greenshot found - Version: 1.2.10.6 - Path: C:\Program Files\Greenshot\Greenshot.exe

.EXAMPLE
    PS> .\Detect-Greenshot.ps1
    Greenshot not found

.NOTES
    File Name      : Detect-Greenshot.ps1
    Author         : ctrlaltnod.com
    Prerequisite   : PowerShell 3.0 or higher
    Creation Date  : September 16, 2025
    Purpose/Change : Greenshot detection for MSP deployment workflows

.LINK
    https://ctrlaltnod.com

#>

# CTRLALTNOD.COM
# Access Granted : CTRL maîtrisé, ALT-ernative validée, NOD confirmé
#
# Copyright © 2025 ctrlaltnod.com - All rights reserved
# 
# ============================================================================
# SCRIPT: Greenshot Detection
# PURPOSE: Enterprise software detection for MSP deployment workflows
# AUTHOR: ctrlaltnod.com
# VERSION: 1.0
# CREATED: September 16, 2025
# UPDATED: September 16, 2025
# ============================================================================

#Requires -Version 3.0

# Initialize script parameters and variables
param(
    [Parameter(Mandatory = $false)]
    [string[]]$SearchPaths = @(
        "C:\Program Files\Greenshot\Greenshot.exe",
        "C:\Program Files (x86)\Greenshot\Greenshot.exe"
    )
)

# Set error handling preferences
$ErrorActionPreference = "SilentlyContinue"

# Main detection logic
try {
    Write-Verbose "Starting Greenshot detection process..."
    Write-Verbose "Search Paths: $($SearchPaths -join ', ')"

    # Search for Greenshot executable in both 64-bit and 32-bit locations
    $GreenshotExe = Get-ChildItem -Path $SearchPaths -ErrorAction SilentlyContinue | Select-Object -First 1

    # Check if executable was found
    if ($null -ne $GreenshotExe) {
        Write-Verbose "Executable found at: $($GreenshotExe.FullName)"

        # Extract version information from executable
        $VersionInfo = Get-Item -Path $GreenshotExe.FullName -ErrorAction SilentlyContinue
        $FileVersion = $VersionInfo.VersionInfo.FileVersion
        $ProductVersion = $VersionInfo.VersionInfo.ProductVersion

        # Determine architecture based on path
        $Architecture = if ($GreenshotExe.FullName -match "Program Files \(x86\)") { "32-bit" } else { "64-bit" }

        # Output success message with version details
        $OutputMessage = "Greenshot found - Version: $FileVersion ($Architecture) - Path: $($GreenshotExe.FullName)"
        Write-Output $OutputMessage
        Write-Verbose "Detection completed successfully"
        Write-Verbose "Architecture: $Architecture"
        Write-Verbose "File Version: $FileVersion"
        Write-Verbose "Product Version: $ProductVersion"

        # Return success exit code for deployment systems
        Exit 0
    }
    else {
        # Software not found - output appropriate message
        $OutputMessage = "Greenshot not found"
        Write-Output $OutputMessage
        Write-Verbose "No installation detected at specified paths"
        Write-Verbose "Searched paths: $($SearchPaths -join ', ')"

        # Return failure exit code for deployment systems
        Exit 1
    }
}
catch {
    # Handle any unexpected errors during execution
    $ErrorMessage = "Error during detection: $($_.Exception.Message)"
    Write-Output $ErrorMessage
    Write-Error $ErrorMessage
    Write-Verbose "Script execution failed with error: $($_.Exception.Message)"

    # Return error exit code
    Exit 1
}
finally {
    # Cleanup and final logging
    Write-Verbose "Greenshot detection script completed"
}

# End of script
# ============================================================================
# ctrlaltnod.com - Your MSP IT Reflex
# Tested scripts, shared expertise, real solutions
# ============================================================================
