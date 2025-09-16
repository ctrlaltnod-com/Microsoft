<#
.SYNOPSIS
    Autodesk DWG TrueView 2026 Detection Script

.DESCRIPTION
    This PowerShell script detects the presence of Autodesk DWG TrueView 2026 software
    on the system and returns version information. Designed for enterprise deployment
    and software inventory management in MSP environments.

.PARAMETER SearchPath
    The path pattern where DWG TrueView 2026 executable should be located.
    Default: "C:\Program Files\Autodesk\DWG TrueView 2026*\dwgviewr.exe"

.OUTPUTS
    System.String
    Returns detection status and version information, or error messages.

.EXAMPLE
    PS> .\Detect-DWGTrueView2026.ps1
    Autodesk DWG TrueView 2026 found - Version: 24.1.51.0

.EXAMPLE
    PS> .\Detect-DWGTrueView2026.ps1
    Autodesk DWG TrueView 2026 not found

.NOTES
    File Name      : Detect-DWGTrueView2026.ps1
    Author         : ctrlaltnod.com
    Prerequisite   : PowerShell 3.0 or higher
    Creation Date  : September 16, 2025
    Purpose/Change : Software detection for MSP deployment workflows

.LINK
    https://ctrlaltnod.com

#>

# CTRLALTNOD.COM
# Access Granted : CTRL maîtrisé, ALT-ernative validée, NOD confirmé
#
# Copyright © 2025 ctrlaltnod.com - All rights reserved
# 
# ============================================================================
# SCRIPT: Autodesk DWG TrueView 2026 Detection
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
    [string]$SearchPath = "C:\Program Files\Autodesk\DWG TrueView 2026*\dwgviewr.exe"
)

# Set error handling preferences
$ErrorActionPreference = "SilentlyContinue"

# Main detection logic
try {
    Write-Verbose "Starting Autodesk DWG TrueView 2026 detection process..."
    Write-Verbose "Search Path: $SearchPath"

    # Search for DWG TrueView executable using wildcard pattern
    $DWGTrueViewExe = Get-ChildItem -Path $SearchPath -ErrorAction SilentlyContinue | Select-Object -First 1

    # Check if executable was found
    if ($null -ne $DWGTrueViewExe) {
        Write-Verbose "Executable found at: $($DWGTrueViewExe.FullName)"

        # Extract version information from executable
        $VersionInfo = Get-Item -Path $DWGTrueViewExe.FullName -ErrorAction SilentlyContinue
        $ProductVersion = $VersionInfo.VersionInfo.ProductVersionRaw

        # Output success message with version details
        $OutputMessage = "Autodesk DWG TrueView 2026 found - Version: $ProductVersion"
        Write-Output $OutputMessage
        Write-Verbose "Detection completed successfully"

        # Return success exit code for deployment systems
        Exit 0
    }
    else {
        # Software not found - output appropriate message
        $OutputMessage = "Autodesk DWG TrueView 2026 not found"
        Write-Output $OutputMessage
        Write-Verbose "No installation detected at specified path"

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
    Write-Verbose "DWG TrueView 2026 detection script completed"
}

# End of script
# ============================================================================
# ctrlaltnod.com - Your MSP IT Reflex
# Tested scripts, shared expertise, real solutions
# ============================================================================
