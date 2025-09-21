<#
.SYNOPSIS
    Active Directory User Audit Report

.DESCRIPTION
    This script generates a comprehensive audit report of all Active Directory users.
    It retrieves user information including creation date, last logon, password last set,
    group memberships, and account status.

.AUTHOR
    IT Admin - ctrlaltnod.com

.VERSION
    1.0

.DATE
    September 2025

.EXAMPLE
    .\Get-ADUserAuditReport-Fixed.ps1

.EXAMPLE
    .\Get-ADUserAuditReport-Fixed.ps1 -ExportPath "C:\Reports\ADUserAudit.csv"

.EXAMPLE
    .\Get-ADUserAuditReport-Fixed.ps1 -OU "OU=Sales,DC=domain,DC=com"

.EXAMPLE
    .\Get-ADUserAuditReport-Fixed.ps1 -IncludeDisabledUsers -InactiveDays 90

.NOTES
    Requires Active Directory PowerShell module
    Must be run on domain controller or machine with RSAT tools installed
    Requires appropriate AD read permissions
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = "C:\Temp\AD-UserAuditReport-$(Get-Date -Format 'yyyyMMdd-HHmm').csv",

    [Parameter(Mandatory=$false)]
    [string]$OU = "",

    [Parameter(Mandatory=$false)]
    [switch]$IncludeDisabledUsers,

    [Parameter(Mandatory=$false)]
    [int]$InactiveDays = 90,

    [Parameter(Mandatory=$false)]
    [switch]$ShowProgress = $true,

    [Parameter(Mandatory=$false)]
    [switch]$IncludeGroups,

    [Parameter(Mandatory=$false)]
    [switch]$CheckAllDCs
)

# Import Active Directory module
function Import-ADModule {
    Write-Host "Checking Active Directory PowerShell module..." -ForegroundColor Yellow

    if (!(Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-Host "Active Directory module not found. Please install RSAT tools." -ForegroundColor Red
        Write-Host "Download from: https://www.microsoft.com/download/details.aspx?id=45520" -ForegroundColor Yellow
        exit 1
    }

    try {
        Import-Module ActiveDirectory -Force
        Write-Host "Active Directory module loaded successfully." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to import Active Directory module: $($_.Exception.Message)"
        exit 1
    }
}

# Function to get domain controllers
function Get-DomainControllers {
    try {
        $DCs = Get-ADDomainController -Filter * | Select-Object -ExpandProperty HostName
        Write-Host "Found $($DCs.Count) domain controllers." -ForegroundColor Green
        return $DCs
    }
    catch {
        Write-Warning "Could not retrieve domain controllers: $($_.Exception.Message)"
        return @($env:LOGONSERVER -replace '\\','')
    }
}

# Function to get real last logon from all DCs - CORRIGÉE
function Get-RealLastLogon {
    param(
        [string]$UserSamAccountName,
        [array]$DomainControllers
    )

    if (!$CheckAllDCs) {
        # Use LastLogonDate from current DC (faster but less accurate)
        try {
            $User = Get-ADUser -Identity $UserSamAccountName -Properties LastLogonDate -ErrorAction SilentlyContinue
            return $User.LastLogonDate
        }
        catch {
            return $null
        }
    }

    $MostRecentLogon = $null
    $MostRecentDate = [DateTime]::MinValue

    foreach ($DC in $DomainControllers) {
        try {
            $User = Get-ADUser -Identity $UserSamAccountName -Server $DC -Properties LastLogon -ErrorAction SilentlyContinue
            if ($User.LastLogon -and $User.LastLogon -gt 0) {
                $LogonDate = [DateTime]::FromFileTime($User.LastLogon)
                if ($LogonDate -gt $MostRecentDate) {
                    $MostRecentDate = $LogonDate
                    $MostRecentLogon = $LogonDate
                }
            }
        }
        catch {
            # Continue to next DC if this one fails
            continue
        }
    }

    return $MostRecentLogon
}

# Function to get user group memberships
function Get-UserGroups {
    param(
        [string]$UserDistinguishedName
    )

    if (!$IncludeGroups) {
        return "Not collected"
    }

    try {
        $Groups = Get-ADUser -Identity $UserDistinguishedName -Properties MemberOf -ErrorAction SilentlyContinue | 
                  Select-Object -ExpandProperty MemberOf |
                  ForEach-Object { 
                      try {
                          (Get-ADGroup -Identity $_ -ErrorAction SilentlyContinue).Name
                      } catch {
                          "Unknown Group"
                      }
                  }
        if ($Groups) {
            return ($Groups -join "; ")
        } else {
            return "No groups"
        }
    }
    catch {
        return "Error retrieving groups"
    }
}

# Function to calculate password age and expiry - CORRIGÉE
function Get-PasswordInfo {
    param(
        [AllowNull()]
        [DateTime]$PasswordLastSet,
        [bool]$PasswordNeverExpires
    )

    if ($PasswordNeverExpires) {
        $DaysSinceChange = if ($PasswordLastSet) { 
            (New-TimeSpan -Start $PasswordLastSet -End (Get-Date)).Days 
        } else { 
            "Never set" 
        }

        return @{
            DaysSincePasswordChange = $DaysSinceChange
            PasswordExpired = $false
            PasswordExpiryDate = "Never"
        }
    }

    try {
        $DefaultPasswordPolicy = Get-ADDefaultDomainPasswordPolicy -ErrorAction SilentlyContinue
        if ($DefaultPasswordPolicy) {
            $MaxPasswordAge = $DefaultPasswordPolicy.MaxPasswordAge.Days
        } else {
            $MaxPasswordAge = 90  # Default fallback
        }

        if ($PasswordLastSet) {
            $DaysSinceChange = (New-TimeSpan -Start $PasswordLastSet -End (Get-Date)).Days
            $ExpiryDate = $PasswordLastSet.AddDays($MaxPasswordAge)
            $IsExpired = (Get-Date) -gt $ExpiryDate

            return @{
                DaysSincePasswordChange = $DaysSinceChange
                PasswordExpired = $IsExpired
                PasswordExpiryDate = $ExpiryDate.ToString("yyyy-MM-dd")
            }
        } else {
            return @{
                DaysSincePasswordChange = "Never set"
                PasswordExpired = $true
                PasswordExpiryDate = "Immediate"
            }
        }
    }
    catch {
        return @{
            DaysSincePasswordChange = "Error"
            PasswordExpired = "Error"
            PasswordExpiryDate = "Error"
        }
    }
}

# Function to determine account status - CORRIGÉE pour accepter $null
function Get-AccountStatus {
    param(
        [bool]$Enabled,
        [AllowNull()]
        [DateTime]$LastLogon,
        [int]$InactiveDays,
        [AllowNull()]
        [DateTime]$Created,
        [bool]$LockedOut
    )

    if ($LockedOut) {
        return "Locked Out"
    }

    if (!$Enabled) {
        return "Disabled"
    }

    if (!$LastLogon -or $LastLogon -eq [DateTime]::MinValue) {
        if ($Created) {
            $DaysSinceCreation = (New-TimeSpan -Start $Created -End (Get-Date)).Days
            if ($DaysSinceCreation -le 7) {
                return "New Account (Never Logged In)"
            } else {
                return "Never Logged In"
            }
        } else {
            return "Never Logged In"
        }
    }

    $DaysInactive = (New-TimeSpan -Start $LastLogon -End (Get-Date)).Days

    if ($DaysInactive -le 30) {
        return "Active"
    } elseif ($DaysInactive -le $InactiveDays) {
        return "Recently Active"
    } else {
        return "Inactive ($DaysInactive days)"
    }
}

# Function to format file size
function Format-FileSize {
    param([long]$Size)

    if ($Size -gt 1TB) { return "{0:N2} TB" -f ($Size / 1TB) }
    elseif ($Size -gt 1GB) { return "{0:N2} GB" -f ($Size / 1GB) }
    elseif ($Size -gt 1MB) { return "{0:N2} MB" -f ($Size / 1MB) }
    elseif ($Size -gt 1KB) { return "{0:N2} KB" -f ($Size / 1KB) }
    else { return "$Size bytes" }
}

# Main script execution
try {
    Write-Host "=== Active Directory User Audit Report - FIXED VERSION ===" -ForegroundColor Cyan
    Write-Host "Starting at: $(Get-Date)" -ForegroundColor Gray
    Write-Host "Script by: ctrlaltnod.com" -ForegroundColor Gray

    # Step 1: Import Active Directory module
    Import-ADModule

    # Step 2: Get domain information
    Write-Host "`nGathering domain information..." -ForegroundColor Yellow

    try {
        $Domain = Get-ADDomain
        $DomainName = $Domain.DNSRoot
        $DomainDN = $Domain.DistinguishedName
        Write-Host "Connected to domain: $DomainName" -ForegroundColor Green

        # Get domain controllers if checking all DCs
        $DomainControllers = @()
        if ($CheckAllDCs) {
            $DomainControllers = Get-DomainControllers
            Write-Host "Will check all $($DomainControllers.Count) domain controllers for accurate last logon." -ForegroundColor Cyan
        } else {
            Write-Host "Using LastLogonDate from current domain controller (faster but may be less accurate)." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Error "Failed to connect to domain: $($_.Exception.Message)"
        exit 1
    }

    # Step 3: Build search parameters
    Write-Host "`nBuilding search parameters..." -ForegroundColor Yellow

    # Determine search base
    $SearchBase = if ($OU) { 
        Write-Host "Searching in OU: $OU" -ForegroundColor Gray
        $OU 
    } else { 
        Write-Host "Searching entire domain: $DomainDN" -ForegroundColor Gray
        $DomainDN 
    }

    # Build filter
    if ($IncludeDisabledUsers) {
        $Filter = "*"
        Write-Host "Including all users (enabled and disabled)" -ForegroundColor Gray
    } else {
        $Filter = "Enabled -eq `$true"
        Write-Host "Including only enabled users" -ForegroundColor Gray
    }

    # Properties to retrieve
    $Properties = @(
        'SamAccountName', 'DisplayName', 'GivenName', 'Surname',
        'UserPrincipalName', 'EmailAddress', 'Description', 
        'Department', 'Title', 'Company', 'Manager',
        'Created', 'LastLogonDate', 'PasswordLastSet', 
        'PasswordNeverExpires', 'PasswordExpired', 'CannotChangePassword',
        'Enabled', 'LockedOut', 'AccountLockoutTime',
        'MemberOf', 'DistinguishedName', 'EmployeeID',
        'Office', 'OfficePhone', 'MobilePhone',
        'LastLogon', 'LogonCount', 'BadLogonCount'
    )

    # Step 4: Get users
    Write-Host "`nRetrieving users from Active Directory..." -ForegroundColor Yellow

    try {
        $Users = Get-ADUser -Filter $Filter -SearchBase $SearchBase -Properties $Properties -ErrorAction Stop
        Write-Host "Found $($Users.Count) users to process." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to retrieve users: $($_.Exception.Message)"
        exit 1
    }

    if ($Users.Count -eq 0) {
        Write-Warning "No users found matching the criteria."
        exit 0
    }

    # Step 5: Process each user
    Write-Host "`nProcessing user information..." -ForegroundColor Yellow

    $Report = @()
    $Counter = 0

    foreach ($User in $Users) {
        $Counter++

        if ($ShowProgress) {
            $PercentComplete = [math]::Round(($Counter / $Users.Count) * 100, 2)
            Write-Progress -Activity "Processing Users" -Status "Processing $($User.SamAccountName) ($Counter of $($Users.Count))" -PercentComplete $PercentComplete
        }

        # Get real last logon if requested - CORRIGÉ
        try {
            $LastLogon = if ($CheckAllDCs) {
                Get-RealLastLogon -UserSamAccountName $User.SamAccountName -DomainControllers $DomainControllers
            } else {
                $User.LastLogonDate
            }
        }
        catch {
            Write-Warning "Could not retrieve last logon for $($User.SamAccountName): $($_.Exception.Message)"
            $LastLogon = $null
        }

        # Get password information - CORRIGÉ
        try {
            $PasswordInfo = Get-PasswordInfo -PasswordLastSet $User.PasswordLastSet -PasswordNeverExpires $User.PasswordNeverExpires
        }
        catch {
            Write-Warning "Could not retrieve password info for $($User.SamAccountName): $($_.Exception.Message)"
            $PasswordInfo = @{
                DaysSincePasswordChange = "Error"
                PasswordExpired = "Error"
                PasswordExpiryDate = "Error"
            }
        }

        # Get account status - CORRIGÉ
        try {
            $AccountStatus = Get-AccountStatus -Enabled $User.Enabled -LastLogon $LastLogon -InactiveDays $InactiveDays -Created $User.Created -LockedOut $User.LockedOut
        }
        catch {
            Write-Warning "Could not determine account status for $($User.SamAccountName): $($_.Exception.Message)"
            $AccountStatus = "Unknown"
        }

        # Get group memberships
        try {
            $GroupMemberships = Get-UserGroups -UserDistinguishedName $User.DistinguishedName
        }
        catch {
            Write-Warning "Could not retrieve groups for $($User.SamAccountName): $($_.Exception.Message)"
            $GroupMemberships = "Error retrieving groups"
        }

        # Get manager name
        $ManagerName = if ($User.Manager) {
            try {
                (Get-ADUser -Identity $User.Manager -ErrorAction SilentlyContinue).DisplayName
            } catch {
                "Unknown"
            }
        } else {
            "None"
        }

        # Calculate days since creation - CORRIGÉ
        $DaysSinceCreation = if ($User.Created) {
            (New-TimeSpan -Start $User.Created -End (Get-Date)).Days
        } else {
            "Unknown"
        }

        # Calculate days since last logon - CORRIGÉ
        $DaysSinceLastLogon = if ($LastLogon -and $LastLogon -ne [DateTime]::MinValue) {
            (New-TimeSpan -Start $LastLogon -End (Get-Date)).Days
        } else {
            "Never"
        }

        # Prepare report object with safe value handling
        $UserReport = [PSCustomObject]@{
            "Sam Account Name" = if ($User.SamAccountName) { $User.SamAccountName } else { "Unknown" }
            "Display Name" = if ($User.DisplayName) { $User.DisplayName } else { "Unknown" }
            "First Name" = if ($User.GivenName) { $User.GivenName } else { "N/A" }
            "Last Name" = if ($User.Surname) { $User.Surname } else { "N/A" }
            "User Principal Name" = if ($User.UserPrincipalName) { $User.UserPrincipalName } else { "N/A" }
            "Email Address" = if ($User.EmailAddress) { $User.EmailAddress } else { "N/A" }
            "Employee ID" = if ($User.EmployeeID) { $User.EmployeeID } else { "N/A" }
            "Description" = if ($User.Description) { $User.Description } else { "N/A" }
            "Department" = if ($User.Department) { $User.Department } else { "N/A" }
            "Title" = if ($User.Title) { $User.Title } else { "N/A" }
            "Company" = if ($User.Company) { $User.Company } else { "N/A" }
            "Office" = if ($User.Office) { $User.Office } else { "N/A" }
            "Office Phone" = if ($User.OfficePhone) { $User.OfficePhone } else { "N/A" }
            "Mobile Phone" = if ($User.MobilePhone) { $User.MobilePhone } else { "N/A" }
            "Manager" = $ManagerName
            "Account Created" = if ($User.Created) { $User.Created.ToString("yyyy-MM-dd HH:mm:ss") } else { "Unknown" }
            "Days Since Created" = $DaysSinceCreation
            "Last Logon Date" = if ($LastLogon -and $LastLogon -ne [DateTime]::MinValue) { $LastLogon.ToString("yyyy-MM-dd HH:mm:ss") } else { "Never" }
            "Days Since Last Logon" = $DaysSinceLastLogon
            "Password Last Set" = if ($User.PasswordLastSet) { $User.PasswordLastSet.ToString("yyyy-MM-dd HH:mm:ss") } else { "Never" }
            "Days Since Password Change" = $PasswordInfo.DaysSincePasswordChange
            "Password Expiry Date" = $PasswordInfo.PasswordExpiryDate
            "Password Expired" = $PasswordInfo.PasswordExpired
            "Password Never Expires" = if ($null -ne $User.PasswordNeverExpires) { $User.PasswordNeverExpires } else { $false }
            "Cannot Change Password" = if ($null -ne $User.CannotChangePassword) { $User.CannotChangePassword } else { $false }
            "Account Status" = $AccountStatus
            "Account Enabled" = if ($null -ne $User.Enabled) { $User.Enabled } else { $false }
            "Account Locked Out" = if ($null -ne $User.LockedOut) { $User.LockedOut } else { $false }
            "Lockout Time" = if ($User.AccountLockoutTime) { $User.AccountLockoutTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "Never" }
            "Logon Count" = if ($null -ne $User.LogonCount) { $User.LogonCount } else { 0 }
            "Bad Logon Count" = if ($null -ne $User.BadLogonCount) { $User.BadLogonCount } else { 0 }
            "Group Memberships" = $GroupMemberships
            "Distinguished Name" = if ($User.DistinguishedName) { $User.DistinguishedName } else { "Unknown" }
        }

        $Report += $UserReport
    }

    # Clear progress bar
    if ($ShowProgress) {
        Write-Progress -Activity "Processing Users" -Completed
    }

    # Step 6: Generate statistics
    Write-Host "`nGenerating report statistics..." -ForegroundColor Yellow

    $Stats = @{
        TotalUsers = $Report.Count
        EnabledUsers = ($Report | Where-Object { $_.'Account Enabled' -eq $true }).Count
        DisabledUsers = ($Report | Where-Object { $_.'Account Enabled' -eq $false }).Count
        LockedOutUsers = ($Report | Where-Object { $_.'Account Locked Out' -eq $true }).Count
        NeverLoggedIn = ($Report | Where-Object { $_.'Last Logon Date' -eq "Never" }).Count
        InactiveUsers = ($Report | Where-Object { $_.'Account Status' -like "Inactive*" }).Count
        ExpiredPasswords = ($Report | Where-Object { $_.'Password Expired' -eq $true }).Count
        PasswordNeverExpires = ($Report | Where-Object { $_.'Password Never Expires' -eq $true }).Count
    }

    # Step 7: Export report
    Write-Host "`nExporting report..." -ForegroundColor Yellow

    # Ensure export directory exists
    $ExportDir = Split-Path $ExportPath -Parent
    if (!(Test-Path $ExportDir)) {
        New-Item -ItemType Directory -Path $ExportDir -Force | Out-Null
    }

    $Report | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8

    # Step 8: Display summary
    Write-Host "`n=== AUDIT REPORT SUMMARY ===" -ForegroundColor Cyan
    Write-Host "Domain: $DomainName" -ForegroundColor White
    Write-Host "Search Base: $SearchBase" -ForegroundColor White
    Write-Host "Generated: $(Get-Date)" -ForegroundColor White
    Write-Host ""
    Write-Host "USER STATISTICS:" -ForegroundColor Yellow
    Write-Host "Total Users: $($Stats.TotalUsers)" -ForegroundColor White
    Write-Host "Enabled Users: $($Stats.EnabledUsers)" -ForegroundColor Green
    Write-Host "Disabled Users: $($Stats.DisabledUsers)" -ForegroundColor Gray
    Write-Host "Locked Out Users: $($Stats.LockedOutUsers)" -ForegroundColor Red
    Write-Host ""
    Write-Host "ACTIVITY ANALYSIS:" -ForegroundColor Yellow
    Write-Host "Never Logged In: $($Stats.NeverLoggedIn)" -ForegroundColor Yellow
    Write-Host "Inactive Users (>$InactiveDays days): $($Stats.InactiveUsers)" -ForegroundColor Red
    Write-Host ""
    Write-Host "PASSWORD SECURITY:" -ForegroundColor Yellow
    Write-Host "Expired Passwords: $($Stats.ExpiredPasswords)" -ForegroundColor Red
    Write-Host "Password Never Expires: $($Stats.PasswordNeverExpires)" -ForegroundColor Yellow

    # Calculate file size
    $FileInfo = Get-Item $ExportPath
    $FileSize = Format-FileSize -Size $FileInfo.Length

    Write-Host "`nREPORT OUTPUT:" -ForegroundColor Yellow
    Write-Host "File Path: $ExportPath" -ForegroundColor Cyan
    Write-Host "File Size: $FileSize" -ForegroundColor Cyan
    Write-Host "Columns: $($Report[0].PSObject.Properties.Count)" -ForegroundColor Cyan

    if ($CheckAllDCs) {
        Write-Host "`nNOTE: Real last logon times retrieved from all $($DomainControllers.Count) domain controllers." -ForegroundColor Green
    } else {
        Write-Host "`nNOTE: Last logon times from current DC only. Use -CheckAllDCs for most accurate data." -ForegroundColor Yellow
    }

    Write-Host "`nCompleted at: $(Get-Date)" -ForegroundColor Gray

    # Step 9: Offer to open report
    $OpenLocation = Read-Host "`nWould you like to open the report location? (Y/N)"
    if ($OpenLocation -match "^[Yy]") {
        Start-Process -FilePath "explorer.exe" -ArgumentList "/select,`"$ExportPath`""
    }

    Write-Host "`nActive Directory User Audit completed successfully!" -ForegroundColor Green
    Write-Host "Thank you for using ctrlaltnod.com scripts!" -ForegroundColor Cyan
}
catch {
    Write-Error "Script execution failed: $($_.Exception.Message)"
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}

# End of script
Write-Host "`nScript execution completed." -ForegroundColor Gray
