<#
.SYNOPSIS
    Audits Exchange Online mailbox sizes and quotas, generating a detailed compliance report.
    
.DESCRIPTION
    Save this script as: Get-MailboxSizeAudit.ps1
    
    QUICK START:
    1. Save this entire script as Get-MailboxSizeAudit.ps1
    2. Open PowerShell as Administrator
    3. Run: .\Get-MailboxSizeAudit.ps1
    4. Sign in when prompted
    5. Find your CSV report in the current directory

.DESCRIPTION
    Audits Exchange Online mailbox sizes and quotas, generating a detailed compliance report.

.DESCRIPTION
    Production-ready script for auditing Exchange Online mailbox usage, quota consumption,
    and archive status. Designed for large tenants with robust error handling, retry logic,
    and comprehensive reporting. READ-ONLY by default.

.PARAMETER OutputFolder
    Directory path for CSV report output. Defaults to current directory.

.PARAMETER IncludeSharedMailboxes
    Include shared, room, and equipment mailboxes in the audit.

.PARAMETER LogPath
    Path to write detailed log file. If not specified, logs to console only.

.PARAMETER MaxRetries
    Maximum number of retry attempts for transient failures. Default: 3

.PARAMETER RetryDelaySeconds
    Base delay in seconds between retries (uses exponential backoff). Default: 5

.EXAMPLE
    .\Get-MailboxSizeAudit.ps1
    Runs audit for user mailboxes only, outputs CSV to current directory.

.EXAMPLE
    .\Get-MailboxSizeAudit.ps1 -OutputFolder "C:\Reports" -IncludeSharedMailboxes -Verbose
    Audits all mailbox types with verbose logging.

.EXAMPLE
    .\Get-MailboxSizeAudit.ps1 -LogPath "C:\Logs\audit.log" -OutputFolder "C:\Reports"
    Runs audit with file logging enabled.

.NOTES
    Author: Senior M365 SysAdmin
    Version: 1.0.0
    Requires: ExchangeOnlineManagement module
    
    REQUIRED PERMISSIONS:
    - Exchange Online: View-Only Recipients role OR
    - Exchange Online: Mail Recipients role (read-only access)
    - Azure AD: Global Reader OR Exchange Administrator
    
    Minimum required cmdlet permissions:
    - Get-Mailbox
    - Get-MailboxStatistics
    - Get-EXOMailbox (for performance)
    - Get-EXOMailboxStatistics (for performance)

.LINK
    https://learn.microsoft.com/powershell/module/exchange/

#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateScript({
        if (-not (Test-Path -Path $_ -PathType Container)) {
            throw "Output folder does not exist: $_"
        }
        $true
    })]
    [string]$OutputFolder = (Get-Location).Path,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeSharedMailboxes,

    [Parameter(Mandatory = $false)]
    [string]$LogPath,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 10)]
    [int]$MaxRetries = 3,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 60)]
    [int]$RetryDelaySeconds = 5
)

#Requires -Version 5.1

# ============================================================================
# SCRIPT CONFIGURATION
# ============================================================================

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'  # Speeds up cmdlet execution significantly

$script:LogFile = $LogPath
$script:StartTime = Get-Date
$script:ExitCode = 0

# ============================================================================
# LOGGING FUNCTIONS
# ============================================================================

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('INFO', 'WARNING', 'ERROR', 'SUCCESS', 'VERBOSE')]
        [string]$Level = 'INFO'
    )
    
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Console output with color coding
    switch ($Level) {
        'ERROR'   { Write-Host $logMessage -ForegroundColor Red }
        'WARNING' { Write-Host $logMessage -ForegroundColor Yellow }
        'SUCCESS' { Write-Host $logMessage -ForegroundColor Green }
        'VERBOSE' { if ($VerbosePreference -ne 'SilentlyContinue') { Write-Host $logMessage -ForegroundColor Cyan } }
        default   { Write-Host $logMessage }
    }
    
    # File logging
    if ($script:LogFile) {
        try {
            Add-Content -Path $script:LogFile -Value $logMessage -ErrorAction Stop
        }
        catch {
            Write-Warning "Failed to write to log file: $_"
        }
    }
}

# ============================================================================
# MODULE MANAGEMENT
# ============================================================================

function Test-RequiredModule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ModuleName
    )
    
    Write-Log "Checking for required module: $ModuleName" -Level VERBOSE
    
    $module = Get-Module -Name $ModuleName -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
    
    if (-not $module) {
        Write-Log "Module '$ModuleName' is not installed." -Level WARNING
        
        $install = Read-Host "Would you like to install '$ModuleName' from PSGallery? (Y/N)"
        if ($install -eq 'Y' -or $install -eq 'y') {
            try {
                Write-Log "Installing module '$ModuleName'..." -Level INFO
                Install-Module -Name $ModuleName -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                Write-Log "Module '$ModuleName' installed successfully." -Level SUCCESS
                return $true
            }
            catch {
                Write-Log "Failed to install module '$ModuleName': $_" -Level ERROR
                return $false
            }
        }
        else {
            Write-Log "Module installation declined. Cannot proceed." -Level ERROR
            return $false
        }
    }
    else {
        Write-Log "Module '$ModuleName' version $($module.Version) is available." -Level VERBOSE
        return $true
    }
}

function Connect-ExchangeOnlineWithRetry {
    [CmdletBinding()]
    param(
        [int]$MaxAttempts = 3
    )
    
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            Write-Log "Attempting to connect to Exchange Online (Attempt $attempt of $MaxAttempts)..." -Level INFO
            
            # Check if already connected
            $existingSession = Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Connected' }
            
            if ($existingSession) {
                Write-Log "Already connected to Exchange Online tenant: $($existingSession.TenantId)" -Level SUCCESS
                return $true
            }
            
            # Connect with modern authentication
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
            
            Write-Log "Successfully connected to Exchange Online." -Level SUCCESS
            return $true
        }
        catch {
            Write-Log "Connection attempt $attempt failed: $_" -Level WARNING
            
            if ($attempt -lt $MaxAttempts) {
                $delay = $RetryDelaySeconds * [Math]::Pow(2, $attempt - 1)
                Write-Log "Retrying in $delay seconds..." -Level INFO
                Start-Sleep -Seconds $delay
            }
            else {
                Write-Log "Failed to connect to Exchange Online after $MaxAttempts attempts." -Level ERROR
                return $false
            }
        }
    }
    
    return $false
}

# ============================================================================
# DATA RETRIEVAL WITH RETRY LOGIC
# ============================================================================

function Invoke-WithRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,
        
        [Parameter(Mandatory = $false)]
        [string]$OperationName = "Operation",
        
        [Parameter(Mandatory = $false)]
        [int]$MaxAttempts = $MaxRetries,
        
        [Parameter(Mandatory = $false)]
        [int]$BaseDelay = $RetryDelaySeconds
    )
    
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            return & $ScriptBlock
        }
        catch {
            $errorMessage = $_.Exception.Message
            
            # Check for throttling or transient errors
            $isTransient = $errorMessage -match 'throttl|timeout|service unavailable|temporarily unavailable|503|429'
            
            if ($attempt -lt $MaxAttempts -and $isTransient) {
                $delay = $BaseDelay * [Math]::Pow(2, $attempt - 1)
                Write-Log "$OperationName failed (attempt $attempt): $errorMessage. Retrying in $delay seconds..." -Level WARNING
                Start-Sleep -Seconds $delay
            }
            else {
                Write-Log "$OperationName failed: $errorMessage" -Level ERROR
                throw
            }
        }
    }
}

# ============================================================================
# MAILBOX DATA PROCESSING
# ============================================================================

function Get-MailboxData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [switch]$IncludeShared
    )
    
    Write-Log "Retrieving mailbox list..." -Level INFO
    
    try {
        # Determine recipient types to include
        $recipientTypes = @('UserMailbox')
        if ($IncludeShared) {
            $recipientTypes += @('SharedMailbox', 'RoomMailbox', 'EquipmentMailbox')
            Write-Log "Including shared, room, and equipment mailboxes in audit." -Level INFO
        }
        
        # Use Get-EXOMailbox for better performance on large tenants
        $mailboxes = Invoke-WithRetry -OperationName "Get mailbox list" -ScriptBlock {
            Get-EXOMailbox -ResultSize Unlimited -Properties DisplayName, UserPrincipalName, RecipientTypeDetails, `
                ProhibitSendQuota, ProhibitSendReceiveQuota, ArchiveStatus, ArchiveQuota | 
                Where-Object { $_.RecipientTypeDetails -in $recipientTypes }
        }
        
        if (-not $mailboxes) {
            Write-Log "No mailboxes found matching criteria." -Level WARNING
            return @()
        }
        
        $mailboxCount = ($mailboxes | Measure-Object).Count
        Write-Log "Found $mailboxCount mailbox(es) to audit." -Level SUCCESS
        
        return $mailboxes
    }
    catch {
        Write-Log "Failed to retrieve mailbox list: $_" -Level ERROR
        throw
    }
}

function Get-MailboxSizeData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Mailbox
    )
    
    try {
        # Get mailbox statistics
        $stats = Invoke-WithRetry -OperationName "Get statistics for $($Mailbox.UserPrincipalName)" -ScriptBlock {
            Get-EXOMailboxStatistics -Identity $Mailbox.UserPrincipalName -ErrorAction Stop
        }
        
        # Parse total item size (format: "123.4 MB (129,499,136 bytes)")
        $sizeInBytes = 0
        if ($stats.TotalItemSize) {
            $sizeString = $stats.TotalItemSize.ToString()
            if ($sizeString -match '\(([0-9,]+) bytes\)') {
                $sizeInBytes = [long]($matches[1] -replace ',', '')
            }
        }
        $sizeInGB = [math]::Round($sizeInBytes / 1GB, 2)
        
        # Parse quota values
        $prohibitSendQuotaGB = Get-QuotaInGB -QuotaValue $Mailbox.ProhibitSendQuota
        $prohibitSendReceiveQuotaGB = Get-QuotaInGB -QuotaValue $Mailbox.ProhibitSendReceiveQuota
        
        # Determine quota status
        $quotaStatus = Get-QuotaStatus -CurrentSizeGB $sizeInGB -ProhibitSendQuotaGB $prohibitSendQuotaGB
        $quotaPercentage = Get-QuotaPercentage -CurrentSizeGB $sizeInGB -ProhibitSendQuotaGB $prohibitSendQuotaGB
        
        # Archive information
        $archiveStatus = if ($Mailbox.ArchiveStatus -eq 'Active') { 'Enabled' } else { 'Disabled' }
        $archiveQuotaGB = if ($Mailbox.ArchiveQuota) { Get-QuotaInGB -QuotaValue $Mailbox.ArchiveQuota } else { 'N/A' }
        
        # Last logon time
        $lastLogonTime = if ($stats.LastLogonTime) { $stats.LastLogonTime.ToString('yyyy-MM-dd HH:mm:ss') } else { 'Never' }
        
        return [PSCustomObject]@{
            DisplayName                  = $Mailbox.DisplayName
            UserPrincipalName            = $Mailbox.UserPrincipalName
            RecipientTypeDetails         = $Mailbox.RecipientTypeDetails
            TotalItemSizeGB              = $sizeInGB
            ItemCount                    = $stats.ItemCount
            ProhibitSendQuotaGB          = $prohibitSendQuotaGB
            ProhibitSendReceiveQuotaGB   = $prohibitSendReceiveQuotaGB
            QuotaUsagePercentage         = $quotaPercentage
            QuotaStatus                  = $quotaStatus
            ArchiveStatus                = $archiveStatus
            ArchiveQuotaGB               = $archiveQuotaGB
            LastLogonTime                = $lastLogonTime
        }
    }
    catch {
        Write-Log "Failed to retrieve statistics for $($Mailbox.UserPrincipalName): $_" -Level ERROR
        
        # Return partial data on error
        return [PSCustomObject]@{
            DisplayName                  = $Mailbox.DisplayName
            UserPrincipalName            = $Mailbox.UserPrincipalName
            RecipientTypeDetails         = $Mailbox.RecipientTypeDetails
            TotalItemSizeGB              = 'ERROR'
            ItemCount                    = 'ERROR'
            ProhibitSendQuotaGB          = Get-QuotaInGB -QuotaValue $Mailbox.ProhibitSendQuota
            ProhibitSendReceiveQuotaGB   = Get-QuotaInGB -QuotaValue $Mailbox.ProhibitSendReceiveQuota
            QuotaUsagePercentage         = 'ERROR'
            QuotaStatus                  = 'ERROR'
            ArchiveStatus                = if ($Mailbox.ArchiveStatus -eq 'Active') { 'Enabled' } else { 'Disabled' }
            ArchiveQuotaGB               = if ($Mailbox.ArchiveQuota) { Get-QuotaInGB -QuotaValue $Mailbox.ArchiveQuota } else { 'N/A' }
            LastLogonTime                = 'ERROR'
        }
    }
}

function Get-QuotaInGB {
    param([string]$QuotaValue)
    
    if ([string]::IsNullOrWhiteSpace($QuotaValue) -or $QuotaValue -eq 'Unlimited') {
        return 'Unlimited'
    }
    
    # Parse quota string (e.g., "50 GB (53,687,091,200 bytes)")
    if ($QuotaValue -match '\(([0-9,]+) bytes\)') {
        $bytes = [long]($matches[1] -replace ',', '')
        return [math]::Round($bytes / 1GB, 2)
    }
    
    return $QuotaValue
}

function Get-QuotaPercentage {
    param(
        [double]$CurrentSizeGB,
        $ProhibitSendQuotaGB
    )
    
    if ($ProhibitSendQuotaGB -eq 'Unlimited' -or $ProhibitSendQuotaGB -eq 0) {
        return 0
    }
    
    try {
        $percentage = [math]::Round(($CurrentSizeGB / [double]$ProhibitSendQuotaGB) * 100, 2)
        return $percentage
    }
    catch {
        return 0
    }
}

function Get-QuotaStatus {
    param(
        [double]$CurrentSizeGB,
        $ProhibitSendQuotaGB
    )
    
    if ($ProhibitSendQuotaGB -eq 'Unlimited') {
        return 'Healthy'
    }
    
    $percentage = Get-QuotaPercentage -CurrentSizeGB $CurrentSizeGB -ProhibitSendQuotaGB $ProhibitSendQuotaGB
    
    if ($percentage -ge 100) {
        return 'OverQuota'
    }
    elseif ($percentage -ge 90) {
        return 'NearQuota'
    }
    else {
        return 'Healthy'
    }
}

# ============================================================================
# REPORT GENERATION
# ============================================================================

function Export-MailboxReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$MailboxData,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )
    
    try {
        Write-Log "Exporting report to: $OutputPath" -Level INFO
        
        $MailboxData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        
        Write-Log "Report exported successfully." -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "Failed to export report: $_" -Level ERROR
        return $false
    }
}

function Show-AuditSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$MailboxData
    )
    
    Write-Log "`n========================================" -Level INFO
    Write-Log "MAILBOX AUDIT SUMMARY" -Level INFO
    Write-Log "========================================" -Level INFO
    
    $total = ($MailboxData | Measure-Object).Count
    $overQuota = ($MailboxData | Where-Object { $_.QuotaStatus -eq 'OverQuota' } | Measure-Object).Count
    $nearQuota = ($MailboxData | Where-Object { $_.QuotaStatus -eq 'NearQuota' } | Measure-Object).Count
    $healthy = ($MailboxData | Where-Object { $_.QuotaStatus -eq 'Healthy' } | Measure-Object).Count
    $errors = ($MailboxData | Where-Object { $_.QuotaStatus -eq 'ERROR' } | Measure-Object).Count
    
    $archiveEnabled = ($MailboxData | Where-Object { $_.ArchiveStatus -eq 'Enabled' } | Measure-Object).Count
    
    Write-Log "Total Mailboxes Audited: $total" -Level INFO
    Write-Log "  - Over Quota (>=100%): $overQuota" -Level $(if ($overQuota -gt 0) { 'ERROR' } else { 'INFO' })
    Write-Log "  - Near Quota (>=90%): $nearQuota" -Level $(if ($nearQuota -gt 0) { 'WARNING' } else { 'INFO' })
    Write-Log "  - Healthy (<90%): $healthy" -Level SUCCESS
    Write-Log "  - Errors During Collection: $errors" -Level $(if ($errors -gt 0) { 'ERROR' } else { 'INFO' })
    Write-Log "" -Level INFO
    Write-Log "Archive Mailboxes Enabled: $archiveEnabled" -Level INFO
    Write-Log "========================================`n" -Level INFO
}

# ============================================================================
# MAIN EXECUTION
# ============================================================================

function Main {
    try {
        Write-Log "========================================" -Level INFO
        Write-Log "Exchange Online Mailbox Audit Script" -Level INFO
        Write-Log "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Level INFO
        Write-Log "========================================`n" -Level INFO
        
        # Validate and load required module
        if (-not (Test-RequiredModule -ModuleName 'ExchangeOnlineManagement')) {
            Write-Log "Required module not available. Exiting." -Level ERROR
            $script:ExitCode = 1
            return
        }
        
        # Import module
        try {
            Import-Module ExchangeOnlineManagement -ErrorAction Stop
            Write-Log "ExchangeOnlineManagement module imported successfully." -Level VERBOSE
        }
        catch {
            Write-Log "Failed to import ExchangeOnlineManagement module: $_" -Level ERROR
            $script:ExitCode = 2
            return
        }
        
        # Connect to Exchange Online
        if (-not (Connect-ExchangeOnlineWithRetry -MaxAttempts $MaxRetries)) {
            Write-Log "Unable to establish Exchange Online connection. Exiting." -Level ERROR
            $script:ExitCode = 3
            return
        }
        
        # Retrieve mailbox list
        $mailboxes = Get-MailboxData -IncludeShared:$IncludeSharedMailboxes
        
        if ($mailboxes.Count -eq 0) {
            Write-Log "No mailboxes to process. Exiting." -Level WARNING
            $script:ExitCode = 0
            return
        }
        
        # Process each mailbox
        Write-Log "Processing mailbox statistics..." -Level INFO
        $results = @()
        $counter = 0
        
        foreach ($mailbox in $mailboxes) {
            $counter++
            Write-Progress -Activity "Auditing Mailboxes" -Status "Processing $counter of $($mailboxes.Count): $($mailbox.UserPrincipalName)" -PercentComplete (($counter / $mailboxes.Count) * 100)
            
            Write-Log "Processing [$counter/$($mailboxes.Count)]: $($mailbox.UserPrincipalName)" -Level VERBOSE
            
            $mailboxData = Get-MailboxSizeData -Mailbox $mailbox
            $results += $mailboxData
        }
        
        Write-Progress -Activity "Auditing Mailboxes" -Completed
        
        # Generate report filename
        $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $reportFileName = "MailboxAudit_$timestamp.csv"
        $reportPath = Join-Path -Path $OutputFolder -ChildPath $reportFileName
        
        # Export report
        if (Export-MailboxReport -MailboxData $results -OutputPath $reportPath) {
            Write-Log "Report file: $reportPath" -Level SUCCESS
        }
        else {
            $script:ExitCode = 4
        }
        
        # Display summary
        Show-AuditSummary -MailboxData $results
        
        # Cleanup
        Write-Log "Disconnecting from Exchange Online..." -Level VERBOSE
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        
        $duration = (Get-Date) - $script:StartTime
        Write-Log "Audit completed in $([math]::Round($duration.TotalMinutes, 2)) minutes." -Level SUCCESS
        
    }
    catch {
        Write-Log "Critical error in main execution: $_" -Level ERROR
        Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Level ERROR
        $script:ExitCode = 99
    }
    finally {
        # Ensure cleanup
        try {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        }
        catch {
            # Suppress disconnect errors
        }
        
        Write-Log "`nScript exiting with code: $script:ExitCode" -Level INFO
        exit $script:ExitCode
    }
}

# ============================================================================
# SCRIPT ENTRY POINT
# ============================================================================

Main
