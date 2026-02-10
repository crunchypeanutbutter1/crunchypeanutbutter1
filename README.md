# Sysadmin Scripts & Projects

A collection of practical sysadmin scripts and automation utilities.

## PowerShell Scripts

### `powershell/Add-UserToGroups.ps1`
Copy AD group membership from one user to another with safer operational controls.

Highlights:
- Supports `-WhatIf` / `-Confirm` (safe change previews)
- Optional removal of target-only groups for strict parity
- Optional CSV reporting for audit/change tracking
- Handles both interactive prompts and non-interactive parameter usage

### `powershell/Get-MailboxSizeAudit.ps1`
Read-only Exchange Online mailbox audit and quota compliance report.

Highlights:
- Retry logic for transient failures and throttling scenarios
- Detailed mailbox/quota reporting with CSV output
- Summary section for over-quota and near-quota analysis
- Optional file logging

### `powershell/Get-ServerHealthSnapshot.ps1`
Collects quick local/remote Windows server health data in one pass.

Highlights:
- Uptime, memory usage, disk usage summaries
- Top memory-consuming processes
- Automatic services that are stopped
- Optional CSV export for daily operational snapshots
