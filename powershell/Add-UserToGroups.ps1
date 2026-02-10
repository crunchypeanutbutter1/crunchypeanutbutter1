#Requires -Modules ActiveDirectory

<#
.SYNOPSIS
    Copies Active Directory group memberships from one user to another.

.DESCRIPTION
    Production-oriented helper for onboarding and access parity tasks.
    The script validates source/target users, safely skips primary groups,
    supports -WhatIf/-Confirm, optional removal of extra memberships,
    and can export an audit CSV for change tracking.

.PARAMETER SourceUser
    SamAccountName, UPN, DN, or GUID for the source user.

.PARAMETER TargetUser
    SamAccountName, UPN, DN, or GUID for the target user.

.PARAMETER IncludeDistributionGroups
    Include non-security groups when copying memberships.

.PARAMETER RemoveGroupsNotOnSource
    Remove target user from groups they are in that source user is not.

.PARAMETER ReportPath
    Optional path to write an action report as CSV.

.PARAMETER PassThru
    Return action objects to the pipeline.

.EXAMPLE
    .\Add-UserToGroups.ps1 -SourceUser jsmith -TargetUser jdoe -WhatIf

.EXAMPLE
    .\Add-UserToGroups.ps1 -SourceUser jsmith -TargetUser jdoe -RemoveGroupsNotOnSource -Confirm:$false
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $false)]
    [string]$SourceUser,

    [Parameter(Mandatory = $false)]
    [string]$TargetUser,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeDistributionGroups,

    [Parameter(Mandatory = $false)]
    [switch]$RemoveGroupsNotOnSource,

    [Parameter(Mandatory = $false)]
    [string]$ReportPath,

    [Parameter(Mandatory = $false)]
    [switch]$PassThru
)

$ErrorActionPreference = 'Stop'

function Resolve-User {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Identity)

    try {
        Get-ADUser -Identity $Identity -Properties MemberOf, SamAccountName, UserPrincipalName, DistinguishedName
    }
    catch {
        throw "Active Directory user '$Identity' was not found."
    }
}

function Get-ComparableGroupSet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$GroupDns,

        [Parameter(Mandatory = $false)]
        [switch]$AllowDistribution
    )

    $result = New-Object System.Collections.Generic.HashSet[string]

    foreach ($groupDn in $GroupDns) {
        try {
            $group = Get-ADGroup -Identity $groupDn -Properties GroupCategory, Name

            if (-not $AllowDistribution -and $group.GroupCategory -ne 'Security') {
                continue
            }

            [void]$result.Add($group.DistinguishedName)
        }
        catch {
            Write-Warning "Skipping unresolved group '$groupDn'. Error: $($_.Exception.Message)"
        }
    }

    return $result
}

if (-not $SourceUser) {
    $SourceUser = Read-Host 'Enter source user (copy memberships FROM)'
}

if (-not $TargetUser) {
    $TargetUser = Read-Host 'Enter target user (copy memberships TO)'
}

$source = Resolve-User -Identity $SourceUser
$target = Resolve-User -Identity $TargetUser

if ($source.DistinguishedName -eq $target.DistinguishedName) {
    throw 'Source and target are the same user. No action taken.'
}

$sourceGroups = Get-ComparableGroupSet -GroupDns @($source.MemberOf) -AllowDistribution:$IncludeDistributionGroups
$targetGroups = Get-ComparableGroupSet -GroupDns @($target.MemberOf) -AllowDistribution:$IncludeDistributionGroups

$groupsToAdd = $sourceGroups.Where({ -not $targetGroups.Contains($_) })
$groupsToRemove = @()

if ($RemoveGroupsNotOnSource) {
    $groupsToRemove = $targetGroups.Where({ -not $sourceGroups.Contains($_) })
}

Write-Host "Source user: $($source.SamAccountName)" -ForegroundColor Cyan
Write-Host "Target user: $($target.SamAccountName)" -ForegroundColor Cyan
Write-Host "Groups to add: $($groupsToAdd.Count)" -ForegroundColor Green
if ($RemoveGroupsNotOnSource) {
    Write-Host "Groups to remove: $($groupsToRemove.Count)" -ForegroundColor Yellow
}

$actionLog = New-Object System.Collections.Generic.List[object]

foreach ($groupDn in $groupsToAdd) {
    $group = Get-ADGroup -Identity $groupDn -Properties Name
    $targetName = "$($target.SamAccountName) -> $($group.Name)"

    if ($PSCmdlet.ShouldProcess($targetName, 'Add user to group')) {
        try {
            Add-ADGroupMember -Identity $groupDn -Members $target.DistinguishedName -ErrorAction Stop
            Write-Host "Added to group: $($group.Name)" -ForegroundColor Green

            $actionLog.Add([PSCustomObject]@{
                Timestamp = Get-Date
                Action    = 'Add'
                GroupName = $group.Name
                GroupDN   = $groupDn
                User      = $target.SamAccountName
                Result    = 'Success'
                Error     = $null
            })
        }
        catch {
            Write-Warning "Failed adding to '$($group.Name)': $($_.Exception.Message)"
            $actionLog.Add([PSCustomObject]@{
                Timestamp = Get-Date
                Action    = 'Add'
                GroupName = $group.Name
                GroupDN   = $groupDn
                User      = $target.SamAccountName
                Result    = 'Failed'
                Error     = $_.Exception.Message
            })
        }
    }
}

foreach ($groupDn in $groupsToRemove) {
    $group = Get-ADGroup -Identity $groupDn -Properties Name
    $targetName = "$($target.SamAccountName) <- $($group.Name)"

    if ($PSCmdlet.ShouldProcess($targetName, 'Remove user from group')) {
        try {
            Remove-ADGroupMember -Identity $groupDn -Members $target.DistinguishedName -Confirm:$false -ErrorAction Stop
            Write-Host "Removed from group: $($group.Name)" -ForegroundColor Yellow

            $actionLog.Add([PSCustomObject]@{
                Timestamp = Get-Date
                Action    = 'Remove'
                GroupName = $group.Name
                GroupDN   = $groupDn
                User      = $target.SamAccountName
                Result    = 'Success'
                Error     = $null
            })
        }
        catch {
            Write-Warning "Failed removing from '$($group.Name)': $($_.Exception.Message)"
            $actionLog.Add([PSCustomObject]@{
                Timestamp = Get-Date
                Action    = 'Remove'
                GroupName = $group.Name
                GroupDN   = $groupDn
                User      = $target.SamAccountName
                Result    = 'Failed'
                Error     = $_.Exception.Message
            })
        }
    }
}

if ($ReportPath) {
    $parent = Split-Path -Path $ReportPath -Parent
    if ($parent -and -not (Test-Path -Path $parent -PathType Container)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }

    $actionLog | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
    Write-Host "Audit report exported to: $ReportPath" -ForegroundColor Cyan
}

if ($PassThru) {
    $actionLog
}

Write-Host 'Completed group membership sync.' -ForegroundColor Green
