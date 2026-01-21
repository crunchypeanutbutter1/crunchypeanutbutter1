<#
.SYNOPSIS
    Adds a user to all groups that another user is a member of.

.DESCRIPTION
    This script gets all AD groups from a source user and adds a target user to those groups.
    It validates that both users exist before executing the operations.

.EXAMPLE
    .\Add-UserToGroups.ps1
#>

# Function to validate if a user exists in Active Directory
function Test-ADUserExists {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Identity
    )
    
    try {
        $user = Get-ADUser -Identity $Identity -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

# Prompt for source user
do {
    $sourceUser = Read-Host "Enter the source user (the user whose groups will be copied FROM)"
    
    if (-not (Test-ADUserExists -Identity $sourceUser)) {
        Write-Host "User '$sourceUser' not found in Active Directory. Please try again." -ForegroundColor Red
    }
} while (-not (Test-ADUserExists -Identity $sourceUser))

Write-Host "Source user '$sourceUser' found." -ForegroundColor Green

# Prompt for target user
do {
    $targetUser = Read-Host "Enter the target user (the user whose will be added TO the groups)"
    
    if (-not (Test-ADUserExists -Identity $targetUser)) {
        Write-Host "User '$targetUser' not found in Active Directory. Please try again." -ForegroundColor Red
    }
} while (-not (Test-ADUserExists -Identity $targetUser))

Write-Host "Target user '$targetUser' found." -ForegroundColor Green

# Confirm operation
Write-Host ""
Write-Host "This will add '$targetUser' to all groups that '$sourceUser' is a member of." -ForegroundColor Yellow
$confirm = Read-Host "Do you want to continue? (Y/N)"

if ($confirm -ne 'Y' -and $confirm -ne 'y') {
    Write-Host "Operation cancelled." -ForegroundColor Yellow
    exit
}

# Get the groups from the source user
try {
    Write-Host "Retrieving groups for '$sourceUser'..." -ForegroundColor Cyan
    $getUserGroups = Get-ADUser -Identity $sourceUser -Properties MemberOf -ErrorAction Stop | Select-Object -ExpandProperty MemberOf
    
    if ($getUserGroups.Count -eq 0) {
        Write-Host "User '$sourceUser' is not a member of any groups." -ForegroundColor Yellow
        exit
    }
    
    Write-Host "Found $($getUserGroups.Count) group(s) to add '$targetUser' to." -ForegroundColor Green
    Write-Host ""
    
    # Add target user to each group
    $getUserGroups | Add-ADGroupMember -Members $targetUser -Verbose
    
    Write-Host ""
    Write-Host "Successfully added '$targetUser' to all groups!" -ForegroundColor Green
}
catch {
    Write-Host "An error occurred: $_" -ForegroundColor Red
    exit 1
}
