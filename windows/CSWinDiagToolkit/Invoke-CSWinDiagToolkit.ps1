[CmdletBinding()]
param(
    [string]$ToolkitRoot = $PSScriptRoot,
    [string]$ProgramFilesSubfolder = 'AdminTools',
    [switch]$NoElevation,
    [switch]$SkipRun,
    [switch]$KeepProgramFilesCopy
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Step {
    param([string]$Message)
    Write-Host "`n=== $Message ===" -ForegroundColor Cyan
}

function Test-IsAdministrator {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Start-ElevatedSelf {
    param([string[]]$ForwardedArgs)

    $scriptPath = $MyInvocation.PSCommandPath
    $pwsh = (Get-Process -Id $PID).Path
    if (-not $pwsh) {
        $pwsh = 'powershell.exe'
    }

    $escapedArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $scriptPath)) + $ForwardedArgs
    $argList = $escapedArgs -join ' '

    Write-Host 'Not currently elevated. Requesting Administrator privileges...' -ForegroundColor Yellow
    Start-Process -FilePath $pwsh -Verb RunAs -ArgumentList $argList | Out-Null
}

function New-EasySpotFolder {
    param([string]$DesktopPath)

    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $folderName = "CSWinDiag-ReadyToSend-$env:COMPUTERNAME-$stamp"
    $path = Join-Path $DesktopPath $folderName
    New-Item -Path $path -ItemType Directory -Force | Out-Null
    return $path
}

function Copy-ToolkitToProgramFiles {
    param(
        [string]$SourceRoot,
        [string]$Subfolder
    )

    $targetRoot = Join-Path $env:ProgramFiles $Subfolder
    if (-not (Test-Path $targetRoot)) {
        New-Item -Path $targetRoot -ItemType Directory -Force | Out-Null
    }

    Copy-Item -Path (Join-Path $SourceRoot '*') -Destination $targetRoot -Recurse -Force
    return $targetRoot
}

function Get-LatestCsWinDiagZip {
    param([string]$SearchRoot)

    return Get-ChildItem -Path $SearchRoot -Filter 'CSWinDiag*.zip' -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
}

if (-not (Test-Path $ToolkitRoot)) {
    throw "ToolkitRoot does not exist: $ToolkitRoot"
}

$csWinDiagPath = Join-Path $ToolkitRoot 'cswindiag.exe'
if (-not (Test-Path $csWinDiagPath)) {
    throw "cswindiag.exe was not found in toolkit root: $ToolkitRoot"
}

if (-not (Test-IsAdministrator)) {
    if ($NoElevation) {
        throw 'Administrator privileges are required. Re-run from an elevated PowerShell session.'
    }

    $forwardArgs = @(
        ('-ToolkitRoot "{0}"' -f $ToolkitRoot),
        ('-ProgramFilesSubfolder "{0}"' -f $ProgramFilesSubfolder)
    )
    if ($SkipRun) { $forwardArgs += '-SkipRun' }
    if ($KeepProgramFilesCopy) { $forwardArgs += '-KeepProgramFilesCopy' }

    Start-ElevatedSelf -ForwardedArgs $forwardArgs
    return
}

Write-Step 'Preparing Program Files runtime location'
$programFilesRoot = Copy-ToolkitToProgramFiles -SourceRoot $ToolkitRoot -Subfolder $ProgramFilesSubfolder
$exePath = Join-Path $programFilesRoot 'cswindiag.exe'

if (-not (Test-Path $exePath)) {
    throw "Failed to stage cswindiag.exe in Program Files: $exePath"
}

if (-not $SkipRun) {
    Write-Step 'Running CSWinDiag (this can take ~3-4 minutes)'
    Push-Location $programFilesRoot
    try {
        & $exePath
    }
    finally {
        Pop-Location
    }
}
else {
    Write-Host 'SkipRun specified. CSWinDiag execution was skipped.' -ForegroundColor Yellow
}

Write-Step 'Collecting output package and creating desktop transfer folder'
$desktop = [Environment]::GetFolderPath('Desktop')
$outputFolder = New-EasySpotFolder -DesktopPath $desktop

$zip = Get-LatestCsWinDiagZip -SearchRoot $programFilesRoot
if (-not $zip) {
    Write-Warning "No CSWinDiag*.zip file found in $programFilesRoot"
}
else {
    Copy-Item -Path $zip.FullName -Destination $outputFolder -Force
}

$notesPath = Join-Path $outputFolder 'README-Transfer.txt'
@"
CSWinDiag collection transfer folder
===================================
Computer Name: $env:COMPUTERNAME
Collected On : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')
Source Folder: $programFilesRoot

Contents:
- CSWinDiag ZIP collection (if found)
- This README for chain-of-custody / transfer notes

Suggested workflow:
1) Copy this full folder to your flash drive.
2) Upload the CSWinDiag ZIP to CrowdStrike Support case portal.
3) Avoid emailing the ZIP directly.
"@ | Set-Content -Path $notesPath -Encoding UTF8

if (-not $KeepProgramFilesCopy) {
    Write-Step 'Cleaning staged toolkit from Program Files'
    Remove-Item -Path $programFilesRoot -Recurse -Force
}

Write-Step 'Done'
Write-Host "Desktop output folder: $outputFolder" -ForegroundColor Green
if ($zip) {
    Write-Host "Collected ZIP: $(Join-Path $outputFolder $zip.Name)" -ForegroundColor Green
}
