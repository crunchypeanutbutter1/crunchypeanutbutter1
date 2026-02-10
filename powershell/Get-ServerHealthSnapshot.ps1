<#
.SYNOPSIS
    Collects a quick server health snapshot for local or remote Windows hosts.

.DESCRIPTION
    Gathers uptime, CPU utilization, memory usage, disk capacity, top processes,
    and stopped automatic services in one report. Designed as a fast triage tool
    for sysadmins doing reactive troubleshooting or daily checks.

.PARAMETER ComputerName
    One or more remote computers. Defaults to localhost.

.PARAMETER Credential
    Credential used for remote CIM/PowerShell calls.

.PARAMETER TopProcesses
    Number of top processes by working set memory to return. Default: 5.

.PARAMETER OutputPath
    Optional path to export flattened snapshot data to CSV.

.EXAMPLE
    .\Get-ServerHealthSnapshot.ps1

.EXAMPLE
    .\Get-ServerHealthSnapshot.ps1 -ComputerName FS01,APP01 -Credential (Get-Credential) -OutputPath C:\Reports\health.csv
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string[]]$ComputerName = @('localhost'),

    [Parameter(Mandatory = $false)]
    [pscredential]$Credential,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 25)]
    [int]$TopProcesses = 5,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath
)

function New-CimSessionSafe {
    param(
        [string]$Name,
        [pscredential]$Cred
    )

    try {
        if ($Name -eq 'localhost' -or $Name -eq $env:COMPUTERNAME) {
            return $null
        }

        if ($Cred) {
            return New-CimSession -ComputerName $Name -Credential $Cred -ErrorAction Stop
        }

        return New-CimSession -ComputerName $Name -ErrorAction Stop
    }
    catch {
        throw "Failed to create CIM session to '$Name'. $($_.Exception.Message)"
    }
}

$allSnapshots = New-Object System.Collections.Generic.List[object]

foreach ($computer in $ComputerName) {
    Write-Verbose "Collecting health data from $computer"
    $session = $null

    try {
        $session = New-CimSessionSafe -Name $computer -Cred $Credential

        $os = Get-CimInstance -ClassName Win32_OperatingSystem -CimSession $session -ErrorAction Stop
        $cpu = Get-CimInstance -ClassName Win32_Processor -CimSession $session -ErrorAction Stop
        $disks = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType = 3" -CimSession $session -ErrorAction Stop

        $boot = [Management.ManagementDateTimeConverter]::ToDateTime($os.LastBootUpTime)
        $uptime = (Get-Date) - $boot

        $totalMemGB = [math]::Round($os.TotalVisibleMemorySize / 1MB, 2)
        $freeMemGB = [math]::Round($os.FreePhysicalMemory / 1MB, 2)
        $usedMemGB = [math]::Round($totalMemGB - $freeMemGB, 2)
        $memPctUsed = if ($totalMemGB -gt 0) { [math]::Round(($usedMemGB / $totalMemGB) * 100, 2) } else { 0 }

        $diskSummary = $disks | ForEach-Object {
            $sizeGB = [math]::Round($_.Size / 1GB, 2)
            $freeGB = [math]::Round($_.FreeSpace / 1GB, 2)
            $usedPct = if ($_.Size -gt 0) { [math]::Round((($_.Size - $_.FreeSpace) / $_.Size) * 100, 2) } else { 0 }
            "{0}: {1}GB free / {2}GB total ({3}% used)" -f $_.DeviceID, $freeGB, $sizeGB, $usedPct
        }

        $processScript = {
            param($Count)
            Get-Process |
                Sort-Object -Property WorkingSet64 -Descending |
                Select-Object -First $Count -Property ProcessName, Id,
                    @{Name='WorkingSetMB';Expression={[math]::Round($_.WorkingSet64 / 1MB, 2)}}, CPU
        }

        $serviceScript = {
            Get-Service |
                Where-Object { $_.StartType -eq 'Automatic' -and $_.Status -ne 'Running' } |
                Select-Object -Property Name, DisplayName, Status
        }

        if ($session) {
            $topProc = Invoke-Command -ComputerName $computer -Credential $Credential -ScriptBlock $processScript -ArgumentList $TopProcesses -ErrorAction Stop
            $stoppedAuto = Invoke-Command -ComputerName $computer -Credential $Credential -ScriptBlock $serviceScript -ErrorAction Stop
        }
        else {
            $topProc = & $processScript $TopProcesses
            $stoppedAuto = & $serviceScript
        }

        $snapshot = [PSCustomObject]@{
            ComputerName              = $computer
            CollectedAt               = Get-Date
            OS                        = $os.Caption
            LastBootTime              = $boot
            UptimeDays                = [math]::Round($uptime.TotalDays, 2)
            CPUCount                  = ($cpu | Measure-Object).Count
            LogicalProcessors         = ($cpu | Measure-Object -Property NumberOfLogicalProcessors -Sum).Sum
            MemoryTotalGB             = $totalMemGB
            MemoryUsedGB              = $usedMemGB
            MemoryUsedPercent         = $memPctUsed
            DiskSummary               = ($diskSummary -join '; ')
            StoppedAutomaticServices  = ($stoppedAuto.Name -join ', ')
            TopProcessesByMemory      = ($topProc | ForEach-Object { "{0}({1}) {2}MB" -f $_.ProcessName, $_.Id, $_.WorkingSetMB }) -join '; '
        }

        $allSnapshots.Add($snapshot)
    }
    catch {
        Write-Warning "Failed to collect from $computer: $($_.Exception.Message)"
        $allSnapshots.Add([PSCustomObject]@{
            ComputerName              = $computer
            CollectedAt               = Get-Date
            OS                        = 'ERROR'
            LastBootTime              = $null
            UptimeDays                = $null
            CPUCount                  = $null
            LogicalProcessors         = $null
            MemoryTotalGB             = $null
            MemoryUsedGB              = $null
            MemoryUsedPercent         = $null
            DiskSummary               = 'ERROR'
            StoppedAutomaticServices  = 'ERROR'
            TopProcessesByMemory      = 'ERROR'
        })
    }
    finally {
        if ($session) {
            $session | Remove-CimSession -ErrorAction SilentlyContinue
        }
    }
}

if ($OutputPath) {
    $parent = Split-Path -Path $OutputPath -Parent
    if ($parent -and -not (Test-Path -Path $parent -PathType Container)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }

    $allSnapshots | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Verbose "Exported snapshot to $OutputPath"
}

$allSnapshots
