#requires -version 3.0
<#
.SYNOPSIS
    Checks Windows/Citrix terminal server health via PowerShell Remoting.
.DESCRIPTION
    Reads target servers from config/servers.txt and settings from config/settings.json,
    validates WinRM reachability, captures CPU/RAM/session/process data and writes
    daily raw CSV files, per-run summaries and log files.
#>
[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'config'),
    [string]$OutputPath = (Join-Path $PSScriptRoot 'output')
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Get-TimeStamp { (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') }

function Ensure-Directory {
    param([Parameter(Mandatory=$true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
}

function Write-Log {
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR')][string]$Level = 'INFO',
        [Parameter(Mandatory=$true)][string]$LogPath
    )
    $line = '{0};{1};{2}' -f (Get-TimeStamp), $Level, $Message
    Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
    if ($Level -eq 'ERROR') { Write-Warning $Message } else { Write-Verbose $Message }
}

function Read-Settings {
    param([Parameter(Mandatory=$true)][string]$Path)
    $defaults = [pscustomobject]@{
        CpuSampleSeconds = 5
        TopProcessCount = 10
        WinRMTimeoutSeconds = 5
        OutputDelimiter = ';'
        IncludeDisconnectedSessions = $true
    }
    if (-not (Test-Path -LiteralPath $Path)) { return $defaults }
    $json = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($property in $defaults.PSObject.Properties.Name) {
        if ($null -eq $json.$property) { $json | Add-Member -NotePropertyName $property -NotePropertyValue $defaults.$property }
    }
    return $json
}

function Read-ServerList {
    param([Parameter(Mandatory=$true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { throw "Serverliste nicht gefunden: $Path" }
    Get-Content -LiteralPath $Path -Encoding UTF8 |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ -and -not $_.StartsWith('#') } |
        Select-Object -Unique
}

function Test-WinRMReachability {
    param(
        [Parameter(Mandatory=$true)][string]$ComputerName,
        [int]$TimeoutSeconds = 5
    )
    $job = Start-Job -ScriptBlock {
        param($Name)
        Test-WSMan -ComputerName $Name -ErrorAction Stop | Out-Null
        $true
    } -ArgumentList $ComputerName
    try {
        if (Wait-Job -Job $job -Timeout $TimeoutSeconds) {
            return [bool](Receive-Job -Job $job -ErrorAction Stop)
        }
        Stop-Job -Job $job -Force -ErrorAction SilentlyContinue
        return $false
    }
    catch { return $false }
    finally { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue }
}

function Invoke-RemoteHealthSample {
    param(
        [Parameter(Mandatory=$true)][string]$ComputerName,
        [Parameter(Mandatory=$true)]$Settings
    )

    Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        param($CpuSampleSeconds, $TopProcessCount, $IncludeDisconnectedSessions)

        function Convert-SessionState {
            param([string]$State)
            switch -Regex ($State) {
                '^(Active|Aktiv)$' { 'Active'; break }
                '^(Disc|Getr|Disconnected)$' { 'Disconnected'; break }
                default { $State }
            }
        }

        function Get-TotalCpuPercent {
            $processor = Get-CimInstance -ClassName Win32_PerfFormattedData_PerfOS_Processor -Filter "Name='_Total'"
            if ($processor -and $null -ne $processor.PercentProcessorTime) {
                return [double]$processor.PercentProcessorTime
            }
            return 0
        }

        $sampleStart = Get-Date
        $cpuStartPercent = Get-TotalCpuPercent
        $processStart = Get-Process | Select-Object Id, ProcessName, CPU
        Start-Sleep -Seconds $CpuSampleSeconds
        $cpuEndPercent = Get-TotalCpuPercent
        $processEnd = Get-Process | Select-Object Id, ProcessName, CPU, SessionId, WorkingSet64
        $sampleEnd = Get-Date
        $elapsedSeconds = [Math]::Max(($sampleEnd - $sampleStart).TotalSeconds, 1)

        $os = Get-CimInstance -ClassName Win32_OperatingSystem
        $totalMemoryMb = [Math]::Round($os.TotalVisibleMemorySize / 1024, 2)
        $freeMemoryMb = [Math]::Round($os.FreePhysicalMemory / 1024, 2)
        $usedMemoryMb = [Math]::Round($totalMemoryMb - $freeMemoryMb, 2)
        $memoryPercent = if ($totalMemoryMb -gt 0) { [Math]::Round(($usedMemoryMb / $totalMemoryMb) * 100, 2) } else { 0 }
        $cpuPercent = [Math]::Round((($cpuStartPercent + $cpuEndPercent) / 2), 2)

        $sessions = @()
        $quserOutput = & quser.exe 2>$null
        if ($LASTEXITCODE -eq 0 -and $quserOutput) {
            foreach ($line in ($quserOutput | Select-Object -Skip 1)) {
                $clean = ($line -replace '^>', '').Trim()
                if (-not $clean) { continue }
                if ($clean -match '^(.+?)\s+(\S+)\s+(\d+)\s+(\S+)\s+') {
                    $state = Convert-SessionState $matches[4]
                    if ($IncludeDisconnectedSessions -or $state -eq 'Active') {
                        $sessions += [pscustomobject]@{ UserName = $matches[1].Trim(); SessionName = $matches[2]; SessionId = [int]$matches[3]; State = $state }
                    }
                }
                elseif ($clean -match '^(.+?)\s+(\d+)\s+(\S+)\s+') {
                    $state = Convert-SessionState $matches[3]
                    if ($IncludeDisconnectedSessions -or $state -eq 'Active') {
                        $sessions += [pscustomobject]@{ UserName = $matches[1].Trim(); SessionName = ''; SessionId = [int]$matches[2]; State = $state }
                    }
                }
            }
        }

        $sessionById = @{}
        foreach ($session in $sessions) { $sessionById[[int]$session.SessionId] = $session }
        $startById = @{}
        foreach ($process in $processStart) { $startById[[int]$process.Id] = $process }

        $topProcesses = @(foreach ($process in $processEnd) {
            $startProcess = $startById[[int]$process.Id]
            $startCpu = if ($startProcess -and $null -ne $startProcess.CPU) { [double]$startProcess.CPU } else { 0 }
            $endCpu = if ($null -ne $process.CPU) { [double]$process.CPU } else { 0 }
            $delta = [Math]::Max($endCpu - $startCpu, 0)
            $cpuPct = [Math]::Round(($delta / $elapsedSeconds) * 100, 2)
            $session = $sessionById[[int]$process.SessionId]
            [pscustomobject]@{
                ProcessId = $process.Id
                ProcessName = $process.ProcessName
                SessionId = $process.SessionId
                UserName = if ($session) { $session.UserName } else { '' }
                ProcessCpuPercent = $cpuPct
                CpuSecondsDelta = [Math]::Round($delta, 2)
                WorkingSetMb = [Math]::Round($process.WorkingSet64 / 1MB, 2)
            }
        }) | Sort-Object ProcessCpuPercent -Descending | Select-Object -First $TopProcessCount

        [pscustomobject]@{
            ComputerName = $env:COMPUTERNAME
            SampleTime = $sampleEnd.ToString('s')
            CpuPercent = $cpuPercent
            MemoryPercent = $memoryPercent
            TotalMemoryMb = $totalMemoryMb
            UsedMemoryMb = $usedMemoryMb
            FreeMemoryMb = $freeMemoryMb
            ActiveSessions = @($sessions | Where-Object { $_.State -eq 'Active' }).Count
            DisconnectedSessions = @($sessions | Where-Object { $_.State -eq 'Disconnected' }).Count
            Sessions = $sessions
            TopProcesses = $topProcesses
        }
    } -ArgumentList ([int]$Settings.CpuSampleSeconds), ([int]$Settings.TopProcessCount), ([bool]$Settings.IncludeDisconnectedSessions)
}

function Convert-HealthResultToRows {
    param(
        [Parameter(Mandatory=$true)][string]$TargetServer,
        [Parameter(Mandatory=$true)]$Result
    )
    foreach ($process in $Result.TopProcesses) {
        [pscustomobject]@{
            Timestamp = $Result.SampleTime
            TargetServer = $TargetServer
            ComputerName = $Result.ComputerName
            CpuPercent = $Result.CpuPercent
            MemoryPercent = $Result.MemoryPercent
            TotalMemoryMb = $Result.TotalMemoryMb
            UsedMemoryMb = $Result.UsedMemoryMb
            FreeMemoryMb = $Result.FreeMemoryMb
            ActiveSessions = $Result.ActiveSessions
            DisconnectedSessions = $Result.DisconnectedSessions
            ProcessRank = $null
            ProcessId = $process.ProcessId
            ProcessName = $process.ProcessName
            ProcessSessionId = $process.SessionId
            ProcessUserName = $process.UserName
            ProcessCpuPercent = $process.ProcessCpuPercent
            ProcessCpuSecondsDelta = $process.CpuSecondsDelta
            ProcessWorkingSetMb = $process.WorkingSetMb
            Status = 'OK'
            ErrorMessage = ''
        }
    }
}

$rawPath = Join-Path $OutputPath 'raw'
$logPath = Join-Path $OutputPath 'logs'
$summaryPath = Join-Path $OutputPath 'summary'
Ensure-Directory $rawPath; Ensure-Directory $logPath; Ensure-Directory $summaryPath

$runId = Get-Date -Format 'yyyyMMdd-HHmmss'
$day = Get-Date -Format 'yyyy-MM-dd'
$logFile = Join-Path $logPath "healthcheck-$day.log"
$rawFile = Join-Path $rawPath "healthcheck-$day.csv"
$summaryFile = Join-Path $summaryPath "summary-$runId.csv"

$settings = Read-Settings -Path (Join-Path $ConfigPath 'settings.json')
$servers = @(Read-ServerList -Path (Join-Path $ConfigPath 'servers.txt'))
$summaryRows = @()
$rawRows = @()

Write-Log -LogPath $logFile -Level INFO -Message "Run $runId gestartet. Server: $($servers.Count)"
foreach ($server in $servers) {
    try {
        Write-Log -LogPath $logFile -Level INFO -Message "Pruefe WinRM: $server"
        if (-not (Test-WinRMReachability -ComputerName $server -TimeoutSeconds ([int]$settings.WinRMTimeoutSeconds))) { throw "WinRM nicht erreichbar oder Timeout nach $($settings.WinRMTimeoutSeconds) Sekunden." }
        $result = Invoke-RemoteHealthSample -ComputerName $server -Settings $settings
        $rows = @(Convert-HealthResultToRows -TargetServer $server -Result $result)
        $rank = 1
        foreach ($row in $rows) { $row.ProcessRank = $rank; $rank++ }
        $rawRows += $rows
        $summaryRows += [pscustomobject]@{
            RunId = $runId; Timestamp = $result.SampleTime; TargetServer = $server; Status = 'OK'; CpuPercent = $result.CpuPercent;
            MemoryPercent = $result.MemoryPercent; ActiveSessions = $result.ActiveSessions; DisconnectedSessions = $result.DisconnectedSessions;
            TopProcess = if ($rows.Count -gt 0) { $rows[0].ProcessName } else { '' }; TopProcessCpuPercent = if ($rows.Count -gt 0) { $rows[0].ProcessCpuPercent } else { 0 }; ErrorMessage = ''
        }
    }
    catch {
        $message = $_.Exception.Message
        Write-Log -LogPath $logFile -Level ERROR -Message "${server}: $message"
        $summaryRows += [pscustomobject]@{ RunId = $runId; Timestamp = (Get-Date).ToString('s'); TargetServer = $server; Status = 'ERROR'; CpuPercent = ''; MemoryPercent = ''; ActiveSessions = ''; DisconnectedSessions = ''; TopProcess = ''; TopProcessCpuPercent = ''; ErrorMessage = $message }
    }
}

if ($rawRows.Count -gt 0) { $rawRows | Export-Csv -LiteralPath $rawFile -Delimiter $settings.OutputDelimiter -NoTypeInformation -Append -Encoding UTF8 }
$summaryRows | Export-Csv -LiteralPath $summaryFile -Delimiter $settings.OutputDelimiter -NoTypeInformation -Encoding UTF8
Write-Log -LogPath $logFile -Level INFO -Message "Run $runId beendet. Summary: $summaryFile Raw: $rawFile"
$summaryRows
