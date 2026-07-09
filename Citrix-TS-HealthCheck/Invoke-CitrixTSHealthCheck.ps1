#requires -version 3.0
<#
.SYNOPSIS
    Collects Citrix/Terminalserver health samples via PowerShell Remoting.
.DESCRIPTION
    Windows PowerShell 5.1 compatible collector for longer CPU peak analysis runs.
    It writes separate CSV files for server samples, process samples, alerts,
    category summaries, optional event context and a per-run summary.
#>
[CmdletBinding()]
param(
    [string]$ServerListPath = (Join-Path $PSScriptRoot 'config\servers.txt'),
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'config\settings.json'),
    [string]$OutputPath = (Join-Path $PSScriptRoot 'output'),
    [int]$DurationMinutes,
    [int]$IntervalSeconds,
    [int]$CpuSampleSeconds,
    [int]$TopProcessCount,
    [int]$AlertTopProcessCount,
    [double]$CpuWarningThreshold,
    [double]$CpuCriticalThreshold,
    [int]$MaxParallel,
    [switch]$IncludeEventLogContext,
    [switch]$AnonymizeUsers,
    [string]$Delimiter
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function New-DefaultSettings {
    [pscustomobject]@{
        DurationMinutes = 480
        IntervalSeconds = 300
        CpuSampleSeconds = 5
        TopProcessCount = 10
        AlertTopProcessCount = 25
        CpuWarningThreshold = 70
        CpuCriticalThreshold = 90
        MaxParallel = 4
        IncludeEventLogContext = $false
        AnonymizeUsers = $false
        OutputDelimiter = ';'
    }
}

function Read-Settings {
    param([string]$Path)
    $settings = New-DefaultSettings
    if (Test-Path -LiteralPath $Path) {
        $json = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($property in $settings.PSObject.Properties.Name) {
            if ($null -ne $json.$property) { $settings.$property = $json.$property }
        }
    }
    return $settings
}

function Merge-ParameterSettings {
    param($Settings)
    if ($PSBoundParameters.ContainsKey('DurationMinutes')) { $Settings.DurationMinutes = $DurationMinutes }
    if ($PSBoundParameters.ContainsKey('IntervalSeconds')) { $Settings.IntervalSeconds = $IntervalSeconds }
    if ($PSBoundParameters.ContainsKey('CpuSampleSeconds')) { $Settings.CpuSampleSeconds = $CpuSampleSeconds }
    if ($PSBoundParameters.ContainsKey('TopProcessCount')) { $Settings.TopProcessCount = $TopProcessCount }
    if ($PSBoundParameters.ContainsKey('AlertTopProcessCount')) { $Settings.AlertTopProcessCount = $AlertTopProcessCount }
    if ($PSBoundParameters.ContainsKey('CpuWarningThreshold')) { $Settings.CpuWarningThreshold = $CpuWarningThreshold }
    if ($PSBoundParameters.ContainsKey('CpuCriticalThreshold')) { $Settings.CpuCriticalThreshold = $CpuCriticalThreshold }
    if ($PSBoundParameters.ContainsKey('MaxParallel')) { $Settings.MaxParallel = $MaxParallel }
    if ($PSBoundParameters.ContainsKey('Delimiter')) { $Settings.OutputDelimiter = $Delimiter }
    if ($PSBoundParameters.ContainsKey('IncludeEventLogContext')) { $Settings.IncludeEventLogContext = [bool]$IncludeEventLogContext }
    if ($PSBoundParameters.ContainsKey('AnonymizeUsers')) { $Settings.AnonymizeUsers = [bool]$AnonymizeUsers }
    return $Settings
}

function Read-ServerList {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { throw "Serverliste nicht gefunden: $Path" }
    Get-Content -LiteralPath $Path -Encoding UTF8 |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ -and -not $_.StartsWith('#') } |
        Select-Object -Unique
}

function Ensure-Directory {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
}

function ConvertTo-InvariantObject {
    param([Parameter(ValueFromPipeline=$true)]$InputObject)
    process {
        $copy = [ordered]@{}
        foreach ($property in $InputObject.PSObject.Properties) {
            $value = $property.Value
            if ($value -is [double] -or $value -is [single] -or $value -is [decimal]) {
                $copy[$property.Name] = ([double]$value).ToString('0.##', [Globalization.CultureInfo]::InvariantCulture)
            }
            else { $copy[$property.Name] = $value }
        }
        [pscustomobject]$copy
    }
}

function Export-Rows {
    param([array]$Rows, [string]$Path, [string]$Delimiter)
    if (-not $Rows -or $Rows.Count -eq 0) { return }
    $Rows | ConvertTo-InvariantObject | Export-Csv -LiteralPath $Path -Delimiter $Delimiter -NoTypeInformation -Append -Encoding UTF8
}

function Get-RemoteSamplerScriptBlock {
    {
        param($CpuSampleSeconds, $TopProcessCount, $AlertTopProcessCount, $CpuWarningThreshold, $CpuCriticalThreshold, $IncludeEventLogContext, $AnonymizeUsers)

        function Get-ProcessCategory {
            param([string]$Name)
            switch -Regex ($Name) {
                '^(MsSense|SenseNdr|MsMpEng|CylanceSvc)$' { 'Security'; break }
                '^(BrokerAgent|CtxGfx|Citrix\.Wem\.Agent\.Service|wfcrun32|wfica32)$' { 'Citrix'; break }
                '^(nexus\.framework\.healthcare|Infoclient|anm_neu)$' { 'Nexus'; break }
                '^(WINWORD|EXCEL|OUTLOOK|POWERPNT)$' { 'Office'; break }
                '^(msedge|msedgewebview2|chrome)$' { 'Browser'; break }
                '^(spoolsv|splwow64)$' { 'Printing'; break }
                '^(wsmprovhost|WmiPrvSE|powershell|pwsh)$' { 'Monitoring'; break }
                '^(svchost|System|explorer|dwm)$' { 'Windows'; break }
                default { 'Other' }
            }
        }

        function Convert-SessionState {
            param([string]$State)
            switch -Regex ($State) {
                '^(Active|Aktiv)$' { 'Active'; break }
                '^(Disc|Getr|Disconnected)$' { 'Disconnected'; break }
                default { $State }
            }
        }

        function Convert-UserName {
            param([string]$UserName, [bool]$Anonymize)
            if ([string]::IsNullOrWhiteSpace($UserName)) { return '' }
            if (-not $Anonymize) { return $UserName }
            $sha = [Security.Cryptography.SHA256]::Create()
            $bytes = [Text.Encoding]::UTF8.GetBytes($UserName.ToLowerInvariant())
            $hash = $sha.ComputeHash($bytes)
            return ([BitConverter]::ToString($hash) -replace '-', '').Substring(0, 16)
        }

        function Get-TotalCpuPercent {
            $processor = Get-CimInstance -ClassName Win32_PerfFormattedData_PerfOS_Processor -Filter "Name='_Total'"
            if ($processor -and $null -ne $processor.PercentProcessorTime) { return [double]$processor.PercentProcessorTime }
            return 0
        }

        function Get-SessionMap {
            $map = @{}
            $sessions = @()
            $quserOutput = & quser.exe 2>$null
            if ($LASTEXITCODE -eq 0 -and $quserOutput) {
                foreach ($line in ($quserOutput | Select-Object -Skip 1)) {
                    $clean = ($line -replace '^>', '').Trim()
                    if (-not $clean) { continue }
                    $parts = $clean -split '\s{2,}'
                    if ($parts.Count -lt 3) { continue }
                    $user = $parts[0].Trim()
                    $sessionName = ''
                    $sessionId = $null
                    $state = ''
                    if ($parts[1] -match '^\d+$') {
                        $sessionId = [int]$parts[1]
                        $state = Convert-SessionState $parts[2]
                    }
                    elseif ($parts.Count -ge 4 -and $parts[2] -match '^\d+$') {
                        $sessionName = $parts[1]
                        $sessionId = [int]$parts[2]
                        $state = Convert-SessionState $parts[3]
                    }
                    if ($null -ne $sessionId) {
                        $session = [pscustomobject]@{ UserName = $user; SessionName = $sessionName; SessionId = $sessionId; State = $state }
                        $sessions += $session
                        $map[$sessionId] = $session
                    }
                }
            }
            [pscustomobject]@{ Map = $map; Sessions = $sessions }
        }

        $sampleStart = Get-Date
        $cpuStartPercent = Get-TotalCpuPercent
        $processStart = Get-Process | Select-Object Id, ProcessName, CPU
        Start-Sleep -Seconds $CpuSampleSeconds
        $cpuEndPercent = Get-TotalCpuPercent
        $processEnd = Get-Process | Select-Object Id, ProcessName, CPU, SessionId, WorkingSet64, PrivateMemorySize64, StartTime -ErrorAction SilentlyContinue
        $sampleEnd = Get-Date
        $elapsedSeconds = [Math]::Max(($sampleEnd - $sampleStart).TotalSeconds, 1)
        $cpuPercent = [Math]::Round((($cpuStartPercent + $cpuEndPercent) / 2), 2)

        $os = Get-CimInstance -ClassName Win32_OperatingSystem
        $totalMemoryMb = [Math]::Round($os.TotalVisibleMemorySize / 1024, 2)
        $freeMemoryMb = [Math]::Round($os.FreePhysicalMemory / 1024, 2)
        $usedMemoryMb = [Math]::Round($totalMemoryMb - $freeMemoryMb, 2)
        $memoryPercent = if ($totalMemoryMb -gt 0) { [Math]::Round(($usedMemoryMb / $totalMemoryMb) * 100, 2) } else { 0 }

        $sessionInfo = Get-SessionMap
        $sessionById = $sessionInfo.Map
        $sessions = @($sessionInfo.Sessions)
        $startById = @{}
        foreach ($process in $processStart) { $startById[[int]$process.Id] = $process }

        $processDeltas = foreach ($process in $processEnd) {
            $startProcess = $startById[[int]$process.Id]
            $startCpu = if ($startProcess -and $null -ne $startProcess.CPU) { [double]$startProcess.CPU } else { 0 }
            $endCpu = if ($null -ne $process.CPU) { [double]$process.CPU } else { 0 }
            $delta = [Math]::Max($endCpu - $startCpu, 0)
            $cpuPct = [Math]::Round(($delta / $elapsedSeconds) * 100, 2)
            $category = Get-ProcessCategory $process.ProcessName
            $startTimeText = ''
            try {
                if ($process.StartTime) { $startTimeText = $process.StartTime.ToString('s') }
            }
            catch { $startTimeText = '' }
            [pscustomobject]@{
                ProcessId = [int]$process.Id
                ProcessName = [string]$process.ProcessName
                SessionId = if ($null -ne $process.SessionId) { [int]$process.SessionId } else { -1 }
                ProcessCpuPercent = $cpuPct
                CpuSecondsDelta = [Math]::Round($delta, 3)
                WorkingSetMb = [Math]::Round($process.WorkingSet64 / 1MB, 2)
                PrivateMemoryMb = if ($null -ne $process.PrivateMemorySize64) { [Math]::Round($process.PrivateMemorySize64 / 1MB, 2) } else { '' }
                StartTime = $startTimeText
                Category = $category
                IsMonitoringRelated = ($category -eq 'Monitoring')
            }
        }

        $isWarning = $cpuPercent -ge $CpuWarningThreshold
        $isCritical = $cpuPercent -ge $CpuCriticalThreshold
        $topLimit = if ($isWarning) { $AlertTopProcessCount } else { $TopProcessCount }
        $specialNames = @('MsSense','SenseNdr','MsMpEng','CylanceSvc','spoolsv','splwow64','nexus.framework.healthcare','Infoclient','anm_neu')
        $selectedIds = @{}
        foreach ($p in ($processDeltas | Sort-Object ProcessCpuPercent -Descending | Select-Object -First $topLimit)) { $selectedIds[$p.ProcessId] = $true }
        if ($isCritical) { foreach ($p in ($processDeltas | Where-Object { $specialNames -contains $_.ProcessName })) { $selectedIds[$p.ProcessId] = $true } }
        $selected = @($processDeltas | Where-Object { $selectedIds.ContainsKey($_.ProcessId) } | Sort-Object ProcessCpuPercent -Descending)

        $cimByPid = @{}
        foreach ($p in $selected) {
            try {
                $cimProcess = Get-CimInstance -ClassName Win32_Process -Filter ("ProcessId={0}" -f $p.ProcessId) -ErrorAction Stop
                if ($cimProcess) { $cimByPid[$p.ProcessId] = $cimProcess }
            } catch { }
        }
        $nameByPid = @{}
        foreach ($p in $processDeltas) { if (-not $nameByPid.ContainsKey($p.ProcessId)) { $nameByPid[$p.ProcessId] = $p.ProcessName } }

        $processSamples = foreach ($p in $selected) {
            $cim = $cimByPid[$p.ProcessId]
            $ownerUser = ''
            $ownerDomain = ''
            if ($cim) {
                try {
                    $owner = Invoke-CimMethod -InputObject $cim -MethodName GetOwner -ErrorAction Stop
                    if ($owner.ReturnValue -eq 0) { $ownerUser = Convert-UserName $owner.User ([bool]$AnonymizeUsers); $ownerDomain = $owner.Domain }
                } catch { }
            }
            $session = $sessionById[$p.SessionId]
            if ([string]::IsNullOrWhiteSpace($ownerUser) -and $session) { $ownerUser = Convert-UserName $session.UserName ([bool]$AnonymizeUsers) }
            [pscustomobject]@{
                ProcessName = $p.ProcessName
                PID = $p.ProcessId
                ProcessSessionId = $p.SessionId
                ProcessUserName = $ownerUser
                ProcessUserDomain = $ownerDomain
                SessionState = if ($session) { $session.State } else { '' }
                ProcessCpuPercent = $p.ProcessCpuPercent
                ProcessCpuSecondsDelta = $p.CpuSecondsDelta
                WorkingSetMB = $p.WorkingSetMb
                PrivateMemoryMB = $p.PrivateMemoryMb
                ProcessPath = if ($cim) { $cim.ExecutablePath } else { '' }
                CommandLine = if ($cim) { $cim.CommandLine } else { '' }
                ParentProcessId = if ($cim) { $cim.ParentProcessId } else { '' }
                ParentProcessName = if ($cim -and $nameByPid.ContainsKey([int]$cim.ParentProcessId)) { $nameByPid[[int]$cim.ParentProcessId] } else { '' }
                StartTime = $p.StartTime
                IsSystemProcess = [string]::IsNullOrWhiteSpace($ownerUser)
                IsUserProcess = -not [string]::IsNullOrWhiteSpace($ownerUser)
                IsMonitoringRelated = $p.IsMonitoringRelated
                Category = $p.Category
            }
        }

        $categoryNames = @('Security','Nexus','Office','Browser','Citrix','Printing','Monitoring','Windows','Other')
        $categorySums = @{}
        foreach ($name in $categoryNames) { $categorySums[$name] = 0.0 }
        foreach ($p in $processDeltas) { $categorySums[$p.Category] = [double]$categorySums[$p.Category] + [double]$p.ProcessCpuPercent }

        $events = @()
        if ($IncludeEventLogContext -and $isCritical) {
            $since = (Get-Date).AddMinutes(-5)
            foreach ($logName in @('System','Application','Microsoft-Windows-Windows Defender/Operational')) {
                try {
                    $events += Get-WinEvent -FilterHashtable @{ LogName = $logName; StartTime = $since } -MaxEvents 25 -ErrorAction Stop |
                        Select-Object @{n='LogName';e={$logName}}, TimeCreated, Id, ProviderName, LevelDisplayName, Message
                } catch { }
            }
        }

        [pscustomobject]@{
            ServerSample = [pscustomobject]@{
                Timestamp = $sampleEnd.ToString('s')
                ComputerName = $env:COMPUTERNAME
                CpuPercent = $cpuPercent
                MemoryPercent = $memoryPercent
                TotalMemoryMB = $totalMemoryMb
                UsedMemoryMB = $usedMemoryMb
                FreeMemoryMB = $freeMemoryMb
                ActiveSessions = @($sessions | Where-Object { $_.State -eq 'Active' }).Count
                DisconnectedSessions = @($sessions | Where-Object { $_.State -eq 'Disconnected' }).Count
                TotalSessions = @($sessions).Count
                Status = 'OK'
                ErrorMessage = ''
                IsCpuWarning = $isWarning
                IsCpuCritical = $isCritical
                TopProcessCpuPercentSum = [Math]::Round((($selected | Measure-Object -Property ProcessCpuPercent -Sum).Sum), 2)
            }
            ProcessSamples = @($processSamples)
            CategorySummary = [pscustomobject]@{
                SecurityCpuPercent = [Math]::Round($categorySums['Security'], 2)
                NexusCpuPercent = [Math]::Round($categorySums['Nexus'], 2)
                OfficeCpuPercent = [Math]::Round($categorySums['Office'], 2)
                BrowserCpuPercent = [Math]::Round($categorySums['Browser'], 2)
                CitrixCpuPercent = [Math]::Round($categorySums['Citrix'], 2)
                PrintingCpuPercent = [Math]::Round($categorySums['Printing'], 2)
                MonitoringCpuPercent = [Math]::Round($categorySums['Monitoring'], 2)
                WindowsCpuPercent = [Math]::Round($categorySums['Windows'], 2)
                OtherCpuPercent = [Math]::Round($categorySums['Other'], 2)
            }
            EventSamples = @($events)
        }
    }
}

function Invoke-ServerCollectionRound {
    param([string[]]$Servers, $Settings, [datetime]$RoundTimestamp)
    $remoteScript = (Get-RemoteSamplerScriptBlock).ToString()
    $jobs = @()
    $results = @()
    foreach ($server in $Servers) {
        while (@($jobs | Where-Object { $_.State -eq 'Running' }).Count -ge [int]$Settings.MaxParallel) {
            $done = Wait-Job -Job $jobs -Any -Timeout 5
            if ($done) { $results += Receive-Job -Job $done; Remove-Job -Job $done -Force; $jobs = @($jobs | Where-Object Id -ne $done.Id) }
        }
        $jobs += Start-Job -ScriptBlock {
            param($Server, $RemoteScript, $SettingsHash)
            try {
                Test-WSMan -ComputerName $Server -ErrorAction Stop | Out-Null
                $remoteBlock = [scriptblock]::Create($RemoteScript)
                $result = Invoke-Command -ComputerName $Server -ScriptBlock $remoteBlock -ArgumentList $SettingsHash.CpuSampleSeconds, $SettingsHash.TopProcessCount, $SettingsHash.AlertTopProcessCount, $SettingsHash.CpuWarningThreshold, $SettingsHash.CpuCriticalThreshold, $SettingsHash.IncludeEventLogContext, $SettingsHash.AnonymizeUsers -ErrorAction Stop
                [pscustomobject]@{ TargetServer = $Server; Status = 'OK'; Result = $result; ErrorMessage = '' }
            }
            catch {
                [pscustomobject]@{ TargetServer = $Server; Status = 'ERROR'; Result = $null; ErrorMessage = $_.Exception.Message }
            }
        } -ArgumentList $server, $remoteScript, $Settings
    }
    while ($jobs.Count -gt 0) {
        $done = Wait-Job -Job $jobs -Any -Timeout 5
        if ($done) { $results += Receive-Job -Job $done; Remove-Job -Job $done -Force; $jobs = @($jobs | Where-Object Id -ne $done.Id) }
    }
    return $results
}

function New-ErrorServerSample {
    param([string]$Server, [string]$Message, [datetime]$Timestamp)
    [pscustomobject]@{
        Timestamp = $Timestamp.ToString('s'); TargetServer = $Server; ComputerName = ''; CpuPercent = ''; MemoryPercent = ''; TotalMemoryMB = ''; UsedMemoryMB = ''; FreeMemoryMB = '';
        ActiveSessions = ''; DisconnectedSessions = ''; TotalSessions = ''; Status = 'ERROR'; ErrorMessage = $Message; IsCpuWarning = $false; IsCpuCritical = $false; TopProcessCpuPercentSum = ''
    }
}

function New-RunSummaryRows {
    param([array]$ServerRows, [array]$ProcessRows, [datetime]$Start, [datetime]$End, [int]$ServerCount)
    foreach ($group in ($ServerRows | Group-Object TargetServer)) {
        $server = $group.Name
        $okRows = @($group.Group | Where-Object Status -eq 'OK')
        $cpuValues = @($okRows | Where-Object { $_.CpuPercent -ne '' } | ForEach-Object { [double]$_.CpuPercent } | Sort-Object)
        $ramValues = @($okRows | Where-Object { $_.MemoryPercent -ne '' } | ForEach-Object { [double]$_.MemoryPercent })
        $sessionValues = @($okRows | Where-Object { $_.TotalSessions -ne '' } | ForEach-Object { [double]$_.TotalSessions })
        $median = if ($cpuValues.Count -gt 0) { if ($cpuValues.Count % 2) { $cpuValues[[int]($cpuValues.Count/2)] } else { ($cpuValues[$cpuValues.Count/2-1] + $cpuValues[$cpuValues.Count/2]) / 2 } } else { '' }
        $topProcesses = @($ProcessRows | Where-Object TargetServer -eq $server | Group-Object ProcessName | ForEach-Object { [pscustomobject]@{ Name=$_.Name; Cpu=(($_.Group | Measure-Object ProcessCpuSecondsDelta -Sum).Sum) } } | Sort-Object Cpu -Descending | Select-Object -First 10 | ForEach-Object { '{0}={1:n2}s' -f $_.Name,$_.Cpu }) -join ', '
        $topCategories = @($ProcessRows | Where-Object TargetServer -eq $server | Group-Object Category | ForEach-Object { [pscustomobject]@{ Name=$_.Name; Cpu=(($_.Group | Measure-Object ProcessCpuSecondsDelta -Sum).Sum) } } | Sort-Object Cpu -Descending | Select-Object -First 5 | ForEach-Object { '{0}={1:n2}s' -f $_.Name,$_.Cpu }) -join ', '
        [pscustomobject]@{
            RunStart = $Start.ToString('s'); RunEnd = $End.ToString('s'); DurationMinutes = [Math]::Round(($End-$Start).TotalMinutes, 2); ServerCount = $ServerCount; TargetServer = $server;
            SuccessfulSamples = $okRows.Count; ErrorSamples = @($group.Group | Where-Object Status -eq 'ERROR').Count; CpuAverage = if ($cpuValues.Count) { [Math]::Round((($cpuValues | Measure-Object -Average).Average),2) } else { '' };
            CpuMedian = if ($median -ne '') { [Math]::Round($median,2) } else { '' }; CpuMaximum = if ($cpuValues.Count) { [Math]::Round((($cpuValues | Measure-Object -Maximum).Maximum),2) } else { '' };
            CpuWarnings = @($okRows | Where-Object { $_.IsCpuWarning -eq $true }).Count; CpuCriticals = @($okRows | Where-Object { $_.IsCpuCritical -eq $true }).Count;
            RamAverage = if ($ramValues.Count) { [Math]::Round((($ramValues | Measure-Object -Average).Average),2) } else { '' }; SessionAverage = if ($sessionValues.Count) { [Math]::Round((($sessionValues | Measure-Object -Average).Average),2) } else { '' };
            TopProcessesByCpuSeconds = $topProcesses; TopCategoriesByCpuSeconds = $topCategories
        }
    }
}

if ((Test-Path -LiteralPath $ConfigPath -PathType Container)) { $ConfigPath = Join-Path $ConfigPath 'settings.json' }
$settings = Merge-ParameterSettings -Settings (Read-Settings -Path $ConfigPath)
$servers = @(Read-ServerList -Path $ServerListPath)
$rawPath = Join-Path $OutputPath 'raw'
$summaryPath = Join-Path $OutputPath 'summary'
$logPath = Join-Path $OutputPath 'logs'
Ensure-Directory $rawPath; Ensure-Directory $summaryPath; Ensure-Directory $logPath

$runStart = Get-Date
$runId = $runStart.ToString('yyyy-MM-dd_HH-mm')
$allServerRows = @()
$allProcessRows = @()
$round = 0
$endAt = $runStart.AddMinutes([int]$settings.DurationMinutes)
if ([int]$settings.DurationMinutes -lt 1) { $endAt = $runStart }

do {
    $round++
    $roundStart = Get-Date
    Write-Host ("[{0}] Runde {1} gestartet. Server={2}, Dauer={3}s, Parallel={4}" -f $roundStart.ToString('HH:mm:ss'), $round, $servers.Count, $settings.CpuSampleSeconds, $settings.MaxParallel)
    $roundResults = Invoke-ServerCollectionRound -Servers $servers -Settings $settings -RoundTimestamp $roundStart
    $day = (Get-Date).ToString('yyyy-MM-dd')
    $serverRows = @(); $processRows = @(); $alertRows = @(); $categoryRows = @(); $eventRows = @()

    foreach ($item in $roundResults) {
        if ($item.Status -ne 'OK') {
            $serverRows += New-ErrorServerSample -Server $item.TargetServer -Message $item.ErrorMessage -Timestamp $roundStart
            continue
        }
        $serverSample = $item.Result.ServerSample
        $serverRow = [pscustomobject]@{
            Timestamp=$serverSample.Timestamp; TargetServer=$item.TargetServer; ComputerName=$serverSample.ComputerName; CpuPercent=$serverSample.CpuPercent; MemoryPercent=$serverSample.MemoryPercent;
            TotalMemoryMB=$serverSample.TotalMemoryMB; UsedMemoryMB=$serverSample.UsedMemoryMB; FreeMemoryMB=$serverSample.FreeMemoryMB; ActiveSessions=$serverSample.ActiveSessions;
            DisconnectedSessions=$serverSample.DisconnectedSessions; TotalSessions=$serverSample.TotalSessions; Status='OK'; ErrorMessage=''; IsCpuWarning=$serverSample.IsCpuWarning; IsCpuCritical=$serverSample.IsCpuCritical; TopProcessCpuPercentSum=$serverSample.TopProcessCpuPercentSum
        }
        $serverRows += $serverRow
        foreach ($p in @($item.Result.ProcessSamples)) {
            $processRows += [pscustomobject]@{ Timestamp=$serverSample.Timestamp; TargetServer=$item.TargetServer; ComputerName=$serverSample.ComputerName; CpuPercent=$serverSample.CpuPercent; AlertLevel=if($serverSample.IsCpuCritical){'Critical'}elseif($serverSample.IsCpuWarning){'Warning'}else{''}; ProcessName=$p.ProcessName; PID=$p.PID; ProcessSessionId=$p.ProcessSessionId; ProcessUserName=$p.ProcessUserName; ProcessUserDomain=$p.ProcessUserDomain; SessionState=$p.SessionState; ProcessCpuPercent=$p.ProcessCpuPercent; ProcessCpuSecondsDelta=$p.ProcessCpuSecondsDelta; WorkingSetMB=$p.WorkingSetMB; PrivateMemoryMB=$p.PrivateMemoryMB; ProcessPath=$p.ProcessPath; CommandLine=$p.CommandLine; ParentProcessId=$p.ParentProcessId; ParentProcessName=$p.ParentProcessName; StartTime=$p.StartTime; IsSystemProcess=$p.IsSystemProcess; IsUserProcess=$p.IsUserProcess; IsMonitoringRelated=$p.IsMonitoringRelated; Category=$p.Category }
        }
        $cat = $item.Result.CategorySummary
        $categoryRows += [pscustomobject]@{ Timestamp=$serverSample.Timestamp; TargetServer=$item.TargetServer; ComputerName=$serverSample.ComputerName; SecurityCpuPercent=$cat.SecurityCpuPercent; NexusCpuPercent=$cat.NexusCpuPercent; OfficeCpuPercent=$cat.OfficeCpuPercent; BrowserCpuPercent=$cat.BrowserCpuPercent; CitrixCpuPercent=$cat.CitrixCpuPercent; PrintingCpuPercent=$cat.PrintingCpuPercent; MonitoringCpuPercent=$cat.MonitoringCpuPercent; WindowsCpuPercent=$cat.WindowsCpuPercent; OtherCpuPercent=$cat.OtherCpuPercent }
        if ($serverSample.IsCpuWarning) {
            $categoryText = @('Security','Nexus','Office','Browser','Citrix','Printing','Monitoring','Windows','Other') | ForEach-Object { $propertyName = "${_}CpuPercent"; '{0}={1}' -f $_,$cat.$propertyName }
            foreach ($p in @($item.Result.ProcessSamples)) { $alertRows += [pscustomobject]@{ Timestamp=$serverSample.Timestamp; TargetServer=$item.TargetServer; ComputerName=$serverSample.ComputerName; AlertLevel=if($serverSample.IsCpuCritical){'Critical'}else{'Warning'}; CpuPercent=$serverSample.CpuPercent; MemoryPercent=$serverSample.MemoryPercent; ActiveSessions=$serverSample.ActiveSessions; DisconnectedSessions=$serverSample.DisconnectedSessions; TotalSessions=$serverSample.TotalSessions; CategorySums=($categoryText -join ', '); ProcessName=$p.ProcessName; PID=$p.PID; ProcessCpuPercent=$p.ProcessCpuPercent; ProcessCpuSecondsDelta=$p.ProcessCpuSecondsDelta; Category=$p.Category; IsMonitoringRelated=$p.IsMonitoringRelated; ProcessUserName=$p.ProcessUserName; ProcessUserDomain=$p.ProcessUserDomain; ProcessSessionId=$p.ProcessSessionId; CommandLine=$p.CommandLine } }
        }
        foreach ($e in @($item.Result.EventSamples)) { $eventRows += [pscustomobject]@{ Timestamp=$serverSample.Timestamp; TargetServer=$item.TargetServer; ComputerName=$serverSample.ComputerName; LogName=$e.LogName; EventTime=$e.TimeCreated; EventId=$e.Id; ProviderName=$e.ProviderName; Level=$e.LevelDisplayName; Message=$e.Message } }
    }

    Export-Rows -Rows $serverRows -Path (Join-Path $rawPath "ServerSamples_$day.csv") -Delimiter $settings.OutputDelimiter
    Export-Rows -Rows $processRows -Path (Join-Path $rawPath "Raw_ProcessSamples_$day.csv") -Delimiter $settings.OutputDelimiter
    Export-Rows -Rows $alertRows -Path (Join-Path $rawPath "AlertSamples_$day.csv") -Delimiter $settings.OutputDelimiter
    Export-Rows -Rows $categoryRows -Path (Join-Path $rawPath "CategorySummary_$day.csv") -Delimiter $settings.OutputDelimiter
    Export-Rows -Rows $eventRows -Path (Join-Path $rawPath "EventContext_$day.csv") -Delimiter $settings.OutputDelimiter
    $allServerRows += $serverRows; $allProcessRows += $processRows
    Write-Host ("[{0}] Runde {1} beendet. OK={2}, Fehler={3}" -f (Get-Date).ToString('HH:mm:ss'), $round, @($serverRows | Where-Object Status -eq 'OK').Count, @($serverRows | Where-Object Status -eq 'ERROR').Count)
    $remaining = [int][Math]::Floor(($endAt - (Get-Date)).TotalSeconds)
    if ($remaining -gt 0) { Start-Sleep -Seconds ([Math]::Min([int]$settings.IntervalSeconds, $remaining)) }
} while ((Get-Date) -lt $endAt)

$runEnd = Get-Date
$runSummaryRows = New-RunSummaryRows -ServerRows $allServerRows -ProcessRows $allProcessRows -Start $runStart -End $runEnd -ServerCount $servers.Count
Export-Rows -Rows $runSummaryRows -Path (Join-Path $summaryPath "RunSummary_$runId.csv") -Delimiter $settings.OutputDelimiter
$runSummaryRows
