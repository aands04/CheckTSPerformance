#requires -version 3.0
<#
.SYNOPSIS
    Collects Citrix/Terminalserver health samples via PowerShell Remoting.
.DESCRIPTION
    Windows PowerShell 5.1 compatible collector for longer CPU peak analysis runs.
    It writes separate CSV files for server samples, process samples, alerts,
    category summaries, optional event context and a per-run summary.

    CPU semantics:
    - ProcessCpuCorePercent is based on one logical processor. 100 means roughly
      one logical processor fully used during the delta window.
    - ProcessCpuServerPercent is normalized to the whole server and is calculated
      as ProcessCpuCorePercent / LogicalProcessorCount.
#>
[CmdletBinding()]
param(
    [string]$ServerListPath,
    [string]$ConfigPath,
    [string]$OutputPath,
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
    [int]$MaxEventsPerAlert,
    [string]$Delimiter
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$script:InvocationPath = $MyInvocation.MyCommand.Path


function Resolve-ScriptRoot {
    if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) { return $PSScriptRoot }
    if (-not [string]::IsNullOrWhiteSpace($script:InvocationPath)) { return (Split-Path -Parent $script:InvocationPath) }
    return (Get-Location).ProviderPath
}

$scriptRoot = Resolve-ScriptRoot
if ([string]::IsNullOrWhiteSpace($ServerListPath)) { $ServerListPath = Join-Path -Path $scriptRoot -ChildPath 'config\servers.txt' }
if ([string]::IsNullOrWhiteSpace($ConfigPath)) { $ConfigPath = Join-Path -Path $scriptRoot -ChildPath 'config\settings.json' }
if ([string]::IsNullOrWhiteSpace($OutputPath)) { $OutputPath = Join-Path -Path $scriptRoot -ChildPath 'output' }

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
        MaxEventsPerAlert = 50
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
    if ($PSBoundParameters.ContainsKey('MaxEventsPerAlert')) { $Settings.MaxEventsPerAlert = $MaxEventsPerAlert }
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

function Write-RunLog {
    param([string]$Path, [string]$Message, [string]$Level = 'INFO')
    $line = '{0};{1};{2}' -f (Get-Date).ToString('s'), $Level, $Message
    Add-Content -LiteralPath $Path -Value $line -Encoding UTF8
}

function Get-RemoteSamplerScriptBlock {
    {
        param($CpuSampleSeconds, $TopProcessCount, $AlertTopProcessCount, $CpuWarningThreshold, $CpuCriticalThreshold, $IncludeEventLogContext, $AnonymizeUsers, $MaxEventsPerAlert)

        function Get-ProcessCategory {
            param([string]$Name, [string]$ParentProcessName = '')
            if ($Name -match '^(MsSense|SenseNdr|MsMpEng|CylanceSvc|CSFalconService|SentinelAgent)$' -or $Name -match '^(Sophos|CarbonBlack|cb|Tanium).*') { return 'Security' }
            if ($Name -match '^(BrokerAgent|CtxGfx|Citrix\.Wem\.Agent\.Service|VUEMUIAgent|VUEMAppCmd|wfcrun32|wfica32|SelfService|Receiver|AuthManSvr)$') { return 'Citrix' }
            if ($Name -match '^(nexus\.framework\.healthcare|Infoclient|anm_neu)$') { return 'Nexus' }
            if ($Name -match '^(WINWORD|EXCEL|OUTLOOK|POWERPNT|ONENOTE|OfficeClickToRun)$') { return 'Office' }
            if ($Name -match '^(msedge|msedgewebview2|chrome|firefox|iexplore)$') { return 'Browser' }
            if ($Name -match '^(Acrobat|AcroRd32|AdobeCollabSync|AdobeARM)$') { return 'Adobe' }
            if ($Name -match '^(spoolsv|splwow64|PrintIsolationHost)$') { return 'Printing' }
            if ($Name -match '^(wsmprovhost|WmiPrvSE|powershell|pwsh)$') { return 'Monitoring' }
            if ($Name -eq 'conhost' -and $ParentProcessName -match '^(powershell|pwsh|wsmprovhost)$') { return 'Monitoring' }
            if ($Name -match '^(svchost|System|Registry|explorer|dwm|SearchIndexer|RuntimeBroker|StartMenuExperienceHost|ShellExperienceHost|taskhostw|sihost|csrss|lsass|services)$') { return 'Windows' }
            return 'Other'
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
        $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
        $logicalProcessorCount = [Math]::Max([int]$computerSystem.NumberOfLogicalProcessors, 1)
        $basicProcesses = @{}
        try {
            foreach ($bp in (Get-CimInstance -ClassName Win32_Process | Select-Object ProcessId, ParentProcessId, Name)) {
                $basicProcesses[[int]$bp.ProcessId] = $bp
            }
        } catch { }
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
            $cpuCorePct = [Math]::Round(($delta / $elapsedSeconds) * 100, 2)
            $cpuServerPct = [Math]::Round(($cpuCorePct / $logicalProcessorCount), 2)
            $basicProcess = $basicProcesses[[int]$process.Id]
            $parentProcessId = if ($basicProcess) { [int]$basicProcess.ParentProcessId } else { 0 }
            $parentBasicProcess = $basicProcesses[$parentProcessId]
            $parentProcessName = if ($parentBasicProcess) { [string]$parentBasicProcess.Name } else { '' }
            $category = Get-ProcessCategory $process.ProcessName $parentProcessName
            $startTimeText = ''
            try {
                if ($process.StartTime) { $startTimeText = $process.StartTime.ToString('s') }
            }
            catch { $startTimeText = '' }
            [pscustomobject]@{
                ProcessId = [int]$process.Id
                ProcessName = [string]$process.ProcessName
                SessionId = if ($null -ne $process.SessionId) { [int]$process.SessionId } else { -1 }
                ProcessCpuCorePercent = $cpuCorePct
                ProcessCpuServerPercent = $cpuServerPct
                CpuSecondsDelta = [Math]::Round($delta, 3)
                WorkingSetMb = [Math]::Round($process.WorkingSet64 / 1MB, 2)
                PrivateMemoryMb = if ($null -ne $process.PrivateMemorySize64) { [Math]::Round($process.PrivateMemorySize64 / 1MB, 2) } else { '' }
                ParentProcessId = $parentProcessId
                ParentProcessName = $parentProcessName
                StartTime = $startTimeText
                Category = $category
                IsMonitoringRelated = ($category -eq 'Monitoring')
            }
        }

        $isWarning = $cpuPercent -ge $CpuWarningThreshold
        $isCritical = $cpuPercent -ge $CpuCriticalThreshold
        $topLimit = if ($isWarning) { $AlertTopProcessCount } else { $TopProcessCount }
        $specialNames = @('Citrix.Wem.Agent.Service','VUEMUIAgent','WmiPrvSE','wsmprovhost')
        $selectedIds = @{}
        foreach ($p in ($processDeltas | Sort-Object ProcessCpuCorePercent -Descending | Select-Object -First $topLimit)) { $selectedIds[$p.ProcessId] = $true }
        if ($isWarning) { foreach ($p in ($processDeltas | Where-Object { $_.Category -in @('Security','Nexus','Printing') -or $specialNames -contains $_.ProcessName })) { $selectedIds[$p.ProcessId] = $true } }
        $selected = @($processDeltas | Where-Object { $selectedIds.ContainsKey($_.ProcessId) } | Sort-Object ProcessCpuCorePercent -Descending)

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
            $rawOwnerUser = $ownerUser
            if ([string]::IsNullOrWhiteSpace($ownerUser) -and $session) { $ownerUser = Convert-UserName $session.UserName ([bool]$AnonymizeUsers); $rawOwnerUser = $session.UserName }
            $isServiceOwner = ($rawOwnerUser -match '^(SYSTEM|LOCAL SERVICE|NETWORK SERVICE)$' -or $ownerDomain -eq 'NT AUTHORITY')
            $knownServiceCategory = ($p.Category -in @('Security','Citrix','Printing','Monitoring','Windows'))
            $isSystemProcess = (($p.SessionId -eq 0) -or $isServiceOwner -or ([string]::IsNullOrWhiteSpace($ownerUser) -and $knownServiceCategory))
            $isUserProcess = (($p.SessionId -gt 0) -and -not [string]::IsNullOrWhiteSpace($ownerUser) -and -not $isServiceOwner)
            [pscustomobject]@{
                ProcessName = $p.ProcessName
                ProcessRank = 0
                ProcessId = $p.ProcessId
                ProcessSessionId = $p.SessionId
                ProcessUserName = $ownerUser
                ProcessUserDomain = $ownerDomain
                SessionState = if ($session) { $session.State } else { '' }
                ProcessCpuCorePercent = $p.ProcessCpuCorePercent
                ProcessCpuServerPercent = $p.ProcessCpuServerPercent
                ProcessCpuSecondsDelta = $p.CpuSecondsDelta
                ProcessWorkingSetMB = $p.WorkingSetMb
                ProcessPrivateMemoryMB = $p.PrivateMemoryMb
                ProcessPath = if ($cim) { $cim.ExecutablePath } else { '' }
                ProcessCommandLine = if ($cim) { $cim.CommandLine } else { '' }
                ParentProcessId = $p.ParentProcessId
                ParentProcessName = $p.ParentProcessName
                ProcessStartTime = $p.StartTime
                IsSystemProcess = $isSystemProcess
                IsUserProcess = $isUserProcess
                IsMonitoringRelated = $p.IsMonitoringRelated
                Category = $p.Category
            }
        }

        $categoryNames = @('Security','Nexus','Office','Browser','Adobe','Citrix','Printing','Monitoring','Windows','Other')
        $categorySums = @{}
        foreach ($name in $categoryNames) { $categorySums[$name] = [pscustomobject]@{ Core = 0.0; Server = 0.0 } }
        foreach ($p in $processDeltas) { $categorySums[$p.Category].Core += [double]$p.ProcessCpuCorePercent; $categorySums[$p.Category].Server += [double]$p.ProcessCpuServerPercent }
        $topCategoryByServer = ($categoryNames | Sort-Object { -1 * [double]$categorySums[$_].Server } | Select-Object -First 1)

        $events = @()
        if ($IncludeEventLogContext -and $isWarning) {
            $since = (Get-Date).AddMinutes(-5)
            foreach ($logName in @('System','Application','Microsoft-Windows-Windows Defender/Operational')) {
                try {
                    $events += Get-WinEvent -FilterHashtable @{ LogName = $logName; StartTime = $since } -MaxEvents $MaxEventsPerAlert -ErrorAction Stop |
                        Select-Object @{n='LogName';e={$logName}}, TimeCreated, Id, ProviderName, LevelDisplayName, Message
                } catch { }
            }
        }

        [pscustomobject]@{
            ServerSample = [pscustomobject]@{
                Timestamp = $sampleEnd.ToString('s')
                ComputerName = $env:COMPUTERNAME
                LogicalProcessorCount = $logicalProcessorCount
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
                AlertSeverity = if ($isCritical) { 'Critical' } elseif ($isWarning) { 'Warning' } else { 'OK' }
                TopCategoryByCpuServerPercent = $topCategoryByServer
                TopProcessCpuCorePercentSum = [Math]::Round((($selected | Measure-Object -Property ProcessCpuCorePercent -Sum).Sum), 2)
                TopProcessCpuServerPercentSum = [Math]::Round((($selected | Measure-Object -Property ProcessCpuServerPercent -Sum).Sum), 2)
            }
            ProcessSamples = @($processSamples)
            CategorySummary = [pscustomobject]@{
                TopCategoryByCpuServerPercent = $topCategoryByServer
                SecurityCpuCorePercent = [Math]::Round($categorySums['Security'].Core, 2); SecurityCpuServerPercent = [Math]::Round($categorySums['Security'].Server, 2)
                NexusCpuCorePercent = [Math]::Round($categorySums['Nexus'].Core, 2); NexusCpuServerPercent = [Math]::Round($categorySums['Nexus'].Server, 2)
                OfficeCpuCorePercent = [Math]::Round($categorySums['Office'].Core, 2); OfficeCpuServerPercent = [Math]::Round($categorySums['Office'].Server, 2)
                BrowserCpuCorePercent = [Math]::Round($categorySums['Browser'].Core, 2); BrowserCpuServerPercent = [Math]::Round($categorySums['Browser'].Server, 2)
                AdobeCpuCorePercent = [Math]::Round($categorySums['Adobe'].Core, 2); AdobeCpuServerPercent = [Math]::Round($categorySums['Adobe'].Server, 2)
                CitrixCpuCorePercent = [Math]::Round($categorySums['Citrix'].Core, 2); CitrixCpuServerPercent = [Math]::Round($categorySums['Citrix'].Server, 2)
                PrintingCpuCorePercent = [Math]::Round($categorySums['Printing'].Core, 2); PrintingCpuServerPercent = [Math]::Round($categorySums['Printing'].Server, 2)
                MonitoringCpuCorePercent = [Math]::Round($categorySums['Monitoring'].Core, 2); MonitoringCpuServerPercent = [Math]::Round($categorySums['Monitoring'].Server, 2)
                WindowsCpuCorePercent = [Math]::Round($categorySums['Windows'].Core, 2); WindowsCpuServerPercent = [Math]::Round($categorySums['Windows'].Server, 2)
                OtherCpuCorePercent = [Math]::Round($categorySums['Other'].Core, 2); OtherCpuServerPercent = [Math]::Round($categorySums['Other'].Server, 2)
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
                $result = Invoke-Command -ComputerName $Server -ScriptBlock $remoteBlock -ArgumentList $SettingsHash.CpuSampleSeconds, $SettingsHash.TopProcessCount, $SettingsHash.AlertTopProcessCount, $SettingsHash.CpuWarningThreshold, $SettingsHash.CpuCriticalThreshold, $SettingsHash.IncludeEventLogContext, $SettingsHash.AnonymizeUsers, $SettingsHash.MaxEventsPerAlert -ErrorAction Stop
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
        Timestamp = $Timestamp.ToString('s'); TargetServer = $Server; ComputerName = ''; LogicalProcessorCount = ''; CpuPercent = ''; MemoryPercent = ''; TotalMemoryMB = ''; UsedMemoryMB = ''; FreeMemoryMB = '';
        ActiveSessions = ''; DisconnectedSessions = ''; TotalSessions = ''; Status = 'ERROR'; ErrorMessage = $Message; IsCpuWarning = $false; IsCpuCritical = $false; AlertSeverity = 'Error';
        TopCategoryByCpuServerPercent = ''; SecurityCpuServerPercent = ''; NexusCpuServerPercent = ''; OfficeCpuServerPercent = ''; BrowserCpuServerPercent = ''; CitrixCpuServerPercent = ''; PrintingCpuServerPercent = ''; MonitoringCpuServerPercent = ''; WindowsCpuServerPercent = ''; OtherCpuServerPercent = ''
    }
}

function Get-PercentileValue {
    param([double[]]$Values, [double]$Percentile)
    if (-not $Values -or $Values.Count -eq 0) { return '' }
    $sorted = @($Values | Sort-Object)
    $index = [Math]::Ceiling(($Percentile / 100) * $sorted.Count) - 1
    $index = [Math]::Max(0, [Math]::Min($index, $sorted.Count - 1))
    return $sorted[$index]
}

function Join-TopCpuGroups {
    param([array]$Rows, [string]$GroupProperty, [int]$Count = 10)
    @($Rows | Group-Object $GroupProperty | ForEach-Object {
        [pscustomobject]@{ Name = $_.Name; Cpu = (($_.Group | Measure-Object ProcessCpuServerPercent -Sum).Sum) }
    } | Sort-Object Cpu -Descending | Select-Object -First $Count | ForEach-Object { '{0}={1:n2}' -f $_.Name,$_.Cpu }) -join ', '
}

function New-RunSummaryRows {
    param([array]$ServerRows, [array]$ProcessRows, [datetime]$Start, [datetime]$End, [int]$ServerCount)
    $categoryNames = @('Security','Nexus','Office','Browser','Adobe','Citrix','Printing','Monitoring','Windows','Other')
    foreach ($group in ($ServerRows | Group-Object TargetServer)) {
        $server = $group.Name
        $okRows = @($group.Group | Where-Object Status -eq 'OK')
        $cpuValues = @($okRows | Where-Object { $_.CpuPercent -ne '' } | ForEach-Object { [double]$_.CpuPercent } | Sort-Object)
        $ramValues = @($okRows | Where-Object { $_.MemoryPercent -ne '' } | ForEach-Object { [double]$_.MemoryPercent })
        $activeValues = @($okRows | Where-Object { $_.ActiveSessions -ne '' } | ForEach-Object { [double]$_.ActiveSessions })
        $discValues = @($okRows | Where-Object { $_.DisconnectedSessions -ne '' } | ForEach-Object { [double]$_.DisconnectedSessions })
        $median = Get-PercentileValue -Values $cpuValues -Percentile 50
        $p95 = Get-PercentileValue -Values $cpuValues -Percentile 95
        $serverProcessRows = @($ProcessRows | Where-Object TargetServer -eq $server)
        $topProc = Join-TopCpuGroups -Rows $serverProcessRows -GroupProperty ProcessName -Count 1
        $top10Proc = Join-TopCpuGroups -Rows $serverProcessRows -GroupProperty ProcessName -Count 10
        $topCat = Join-TopCpuGroups -Rows $serverProcessRows -GroupProperty Category -Count 1
        $categoryTotals = @{}
        foreach ($category in $categoryNames) { $categoryTotals[$category] = (($serverProcessRows | Where-Object Category -eq $category | Measure-Object ProcessCpuServerPercent -Sum).Sum) }
        [pscustomobject]@{
            RunStart = $Start.ToString('s'); RunEnd = $End.ToString('s'); DurationMinutes = [Math]::Round(($End-$Start).TotalMinutes, 2); ServerCount = $ServerCount; TargetServer = $server;
            SampleCount = $okRows.Count; ErrorCount = @($group.Group | Where-Object Status -eq 'ERROR').Count;
            CpuAverage = if ($cpuValues.Count) { [Math]::Round((($cpuValues | Measure-Object -Average).Average),2) } else { '' }; CpuMedian = if ($median -ne '') { [Math]::Round($median,2) } else { '' }; CpuMax = if ($cpuValues.Count) { [Math]::Round((($cpuValues | Measure-Object -Maximum).Maximum),2) } else { '' }; CpuP95 = if ($p95 -ne '') { [Math]::Round($p95,2) } else { '' };
            CpuWarningCount = @($okRows | Where-Object { $_.IsCpuWarning -eq $true }).Count; CpuCriticalCount = @($okRows | Where-Object { $_.IsCpuCritical -eq $true }).Count;
            MemoryAverage = if ($ramValues.Count) { [Math]::Round((($ramValues | Measure-Object -Average).Average),2) } else { '' }; MemoryMax = if ($ramValues.Count) { [Math]::Round((($ramValues | Measure-Object -Maximum).Maximum),2) } else { '' };
            ActiveSessionsAverage = if ($activeValues.Count) { [Math]::Round((($activeValues | Measure-Object -Average).Average),2) } else { '' }; ActiveSessionsMax = if ($activeValues.Count) { [Math]::Round((($activeValues | Measure-Object -Maximum).Maximum),2) } else { '' };
            DisconnectedSessionsAverage = if ($discValues.Count) { [Math]::Round((($discValues | Measure-Object -Average).Average),2) } else { '' };
            TopProcessByTotalCpuServerPercent = $topProc; Top10ProcessesByTotalCpuServerPercent = $top10Proc; TopCategoryByTotalCpuServerPercent = $topCat;
            SecurityTotalCpuServerPercent = [Math]::Round($categoryTotals['Security'],2); NexusTotalCpuServerPercent = [Math]::Round($categoryTotals['Nexus'],2); OfficeTotalCpuServerPercent = [Math]::Round($categoryTotals['Office'],2); BrowserTotalCpuServerPercent = [Math]::Round($categoryTotals['Browser'],2); AdobeTotalCpuServerPercent = [Math]::Round($categoryTotals['Adobe'],2); CitrixTotalCpuServerPercent = [Math]::Round($categoryTotals['Citrix'],2); PrintingTotalCpuServerPercent = [Math]::Round($categoryTotals['Printing'],2); MonitoringTotalCpuServerPercent = [Math]::Round($categoryTotals['Monitoring'],2); WindowsTotalCpuServerPercent = [Math]::Round($categoryTotals['Windows'],2); OtherTotalCpuServerPercent = [Math]::Round($categoryTotals['Other'],2)
        }
    }
}

function New-RunSummaryText {
    param([array]$SummaryRows, [datetime]$Start, [datetime]$End)
    $lines = @()
    $lines += "Citrix-TS-HealthCheck RunSummary"
    $lines += "Start: $($Start.ToString('s'))"
    $lines += "Ende : $($End.ToString('s'))"
    $lines += ''
    foreach ($row in $SummaryRows) {
        $lines += "Server $($row.TargetServer): CPU avg=$($row.CpuAverage) max=$($row.CpuMax) p95=$($row.CpuP95), Warnings=$($row.CpuWarningCount), Criticals=$($row.CpuCriticalCount), TopCategory=$($row.TopCategoryByTotalCpuServerPercent)"
        if ([double]($row.CpuMax -as [double]) -ge 70) {
            if ([double]($row.SecurityTotalCpuServerPercent -as [double]) -gt 0) { $lines += '  Bewertung: CPU hoch + Security hoch: Defender/MDE/Cylance pruefen.' }
            if ([double]($row.NexusTotalCpuServerPercent -as [double]) -gt 0) { $lines += '  Bewertung: CPU hoch + Nexus hoch: Nexus/Fachanwendung pruefen.' }
            if ([double]($row.PrintingTotalCpuServerPercent -as [double]) -gt 0) { $lines += '  Bewertung: CPU hoch + Printing hoch: Druck/Spooler pruefen.' }
            if ([double]($row.CitrixTotalCpuServerPercent -as [double]) -gt 0) { $lines += '  Bewertung: CPU hoch + Citrix/WEM hoch: WEM-Agent/VUEM pruefen.' }
            if ([double]($row.MonitoringTotalCpuServerPercent -as [double]) -gt 0) { $lines += '  Bewertung: CPU hoch + Monitoring hoch: WMI/Monitoring-Overhead beachten.' }
            if ([double]($row.ActiveSessionsAverage -as [double]) -lt 10) { $lines += '  Bewertung: CPU hoch + wenige Sessions: eher Prozess-/Dienstproblem.' } else { $lines += '  Bewertung: CPU hoch + viele Sessions: eher Sizing/Kapazitaet pruefen.' }
        }
    }
    return ($lines -join [Environment]::NewLine)
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
$runLogFile = Join-Path $logPath "RunLog_$runId.log"
$allServerRows = @()
$allProcessRows = @()
$round = 0
$endAt = $runStart.AddMinutes([int]$settings.DurationMinutes)
if ([int]$settings.DurationMinutes -lt 1) { $endAt = $runStart }
$runError = $null

try {
    Write-RunLog -Path $runLogFile -Message "Run gestartet. Server=$($servers.Count), DurationMinutes=$($settings.DurationMinutes), IntervalSeconds=$($settings.IntervalSeconds), MaxParallel=$($settings.MaxParallel)"
    do {
        $round++
        $roundStart = Get-Date
        $message = "Runde $round gestartet. Server=$($servers.Count), CpuSampleSeconds=$($settings.CpuSampleSeconds), Parallel=$($settings.MaxParallel)"
        Write-Host ("[{0}] {1}" -f $roundStart.ToString('HH:mm:ss'), $message)
        Write-RunLog -Path $runLogFile -Message $message
        $roundResults = @(Invoke-ServerCollectionRound -Servers $servers -Settings $settings -RoundTimestamp $roundStart)
        $day = (Get-Date).ToString('yyyy-MM-dd')
        $serverRows = @(); $processRows = @(); $alertRows = @(); $categoryRows = @(); $eventRows = @()

        foreach ($item in $roundResults) {
            if ($item.Status -ne 'OK') {
                $serverRows += New-ErrorServerSample -Server $item.TargetServer -Message $item.ErrorMessage -Timestamp $roundStart
                Write-RunLog -Path $runLogFile -Level 'ERROR' -Message "$($item.TargetServer): $($item.ErrorMessage)"
                continue
            }
            $serverSample = $item.Result.ServerSample
            $cat = $item.Result.CategorySummary
            $serverRow = [pscustomobject]@{
                Timestamp=$serverSample.Timestamp; TargetServer=$item.TargetServer; ComputerName=$serverSample.ComputerName; LogicalProcessorCount=$serverSample.LogicalProcessorCount; CpuPercent=$serverSample.CpuPercent; MemoryPercent=$serverSample.MemoryPercent;
                TotalMemoryMB=$serverSample.TotalMemoryMB; UsedMemoryMB=$serverSample.UsedMemoryMB; FreeMemoryMB=$serverSample.FreeMemoryMB; ActiveSessions=$serverSample.ActiveSessions; DisconnectedSessions=$serverSample.DisconnectedSessions; TotalSessions=$serverSample.TotalSessions;
                Status='OK'; ErrorMessage=''; IsCpuWarning=$serverSample.IsCpuWarning; IsCpuCritical=$serverSample.IsCpuCritical; TopCategoryByCpuServerPercent=$serverSample.TopCategoryByCpuServerPercent;
                SecurityCpuServerPercent=$cat.SecurityCpuServerPercent; NexusCpuServerPercent=$cat.NexusCpuServerPercent; OfficeCpuServerPercent=$cat.OfficeCpuServerPercent; BrowserCpuServerPercent=$cat.BrowserCpuServerPercent; CitrixCpuServerPercent=$cat.CitrixCpuServerPercent; PrintingCpuServerPercent=$cat.PrintingCpuServerPercent; MonitoringCpuServerPercent=$cat.MonitoringCpuServerPercent; WindowsCpuServerPercent=$cat.WindowsCpuServerPercent; OtherCpuServerPercent=$cat.OtherCpuServerPercent
            }
            $serverRows += $serverRow
            $rank = 1
            foreach ($p in @($item.Result.ProcessSamples | Sort-Object ProcessCpuCorePercent -Descending)) {
                $processRows += [pscustomobject]@{
                    Timestamp=$serverSample.Timestamp; TargetServer=$item.TargetServer; ComputerName=$serverSample.ComputerName; LogicalProcessorCount=$serverSample.LogicalProcessorCount; CpuPercent=$serverSample.CpuPercent; MemoryPercent=$serverSample.MemoryPercent; ActiveSessions=$serverSample.ActiveSessions; DisconnectedSessions=$serverSample.DisconnectedSessions;
                    ProcessRank=$rank; ProcessId=$p.ProcessId; ParentProcessId=$p.ParentProcessId; ParentProcessName=$p.ParentProcessName; ProcessName=$p.ProcessName; ProcessSessionId=$p.ProcessSessionId; ProcessUserName=$p.ProcessUserName; ProcessUserDomain=$p.ProcessUserDomain; SessionState=$p.SessionState;
                    ProcessCpuCorePercent=$p.ProcessCpuCorePercent; ProcessCpuServerPercent=$p.ProcessCpuServerPercent; ProcessCpuSecondsDelta=$p.ProcessCpuSecondsDelta; ProcessWorkingSetMB=$p.ProcessWorkingSetMB; ProcessPrivateMemoryMB=$p.ProcessPrivateMemoryMB; ProcessPath=$p.ProcessPath; ProcessCommandLine=$p.ProcessCommandLine; ProcessStartTime=$p.ProcessStartTime;
                    Category=$p.Category; IsSystemProcess=$p.IsSystemProcess; IsUserProcess=$p.IsUserProcess; IsMonitoringRelated=$p.IsMonitoringRelated; Status='OK'; ErrorMessage=''
                }
                $rank++
            }
            $categoryRows += [pscustomobject]@{
                Timestamp=$serverSample.Timestamp; TargetServer=$item.TargetServer; ComputerName=$serverSample.ComputerName; LogicalProcessorCount=$serverSample.LogicalProcessorCount; CpuPercent=$serverSample.CpuPercent; ActiveSessions=$serverSample.ActiveSessions; DisconnectedSessions=$serverSample.DisconnectedSessions; TotalSessions=$serverSample.TotalSessions;
                SecurityCpuCorePercent=$cat.SecurityCpuCorePercent; SecurityCpuServerPercent=$cat.SecurityCpuServerPercent; NexusCpuCorePercent=$cat.NexusCpuCorePercent; NexusCpuServerPercent=$cat.NexusCpuServerPercent; OfficeCpuCorePercent=$cat.OfficeCpuCorePercent; OfficeCpuServerPercent=$cat.OfficeCpuServerPercent; BrowserCpuCorePercent=$cat.BrowserCpuCorePercent; BrowserCpuServerPercent=$cat.BrowserCpuServerPercent; AdobeCpuCorePercent=$cat.AdobeCpuCorePercent; AdobeCpuServerPercent=$cat.AdobeCpuServerPercent; CitrixCpuCorePercent=$cat.CitrixCpuCorePercent; CitrixCpuServerPercent=$cat.CitrixCpuServerPercent; PrintingCpuCorePercent=$cat.PrintingCpuCorePercent; PrintingCpuServerPercent=$cat.PrintingCpuServerPercent; MonitoringCpuCorePercent=$cat.MonitoringCpuCorePercent; MonitoringCpuServerPercent=$cat.MonitoringCpuServerPercent; WindowsCpuCorePercent=$cat.WindowsCpuCorePercent; WindowsCpuServerPercent=$cat.WindowsCpuServerPercent; OtherCpuCorePercent=$cat.OtherCpuCorePercent; OtherCpuServerPercent=$cat.OtherCpuServerPercent; TopCategoryByCpuServerPercent=$cat.TopCategoryByCpuServerPercent
            }
            if ($serverSample.IsCpuWarning) {
                foreach ($p in @($item.Result.ProcessSamples)) {
                    $alertRows += [pscustomobject]@{
                        Timestamp=$serverSample.Timestamp; TargetServer=$item.TargetServer; ComputerName=$serverSample.ComputerName; LogicalProcessorCount=$serverSample.LogicalProcessorCount; AlertSeverity=$serverSample.AlertSeverity; AlertTopCategory=$cat.TopCategoryByCpuServerPercent; CpuPercent=$serverSample.CpuPercent; MemoryPercent=$serverSample.MemoryPercent; ActiveSessions=$serverSample.ActiveSessions; DisconnectedSessions=$serverSample.DisconnectedSessions; TotalSessions=$serverSample.TotalSessions;
                        SecurityCpuServerPercent=$cat.SecurityCpuServerPercent; NexusCpuServerPercent=$cat.NexusCpuServerPercent; OfficeCpuServerPercent=$cat.OfficeCpuServerPercent; BrowserCpuServerPercent=$cat.BrowserCpuServerPercent; CitrixCpuServerPercent=$cat.CitrixCpuServerPercent; PrintingCpuServerPercent=$cat.PrintingCpuServerPercent; MonitoringCpuServerPercent=$cat.MonitoringCpuServerPercent; WindowsCpuServerPercent=$cat.WindowsCpuServerPercent; OtherCpuServerPercent=$cat.OtherCpuServerPercent;
                        ProcessName=$p.ProcessName; ProcessId=$p.ProcessId; ProcessCpuCorePercent=$p.ProcessCpuCorePercent; ProcessCpuServerPercent=$p.ProcessCpuServerPercent; ProcessCpuSecondsDelta=$p.ProcessCpuSecondsDelta; Category=$p.Category; IsMonitoringRelated=$p.IsMonitoringRelated; ProcessUserName=$p.ProcessUserName; ProcessUserDomain=$p.ProcessUserDomain; ProcessSessionId=$p.ProcessSessionId; ProcessCommandLine=$p.ProcessCommandLine
                    }
                }
            }
            foreach ($e in @($item.Result.EventSamples)) { $eventRows += [pscustomobject]@{ Timestamp=$serverSample.Timestamp; TargetServer=$item.TargetServer; ComputerName=$serverSample.ComputerName; LogName=$e.LogName; EventTime=$e.TimeCreated; EventId=$e.Id; ProviderName=$e.ProviderName; Level=$e.LevelDisplayName; Message=$e.Message } }
        }

        Export-Rows -Rows $serverRows -Path (Join-Path $rawPath "ServerSamples_$day.csv") -Delimiter $settings.OutputDelimiter
        Export-Rows -Rows $processRows -Path (Join-Path $rawPath "Raw_ProcessSamples_$day.csv") -Delimiter $settings.OutputDelimiter
        Export-Rows -Rows $alertRows -Path (Join-Path $rawPath "AlertSamples_$day.csv") -Delimiter $settings.OutputDelimiter
        Export-Rows -Rows $categoryRows -Path (Join-Path $rawPath "CategorySummary_$day.csv") -Delimiter $settings.OutputDelimiter
        Export-Rows -Rows $eventRows -Path (Join-Path $rawPath "EventContext_$day.csv") -Delimiter $settings.OutputDelimiter
        $allServerRows += $serverRows; $allProcessRows += $processRows
        $doneMessage = "Runde $round beendet. OK=$(@($serverRows | Where-Object Status -eq 'OK').Count), Fehler=$(@($serverRows | Where-Object Status -eq 'ERROR').Count)"
        Write-Host ("[{0}] {1}" -f (Get-Date).ToString('HH:mm:ss'), $doneMessage)
        Write-RunLog -Path $runLogFile -Message $doneMessage
        $remaining = [int][Math]::Floor(($endAt - (Get-Date)).TotalSeconds)
        if ($remaining -gt 0) { Start-Sleep -Seconds ([Math]::Min([int]$settings.IntervalSeconds, $remaining)) }
    } while ((Get-Date) -lt $endAt)
}
catch {
    $runError = $_.Exception.Message
    Write-RunLog -Path $runLogFile -Level 'ERROR' -Message "Runfehler: $runError"
    Write-Warning $runError
}
finally {
    $runEnd = Get-Date
    try {
        $runSummaryRows = @(New-RunSummaryRows -ServerRows $allServerRows -ProcessRows $allProcessRows -Start $runStart -End $runEnd -ServerCount $servers.Count)
        $summaryCsv = Join-Path $summaryPath "RunSummary_$runId.csv"
        $summaryTxt = Join-Path $summaryPath "RunSummary_$runId.txt"
        Export-Rows -Rows $runSummaryRows -Path $summaryCsv -Delimiter $settings.OutputDelimiter
        New-RunSummaryText -SummaryRows $runSummaryRows -Start $runStart -End $runEnd | Set-Content -LiteralPath $summaryTxt -Encoding UTF8
        Write-RunLog -Path $runLogFile -Message "RunSummary geschrieben: $summaryCsv"
        $runSummaryRows
    }
    catch {
        Write-RunLog -Path $runLogFile -Level 'ERROR' -Message "RunSummary konnte nicht geschrieben werden: $($_.Exception.Message)"
        Write-Warning $_.Exception.Message
    }
}
