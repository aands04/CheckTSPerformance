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
    [int]$MaxForcedProcessesPerCategory,
    [switch]$AutoDefenderPerfRecording,
    [double]$DefenderPerfTriggerServerCpuPercent,
    [int]$DefenderPerfRecordingSeconds,
    [int]$DefenderPerfCooldownMinutes,
    [int]$MaxConcurrentDefenderPerfRecordings,
    [switch]$IncludeWemEventContext,
    [double]$WemTriggerServerCpuPercent,
    [int]$WemEventWindowMinutes,
    [switch]$IncludeWemLogTail,
    [int]$WemLogTailLines,
    [switch]$IncludeEventLogContext,
    [switch]$IncludeEventContext,
    [switch]$AnonymizeUsers,
    [int]$MaxEventsPerAlert,
    [string]$ImageVersion,
    [string]$Notes,
    [string]$RunId,
    [switch]$IncludeScheduledTaskInventory,
    [switch]$IncludeCylanceHealth,
    [string[]]$TaskNamesToCheck,
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
        MaxForcedProcessesPerCategory = 10
        AutoDefenderPerfRecording = $false
        DefenderPerfTriggerServerCpuPercent = 10
        DefenderPerfRecordingSeconds = 900
        DefenderPerfCooldownMinutes = 120
        MaxConcurrentDefenderPerfRecordings = 2
        IncludeWemEventContext = $false
        WemTriggerServerCpuPercent = 10
        WemEventWindowMinutes = 10
        IncludeWemLogTail = $false
        WemLogTailLines = 200
        IncludeEventLogContext = $false
        IncludeEventContext = $false
        AnonymizeUsers = $false
        MaxEventsPerAlert = 50
        ImageVersion = ''
        Notes = ''
        RunId = ''
        IncludeScheduledTaskInventory = $false
        IncludeCylanceHealth = $false
        TaskNamesToCheck = @('nWizard_{B2FE1952-0186-46C3-BAEC-A80AA35AC5B8}','Adobe Acrobat Update Task','MicrosoftEdgeUpdateTaskMachineUA','Launch Adobe CCXProcess','LexwareAppSysOpt','Office Automatic Updates 2.0','Office Feature Updates','Office Feature Updates Logon','BackgroundDownload')
        OutputDelimiter = ';'
    }
}

function Read-Settings {
    param([string]$Path)
    $settings = New-DefaultSettings
    if (Test-Path -LiteralPath $Path) {
        $json = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($property in $settings.PSObject.Properties.Name) {
            if ($json.PSObject.Properties.Name -contains $property -and $null -ne $json.$property) { $settings.$property = $json.$property }
        }
    }
    return $settings
}


function Ensure-SettingProperty {
    param($Settings, [string]$Name, $DefaultValue)
    if (-not ($Settings.PSObject.Properties.Name -contains $Name)) {
        $Settings | Add-Member -MemberType NoteProperty -Name $Name -Value $DefaultValue
    }
    elseif ($null -eq $Settings.$Name) {
        $Settings.$Name = $DefaultValue
    }
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
    if ($PSBoundParameters.ContainsKey('MaxForcedProcessesPerCategory')) { $Settings.MaxForcedProcessesPerCategory = $MaxForcedProcessesPerCategory }
    if ($PSBoundParameters.ContainsKey('AutoDefenderPerfRecording')) { $Settings.AutoDefenderPerfRecording = [bool]$AutoDefenderPerfRecording }
    if ($PSBoundParameters.ContainsKey('DefenderPerfTriggerServerCpuPercent')) { $Settings.DefenderPerfTriggerServerCpuPercent = $DefenderPerfTriggerServerCpuPercent }
    if ($PSBoundParameters.ContainsKey('DefenderPerfRecordingSeconds')) { $Settings.DefenderPerfRecordingSeconds = $DefenderPerfRecordingSeconds }
    if ($PSBoundParameters.ContainsKey('DefenderPerfCooldownMinutes')) { $Settings.DefenderPerfCooldownMinutes = $DefenderPerfCooldownMinutes }
    if ($PSBoundParameters.ContainsKey('MaxConcurrentDefenderPerfRecordings')) { $Settings.MaxConcurrentDefenderPerfRecordings = $MaxConcurrentDefenderPerfRecordings }
    if ($PSBoundParameters.ContainsKey('IncludeWemEventContext')) { $Settings.IncludeWemEventContext = [bool]$IncludeWemEventContext }
    if ($PSBoundParameters.ContainsKey('WemTriggerServerCpuPercent')) { $Settings.WemTriggerServerCpuPercent = $WemTriggerServerCpuPercent }
    if ($PSBoundParameters.ContainsKey('WemEventWindowMinutes')) { $Settings.WemEventWindowMinutes = $WemEventWindowMinutes }
    if ($PSBoundParameters.ContainsKey('IncludeWemLogTail')) { $Settings.IncludeWemLogTail = [bool]$IncludeWemLogTail }
    if ($PSBoundParameters.ContainsKey('WemLogTailLines')) { $Settings.WemLogTailLines = $WemLogTailLines }
    if ($PSBoundParameters.ContainsKey('Delimiter')) { $Settings.OutputDelimiter = $Delimiter }
    if ($PSBoundParameters.ContainsKey('IncludeEventLogContext')) { $Settings.IncludeEventLogContext = [bool]$IncludeEventLogContext }
    if ($PSBoundParameters.ContainsKey('IncludeEventContext')) { $Settings.IncludeEventLogContext = [bool]$IncludeEventContext; $Settings.IncludeEventContext = [bool]$IncludeEventContext }
    if ($PSBoundParameters.ContainsKey('AnonymizeUsers')) { $Settings.AnonymizeUsers = [bool]$AnonymizeUsers }
    if ($PSBoundParameters.ContainsKey('MaxEventsPerAlert')) { $Settings.MaxEventsPerAlert = $MaxEventsPerAlert }
    if ($PSBoundParameters.ContainsKey('ImageVersion')) { $Settings.ImageVersion = $ImageVersion }
    if ($PSBoundParameters.ContainsKey('Notes')) { $Settings.Notes = $Notes }
    if ($PSBoundParameters.ContainsKey('RunId')) { $Settings.RunId = $RunId }
    if ($PSBoundParameters.ContainsKey('IncludeScheduledTaskInventory')) { $Settings.IncludeScheduledTaskInventory = [bool]$IncludeScheduledTaskInventory }
    if ($PSBoundParameters.ContainsKey('IncludeCylanceHealth')) { $Settings.IncludeCylanceHealth = [bool]$IncludeCylanceHealth }
    if ($PSBoundParameters.ContainsKey('TaskNamesToCheck')) { $Settings.TaskNamesToCheck = $TaskNamesToCheck }
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
        param($CpuSampleSeconds, $TopProcessCount, $AlertTopProcessCount, $CpuWarningThreshold, $CpuCriticalThreshold, $IncludeEventLogContext, $AnonymizeUsers, $MaxEventsPerAlert, $MaxForcedProcessesPerCategory)

        function Get-ProcessCategory {
            param([string]$Name, [string]$ParentProcessName = '', [string]$Path = '', [string]$CommandLine = '')
            if ($Name -match '^(MsSense|SenseNdr|MsMpEng|NisSrv|SecurityHealthService|CylanceSvc|CylanceUI|CSFalconService|SentinelAgent)$' -or $Name -match '^(Sophos|CarbonBlack|cb|Tanium).*') { return 'Security' }
            if ($Name -match '^(Citrix\.Wem\.Agent\.Service|VUEMUIAgent|VUEMAppCmd)$') { return 'Wem' }
            if ($Name -match '^(BrokerAgent|CtxGfx|ctxsvc|concentr|UserProfileManager|wfcrun32|wfica32|SelfService|Receiver|AuthManSvr)$' -or $Path -like '*\Citrix\*') { return 'Citrix' }
            if ($Name -match '^(nexus\.framework\.healthcare|Infoclient|anm_neu|mc)$' -or $Path -like '*\Nexus\Prog\*' -or ($Name -eq 'cefsharp.browsersubprocess' -and $Path -like '*\Nexus\Prog\*')) { return 'Nexus' }
            if ($Name -match '^(WINWORD|EXCEL|OUTLOOK|POWERPNT|ONENOTE|OfficeClickToRun|sdxhelper|OfficeC2RClient)$') { return 'Office' }
            if ($Name -match '^(msedge|msedgewebview2|MicrosoftEdgeUpdate)$') { return 'Edge' }
            if ($Name -match '^(chrome|firefox|iexplore)$') { return 'Browser' }
            if ($Name -match '^(Acrobat|AcroRd32|AcroCEF|AdobeIPCBroker|armsvc|AdobeARM|AdobeCollabSync|CCXProcess)$' -or $Path -like '*\Adobe\*') { return 'Adobe' }
            if ($Name -match '^(spoolsv|splwow64|PrintIsolationHost)$') { return 'Printing' }
            if ($Name -match '^(wsmprovhost|WmiPrvSE|powershell|pwsh|uberAgent|uberAgentSvc|uberAgentHelper)$') { return 'Monitoring' }
            if ($Name -eq 'conhost' -and $ParentProcessName -match '^(powershell|pwsh|wsmprovhost)$') { return 'Monitoring' }
            if ($Name -match '^(svchost|System|Registry|explorer|dwm|spoolsv|SearchApp|SearchIndexer|RuntimeBroker|StartMenuExperienceHost|ShellExperienceHost|taskhostw|sihost|csrss|lsass|services)$') { return 'Windows' }
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
            foreach ($bp in (Get-CimInstance -ClassName Win32_Process | Select-Object ProcessId, ParentProcessId, Name, ExecutablePath, CommandLine)) {
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
            $pathForCategory = if ($basicProcess) { [string]$basicProcess.ExecutablePath } else { '' }
            $cmdForCategory = if ($basicProcess) { [string]$basicProcess.CommandLine } else { '' }
            $category = Get-ProcessCategory $process.ProcessName $parentProcessName $pathForCategory $cmdForCategory
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
        $coreForcedNames = @('MsMpEng','MsSense','CylanceSvc','Citrix.Wem.Agent.Service','VUEMUIAgent')
        $forcedProcessRules = @(
            [pscustomobject]@{ Category='Security'; Reason='ForcedSecurity'; Names=@('MsMpEng','MsSense','SenseNdr','CylanceSvc') },
            [pscustomobject]@{ Category='Wem'; Reason='ForcedWEM'; Names=@('Citrix.Wem.Agent.Service','VUEMUIAgent') },
            [pscustomobject]@{ Category='Citrix'; Reason='ForcedCitrix'; Names=@('BrokerAgent','ctxsvc','UserProfileManager') },
            [pscustomobject]@{ Category='Nexus'; Reason='ForcedNexus'; Names=@('nexus.framework.healthcare','mc','anm_neu') },
            [pscustomobject]@{ Category='Adobe'; Reason='ForcedAdobe'; Names=@('Acrobat','AcroCEF') },
            [pscustomobject]@{ Category='Edge'; Reason='ForcedEdge'; Names=@('msedge','msedgewebview2') },
            [pscustomobject]@{ Category='Office'; Reason='ForcedOffice'; Names=@('OUTLOOK','EXCEL') },
            [pscustomobject]@{ Category='Printing'; Reason='ForcedPrinting'; Names=@('spoolsv') },
            [pscustomobject]@{ Category='Monitoring'; Reason='ForcedMonitoring'; Names=@('WmiPrvSE','wsmprovhost') }
        )
        $rawIds = @{}
        $alertIds = @{}
        $inclusionReasons = @{}
        foreach ($p in ($processDeltas | Sort-Object ProcessCpuCorePercent -Descending | Select-Object -First $TopProcessCount)) {
            $rawIds[$p.ProcessId] = $true
            if (-not $inclusionReasons.ContainsKey($p.ProcessId)) { $inclusionReasons[$p.ProcessId] = 'TopCpu' }
        }
        if ($isWarning) {
            foreach ($p in ($processDeltas | Sort-Object ProcessCpuCorePercent -Descending | Select-Object -First $AlertTopProcessCount)) {
                $alertIds[$p.ProcessId] = $true
                $inclusionReasons[$p.ProcessId] = 'TopCpu'
            }
            foreach ($rule in $forcedProcessRules) {
                $forcedCount = 0
                foreach ($p in ($processDeltas | Where-Object { $_.ProcessName -in $rule.Names } | Sort-Object ProcessCpuCorePercent -Descending)) {
                    if ($forcedCount -ge [int]$MaxForcedProcessesPerCategory) { break }
                    if (($p.ProcessCpuCorePercent -le 0) -and -not ($p.ProcessName -in $coreForcedNames)) { continue }
                    $alertIds[$p.ProcessId] = $true
                    if (-not $inclusionReasons.ContainsKey($p.ProcessId) -or $inclusionReasons[$p.ProcessId] -eq 'TopCpu') { $inclusionReasons[$p.ProcessId] = $rule.Reason }
                    $forcedCount++
                }
            }
        }
        $selectedIds = @{}
        foreach ($id in $rawIds.Keys) { $selectedIds[$id] = $true }
        foreach ($id in $alertIds.Keys) { $selectedIds[$id] = $true }
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
                InclusionReason = $inclusionReasons[$p.ProcessId]
                IsRawSample = $rawIds.ContainsKey($p.ProcessId)
                IsAlertSample = $alertIds.ContainsKey($p.ProcessId)
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

        $categoryNames = @('Security','Nexus','Office','Browser','Adobe','Edge','Citrix','Wem','Printing','Monitoring','Windows','Other')
        $categorySums = @{}
        foreach ($name in $categoryNames) { $categorySums[$name] = [pscustomobject]@{ Core = 0.0; Server = 0.0 } }
        foreach ($p in $processDeltas) { $categorySums[$p.Category].Core += [double]$p.ProcessCpuCorePercent; $categorySums[$p.Category].Server += [double]$p.ProcessCpuServerPercent }
        $topCategoryByServer = ($categoryNames | Sort-Object { -1 * [double]$categorySums[$_].Server } | Select-Object -First 1)

        $events = @()
        if ($IncludeEventLogContext -and $isWarning) {
            $since = (Get-Date).AddMinutes(-10)
            foreach ($logName in @('Microsoft-Windows-TaskScheduler/Operational','System','Application','Microsoft-Windows-WMI-Activity/Operational','Microsoft-Windows-Windows Defender/Operational','Microsoft-Windows-SENSE/Operational','Citrix WEM Agent')) {
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
                UserSessionsTotal = (@($sessions | Where-Object { $_.State -eq 'Active' }).Count + @($sessions | Where-Object { $_.State -eq 'Disconnected' }).Count)
                RawSessionCount = @($sessions).Count
                TotalSessions = (@($sessions | Where-Object { $_.State -eq 'Active' }).Count + @($sessions | Where-Object { $_.State -eq 'Disconnected' }).Count)
                Status = 'OK'
                ErrorMessage = ''
                IsCpuWarning = $isWarning
                IsCpuCritical = $isCritical
                AlertSeverity = if ($isCritical) { 'Critical' } elseif ($isWarning) { 'Warning' } else { 'OK' }
                TopCategoryByCpuServerPercent = $topCategoryByServer
                TopProcessCpuCorePercentSum = [Math]::Round(((@($processDeltas | Sort-Object ProcessCpuCorePercent -Descending | Select-Object -First $TopProcessCount) | Measure-Object -Property ProcessCpuCorePercent -Sum).Sum), 2)
                TopProcessCpuServerPercentSum = [Math]::Round(((@($processDeltas | Sort-Object ProcessCpuCorePercent -Descending | Select-Object -First $TopProcessCount) | Measure-Object -Property ProcessCpuServerPercent -Sum).Sum), 2)
            }
            ProcessSamples = @($processSamples)
            CategorySummary = [pscustomobject]@{
                TopCategoryByCpuServerPercent = $topCategoryByServer
                SecurityCpuCorePercent = [Math]::Round($categorySums['Security'].Core, 2); SecurityCpuServerPercent = [Math]::Round($categorySums['Security'].Server, 2)
                NexusCpuCorePercent = [Math]::Round($categorySums['Nexus'].Core, 2); NexusCpuServerPercent = [Math]::Round($categorySums['Nexus'].Server, 2)
                OfficeCpuCorePercent = [Math]::Round($categorySums['Office'].Core, 2); OfficeCpuServerPercent = [Math]::Round($categorySums['Office'].Server, 2)
                BrowserCpuCorePercent = [Math]::Round($categorySums['Browser'].Core, 2); BrowserCpuServerPercent = [Math]::Round($categorySums['Browser'].Server, 2)
                AdobeCpuCorePercent = [Math]::Round($categorySums['Adobe'].Core, 2); AdobeCpuServerPercent = [Math]::Round($categorySums['Adobe'].Server, 2)
                EdgeCpuCorePercent = [Math]::Round($categorySums['Edge'].Core, 2); EdgeCpuServerPercent = [Math]::Round($categorySums['Edge'].Server, 2)
                CitrixCpuCorePercent = [Math]::Round($categorySums['Citrix'].Core, 2); CitrixCpuServerPercent = [Math]::Round($categorySums['Citrix'].Server, 2)
                WemCpuCorePercent = [Math]::Round($categorySums['Wem'].Core, 2); WemCpuServerPercent = [Math]::Round($categorySums['Wem'].Server, 2)
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
                $result = Invoke-Command -ComputerName $Server -ScriptBlock $remoteBlock -ArgumentList $SettingsHash.CpuSampleSeconds, $SettingsHash.TopProcessCount, $SettingsHash.AlertTopProcessCount, $SettingsHash.CpuWarningThreshold, $SettingsHash.CpuCriticalThreshold, $SettingsHash.IncludeEventLogContext, $SettingsHash.AnonymizeUsers, $SettingsHash.MaxEventsPerAlert, $(if ($SettingsHash.PSObject.Properties.Name -contains 'MaxForcedProcessesPerCategory') { $SettingsHash.MaxForcedProcessesPerCategory } else { 10 }) -ErrorAction Stop
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

function Invoke-ScheduledTaskInventory {
    param([string[]]$Servers, [string[]]$TaskNames, [string]$RunId)
    $rows = @()
    foreach ($server in $Servers) {
        try {
            $remoteRows = Invoke-Command -ComputerName $server -ScriptBlock {
                param($TaskNames, $RunId)
                foreach ($taskName in $TaskNames) {
                    $tasks = @(Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object { $_.TaskName -eq $taskName })
                    if (-not $tasks -or $tasks.Count -eq 0) {
                        [pscustomobject]@{ RunId=$RunId; Server=$env:COMPUTERNAME; TaskName=$taskName; TaskPath=''; State='NOT_FOUND'; Enabled=''; LastRunTime=''; LastTaskResult=''; NextRunTime=''; Author=''; Description=''; Actions=''; Triggers=''; Status='OK'; ErrorMessage='' }
                        continue
                    }
                    foreach ($task in $tasks) {
                        $info = $null
                        try { $info = Get-ScheduledTaskInfo -TaskName $task.TaskName -TaskPath $task.TaskPath -ErrorAction Stop } catch { }
                        [pscustomobject]@{
                            RunId=$RunId; Server=$env:COMPUTERNAME; TaskName=$task.TaskName; TaskPath=$task.TaskPath; State=$task.State; Enabled=($task.Settings.Enabled);
                            LastRunTime=if($info){$info.LastRunTime}else{''}; LastTaskResult=if($info){$info.LastTaskResult}else{''}; NextRunTime=if($info){$info.NextRunTime}else{''};
                            Author=$task.Author; Description=$task.Description; Actions=($task.Actions | Out-String).Trim(); Triggers=($task.Triggers | Out-String).Trim(); Status='OK'; ErrorMessage=''
                        }
                    }
                }
            } -ArgumentList $TaskNames, $RunId -ErrorAction Stop
            $rows += $remoteRows
        }
        catch {
            foreach ($taskName in $TaskNames) { $rows += [pscustomobject]@{ RunId=$RunId; Server=$server; TaskName=$taskName; TaskPath=''; State='ERROR'; Enabled=''; LastRunTime=''; LastTaskResult=''; NextRunTime=''; Author=''; Description=''; Actions=''; Triggers=''; Status='ERROR'; ErrorMessage=$_.Exception.Message } }
        }
    }
    return $rows
}

function Invoke-CylanceHealthCollection {
    param([string[]]$Servers, [string]$RunId)
    $rows = @()
    foreach ($server in $Servers) {
        try {
            $rows += Invoke-Command -ComputerName $server -ScriptBlock {
                param($RunId)
                function Get-FirstJsonValue {
                    param($Object, [string[]]$Names)
                    if ($null -eq $Object) { return '' }
                    foreach ($name in $Names) {
                        foreach ($property in $Object.PSObject.Properties) {
                            if ($property.Name -ieq $name -and $null -ne $property.Value -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) { return [string]$property.Value }
                        }
                    }
                    foreach ($property in $Object.PSObject.Properties) {
                        if ($property.Value -is [pscustomobject]) {
                            $nested = Get-FirstJsonValue -Object $property.Value -Names $Names
                            if ($nested) { return $nested }
                        }
                    }
                    return ''
                }
                function Test-RegistryPropertyExists {
                    param([string[]]$Paths, [string[]]$Names)
                    foreach ($path in $Paths) {
                        try {
                            $item = Get-ItemProperty -LiteralPath $path -ErrorAction Stop
                            foreach ($name in $Names) { if ($item.PSObject.Properties.Name -contains $name) { return $true } }
                        } catch { }
                    }
                    return $false
                }
                $statusPath = 'C:\ProgramData\Cylance\Status\Status.json'
                $statusJson = $null
                $errorMessage = ''
                try { if (Test-Path -LiteralPath $statusPath) { $statusJson = Get-Content -LiteralPath $statusPath -Raw | ConvertFrom-Json } else { $errorMessage = 'Status.json nicht gefunden.' } } catch { $errorMessage = $_.Exception.Message }
                $svc = Get-Service -Name CylanceSvc -ErrorAction SilentlyContinue
                $protection = ''
                try { $protection = (& sc.exe qprotection CylanceSvc 2>$null | Out-String).Trim() } catch { }
                $deviceName = Get-FirstJsonValue -Object $statusJson -Names @('StatusDeviceName','DeviceName','device_name','deviceName','HostName','hostname','ComputerName')
                $serial = Get-FirstJsonValue -Object $statusJson -Names @('SerialNumber','serial_number','serialNumber','AgentId','AgentID','DeviceId','device_id')
                $agentVersion = Get-FirstJsonValue -Object $statusJson -Names @('AgentVersion','agent_version','agentVersion','ProductVersion','Version')
                $policyName = Get-FirstJsonValue -Object $statusJson -Names @('PolicyName','policy_name','policyName','Policy')
                $lastCommunicated = Get-FirstJsonValue -Object $statusJson -Names @('LastCommunicated','last_communicated','lastCommunicated','LastCheckIn')
                $lastBackgroundScan = Get-FirstJsonValue -Object $statusJson -Names @('LastBackgroundScan','last_background_scan','lastBackgroundScan')
                $drivesScanned = Get-FirstJsonValue -Object $statusJson -Names @('DrivesScanned','drives_scanned','drivesScanned')
                $threatCount = Get-FirstJsonValue -Object $statusJson -Names @('ThreatCount','threat_count','threatCount')
                $registryPaths = @('HKLM:\SOFTWARE\Cylance\Desktop','HKLM:\SOFTWARE\WOW6432Node\Cylance\Desktop')
                $fpExists = (Test-RegistryPropertyExists -Paths $registryPaths -Names @('FP','FPExists')) -or (Test-Path 'C:\ProgramData\Cylance\Desktop\config_defaults.xml')
                $fpMaskExists = (Test-RegistryPropertyExists -Paths $registryPaths -Names @('FPMask','FP_Mask','FPMaskExists')) -or (Test-Path 'C:\ProgramData\Cylance\Desktop\fp.mask')
                $fpVersionExists = (Test-RegistryPropertyExists -Paths $registryPaths -Names @('FPVersion','FP_Version','FPVersionExists')) -or (Test-Path 'C:\ProgramData\Cylance\Desktop\fp.version')
                $rowStatus = if ($errorMessage) { 'ERROR' } elseif ([string]::IsNullOrWhiteSpace($deviceName) -or [string]::IsNullOrWhiteSpace($serial)) { 'PARTIAL' } else { 'OK' }
                $rowError = if ($rowStatus -eq 'PARTIAL') { 'Identitaetsfelder fehlen oder konnten nicht aus Status.json gelesen werden.' } else { $errorMessage }
                [pscustomobject]@{
                    RunId=$RunId; Server=$env:COMPUTERNAME; ServiceStatus=if($svc){$svc.Status}else{'NOT_FOUND'}; ServiceCanStop=if($svc){$svc.CanStop}else{''}; ServiceProtection=$protection;
                    StatusDeviceName=$deviceName; SerialNumber=$serial; AgentVersion=$agentVersion; PolicyName=$policyName;
                    LastCommunicated=$lastCommunicated; LastBackgroundScan=$lastBackgroundScan; DrivesScanned=$drivesScanned; ThreatCount=$threatCount;
                    FPExists=$fpExists; FPMaskExists=$fpMaskExists; FPVersionExists=$fpVersionExists;
                    Status=$rowStatus; ErrorMessage=$rowError
                }
            } -ArgumentList $RunId -ErrorAction Stop
        }
        catch { $rows += [pscustomobject]@{ RunId=$RunId; Server=$server; ServiceStatus='ERROR'; ServiceCanStop=''; ServiceProtection=''; StatusDeviceName=''; SerialNumber=''; AgentVersion=''; PolicyName=''; LastCommunicated=''; LastBackgroundScan=''; DrivesScanned=''; ThreatCount=''; FPExists=''; FPMaskExists=''; FPVersionExists=''; Status='ERROR'; ErrorMessage=$_.Exception.Message } }
    }
    return $rows
}

function New-CylanceDuplicateRows {
    param([array]$Rows, [string]$RunId)
    $dupes = @()
    foreach ($field in @('SerialNumber','StatusDeviceName')) {
        $readRows = @($Rows | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.$field) })
        if ($readRows.Count -eq 0) {
            $dupes += [pscustomobject]@{ RunId=$RunId; DuplicateField=$field; DuplicateValue=''; Servers=''; Status='NOT_EVALUATED'; ErrorMessage="Keine Werte fuer $field gelesen; Duplikatpruefung nicht aussagekraeftig." }
            continue
        }
        $dupes += $readRows | Group-Object $field | Where-Object Count -gt 1 | ForEach-Object { [pscustomobject]@{ RunId=$RunId; DuplicateField=$field; DuplicateValue=$_.Name; Servers=(($_.Group | ForEach-Object Server) -join ','); Status='DUPLICATE'; ErrorMessage='' } }
    }
    return $dupes
}


function Start-DefenderPerfRecordingJob {
    param([string]$Server, [string]$RunId, [string]$OutputPath, [int]$Seconds, [string]$TriggerProcess, [double]$TriggerCpu, [datetime]$TriggerTime)
    Start-Job -ScriptBlock {
        param($Server, $RunId, $OutputPath, $Seconds, $TriggerProcess, $TriggerCpu, $TriggerTime)
        $safeServer = $Server -replace '[^A-Za-z0-9_.-]', '_'
        $triggerStamp = $TriggerTime.ToString('yyyyMMdd_HHmmss')
        $centralFolder = Join-Path (Join-Path (Join-Path $OutputPath 'DefenderPerf') $RunId) $safeServer
        $session = $null
        try {
            if (-not (Test-Path -LiteralPath $centralFolder)) { New-Item -ItemType Directory -Path $centralFolder -Force | Out-Null }
            $session = New-PSSession -ComputerName $Server -ErrorAction Stop
            $remote = Invoke-Command -Session $session -ScriptBlock {
                param($RunId, $Seconds, $TriggerStamp, $SafeServer)
                $remoteFolder = Join-Path 'C:\ProgramData\CitrixTSHealthCheck\DefenderPerf' $RunId
                if (-not (Test-Path -LiteralPath $remoteFolder)) { New-Item -ItemType Directory -Path $remoteFolder -Force | Out-Null }
                $baseName = "DefenderPerfRecording_{0}_{1}_{2}" -f $RunId, $SafeServer, $TriggerStamp
                $etlPath = Join-Path $remoteFolder ($baseName + '.etl')
                $logPath = Join-Path $remoteFolder ($baseName + '.log')
                $reportTxtPath = Join-Path $remoteFolder ("DefenderPerfReport_{0}_{1}.txt" -f $RunId, $SafeServer)
                $reportJsonPath = Join-Path $remoteFolder ("DefenderPerfReport_{0}_{1}_raw.json" -f $RunId, $SafeServer)
                $status = 'OK'
                $errorMessage = ''
                try {
                    & { New-MpPerformanceRecording -RecordTo $etlPath -Seconds $Seconds -ErrorAction Stop } *> $logPath
                }
                catch {
                    $status = 'ERROR'
                    $errorMessage = "New-MpPerformanceRecording fehlgeschlagen: $($_.Exception.Message)"
                    try { $errorMessage | Add-Content -LiteralPath $logPath -Encoding UTF8 } catch { }
                }
                try {
                    $report = Get-MpPerformanceReport -Path $etlPath -ErrorAction Stop
                    $report | Out-String -Width 4096 | Set-Content -LiteralPath $reportTxtPath -Encoding UTF8
                    $report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $reportJsonPath -Encoding UTF8
                }
                catch {
                    if ($status -eq 'OK') { $status = 'PARTIAL' }
                    $reportError = "Get-MpPerformanceReport fehlgeschlagen: $($_.Exception.Message)"
                    $errorMessage = @($errorMessage, $reportError | Where-Object { $_ }) -join ' | '
                    $reportError | Set-Content -LiteralPath $reportTxtPath -Encoding UTF8
                    [pscustomobject]@{ Error=$reportError } | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $reportJsonPath -Encoding UTF8
                }
                [pscustomobject]@{
                    ComputerName=$env:COMPUTERNAME; RemoteFolder=$remoteFolder; EtlPath=$etlPath; LogPath=$logPath;
                    ReportPath=$reportTxtPath; RawJsonPath=$reportJsonPath; Status=$status; ErrorMessage=$errorMessage
                }
            } -ArgumentList $RunId, $Seconds, $triggerStamp, $safeServer -ErrorAction Stop

            $copied = @{}
            foreach ($remotePath in @($remote.EtlPath, $remote.LogPath, $remote.ReportPath, $remote.RawJsonPath)) {
                if ([string]::IsNullOrWhiteSpace([string]$remotePath)) { continue }
                $fileName = Split-Path -Path $remotePath -Leaf
                $destination = Join-Path $centralFolder $fileName
                try {
                    Copy-Item -FromSession $session -LiteralPath $remotePath -Destination $destination -Force -ErrorAction Stop
                    $copied[$remotePath] = $destination
                }
                catch {
                    if ($remote.Status -eq 'OK') { $remote.Status = 'PARTIAL' }
                    $copyError = "Kopieren fehlgeschlagen ($remotePath): $($_.Exception.Message)"
                    $remote.ErrorMessage = @($remote.ErrorMessage, $copyError | Where-Object { $_ }) -join ' | '
                }
            }
            $etlResultPath = if ($copied.ContainsKey($remote.EtlPath)) { $copied[$remote.EtlPath] } else { $remote.EtlPath }
            $logResultPath = if ($copied.ContainsKey($remote.LogPath)) { $copied[$remote.LogPath] } else { $remote.LogPath }
            $reportResultPath = if ($copied.ContainsKey($remote.ReportPath)) { $copied[$remote.ReportPath] } else { $remote.ReportPath }
            $rawJsonResultPath = if ($copied.ContainsKey($remote.RawJsonPath)) { $copied[$remote.RawJsonPath] } else { $remote.RawJsonPath }
            [pscustomobject]@{
                RunId=$RunId; Server=$Server; TriggerTime=$TriggerTime.ToString('s'); TriggerProcess=$TriggerProcess; TriggerCpu=$TriggerCpu;
                EtlPath=$etlResultPath; LogPath=$logResultPath; ReportPath=$reportResultPath; RawJsonPath=$rawJsonResultPath;
                RemoteFolder=$remote.RemoteFolder; CentralFolder=$centralFolder; Status=$remote.Status; ErrorMessage=$remote.ErrorMessage
            }
        }
        catch {
            [pscustomobject]@{ RunId=$RunId; Server=$Server; TriggerTime=$TriggerTime.ToString('s'); TriggerProcess=$TriggerProcess; TriggerCpu=$TriggerCpu; EtlPath=''; LogPath=''; ReportPath=''; RawJsonPath=''; RemoteFolder=''; CentralFolder=$centralFolder; Status='ERROR'; ErrorMessage=$_.Exception.Message }
        }
        finally {
            if ($session) { Remove-PSSession -Session $session -ErrorAction SilentlyContinue }
        }
    } -ArgumentList $Server, $RunId, $OutputPath, $Seconds, $TriggerProcess, $TriggerCpu, $TriggerTime
}

function Start-WemEventContextJob {
    param([string]$Server, [string]$RunId, [string]$AlertId, [datetime]$TriggerTime, [int]$WindowMinutes, [bool]$IncludeLogTail, [int]$TailLines)
    Start-Job -ScriptBlock {
        param($Server, $RunId, $AlertId, $TriggerTime, $WindowMinutes, $IncludeLogTail, $TailLines)
        try {
            Invoke-Command -ComputerName $Server -ScriptBlock {
                param($RunId, $AlertId, $TriggerTime, $WindowMinutes, $IncludeLogTail, $TailLines)
                $start = $TriggerTime.AddMinutes(-1 * $WindowMinutes)
                $end = $TriggerTime.AddMinutes($WindowMinutes)
                $logNames = @()
                try { $logNames += Get-WinEvent -ListLog '*WEM*' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty LogName } catch { }
                $logNames += @('WEM Agent Service','Citrix WEM Agent Service','Norskale Agent Service')
                $logNames = @($logNames | Where-Object { $_ } | Select-Object -Unique)
                foreach ($logName in $logNames) {
                    try {
                        Get-WinEvent -FilterHashtable @{ LogName=$logName; StartTime=$start; EndTime=$end } -MaxEvents 200 -ErrorAction Stop | ForEach-Object {
                            [pscustomobject]@{ RecordType='Event'; RunId=$RunId; Server=$env:COMPUTERNAME; AlertId=$AlertId; TimeCreated=$_.TimeCreated; LogName=$logName; ProviderName=$_.ProviderName; EventId=$_.Id; Level=$_.LevelDisplayName; Message=$_.Message; FileName=''; Content='' }
                        }
                    } catch {
                        [pscustomobject]@{ RecordType='Event'; RunId=$RunId; Server=$env:COMPUTERNAME; AlertId=$AlertId; TimeCreated=''; LogName=$logName; ProviderName=''; EventId=''; Level=''; Message="WEM Log nicht lesbar oder nicht vorhanden: $($_.Exception.Message)"; FileName=''; Content='' }
                    }
                }
                if ($IncludeLogTail) {
                    $paths = @('C:\Program Files (x86)\Norskale','C:\Program Files (x86)\Citrix\Workspace Environment Management Agent','C:\ProgramData\Norskale','C:\ProgramData\Citrix\WEM')
                    foreach ($base in $paths) {
                        try {
                            Get-ChildItem -LiteralPath $base -Filter '*.log' -Recurse -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 5 | ForEach-Object {
                                [pscustomobject]@{ RecordType='Tail'; RunId=$RunId; Server=$env:COMPUTERNAME; AlertId=$AlertId; TimeCreated=''; LogName=''; ProviderName=''; EventId=''; Level=''; Message=''; FileName=$_.FullName; Content=((Get-Content -LiteralPath $_.FullName -Tail $TailLines -ErrorAction SilentlyContinue) -join [Environment]::NewLine) }
                            }
                        } catch { }
                    }
                }
            } -ArgumentList $RunId, $AlertId, $TriggerTime, $WindowMinutes, $IncludeLogTail, $TailLines -ErrorAction Stop
        }
        catch { [pscustomobject]@{ RecordType='Event'; RunId=$RunId; Server=$Server; AlertId=$AlertId; TimeCreated=''; LogName=''; ProviderName=''; EventId=''; Level=''; Message=$_.Exception.Message; FileName=''; Content='' } }
    } -ArgumentList $Server, $RunId, $AlertId, $TriggerTime, $WindowMinutes, $IncludeLogTail, $TailLines
}

function New-ErrorServerSample {
    param([string]$Server, [string]$Message, [datetime]$Timestamp, [string]$RunId)
    [pscustomobject]@{
        RunId = $RunId; Timestamp = $Timestamp.ToString('s'); TargetServer = $Server; ComputerName = ''; LogicalProcessorCount = ''; CpuPercent = ''; MemoryPercent = ''; TotalMemoryMB = ''; UsedMemoryMB = ''; FreeMemoryMB = '';
        ActiveSessions = ''; DisconnectedSessions = ''; UserSessionsTotal = ''; RawSessionCount = ''; TotalSessions = ''; Status = 'ERROR'; ErrorMessage = $Message; IsCpuWarning = $false; IsCpuCritical = $false; AlertSeverity = 'Error';
        TopCategoryByCpuServerPercent = ''; SecurityCpuServerPercent = ''; NexusCpuServerPercent = ''; OfficeCpuServerPercent = ''; BrowserCpuServerPercent = ''; AdobeCpuServerPercent = ''; EdgeCpuServerPercent = ''; CitrixCpuServerPercent = ''; WemCpuServerPercent = ''; PrintingCpuServerPercent = ''; MonitoringCpuServerPercent = ''; WindowsCpuServerPercent = ''; OtherCpuServerPercent = ''
    }
}


function Get-SumValue {
    param([array]$Rows, [string]$PropertyName)
    $sum = 0.0
    foreach ($row in @($Rows)) {
        if ($null -ne $row -and ($row.PSObject.Properties.Name -contains $PropertyName) -and $row.$PropertyName -ne '') {
            try { $sum += [double]$row.$PropertyName } catch { }
        }
    }
    return $sum
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
        [pscustomobject]@{ Name = $_.Name; Cpu = (Get-SumValue -Rows $_.Group -PropertyName 'ProcessCpuServerPercent') }
    } | Sort-Object Cpu -Descending | Select-Object -First $Count | ForEach-Object { '{0}={1:n2}' -f $_.Name,$_.Cpu }) -join ', '
}


function New-CategoryAggregateRows {
    param([array]$ProcessRows, [string]$RunId)
    foreach ($group in ($ProcessRows | Group-Object TargetServer, Category)) {
        $rows = @($group.Group)
        if ($rows.Count -eq 0) { continue }
        $server = $rows[0].TargetServer
        $category = $rows[0].Category
        $coreValues = @($rows | ForEach-Object { [double]$_.ProcessCpuCorePercent })
        $serverValues = @($rows | ForEach-Object { [double]$_.ProcessCpuServerPercent })
        $wsValues = @($rows | Where-Object { $_.ProcessWorkingSetMB -ne '' } | ForEach-Object { [double]$_.ProcessWorkingSetMB })
        [pscustomobject]@{
            RunId = $RunId
            Server = $server
            TargetServer = $server
            Category = $category
            Samples = $rows.Count
            CpuCoreSecondsSum = [Math]::Round((Get-SumValue -Rows $rows -PropertyName 'ProcessCpuSecondsDelta'), 3)
            AvgCpuCorePercent = [Math]::Round((($coreValues | Measure-Object -Average).Average), 2)
            MaxCpuCorePercent = [Math]::Round((($coreValues | Measure-Object -Maximum).Maximum), 2)
            AvgCpuServerPercent = [Math]::Round((($serverValues | Measure-Object -Average).Average), 2)
            MaxCpuServerPercent = [Math]::Round((($serverValues | Measure-Object -Maximum).Maximum), 2)
            WorkingSetMBAverage = if ($wsValues.Count) { [Math]::Round((($wsValues | Measure-Object -Average).Average), 2) } else { '' }
            WorkingSetMBMax = if ($wsValues.Count) { [Math]::Round((($wsValues | Measure-Object -Maximum).Maximum), 2) } else { '' }
            ProcessCountDistinct = @($rows | Select-Object -ExpandProperty ProcessId -Unique).Count
            TopProcessesByCpu = Join-TopCpuGroups -Rows $rows -GroupProperty ProcessName -Count 10
        }
    }
}

function New-RunSummaryRows {
    param([array]$ServerRows, [array]$ProcessRows, [datetime]$Start, [datetime]$End, [int]$ServerCount)
    $categoryNames = @('Security','Nexus','Office','Browser','Adobe','Edge','Citrix','Wem','Printing','Monitoring','Windows','Other')
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
        foreach ($category in $categoryNames) { $categoryTotals[$category] = (Get-SumValue -Rows @($serverProcessRows | Where-Object Category -eq $category) -PropertyName 'ProcessCpuServerPercent') }
        [pscustomobject]@{
            RunId = $runId; StartTime = $Start.ToString('s'); EndTime = $End.ToString('s'); IntervalSeconds = $settings.IntervalSeconds; CpuSampleSeconds = $settings.CpuSampleSeconds; ServerList = ($servers -join ','); ImageVersion = $settings.ImageVersion; Notes = $settings.Notes; ScriptVersion = '1.0.0'; OutputPath = $OutputPath; RunStart = $Start.ToString('s'); RunEnd = $End.ToString('s'); DurationMinutes = [Math]::Round(($End-$Start).TotalMinutes, 2); ServerCount = $ServerCount; TargetServer = $server;
            SampleCount = $okRows.Count; ErrorCount = @($group.Group | Where-Object Status -eq 'ERROR').Count;
            CpuAverage = if ($cpuValues.Count) { [Math]::Round((($cpuValues | Measure-Object -Average).Average),2) } else { '' }; CpuMedian = if ($median -ne '') { [Math]::Round($median,2) } else { '' }; CpuMax = if ($cpuValues.Count) { [Math]::Round((($cpuValues | Measure-Object -Maximum).Maximum),2) } else { '' }; CpuP95 = if ($p95 -ne '') { [Math]::Round($p95,2) } else { '' };
            CpuWarningCount = @($okRows | Where-Object { $_.IsCpuWarning -eq $true }).Count; CpuCriticalCount = @($okRows | Where-Object { $_.IsCpuCritical -eq $true }).Count;
            MemoryAverage = if ($ramValues.Count) { [Math]::Round((($ramValues | Measure-Object -Average).Average),2) } else { '' }; MemoryMax = if ($ramValues.Count) { [Math]::Round((($ramValues | Measure-Object -Maximum).Maximum),2) } else { '' };
            ActiveSessionsAverage = if ($activeValues.Count) { [Math]::Round((($activeValues | Measure-Object -Average).Average),2) } else { '' }; ActiveSessionsMax = if ($activeValues.Count) { [Math]::Round((($activeValues | Measure-Object -Maximum).Maximum),2) } else { '' };
            DisconnectedSessionsAverage = if ($discValues.Count) { [Math]::Round((($discValues | Measure-Object -Average).Average),2) } else { '' };
            TopProcessByTotalCpuServerPercent = $topProc; Top10ProcessesByTotalCpuServerPercent = $top10Proc; TopCategoryByTotalCpuServerPercent = $topCat;
            SecurityTotalCpuServerPercent = [Math]::Round($categoryTotals['Security'],2); NexusTotalCpuServerPercent = [Math]::Round($categoryTotals['Nexus'],2); OfficeTotalCpuServerPercent = [Math]::Round($categoryTotals['Office'],2); BrowserTotalCpuServerPercent = [Math]::Round($categoryTotals['Browser'],2); AdobeTotalCpuServerPercent = [Math]::Round($categoryTotals['Adobe'],2); EdgeTotalCpuServerPercent = [Math]::Round($categoryTotals['Edge'],2); CitrixTotalCpuServerPercent = [Math]::Round($categoryTotals['Citrix'],2); WemTotalCpuServerPercent = [Math]::Round($categoryTotals['Wem'],2); PrintingTotalCpuServerPercent = [Math]::Round($categoryTotals['Printing'],2); MonitoringTotalCpuServerPercent = [Math]::Round($categoryTotals['Monitoring'],2); WindowsTotalCpuServerPercent = [Math]::Round($categoryTotals['Windows'],2); OtherTotalCpuServerPercent = [Math]::Round($categoryTotals['Other'],2)
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


function Receive-DetailDiagnosticJobs {
    param([array]$DefenderJobs, [array]$WemJobs, [string]$RawPath, [string]$RunOutputPath, [string]$RunId, [string]$Delimiter)
    foreach ($job in @($DefenderJobs | Where-Object { $_.State -ne 'Running' })) {
        try {
            $rows = @(Receive-Job -Job $job -ErrorAction SilentlyContinue)
            Export-Rows -Rows $rows -Path (Join-Path $RawPath "DefenderPerfRecordings_$RunId.csv") -Delimiter $Delimiter
            Export-Rows -Rows $rows -Path (Join-Path $RunOutputPath 'DefenderPerfRecordings.csv') -Delimiter $Delimiter
        } finally { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue }
    }
    foreach ($job in @($WemJobs | Where-Object { $_.State -ne 'Running' })) {
        try {
            $rows = @(Receive-Job -Job $job -ErrorAction SilentlyContinue)
            $eventRows = @($rows | Where-Object { $_.RecordType -eq 'Event' } | Select-Object RunId,Server,AlertId,TimeCreated,LogName,ProviderName,EventId,Level,Message)
            Export-Rows -Rows $eventRows -Path (Join-Path $RawPath "WemEventContext_$RunId.csv") -Delimiter $Delimiter
            Export-Rows -Rows $eventRows -Path (Join-Path $RunOutputPath 'WemEventContext.csv') -Delimiter $Delimiter
            foreach ($tail in @($rows | Where-Object { $_.RecordType -eq 'Tail' -and $_.Content })) {
                $safeServer = $tail.Server -replace '[^A-Za-z0-9_.-]', '_'
                $stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
                $tailPath = Join-Path $RawPath ("WemLogTail_{0}_{1}_{2}.log" -f $RunId, $safeServer, $stamp)
                $tail.Content | Set-Content -LiteralPath $tailPath -Encoding UTF8
            }
        } finally { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue }
    }
}

if ((Test-Path -LiteralPath $ConfigPath -PathType Container)) { $ConfigPath = Join-Path $ConfigPath 'settings.json' }
$settings = Merge-ParameterSettings -Settings (Read-Settings -Path $ConfigPath)
Ensure-SettingProperty -Settings $settings -Name 'MaxForcedProcessesPerCategory' -DefaultValue 10
Ensure-SettingProperty -Settings $settings -Name 'AutoDefenderPerfRecording' -DefaultValue $false
Ensure-SettingProperty -Settings $settings -Name 'DefenderPerfTriggerServerCpuPercent' -DefaultValue 10
Ensure-SettingProperty -Settings $settings -Name 'DefenderPerfRecordingSeconds' -DefaultValue 900
Ensure-SettingProperty -Settings $settings -Name 'DefenderPerfCooldownMinutes' -DefaultValue 120
Ensure-SettingProperty -Settings $settings -Name 'MaxConcurrentDefenderPerfRecordings' -DefaultValue 2
Ensure-SettingProperty -Settings $settings -Name 'IncludeWemEventContext' -DefaultValue $false
Ensure-SettingProperty -Settings $settings -Name 'WemTriggerServerCpuPercent' -DefaultValue 10
Ensure-SettingProperty -Settings $settings -Name 'WemEventWindowMinutes' -DefaultValue 10
Ensure-SettingProperty -Settings $settings -Name 'IncludeWemLogTail' -DefaultValue $false
Ensure-SettingProperty -Settings $settings -Name 'WemLogTailLines' -DefaultValue 200
if ($settings.IncludeEventContext) { $settings.IncludeEventLogContext = $true }
$servers = @(Read-ServerList -Path $ServerListPath)
$rawPath = Join-Path $OutputPath 'raw'
$summaryPath = Join-Path $OutputPath 'summary'
$logPath = Join-Path $OutputPath 'logs'
Ensure-Directory $rawPath; Ensure-Directory $summaryPath; Ensure-Directory $logPath

$runStart = Get-Date
if ([string]::IsNullOrWhiteSpace($settings.RunId)) { $runId = '{0}_{1}' -f $runStart.ToString('yyyyMMdd_HHmmss'), ([guid]::NewGuid().ToString('N').Substring(0,8)) } else { $runId = $settings.RunId }
$runsPath = Join-Path $OutputPath 'runs'
Ensure-Directory $runsPath
$runOutputPath = Join-Path $runsPath $runId
Ensure-Directory $runOutputPath
$runLogFile = Join-Path $logPath "RunLog_$runId.log"
$allServerRows = @()
$allProcessRows = @()
$defenderJobs = @()
$wemJobs = @()
$defenderLastTriggerByServer = @{}
$defenderTriggeredServers = @{}
$defenderTriggered = $false
$wemTriggered = $false
$round = 0
$endAt = $runStart.AddMinutes([int]$settings.DurationMinutes)
if ([int]$settings.DurationMinutes -lt 1) { $endAt = $runStart }
$runError = $null

try {
    Write-RunLog -Path $runLogFile -Message "Run gestartet. Server=$($servers.Count), DurationMinutes=$($settings.DurationMinutes), IntervalSeconds=$($settings.IntervalSeconds), MaxParallel=$($settings.MaxParallel)"
    if ($settings.IncludeScheduledTaskInventory) {
        Write-RunLog -Path $runLogFile -Message 'ScheduledTaskInventory gestartet.'
        $taskRows = @(Invoke-ScheduledTaskInventory -Servers $servers -TaskNames $settings.TaskNamesToCheck -RunId $runId)
        Export-Rows -Rows $taskRows -Path (Join-Path $rawPath "ScheduledTaskInventory_$runId.csv") -Delimiter $settings.OutputDelimiter
        Export-Rows -Rows $taskRows -Path (Join-Path $runOutputPath 'ScheduledTaskInventory.csv') -Delimiter $settings.OutputDelimiter
        Write-RunLog -Path $runLogFile -Message "ScheduledTaskInventory beendet. Zeilen=$($taskRows.Count)"
    }
    if ($settings.IncludeCylanceHealth) {
        Write-RunLog -Path $runLogFile -Message 'CylanceHealth gestartet.'
        $cylanceRows = @(Invoke-CylanceHealthCollection -Servers $servers -RunId $runId)
        $duplicateRows = @(New-CylanceDuplicateRows -Rows $cylanceRows -RunId $runId)
        Export-Rows -Rows $cylanceRows -Path (Join-Path $rawPath "CylanceHealth_$runId.csv") -Delimiter $settings.OutputDelimiter
        Export-Rows -Rows $cylanceRows -Path (Join-Path $runOutputPath 'CylanceHealth.csv') -Delimiter $settings.OutputDelimiter
        Export-Rows -Rows $duplicateRows -Path (Join-Path $rawPath "CylanceHealth_Duplicates_$runId.csv") -Delimiter $settings.OutputDelimiter
        Export-Rows -Rows $duplicateRows -Path (Join-Path $runOutputPath 'CylanceHealth_Duplicates.csv') -Delimiter $settings.OutputDelimiter
        Write-RunLog -Path $runLogFile -Message "CylanceHealth beendet. Zeilen=$($cylanceRows.Count), Duplikate=$($duplicateRows.Count)"
    }
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
                $errorRow = New-ErrorServerSample -Server $item.TargetServer -Message $item.ErrorMessage -Timestamp $roundStart -RunId $runId
                $serverRows += $errorRow
                Write-RunLog -Path $runLogFile -Level 'ERROR' -Message "$($item.TargetServer): $($item.ErrorMessage)"
                continue
            }
            $serverSample = $item.Result.ServerSample
            $cat = $item.Result.CategorySummary
            $serverRow = [pscustomobject]@{
                RunId=$runId; Timestamp=$serverSample.Timestamp; Server=$item.TargetServer; TargetServer=$item.TargetServer; ComputerName=$serverSample.ComputerName; LogicalProcessorCount=$serverSample.LogicalProcessorCount; CpuPercent=$serverSample.CpuPercent; MemoryPercent=$serverSample.MemoryPercent;
                TotalMemoryMB=$serverSample.TotalMemoryMB; UsedMemoryMB=$serverSample.UsedMemoryMB; FreeMemoryMB=$serverSample.FreeMemoryMB; ActiveSessions=$serverSample.ActiveSessions; DisconnectedSessions=$serverSample.DisconnectedSessions; UserSessionsTotal=$serverSample.UserSessionsTotal; RawSessionCount=$serverSample.RawSessionCount; TotalSessionsRaw=$serverSample.RawSessionCount; TotalSessions=$serverSample.TotalSessions;
                Status='OK'; ErrorMessage=''; IsCpuWarning=$serverSample.IsCpuWarning; IsCpuCritical=$serverSample.IsCpuCritical; AlertSeverity=$serverSample.AlertSeverity; TopCategoryByCpuServerPercent=$serverSample.TopCategoryByCpuServerPercent;
                SecurityCpuServerPercent=$cat.SecurityCpuServerPercent; NexusCpuServerPercent=$cat.NexusCpuServerPercent; OfficeCpuServerPercent=$cat.OfficeCpuServerPercent; BrowserCpuServerPercent=$cat.BrowserCpuServerPercent; AdobeCpuServerPercent=$cat.AdobeCpuServerPercent; EdgeCpuServerPercent=$cat.EdgeCpuServerPercent; CitrixCpuServerPercent=$cat.CitrixCpuServerPercent; WemCpuServerPercent=$cat.WemCpuServerPercent; PrintingCpuServerPercent=$cat.PrintingCpuServerPercent; MonitoringCpuServerPercent=$cat.MonitoringCpuServerPercent; WindowsCpuServerPercent=$cat.WindowsCpuServerPercent; OtherCpuServerPercent=$cat.OtherCpuServerPercent
            }
            $serverRows += $serverRow

            $sampleTime = [datetime]$serverSample.Timestamp
            $msMpEng = @($item.Result.ProcessSamples | Where-Object { $_.ProcessName -eq 'MsMpEng' } | Sort-Object ProcessCpuServerPercent -Descending | Select-Object -First 1)
            if ($settings.AutoDefenderPerfRecording -and $msMpEng -and [double]$msMpEng.ProcessCpuServerPercent -ge [double]$settings.DefenderPerfTriggerServerCpuPercent) {
                $runningDefenderJobs = @($defenderJobs | Where-Object { $_.State -eq 'Running' }).Count
                $cooldownOk = (-not $defenderLastTriggerByServer.ContainsKey($item.TargetServer)) -or (($sampleTime - $defenderLastTriggerByServer[$item.TargetServer]).TotalMinutes -ge [int]$settings.DefenderPerfCooldownMinutes)
                if (-not $defenderTriggeredServers.ContainsKey($item.TargetServer) -and $cooldownOk -and $runningDefenderJobs -lt [int]$settings.MaxConcurrentDefenderPerfRecordings) {
                    $defenderLastTriggerByServer[$item.TargetServer] = $sampleTime
                    $defenderTriggeredServers[$item.TargetServer] = $true
                    $defenderTriggered = $true
                    $defenderJobs += Start-DefenderPerfRecordingJob -Server $item.TargetServer -RunId $runId -OutputPath $OutputPath -Seconds ([int]$settings.DefenderPerfRecordingSeconds) -TriggerProcess 'MsMpEng' -TriggerCpu ([double]$msMpEng.ProcessCpuServerPercent) -TriggerTime $sampleTime
                    Export-Rows -Rows @([pscustomobject]@{ RunId=$runId; Server=$item.TargetServer; TriggerTime=$sampleTime.ToString('s'); TriggerProcess='MsMpEng'; TriggerCpu=[double]$msMpEng.ProcessCpuServerPercent; EtlPath=''; LogPath=''; ReportPath=''; RawJsonPath=''; RemoteFolder=''; CentralFolder=(Join-Path (Join-Path (Join-Path $OutputPath 'DefenderPerf') $runId) ($item.TargetServer -replace '[^A-Za-z0-9_.-]', '_')); Status='STARTED'; ErrorMessage='' }) -Path (Join-Path $rawPath "DefenderPerfRecordings_$runId.csv") -Delimiter $settings.OutputDelimiter
                    Write-RunLog -Path $runLogFile -Message "Defender Performance Recording getriggert: $($item.TargetServer), MsMpEng=$($msMpEng.ProcessCpuServerPercent)%"
                }
            }
            $wemCpu = [double]$cat.WemCpuServerPercent
            $wemProc = @($item.Result.ProcessSamples | Where-Object { $_.ProcessName -in @('Citrix.Wem.Agent.Service','VUEMUIAgent') } | Sort-Object ProcessCpuServerPercent -Descending | Select-Object -First 1)
            if ($settings.IncludeWemEventContext -and (($wemCpu -ge [double]$settings.WemTriggerServerCpuPercent) -or ($wemProc -and [double]$wemProc.ProcessCpuServerPercent -ge [double]$settings.WemTriggerServerCpuPercent))) {
                $wemTriggered = $true
                $wemAlertId = '{0}_{1}_{2}_WEM' -f $runId, ($item.TargetServer -replace '[^A-Za-z0-9_.-]', '_'), ($serverSample.Timestamp -replace '[:]', '')
                $wemJobs += Start-WemEventContextJob -Server $item.TargetServer -RunId $runId -AlertId $wemAlertId -TriggerTime $sampleTime -WindowMinutes ([int]$settings.WemEventWindowMinutes) -IncludeLogTail ([bool]$settings.IncludeWemLogTail) -TailLines ([int]$settings.WemLogTailLines)
                Write-RunLog -Path $runLogFile -Message "WEM Event Context getriggert: $($item.TargetServer), WEM=$wemCpu%"
            }
            $rank = 1
            $rankByProcessId = @{}
            foreach ($rankedProcess in @($item.Result.ProcessSamples | Sort-Object ProcessCpuCorePercent -Descending)) {
                $rankByProcessId[$rankedProcess.ProcessId] = $rank
                $rank++
            }
            foreach ($p in @($item.Result.ProcessSamples | Where-Object { $_.IsRawSample } | Sort-Object ProcessCpuCorePercent -Descending)) {
                $processRows += [pscustomobject]@{
                    RunId=$runId; Timestamp=$serverSample.Timestamp; Server=$item.TargetServer; TargetServer=$item.TargetServer; ComputerName=$serverSample.ComputerName; LogicalProcessorCount=$serverSample.LogicalProcessorCount; CpuPercent=$serverSample.CpuPercent; MemoryPercent=$serverSample.MemoryPercent; ActiveSessions=$serverSample.ActiveSessions; DisconnectedSessions=$serverSample.DisconnectedSessions;
                    ProcessRank=$rankByProcessId[$p.ProcessId]; InclusionReason=$p.InclusionReason; ProcessId=$p.ProcessId; ParentProcessId=$p.ParentProcessId; ParentProcessName=$p.ParentProcessName; ProcessName=$p.ProcessName; ProcessSessionId=$p.ProcessSessionId; ProcessUserName=$p.ProcessUserName; ProcessUserDomain=$p.ProcessUserDomain; SessionState=$p.SessionState;
                    ProcessCpuCorePercent=$p.ProcessCpuCorePercent; ProcessCpuServerPercent=$p.ProcessCpuServerPercent; ProcessCpuSecondsDelta=$p.ProcessCpuSecondsDelta; ProcessWorkingSetMB=$p.ProcessWorkingSetMB; ProcessPrivateMemoryMB=$p.ProcessPrivateMemoryMB; ProcessPath=$p.ProcessPath; ProcessCommandLine=$p.ProcessCommandLine; ProcessStartTime=$p.ProcessStartTime;
                    Category=$p.Category; IsSystemProcess=$p.IsSystemProcess; IsUserProcess=$p.IsUserProcess; IsMonitoringRelated=$p.IsMonitoringRelated; Status='OK'; ErrorMessage=''
                }
            }
            $categoryRows += [pscustomobject]@{
                RunId=$runId; Timestamp=$serverSample.Timestamp; Server=$item.TargetServer; TargetServer=$item.TargetServer; ComputerName=$serverSample.ComputerName; LogicalProcessorCount=$serverSample.LogicalProcessorCount; CpuPercent=$serverSample.CpuPercent; ActiveSessions=$serverSample.ActiveSessions; DisconnectedSessions=$serverSample.DisconnectedSessions; UserSessionsTotal=$serverSample.UserSessionsTotal; RawSessionCount=$serverSample.RawSessionCount; TotalSessionsRaw=$serverSample.RawSessionCount; TotalSessions=$serverSample.TotalSessions;
                SecurityCpuCorePercent=$cat.SecurityCpuCorePercent; SecurityCpuServerPercent=$cat.SecurityCpuServerPercent; NexusCpuCorePercent=$cat.NexusCpuCorePercent; NexusCpuServerPercent=$cat.NexusCpuServerPercent; OfficeCpuCorePercent=$cat.OfficeCpuCorePercent; OfficeCpuServerPercent=$cat.OfficeCpuServerPercent; BrowserCpuCorePercent=$cat.BrowserCpuCorePercent; BrowserCpuServerPercent=$cat.BrowserCpuServerPercent; AdobeCpuCorePercent=$cat.AdobeCpuCorePercent; AdobeCpuServerPercent=$cat.AdobeCpuServerPercent; EdgeCpuCorePercent=$cat.EdgeCpuCorePercent; EdgeCpuServerPercent=$cat.EdgeCpuServerPercent; CitrixCpuCorePercent=$cat.CitrixCpuCorePercent; CitrixCpuServerPercent=$cat.CitrixCpuServerPercent; WemCpuCorePercent=$cat.WemCpuCorePercent; WemCpuServerPercent=$cat.WemCpuServerPercent; PrintingCpuCorePercent=$cat.PrintingCpuCorePercent; PrintingCpuServerPercent=$cat.PrintingCpuServerPercent; MonitoringCpuCorePercent=$cat.MonitoringCpuCorePercent; MonitoringCpuServerPercent=$cat.MonitoringCpuServerPercent; WindowsCpuCorePercent=$cat.WindowsCpuCorePercent; WindowsCpuServerPercent=$cat.WindowsCpuServerPercent; OtherCpuCorePercent=$cat.OtherCpuCorePercent; OtherCpuServerPercent=$cat.OtherCpuServerPercent; TopCategoryByCpuServerPercent=$cat.TopCategoryByCpuServerPercent
            }
            if ($serverSample.IsCpuWarning) {
                foreach ($p in @($item.Result.ProcessSamples | Where-Object { $_.IsAlertSample })) {
                    $alertId = '{0}_{1}_{2}' -f $runId, ($item.TargetServer -replace '[^A-Za-z0-9_.-]', '_'), ($serverSample.Timestamp -replace '[:]', '')
                    $alertRows += [pscustomobject]@{
                        AlertId=$alertId; RunId=$runId; Timestamp=$serverSample.Timestamp; Server=$item.TargetServer; TargetServer=$item.TargetServer; ComputerName=$serverSample.ComputerName; LogicalProcessorCount=$serverSample.LogicalProcessorCount; AlertSeverity=$serverSample.AlertSeverity; AlertTopCategory=$cat.TopCategoryByCpuServerPercent; CpuTotalPercent=$serverSample.CpuPercent; CpuPercent=$serverSample.CpuPercent; MemoryPercent=$serverSample.MemoryPercent; ActiveSessions=$serverSample.ActiveSessions; DisconnectedSessions=$serverSample.DisconnectedSessions; UserSessionsTotal=$serverSample.UserSessionsTotal; RawSessionCount=$serverSample.RawSessionCount; TotalSessionsRaw=$serverSample.RawSessionCount; TotalSessions=$serverSample.TotalSessions;
                        SecurityCpuServerPercent=$cat.SecurityCpuServerPercent; NexusCpuServerPercent=$cat.NexusCpuServerPercent; OfficeCpuServerPercent=$cat.OfficeCpuServerPercent; BrowserCpuServerPercent=$cat.BrowserCpuServerPercent; AdobeCpuServerPercent=$cat.AdobeCpuServerPercent; EdgeCpuServerPercent=$cat.EdgeCpuServerPercent; CitrixCpuServerPercent=$cat.CitrixCpuServerPercent; WemCpuServerPercent=$cat.WemCpuServerPercent; PrintingCpuServerPercent=$cat.PrintingCpuServerPercent; MonitoringCpuServerPercent=$cat.MonitoringCpuServerPercent; WindowsCpuServerPercent=$cat.WindowsCpuServerPercent; OtherCpuServerPercent=$cat.OtherCpuServerPercent;
                        ProcessRank=$rankByProcessId[$p.ProcessId]; InclusionReason=$p.InclusionReason; ProcessName=$p.ProcessName; ProcessId=$p.ProcessId; ParentProcessName=$p.ParentProcessName; ProcessCpuCorePercent=$p.ProcessCpuCorePercent; ProcessCpuServerPercent=$p.ProcessCpuServerPercent; ProcessCpuSecondsDelta=$p.ProcessCpuSecondsDelta; Category=$p.Category; IsMonitoringRelated=$p.IsMonitoringRelated; UserName=$p.ProcessUserName; SessionId=$p.ProcessSessionId; ProcessUserName=$p.ProcessUserName; ProcessUserDomain=$p.ProcessUserDomain; ProcessSessionId=$p.ProcessSessionId; WorkingSetMB=$p.ProcessWorkingSetMB; Path=$p.ProcessPath; CommandLine=$p.ProcessCommandLine; ProcessCommandLine=$p.ProcessCommandLine
                    }
                }
            }
            foreach ($e in @($item.Result.EventSamples)) { $eventRows += [pscustomobject]@{ RunId=$runId; Timestamp=$serverSample.Timestamp; Server=$item.TargetServer; TargetServer=$item.TargetServer; ComputerName=$serverSample.ComputerName; LogName=$e.LogName; EventTime=$e.TimeCreated; EventId=$e.Id; ProviderName=$e.ProviderName; Level=$e.LevelDisplayName; Message=$e.Message } }
        }

        Export-Rows -Rows $serverRows -Path (Join-Path $rawPath "ServerSamples_$runId.csv") -Delimiter $settings.OutputDelimiter
        Export-Rows -Rows $processRows -Path (Join-Path $rawPath "Raw_ProcessSamples_$runId.csv") -Delimiter $settings.OutputDelimiter
        Export-Rows -Rows $alertRows -Path (Join-Path $rawPath "AlertSamples_$runId.csv") -Delimiter $settings.OutputDelimiter
        Export-Rows -Rows $categoryRows -Path (Join-Path $rawPath "CategorySummary_$runId.csv") -Delimiter $settings.OutputDelimiter
        Export-Rows -Rows $eventRows -Path (Join-Path $rawPath "EventContext_$runId.csv") -Delimiter $settings.OutputDelimiter
        Export-Rows -Rows $serverRows -Path (Join-Path $runOutputPath 'ServerSamples.csv') -Delimiter $settings.OutputDelimiter
        Export-Rows -Rows $processRows -Path (Join-Path $runOutputPath 'Raw_ProcessSamples.csv') -Delimiter $settings.OutputDelimiter
        Export-Rows -Rows $alertRows -Path (Join-Path $runOutputPath 'AlertSamples.csv') -Delimiter $settings.OutputDelimiter
        Export-Rows -Rows $categoryRows -Path (Join-Path $runOutputPath 'CategorySummary.csv') -Delimiter $settings.OutputDelimiter
        Export-Rows -Rows $eventRows -Path (Join-Path $runOutputPath 'EventContext.csv') -Delimiter $settings.OutputDelimiter
        Receive-DetailDiagnosticJobs -DefenderJobs $defenderJobs -WemJobs $wemJobs -RawPath $rawPath -RunOutputPath $runOutputPath -RunId $runId -Delimiter $settings.OutputDelimiter
        $defenderJobs = @($defenderJobs | Where-Object { $_.State -eq 'Running' })
        $wemJobs = @($wemJobs | Where-Object { $_.State -eq 'Running' })
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
        $categoryAggregateRows = @(New-CategoryAggregateRows -ProcessRows $allProcessRows -RunId $runId)
        Export-Rows -Rows $categoryAggregateRows -Path (Join-Path $rawPath "CategorySummaryAggregated_$runId.csv") -Delimiter $settings.OutputDelimiter
        Export-Rows -Rows $categoryAggregateRows -Path (Join-Path $runOutputPath 'CategorySummaryAggregated.csv') -Delimiter $settings.OutputDelimiter
        Receive-DetailDiagnosticJobs -DefenderJobs $defenderJobs -WemJobs $wemJobs -RawPath $rawPath -RunOutputPath $runOutputPath -RunId $runId -Delimiter $settings.OutputDelimiter
        $runSummaryRows = @(New-RunSummaryRows -ServerRows $allServerRows -ProcessRows $allProcessRows -Start $runStart -End $runEnd -ServerCount $servers.Count)
        foreach ($summaryRow in $runSummaryRows) { $summaryRow | Add-Member -MemberType NoteProperty -Name DefenderPerfRecordingTriggered -Value $defenderTriggered -Force; $summaryRow | Add-Member -MemberType NoteProperty -Name WemEventContextTriggered -Value $wemTriggered -Force }
        $summaryCsv = Join-Path $summaryPath "RunSummary_$runId.csv"
        $summaryTxt = Join-Path $summaryPath "RunSummary_$runId.txt"
        Export-Rows -Rows $runSummaryRows -Path $summaryCsv -Delimiter $settings.OutputDelimiter
        Export-Rows -Rows $runSummaryRows -Path (Join-Path $runOutputPath 'RunSummary.csv') -Delimiter $settings.OutputDelimiter
        New-RunSummaryText -SummaryRows $runSummaryRows -Start $runStart -End $runEnd | Set-Content -LiteralPath $summaryTxt -Encoding UTF8
        New-RunSummaryText -SummaryRows $runSummaryRows -Start $runStart -End $runEnd | Set-Content -LiteralPath (Join-Path $runOutputPath 'RunSummary.txt') -Encoding UTF8
        Write-RunLog -Path $runLogFile -Message "RunSummary geschrieben: $summaryCsv"
        $runSummaryRows
    }
    catch {
        Write-RunLog -Path $runLogFile -Level 'ERROR' -Message "RunSummary konnte nicht geschrieben werden: $($_.Exception.Message)"
        Write-Warning $_.Exception.Message
    }
}
