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
    [int]$DefenderPerfFinalWaitSeconds,
    [int]$FinalizationTimeoutSeconds,
    [string]$DefenderPerfLocalRoot,
    [bool]$DefenderPerfCopyToOutputPath,
    [switch]$IncludeWemEventContext,
    [switch]$IncludeWemCpuSpikeProtectionEvents,
    [double]$WemTriggerServerCpuPercent,
    [int]$WemEventWindowMinutes,
    [switch]$IncludeWemLogTail,
    [int]$WemLogTailLines,
    [int]$WemPriorityLoweringSeconds,
    [int]$WemEventContextCooldownMinutes,
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
# Snapshot der scriptweiten BoundParameters: innerhalb von Funktionen haette `$PSBoundParameters`
# sonst nur die Funktionsparameter und CLI-/GUI-Werte wuerden von settings.json ueberschrieben.
$script:CliBoundParameters = @{}
foreach ($key in $PSBoundParameters.Keys) { $script:CliBoundParameters[$key] = $PSBoundParameters[$key] }
$script:SettingSources = @{}


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
        DefenderPerfFinalWaitSeconds = 120
        FinalizationTimeoutSeconds = 300
        DefenderPerfLocalRoot = 'C:\ProgramData\CitrixTSHealthCheck\DefenderPerf'
        DefenderPerfCopyToOutputPath = $true
        IncludeWemEventContext = $false
        IncludeWemCpuSpikeProtectionEvents = $false
        WemTriggerServerCpuPercent = 10
        WemEventWindowMinutes = 10
        IncludeWemLogTail = $false
        WemLogTailLines = 200
        WemPriorityLoweringSeconds = 180
        WemEventContextCooldownMinutes = 10
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
    $script:SettingSources = @{}
    foreach ($property in $settings.PSObject.Properties.Name) { $script:SettingSources[$property] = 'Default' }
    if (Test-Path -LiteralPath $Path) {
        $json = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($property in $settings.PSObject.Properties.Name) {
            if ($json.PSObject.Properties.Name -contains $property -and $null -ne $json.$property) {
                $settings.$property = $json.$property
                $script:SettingSources[$property] = 'Config'
            }
        }
    }
    return $settings
}


function Ensure-SettingProperty {
    param($Settings, [string]$Name, $DefaultValue)
    if (-not ($Settings.PSObject.Properties.Name -contains $Name)) {
        $Settings | Add-Member -MemberType NoteProperty -Name $Name -Value $DefaultValue
        $script:SettingSources[$Name] = 'Default'
    }
    elseif ($null -eq $Settings.$Name) {
        $Settings.$Name = $DefaultValue
        if (-not $script:SettingSources.ContainsKey($Name)) { $script:SettingSources[$Name] = 'Default' }
    }
}

function Set-SettingFromBoundParameter {
    param($Settings, [string]$SettingName, [string]$ParameterName = $SettingName, [switch]$AsBoolean)
    if ($script:CliBoundParameters.ContainsKey($ParameterName)) {
        if ($AsBoolean) { $Settings.$SettingName = [bool]$script:CliBoundParameters[$ParameterName] }
        else { $Settings.$SettingName = $script:CliBoundParameters[$ParameterName] }
        $script:SettingSources[$SettingName] = 'CLI'
    }
}

function Merge-ParameterSettings {
    param($Settings)
    foreach ($name in @('DurationMinutes','IntervalSeconds','CpuSampleSeconds','TopProcessCount','AlertTopProcessCount','CpuWarningThreshold','CpuCriticalThreshold','MaxParallel','MaxForcedProcessesPerCategory','DefenderPerfTriggerServerCpuPercent','DefenderPerfRecordingSeconds','DefenderPerfCooldownMinutes','MaxConcurrentDefenderPerfRecordings','DefenderPerfFinalWaitSeconds','FinalizationTimeoutSeconds','DefenderPerfLocalRoot','WemTriggerServerCpuPercent','WemEventWindowMinutes','WemLogTailLines','WemPriorityLoweringSeconds','WemEventContextCooldownMinutes','MaxEventsPerAlert','ImageVersion','Notes','RunId','TaskNamesToCheck')) {
        Set-SettingFromBoundParameter -Settings $Settings -SettingName $name
    }
    foreach ($name in @('AutoDefenderPerfRecording','DefenderPerfCopyToOutputPath','IncludeWemEventContext','IncludeWemCpuSpikeProtectionEvents','IncludeWemLogTail','IncludeEventLogContext','AnonymizeUsers','IncludeScheduledTaskInventory','IncludeCylanceHealth')) {
        Set-SettingFromBoundParameter -Settings $Settings -SettingName $name -AsBoolean
    }
    if ($script:CliBoundParameters.ContainsKey('Delimiter')) {
        $Settings.OutputDelimiter = $script:CliBoundParameters['Delimiter']
        $script:SettingSources['OutputDelimiter'] = 'CLI'
    }
    if ($script:CliBoundParameters.ContainsKey('IncludeEventContext')) {
        $Settings.IncludeEventLogContext = [bool]$script:CliBoundParameters['IncludeEventContext']
        $Settings.IncludeEventContext = [bool]$script:CliBoundParameters['IncludeEventContext']
        $script:SettingSources['IncludeEventLogContext'] = 'CLI'
        $script:SettingSources['IncludeEventContext'] = 'CLI'
    }
    return $Settings
}

function Get-SettingSource {
    param([string]$Name)
    if ($script:SettingSources.ContainsKey($Name)) { return $script:SettingSources[$Name] }
    return 'Default'
}

function Write-EffectiveConfiguration {
    param($Settings, [string]$RunLogFile, [datetime]$HardEndTime, [int]$MaxRounds)
    Write-RunLog -Path $RunLogFile -Message 'Effective Configuration:'
    foreach ($name in @('DurationMinutes','IntervalSeconds','CpuSampleSeconds','TopProcessCount','AlertTopProcessCount','CpuWarningThreshold','CpuCriticalThreshold','MaxParallel','MaxForcedProcessesPerCategory','DefenderPerfTriggerServerCpuPercent','DefenderPerfRecordingSeconds','DefenderPerfCooldownMinutes','MaxConcurrentDefenderPerfRecordings','DefenderPerfFinalWaitSeconds','FinalizationTimeoutSeconds','WemTriggerServerCpuPercent','WemEventWindowMinutes','WemPriorityLoweringSeconds','WemEventContextCooldownMinutes','IncludeWemEventContext','IncludeWemCpuSpikeProtectionEvents','AutoDefenderPerfRecording')) {
        Write-RunLog -Path $RunLogFile -Message ("Effective Configuration: {0}={1} Source={2}" -f $name, $Settings.$name, (Get-SettingSource -Name $name))
    }
    Write-RunLog -Path $RunLogFile -Message ("Effective Configuration: MaxRounds={0} Source=Calculated" -f $MaxRounds)
    Write-RunLog -Path $RunLogFile -Message ("Effective Configuration: HardEndTime={0} Source=Calculated" -f $HardEndTime.ToString('s'))
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
            if ($property.Name -in @('PSComputerName','RunspaceId','PSShowComputerName')) { continue }
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


$script:DefenderPerfRecordingColumns = @(
    'RunId','Server','TriggerTimestamp','TriggerProcessName','TriggerProcessCpuServerPercent','TriggerThreshold','TriggerReason',
    'Status','StartTime','EndTime','DurationSeconds','TriggerToRecordingStartSeconds','EtlLocalPath','LogLocalPath','ReportTxtLocalPath','ReportJsonLocalPath',
    'EtlCentralPath','LogCentralPath','ReportTxtCentralPath','ReportJsonCentralPath','ReportGenerationStatus','CopyStatus','ErrorMessage'
)

function ConvertTo-DefenderPerfRecordingRow {
    param([Parameter(ValueFromPipeline=$true)]$InputObject)
    process {
        $row = [ordered]@{}
        foreach ($column in $script:DefenderPerfRecordingColumns) {
            if ($null -ne $InputObject -and $InputObject.PSObject.Properties.Name -contains $column) { $row[$column] = $InputObject.$column }
            else { $row[$column] = '' }
        }
        [pscustomobject]$row
    }
}

function New-DefenderPerfRecordingRow {
    param([hashtable]$Values)
    $row = [ordered]@{}
    foreach ($column in $script:DefenderPerfRecordingColumns) {
        if ($Values.ContainsKey($column)) { $row[$column] = $Values[$column] }
        else { $row[$column] = '' }
    }
    [pscustomobject]$row
}

function Get-DefenderRecordingStatusRank {
    param([string]$Status)
    switch ($Status) {
        'Pending' { return 1 }
        'Running' { return 2 }
        'Completed' { return 3 }
        'ReportFailed' { return 4 }
        'CopyFailed' { return 5 }
        'Failed' { return 6 }
        'TimedOut' { return 7 }
        default { return 0 }
    }
}

function Export-DefenderPerfRecordingRows {
    param([array]$Rows, [string]$RawPath, [string]$RunOutputPath, [string]$RunId, [string]$Delimiter, [string]$RunLogFile = '')
    if (-not $Rows -or $Rows.Count -eq 0) { return }
    $normalizedRows = @($Rows | ConvertTo-DefenderPerfRecordingRow)
    foreach ($path in @((Join-Path $RawPath "DefenderPerfRecordings_$RunId.csv"), (Join-Path $RunOutputPath 'DefenderPerfRecordings.csv'))) {
        try {
            $existingRows = @()
            if (Test-Path -LiteralPath $path) { $existingRows = @(Import-Csv -LiteralPath $path -Delimiter $Delimiter -ErrorAction Stop | ConvertTo-DefenderPerfRecordingRow) }
            $combinedRows = @($existingRows + $normalizedRows)
            $dedupedRows = @(
                $combinedRows |
                    Group-Object { '{0}|{1}|{2}|{3}' -f $_.RunId, $_.Server, $_.TriggerTimestamp, $_.TriggerProcessName } |
                    ForEach-Object {
                        @($_.Group | Sort-Object @{ Expression = { Get-DefenderRecordingStatusRank -Status $_.Status } }, @{ Expression = { $_.EndTime } }, @{ Expression = { $_.StartTime } } | Select-Object -Last 1)
                    }
            )
            $dedupedRows | ConvertTo-InvariantObject | Export-Csv -LiteralPath $path -Delimiter $Delimiter -NoTypeInformation -Encoding UTF8
        }
        catch {
            if ($RunLogFile) { Write-RunLog -Path $RunLogFile -Level 'WARN' -Message "DefenderPerfRecordings Export fehlgeschlagen ($path): $($_.Exception.Message)" }
        }
    }
}

function Write-RunLog {
    param([string]$Path, [string]$Message, [string]$Level = 'INFO')
    $line = '{0};{1};{2}' -f (Get-Date).ToString('s'), $Level, $Message
    Add-Content -LiteralPath $Path -Value $line -Encoding UTF8
}

function Get-SafePropertyValue {
    param($Object, [string]$PropertyName, $DefaultValue = $null)
    try {
        if ($null -eq $Object -or [string]::IsNullOrWhiteSpace($PropertyName)) { return $DefaultValue }
        $property = $Object.PSObject.Properties[$PropertyName]
        if ($null -eq $property) { return $DefaultValue }
        if ($null -eq $property.Value) { return $DefaultValue }
        return $property.Value
    }
    catch { return $DefaultValue }
}

function Get-RemoteSamplerScriptBlock {
    {
        param($CpuSampleSeconds, $TopProcessCount, $AlertTopProcessCount, $CpuWarningThreshold, $CpuCriticalThreshold, $IncludeEventLogContext, $AnonymizeUsers, $MaxEventsPerAlert, $MaxForcedProcessesPerCategory)

        function Get-SafePropertyValue {
            param($Object, [string]$PropertyName, $DefaultValue = $null)
            try {
                if ($null -eq $Object -or [string]::IsNullOrWhiteSpace($PropertyName)) { return $DefaultValue }
                $property = $Object.PSObject.Properties[$PropertyName]
                if ($null -eq $property) { return $DefaultValue }
                if ($null -eq $property.Value) { return $DefaultValue }
                return $property.Value
            }
            catch { return $DefaultValue }
        }

        function New-StableProcessSnapshot {
            param($Process)
            $processId = Get-SafePropertyValue -Object $Process -PropertyName 'Id'
            if ($null -eq $processId) { return $null }
            $startTimeText = $null
            $startTime = Get-SafePropertyValue -Object $Process -PropertyName 'StartTime'
            if ($startTime) { try { $startTimeText = ([datetime]$startTime).ToString('s') } catch { $startTimeText = $null } }
            [pscustomobject]@{
                Id = [int]$processId
                ProcessId = [int]$processId
                ProcessName = [string](Get-SafePropertyValue -Object $Process -PropertyName 'ProcessName' -DefaultValue '')
                CPU = Get-SafePropertyValue -Object $Process -PropertyName 'CPU' -DefaultValue 0
                SessionId = Get-SafePropertyValue -Object $Process -PropertyName 'SessionId' -DefaultValue $null
                WorkingSet64 = Get-SafePropertyValue -Object $Process -PropertyName 'WorkingSet64' -DefaultValue $null
                PrivateMemorySize64 = Get-SafePropertyValue -Object $Process -PropertyName 'PrivateMemorySize64' -DefaultValue $null
                ProcessStartTime = $startTimeText
                StartTime = $startTimeText
                PriorityClass = Get-SafePropertyValue -Object $Process -PropertyName 'PriorityClass' -DefaultValue $null
                BasePriority = Get-SafePropertyValue -Object $Process -PropertyName 'BasePriority' -DefaultValue $null
            }
        }

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
        $processStart = @(Get-Process -ErrorAction SilentlyContinue | ForEach-Object { New-StableProcessSnapshot -Process $_ } | Where-Object { $_ })
        Start-Sleep -Seconds $CpuSampleSeconds
        $cpuEndPercent = Get-TotalCpuPercent
        $processEnd = @(Get-Process -ErrorAction SilentlyContinue | ForEach-Object { New-StableProcessSnapshot -Process $_ } | Where-Object { $_ })
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
            $startTimeText = Get-SafePropertyValue -Object $process -PropertyName 'ProcessStartTime' -DefaultValue $null
            [pscustomobject]@{
                ProcessId = [int]$process.Id
                ProcessName = [string]$process.ProcessName
                SessionId = if ($null -ne (Get-SafePropertyValue -Object $process -PropertyName 'SessionId' -DefaultValue $null)) { [int](Get-SafePropertyValue -Object $process -PropertyName 'SessionId') } else { -1 }
                ProcessCpuCorePercent = $cpuCorePct
                ProcessCpuServerPercent = $cpuServerPct
                CpuSecondsDelta = [Math]::Round($delta, 3)
                WorkingSetMb = if ($null -ne (Get-SafePropertyValue -Object $process -PropertyName 'WorkingSet64' -DefaultValue $null)) { [Math]::Round(([double](Get-SafePropertyValue -Object $process -PropertyName 'WorkingSet64')) / 1MB, 2) } else { $null }
                PrivateMemoryMb = if ($null -ne (Get-SafePropertyValue -Object $process -PropertyName 'PrivateMemorySize64' -DefaultValue $null)) { [Math]::Round(([double](Get-SafePropertyValue -Object $process -PropertyName 'PrivateMemorySize64')) / 1MB, 2) } else { $null }
                ParentProcessId = $parentProcessId
                ParentProcessName = $parentProcessName
                StartTime = $startTimeText
                PriorityClass = Get-SafePropertyValue -Object $process -PropertyName 'PriorityClass' -DefaultValue $null
                BasePriority = Get-SafePropertyValue -Object $process -PropertyName 'BasePriority' -DefaultValue $null
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
                PriorityClass = Get-SafePropertyValue -Object $p -PropertyName 'PriorityClass' -DefaultValue $null
                BasePriority = Get-SafePropertyValue -Object $p -PropertyName 'BasePriority' -DefaultValue $null
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
                [pscustomobject]@{ TargetServer = $Server; Status = 'ERROR'; Result = $null; ErrorMessage = $_.Exception.Message; ErrorType = $_.Exception.GetType().FullName; ScriptStackTrace = $_.ScriptStackTrace; ObjectType = ''; PropertyNames = ''; ProcessName = ''; ProcessId = '' }
            }
        } -ArgumentList $server, $remoteScript, $Settings
    }
    while ($jobs.Count -gt 0) {
        $done = Wait-Job -Job $jobs -Any -Timeout 5
        if ($done) { $results += Receive-Job -Job $done; Remove-Job -Job $done -Force; $jobs = @($jobs | Where-Object Id -ne $done.Id) }
    }
    return $results
}


function Normalize-TaskNameList {
    param([string[]]$TaskNames)
    $normalized = @()
    foreach ($entry in @($TaskNames)) {
        if ([string]::IsNullOrWhiteSpace([string]$entry)) { continue }
        $entryText = [string]$entry
        $quotedMatches = [regex]::Matches($entryText, '"([^"]+)"|''([^'']+)''')
        if ($quotedMatches.Count -gt 1) {
            foreach ($match in $quotedMatches) {
                $clean = if ($match.Groups[1].Success) { $match.Groups[1].Value } else { $match.Groups[2].Value }
                $clean = $clean.Trim()
                if (-not [string]::IsNullOrWhiteSpace($clean)) { $normalized += $clean }
            }
            continue
        }
        $parts = @($entryText -split "[,;`r`n]+")
        foreach ($part in $parts) {
            $clean = ([string]$part).Trim().Trim('\"').Trim("'").Trim()
            if (-not [string]::IsNullOrWhiteSpace($clean)) { $normalized += $clean }
        }
    }
    return @($normalized | Select-Object -Unique)
}

function Invoke-ScheduledTaskInventory {
    param([string[]]$Servers, [string[]]$TaskNames, [string]$RunId)
    $taskNameList = @(Normalize-TaskNameList -TaskNames $TaskNames)
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
            } -ArgumentList $taskNameList, $RunId -ErrorAction Stop
            $rows += $remoteRows
        }
        catch {
            foreach ($taskName in $taskNameList) { $rows += [pscustomobject]@{ RunId=$RunId; Server=$server; TaskName=$taskName; TaskPath=''; State='ERROR'; Enabled=''; LastRunTime=''; LastTaskResult=''; NextRunTime=''; Author=''; Description=''; Actions=''; Triggers=''; Status='ERROR'; ErrorMessage=$_.Exception.Message } }
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
    param(
        [string]$Server,
        [string]$RunId,
        [string]$OutputPath,
        [string]$LocalRoot,
        [bool]$CopyToOutputPath,
        [int]$Seconds,
        [string]$TriggerProcess,
        [double]$TriggerCpu,
        [double]$TriggerThreshold,
        [datetime]$TriggerTime
    )
    $job = Start-Job -ScriptBlock {
        param($Server, $RunId, $OutputPath, $LocalRoot, $CopyToOutputPath, $Seconds, $TriggerProcess, $TriggerCpu, $TriggerThreshold, $TriggerTime)
        $jobStartTime = Get-Date
        $safeServer = $Server -replace '[^A-Za-z0-9_.-]', '_'
        $triggerStamp = $TriggerTime.ToString('yyyyMMdd_HHmmss')
        $centralFolder = Join-Path (Join-Path (Join-Path $OutputPath 'DefenderPerf') $RunId) $safeServer
        $session = $null
        try {
            if ($CopyToOutputPath -and -not (Test-Path -LiteralPath $centralFolder)) { New-Item -ItemType Directory -Path $centralFolder -Force | Out-Null }
            $session = New-PSSession -ComputerName $Server -ErrorAction Stop
            $remote = Invoke-Command -Session $session -ScriptBlock {
                param($RunId, $LocalRoot, $Seconds, $TriggerStamp, $SafeServer)
                $remoteFolder = Join-Path (Join-Path $LocalRoot $RunId) $SafeServer
                if (-not (Test-Path -LiteralPath $remoteFolder)) { New-Item -ItemType Directory -Path $remoteFolder -Force | Out-Null }
                $recordingBase = "DefenderPerfRecording_{0}_{1}_{2}" -f $RunId, $SafeServer, $TriggerStamp
                $reportBase = "DefenderPerfReport_{0}_{1}_{2}" -f $RunId, $SafeServer, $TriggerStamp
                $etlPath = Join-Path $remoteFolder ($recordingBase + '.etl')
                $logPath = Join-Path $remoteFolder ($recordingBase + '.log')
                $reportTxtPath = Join-Path $remoteFolder ($reportBase + '.txt')
                $reportJsonPath = Join-Path $remoteFolder ($reportBase + '_raw.json')
                $status = 'Completed'
                $reportStatus = 'NotStarted'
                $errorMessage = ''
                try {
                    & { New-MpPerformanceRecording -RecordTo $etlPath -Seconds $Seconds -ErrorAction Stop } *> $logPath
                }
                catch {
                    $status = 'Failed'
                    $errorMessage = "New-MpPerformanceRecording fehlgeschlagen: $($_.Exception.Message)"
                    try { $errorMessage | Add-Content -LiteralPath $logPath -Encoding UTF8 } catch { }
                }
                try {
                    $reportArgs = @{ Path=$etlPath; TopFiles=50; TopPaths=50; TopExtensions=50; TopProcesses=50; TopScans=100; ErrorAction='Stop' }
                    $report = Get-MpPerformanceReport @reportArgs
                    $report | Out-String -Width 4096 | Set-Content -LiteralPath $reportTxtPath -Encoding UTF8
                    $rawReport = Get-MpPerformanceReport @reportArgs -Raw
                    $rawReport | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $reportJsonPath -Encoding UTF8
                    $reportStatus = 'Completed'
                }
                catch {
                    if ($status -eq 'Completed') { $status = 'ReportFailed' }
                    $reportStatus = 'Failed'
                    $reportError = "Get-MpPerformanceReport fehlgeschlagen: $($_.Exception.Message)"
                    $errorMessage = @($errorMessage, $reportError | Where-Object { $_ }) -join ' | '
                    $reportError | Set-Content -LiteralPath $reportTxtPath -Encoding UTF8
                    [pscustomobject]@{ Error=$reportError } | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $reportJsonPath -Encoding UTF8
                }
                [pscustomobject]@{
                    ComputerName=$env:COMPUTERNAME; RemoteFolder=$remoteFolder; EtlLocalPath=$etlPath; LogLocalPath=$logPath;
                    ReportTxtLocalPath=$reportTxtPath; ReportJsonLocalPath=$reportJsonPath; Status=$status; ReportGenerationStatus=$reportStatus; ErrorMessage=$errorMessage
                }
            } -ArgumentList $RunId, $LocalRoot, $Seconds, $triggerStamp, $safeServer -ErrorAction Stop

            $copyStatus = if ($CopyToOutputPath) { 'Pending' } else { 'Skipped' }
            $centralPaths = @{ Etl=''; Log=''; ReportTxt=''; ReportJson='' }
            if ($CopyToOutputPath) {
                $copyStatus = 'Completed'
                foreach ($entry in @(
                    @{ Key='Etl'; RemotePath=$remote.EtlLocalPath },
                    @{ Key='Log'; RemotePath=$remote.LogLocalPath },
                    @{ Key='ReportTxt'; RemotePath=$remote.ReportTxtLocalPath },
                    @{ Key='ReportJson'; RemotePath=$remote.ReportJsonLocalPath }
                )) {
                    $remotePath = [string]$entry['RemotePath']
                    if ([string]::IsNullOrWhiteSpace($remotePath)) { continue }
                    $fileName = Split-Path -Path $remotePath -Leaf
                    $destination = Join-Path $centralFolder $fileName
                    try {
                        Copy-Item -FromSession $session -LiteralPath $remotePath -Destination $destination -Force -ErrorAction Stop
                        $centralPaths[$entry['Key']] = $destination
                    }
                    catch {
                        $copyStatus = 'Failed'
                        if ($remote.Status -eq 'Completed') { $remote.Status = 'CopyFailed' }
                        $copyError = "Kopieren fehlgeschlagen ($remotePath): $($_.Exception.Message)"
                        $remote.ErrorMessage = @($remote.ErrorMessage, $copyError | Where-Object { $_ }) -join ' | '
                    }
                }
            }
            $etlResultPath = if ($centralPaths['Etl']) { $centralPaths['Etl'] } else { $remote.EtlLocalPath }
            $logResultPath = if ($centralPaths['Log']) { $centralPaths['Log'] } else { $remote.LogLocalPath }
            $reportResultPath = if ($centralPaths['ReportTxt']) { $centralPaths['ReportTxt'] } else { $remote.ReportTxtLocalPath }
            $rawJsonResultPath = if ($centralPaths['ReportJson']) { $centralPaths['ReportJson'] } else { $remote.ReportJsonLocalPath }
            $jobEndTime = Get-Date
            [pscustomobject][ordered]@{
                RunId=$RunId; Server=$Server; TriggerTimestamp=$TriggerTime.ToString('s');
                TriggerProcessName=$TriggerProcess; TriggerProcessCpuServerPercent=$TriggerCpu; TriggerThreshold=$TriggerThreshold; TriggerReason='MsMpEng ProcessCpuServerPercent >= DefenderPerfTriggerServerCpuPercent';
                Status=$remote.Status; StartTime=$jobStartTime.ToString('s'); EndTime=$jobEndTime.ToString('s'); DurationSeconds=[Math]::Round(($jobEndTime - $jobStartTime).TotalSeconds, 1); TriggerToRecordingStartSeconds=[Math]::Round(($jobStartTime - $TriggerTime).TotalSeconds, 1);
                EtlLocalPath=$remote.EtlLocalPath; LogLocalPath=$remote.LogLocalPath; ReportTxtLocalPath=$remote.ReportTxtLocalPath; ReportJsonLocalPath=$remote.ReportJsonLocalPath;
                EtlCentralPath=$centralPaths['Etl']; LogCentralPath=$centralPaths['Log']; ReportTxtCentralPath=$centralPaths['ReportTxt']; ReportJsonCentralPath=$centralPaths['ReportJson'];
                ReportGenerationStatus=$remote.ReportGenerationStatus; CopyStatus=$copyStatus; ErrorMessage=$remote.ErrorMessage
            }
        }
        catch {
            $jobEndTime = Get-Date
            [pscustomobject][ordered]@{ RunId=$RunId; Server=$Server; TriggerTimestamp=$TriggerTime.ToString('s'); TriggerProcessName=$TriggerProcess; TriggerProcessCpuServerPercent=$TriggerCpu; TriggerThreshold=$TriggerThreshold; TriggerReason='MsMpEng ProcessCpuServerPercent >= DefenderPerfTriggerServerCpuPercent'; Status='Failed'; StartTime=$jobStartTime.ToString('s'); EndTime=$jobEndTime.ToString('s'); DurationSeconds=[Math]::Round(($jobEndTime - $jobStartTime).TotalSeconds, 1); TriggerToRecordingStartSeconds=[Math]::Round(($jobStartTime - $TriggerTime).TotalSeconds, 1); EtlLocalPath=''; LogLocalPath=''; ReportTxtLocalPath=''; ReportJsonLocalPath=''; EtlCentralPath=''; LogCentralPath=''; ReportTxtCentralPath=''; ReportJsonCentralPath=''; ReportGenerationStatus='Failed'; CopyStatus='NotStarted'; ErrorMessage=$_.Exception.Message }
        }
        finally {
            if ($session) { Remove-PSSession -Session $session -ErrorAction SilentlyContinue }
        }
    } -ArgumentList $Server, $RunId, $OutputPath, $LocalRoot, $CopyToOutputPath, $Seconds, $TriggerProcess, $TriggerCpu, $TriggerThreshold, $TriggerTime
    $job | Add-Member -MemberType NoteProperty -Name DiagnosticType -Value 'Defender' -Force
    $job | Add-Member -MemberType NoteProperty -Name TargetServer -Value $Server -Force
    $job | Add-Member -MemberType NoteProperty -Name TriggerTime -Value $TriggerTime -Force
    $job | Add-Member -MemberType NoteProperty -Name TriggerProcess -Value $TriggerProcess -Force
    $job | Add-Member -MemberType NoteProperty -Name TriggerCpu -Value $TriggerCpu -Force
    $job | Add-Member -MemberType NoteProperty -Name TriggerThreshold -Value $TriggerThreshold -Force
    return $job
}

function Start-WemEventContextJob {
    param(
        [string]$Server,
        [string]$RunId,
        [string]$AlertId,
        [datetime]$TriggerTime,
        [int]$WindowMinutes,
        [bool]$IncludeLogTail,
        [int]$TailLines,
        [int]$MaxEventsPerAlert,
        [string]$TriggerProcessName,
        [double]$TriggerProcessCpuServerPercent,
        [double]$TriggerThreshold
    )
    $job = Start-Job -ScriptBlock {
        param($Server, $RunId, $AlertId, $TriggerTime, $WindowMinutes, $IncludeLogTail, $TailLines, $MaxEventsPerAlert, $TriggerProcessName, $TriggerProcessCpuServerPercent, $TriggerThreshold)
        try {
            Invoke-Command -ComputerName $Server -ScriptBlock {
                param($RunId, $AlertId, $TriggerTime, $WindowMinutes, $IncludeLogTail, $TailLines, $MaxEventsPerAlert, $TriggerProcessName, $TriggerProcessCpuServerPercent, $TriggerThreshold)
                $start = $TriggerTime.AddMinutes(-1 * $WindowMinutes)
                $end = $TriggerTime.AddMinutes($WindowMinutes)
                $logNames = @()
                try { $logNames += Get-WinEvent -ListLog '*WEM*' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty LogName } catch { }
                $logNames += @('WEM Agent Service','Citrix WEM Agent Service','Norskale Agent Service')
                $logNames = @($logNames | Where-Object { $_ } | Select-Object -Unique)
                $remainingEvents = [Math]::Max(0, [int]$MaxEventsPerAlert)
                foreach ($logName in $logNames) {
                    if ($remainingEvents -le 0) { break }
                    try {
                        $events = @(Get-WinEvent -FilterHashtable @{ LogName=$logName; StartTime=$start; EndTime=$end } -MaxEvents $remainingEvents -ErrorAction Stop)
                        foreach ($event in $events) {
                            [pscustomobject]@{ RecordType='Event'; RunId=$RunId; Server=$env:COMPUTERNAME; AlertId=$AlertId; TriggerProcessName=$TriggerProcessName; TriggerProcessCpuServerPercent=$TriggerProcessCpuServerPercent; TriggerThreshold=$TriggerThreshold; TriggerTimestamp=$TriggerTime.ToString('s'); TimeCreated=$event.TimeCreated; LogName=$logName; ProviderName=$event.ProviderName; EventId=$event.Id; Level=$event.LevelDisplayName; Message=$event.Message; FileName=''; Content='' }
                        }
                        $remainingEvents -= $events.Count
                    } catch {
                        [pscustomobject]@{ RecordType='Error'; RunId=$RunId; Server=$env:COMPUTERNAME; AlertId=$AlertId; TriggerProcessName=$TriggerProcessName; TriggerProcessCpuServerPercent=$TriggerProcessCpuServerPercent; TriggerThreshold=$TriggerThreshold; TriggerTimestamp=$TriggerTime.ToString('s'); TimeCreated=''; LogName=$logName; ProviderName=''; EventId=''; Level=''; Message="WEM Log nicht lesbar oder nicht vorhanden: $($_.Exception.Message)"; FileName=''; Content='' }
                    }
                }
                if ($IncludeLogTail) {
                    $paths = @('C:\Program Files (x86)\Norskale','C:\Program Files (x86)\Citrix\Workspace Environment Management Agent','C:\ProgramData\Norskale','C:\ProgramData\Citrix\WEM')
                    foreach ($base in $paths) {
                        try {
                            Get-ChildItem -LiteralPath $base -Filter '*.log' -Recurse -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 5 | ForEach-Object {
                                [pscustomobject]@{ RecordType='Tail'; RunId=$RunId; Server=$env:COMPUTERNAME; AlertId=$AlertId; TriggerProcessName=$TriggerProcessName; TriggerProcessCpuServerPercent=$TriggerProcessCpuServerPercent; TriggerThreshold=$TriggerThreshold; TriggerTimestamp=$TriggerTime.ToString('s'); TimeCreated=''; LogName=''; ProviderName=''; EventId=''; Level=''; Message=''; FileName=$_.FullName; Content=((Get-Content -LiteralPath $_.FullName -Tail $TailLines -ErrorAction SilentlyContinue) -join [Environment]::NewLine) }
                            }
                        } catch { }
                    }
                }
            } -ArgumentList $RunId, $AlertId, $TriggerTime, $WindowMinutes, $IncludeLogTail, $TailLines, $MaxEventsPerAlert, $TriggerProcessName, $TriggerProcessCpuServerPercent, $TriggerThreshold -ErrorAction Stop
        }
        catch { [pscustomobject]@{ RecordType='Error'; RunId=$RunId; Server=$Server; AlertId=$AlertId; TriggerProcessName=$TriggerProcessName; TriggerProcessCpuServerPercent=$TriggerProcessCpuServerPercent; TriggerThreshold=$TriggerThreshold; TriggerTimestamp=$TriggerTime.ToString('s'); TimeCreated=''; LogName=''; ProviderName=''; EventId=''; Level=''; Message=$_.Exception.Message; FileName=''; Content='' } }
    } -ArgumentList $Server, $RunId, $AlertId, $TriggerTime, $WindowMinutes, $IncludeLogTail, $TailLines, $MaxEventsPerAlert, $TriggerProcessName, $TriggerProcessCpuServerPercent, $TriggerThreshold
    $job | Add-Member -MemberType NoteProperty -Name DiagnosticType -Value 'WEM' -Force
    $job | Add-Member -MemberType NoteProperty -Name TargetServer -Value $Server -Force
    $job | Add-Member -MemberType NoteProperty -Name TriggerTime -Value $TriggerTime -Force
    $job | Add-Member -MemberType NoteProperty -Name AlertId -Value $AlertId -Force
    return $job
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

function ConvertTo-NullableDouble {
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [double] -or $Value -is [single] -or $Value -is [decimal] -or $Value -is [int] -or $Value -is [long]) { return [double]$Value }
    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    try { return [double]::Parse($text, [Globalization.CultureInfo]::InvariantCulture) } catch { }
    try { return [double]::Parse($text, [Globalization.CultureInfo]::CurrentCulture) } catch { }
    return $null
}

function Get-NumericValues {
    param([array]$Rows, [string]$PropertyName, [switch]$Sort)
    $values = @()
    foreach ($row in @($Rows)) {
        if ($null -eq $row -or -not ($row.PSObject.Properties.Name -contains $PropertyName)) { continue }
        $number = ConvertTo-NullableDouble -Value $row.$PropertyName
        if ($null -ne $number) { $values += $number }
    }
    if ($Sort) { return @($values | Sort-Object) }
    return @($values)
}

function Get-PercentileValue {
    param([double[]]$Values, [double]$Percentile)
    if (-not $Values -or $Values.Count -eq 0) { return $null }
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
        $coreValues = @(Get-NumericValues -Rows $rows -PropertyName 'ProcessCpuCorePercent')
        $serverValues = @(Get-NumericValues -Rows $rows -PropertyName 'ProcessCpuServerPercent')
        $wsValues = @(Get-NumericValues -Rows $rows -PropertyName 'ProcessWorkingSetMB')
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
    param([array]$ServerRows, [array]$ProcessRows, [datetime]$Start, [datetime]$End, [int]$ServerCount, [string]$EndReason = 'Completed', [int]$PlannedDurationMinutes = 0, [double]$ActualDurationMinutes = 0, [int]$PlannedRounds = 0, [int]$CompletedRounds = 0)
    $categoryNames = @('Security','Nexus','Office','Browser','Adobe','Edge','Citrix','Wem','Printing','Monitoring','Windows','Other')
    foreach ($group in ($ServerRows | Group-Object TargetServer)) {
        $server = $group.Name
        $okRows = @($group.Group | Where-Object Status -eq 'OK')
        $cpuValues = @(Get-NumericValues -Rows $okRows -PropertyName 'CpuPercent' -Sort)
        $ramValues = @(Get-NumericValues -Rows $okRows -PropertyName 'MemoryPercent')
        $activeValues = @(Get-NumericValues -Rows $okRows -PropertyName 'ActiveSessions')
        $discValues = @(Get-NumericValues -Rows $okRows -PropertyName 'DisconnectedSessions')
        $median = Get-PercentileValue -Values $cpuValues -Percentile 50
        $p95 = Get-PercentileValue -Values $cpuValues -Percentile 95
        $serverProcessRows = @($ProcessRows | Where-Object TargetServer -eq $server)
        $topProc = Join-TopCpuGroups -Rows $serverProcessRows -GroupProperty ProcessName -Count 1
        $top10Proc = Join-TopCpuGroups -Rows $serverProcessRows -GroupProperty ProcessName -Count 10
        $topCat = Join-TopCpuGroups -Rows $serverProcessRows -GroupProperty Category -Count 1
        $categoryTotals = @{}
        foreach ($category in $categoryNames) { $categoryTotals[$category] = (Get-SumValue -Rows @($serverProcessRows | Where-Object Category -eq $category) -PropertyName 'ProcessCpuServerPercent') }
        [pscustomobject]@{
            RunId = $runId; StartTime = $Start.ToString('s'); EndTime = $End.ToString('s'); EndReason = $EndReason; PlannedDurationMinutes = $PlannedDurationMinutes; ActualDurationMinutes = $ActualDurationMinutes; PlannedRounds = $PlannedRounds; CompletedRounds = $CompletedRounds; IntervalSeconds = $settings.IntervalSeconds; CpuSampleSeconds = $settings.CpuSampleSeconds; ServerList = ($servers -join ','); ImageVersion = $settings.ImageVersion; Notes = $settings.Notes; ScriptVersion = '1.0.0'; OutputPath = $OutputPath; RunStart = $Start.ToString('s'); RunEnd = $End.ToString('s'); DurationMinutes = [Math]::Round(($End-$Start).TotalMinutes, 2); ServerCount = $ServerCount; TargetServer = $server;
            SampleCount = $okRows.Count; ErrorCount = @($group.Group | Where-Object Status -eq 'ERROR').Count;
            CpuAverage = if ($cpuValues.Count) { [Math]::Round((($cpuValues | Measure-Object -Average).Average),2) } else { '' }; CpuMedian = if ($null -ne $median) { [Math]::Round($median,2) } else { '' }; CpuMax = if ($cpuValues.Count) { [Math]::Round((($cpuValues | Measure-Object -Maximum).Maximum),2) } else { '' }; CpuP95 = if ($null -ne $p95) { [Math]::Round($p95,2) } else { '' };
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
    param([array]$DefenderJobs, [array]$WemJobs, [string]$RawPath, [string]$RunOutputPath, [string]$RunId, [string]$Delimiter, [string]$RunLogFile = '')
    foreach ($job in @($DefenderJobs | Where-Object { $_.State -ne 'Running' })) {
        try {
            $rows = @(Receive-Job -Job $job -ErrorAction SilentlyContinue)
            Export-DefenderPerfRecordingRows -Rows $rows -RawPath $RawPath -RunOutputPath $RunOutputPath -RunId $RunId -Delimiter $Delimiter -RunLogFile $RunLogFile
        } finally { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue }
    }
    foreach ($job in @($WemJobs | Where-Object { $_.State -ne 'Running' })) {
        try {
            $rows = @(Receive-Job -Job $job -ErrorAction SilentlyContinue)
            $eventRows = @($rows | Where-Object { $_.RecordType -eq 'Event' } | Select-Object RunId,Server,AlertId,TriggerProcessName,TriggerProcessCpuServerPercent,TriggerThreshold,TriggerTimestamp,TimeCreated,LogName,ProviderName,EventId,Level,Message)
            foreach ($errorRow in @($rows | Where-Object { $_.RecordType -eq 'Error' })) { if ($RunLogFile) { Write-RunLog -Path $RunLogFile -Level 'WARN' -Message "WEM EventContext: $($errorRow.Server): $($errorRow.Message)" } }
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


function Complete-DetailDiagnosticJobs {
    param(
        [array]$DefenderJobs,
        [array]$WemJobs,
        [string]$RawPath,
        [string]$RunOutputPath,
        [string]$RunId,
        [string]$Delimiter,
        [int]$DefenderFinalWaitSeconds,
        [int]$FinalizationTimeoutSeconds,
        [string]$RunLogFile
    )
    $finalDeadline = (Get-Date).AddSeconds([Math]::Max(0, $FinalizationTimeoutSeconds))
    $defenderDeadline = (Get-Date).AddSeconds([Math]::Max(0, $DefenderFinalWaitSeconds))
    do {
        Receive-DetailDiagnosticJobs -DefenderJobs $DefenderJobs -WemJobs $WemJobs -RawPath $RawPath -RunOutputPath $RunOutputPath -RunId $RunId -Delimiter $Delimiter -RunLogFile $RunLogFile
        $DefenderJobs = @($DefenderJobs | Where-Object { $_.State -eq 'Running' })
        $WemJobs = @($WemJobs | Where-Object { $_.State -eq 'Running' })
        if (($DefenderJobs.Count + $WemJobs.Count) -eq 0) { break }
        if ((Get-Date) -ge $finalDeadline) { break }
        if ((Get-Date) -ge $defenderDeadline -and $DefenderJobs.Count -gt 0) { break }
        Start-Sleep -Seconds 2
    } while ($true)

    foreach ($job in @($DefenderJobs | Where-Object { $_.State -eq 'Running' })) {
        $server = if ($job.PSObject.Properties.Name -contains 'TargetServer') { $job.TargetServer } else { '' }
        $triggerTime = if ($job.PSObject.Properties.Name -contains 'TriggerTime') { ([datetime]$job.TriggerTime).ToString('s') } else { (Get-Date).ToString('s') }
        $triggerProcess = if ($job.PSObject.Properties.Name -contains 'TriggerProcess') { $job.TriggerProcess } else { '' }
        $triggerCpu = if ($job.PSObject.Properties.Name -contains 'TriggerCpu') { $job.TriggerCpu } else { '' }
        $centralFolder = if ($server) { Join-Path (Join-Path (Join-Path $OutputPath 'DefenderPerf') $RunId) ($server -replace '[^A-Za-z0-9_.-]', '_') } else { '' }
        $triggerThreshold = if ($job.PSObject.Properties.Name -contains 'TriggerThreshold') { $job.TriggerThreshold } else { '' }
        $timeoutRow = New-DefenderPerfRecordingRow -Values @{ RunId=$RunId; Server=$server; TriggerTimestamp=$triggerTime; TriggerProcessName=$triggerProcess; TriggerProcessCpuServerPercent=$triggerCpu; TriggerThreshold=$triggerThreshold; TriggerReason='MsMpEng ProcessCpuServerPercent >= DefenderPerfTriggerServerCpuPercent'; Status='TimedOut'; StartTime=$triggerTime; EndTime=(Get-Date).ToString('s'); DurationSeconds=''; ReportGenerationStatus='TimedOut'; CopyStatus='NotStarted'; ErrorMessage='Defender Performance Recording lief nach Ende des Hauptlaufs noch und wurde nicht weiter abgewartet.' }
        Export-DefenderPerfRecordingRows -Rows @($timeoutRow) -RawPath $RawPath -RunOutputPath $RunOutputPath -RunId $RunId -Delimiter $Delimiter -RunLogFile $RunLogFile
        Write-RunLog -Path $RunLogFile -Level 'WARN' -Message "Defender Performance Recording timed out: $server"
        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
    }
    foreach ($job in @($WemJobs | Where-Object { $_.State -eq 'Running' })) {
        $server = if ($job.PSObject.Properties.Name -contains 'TargetServer') { $job.TargetServer } else { '' }
        Write-RunLog -Path $RunLogFile -Level 'WARN' -Message "WEM Detaildiagnose wegen FinalizationTimeout abgebrochen: $server"
        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
    }
}


function Import-CsvIfExists {
    param([string]$Path, [string]$Delimiter)
    if (Test-Path -LiteralPath $Path) {
        try { return @(Import-Csv -LiteralPath $Path -Delimiter $Delimiter -ErrorAction Stop) } catch { return @() }
    }
    return @()
}

function Export-EmptyCsv {
    param([string]$Path, [string[]]$Columns, [string]$Delimiter)
    if (Test-Path -LiteralPath $Path) { return }
    $row = [ordered]@{}
    foreach ($column in $Columns) { $row[$column] = $null }
    [pscustomobject]$row | Export-Csv -LiteralPath $Path -Delimiter $Delimiter -NoTypeInformation -Encoding UTF8
    $content = @(Get-Content -LiteralPath $Path -Encoding UTF8)
    if ($content.Count -gt 0) { $content[0] | Set-Content -LiteralPath $Path -Encoding UTF8 }
}

function Get-WemCpuSpikeProtectionColumns {
    @('RunId','Server','TimeCreated','EventId','RecordId','LogName','ProviderName','Level','ProcessName','ProcessId','UserName','SessionId','CpuPercent','OldPriority','NewPriority','Action','Success','Message','RawEventXml','ParseError','QueryError')
}

function Get-WemCpuSpikeProtectionSummaryColumns {
    @('RunId','Server','ProcessName','UserName','TriggerCount','PriorityChangeCount','FailedPriorityChangeCount','FirstTriggerTime','LastTriggerTime','MaximumProcessCpuPercent','AverageProcessCpuPercent','MaximumServerCpuAtTrigger','AverageServerCpuAtTrigger','MostFrequentNewPriority','SamplesDuringProtection','SamplesOutsideProtection')
}

function Get-WemCpuSpikeProtectionQueryDiagnosticsColumns {
    @('RunId','Server','CandidateLogName','LogExists','Selected','QuerySucceeded','EventsReturned','ErrorType','ErrorMessage')
}

function Invoke-WemCpuSpikeProtectionEventCollection {
    param([string[]]$Servers, [string]$RunId, [datetime]$StartTime, [datetime]$EndTime)
    $rows = @()
    $diagnostics = @()
    foreach ($server in $Servers) {
        try {
            $result = Invoke-Command -ComputerName $server -ScriptBlock {
                param($RunId, $StartTime, $EndTime)
                $eventRows = @()
                $diagRows = @()
                $eventIds = @(7001,7002,7003,7004)
                $knownCandidates = @('WEM Agent Service','Citrix WEM Agent Service','Norskale Agent Service')
                $availableLogs = @()
                try { $availableLogs = @(Get-WinEvent -ListLog '*WEM*' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty LogName) } catch { $availableLogs = @() }
                foreach ($known in $knownCandidates) {
                    if (@($availableLogs | Where-Object { $_ -eq $known }).Count -gt 0) { continue }
                    try {
                        $knownLog = Get-WinEvent -ListLog $known -ErrorAction Stop
                        if ($knownLog -and $knownLog.LogName) { $availableLogs += $knownLog.LogName }
                    } catch { }
                }
                $candidateLogs = @($availableLogs + $knownCandidates | Where-Object { $_ } | Select-Object -Unique)
                foreach ($candidate in $candidateLogs) {
                    $exists = @($availableLogs | Where-Object { $_ -eq $candidate }).Count -gt 0
                    $selected = $exists
                    if (-not $exists) {
                        $diagRows += [pscustomobject]@{ RunId=$RunId; Server=$env:COMPUTERNAME; CandidateLogName=$candidate; LogExists=$false; Selected=$false; QuerySucceeded=$false; EventsReturned=0; ErrorType=''; ErrorMessage='' }
                        continue
                    }
                    try {
                        $events = @(Get-WinEvent -FilterHashtable @{ LogName=$candidate; Id=$eventIds; StartTime=$StartTime; EndTime=$EndTime } -ErrorAction Stop)
                        $diagRows += [pscustomobject]@{ RunId=$RunId; Server=$env:COMPUTERNAME; CandidateLogName=$candidate; LogExists=$true; Selected=$selected; QuerySucceeded=$true; EventsReturned=$events.Count; ErrorType=''; ErrorMessage='' }
                        foreach ($event in $events) {
                            $xmlText = ''
                            $parseError = ''
                            $data = @{}
                            try {
                                $xmlText = $event.ToXml()
                                [xml]$xml = $xmlText
                                foreach ($d in @($xml.Event.EventData.Data)) {
                                    $name = if ($d.Name) { [string]$d.Name } else { '' }
                                    if ($name) { $data[$name] = [string]$d.'#text' }
                                }
                            } catch { $parseError = $_.Exception.Message }
                            function Find-DataValue($Map, [string[]]$Names) {
                                foreach ($n in $Names) { if ($Map.ContainsKey($n) -and -not [string]::IsNullOrWhiteSpace([string]$Map[$n])) { return [string]$Map[$n] } }
                                return ''
                            }
                            $action = switch ([int]$event.Id) { 7001 { 'SpikeProtectionInitialized' } 7002 { 'SpikeProtectionInitialized' } 7003 { 'PriorityChanged' } 7004 { 'PriorityChangeFailed' } default { 'Other' } }
                            $eventRows += [pscustomobject]@{
                                RunId=$RunId; Server=$env:COMPUTERNAME; TimeCreated=$event.TimeCreated.ToString('s'); EventId=$event.Id; RecordId=$event.RecordId; LogName=$event.LogName; ProviderName=$event.ProviderName; Level=$event.LevelDisplayName;
                                ProcessName=(Find-DataValue $data @('ProcessName','Process','ImageName','ApplicationName','Name')); ProcessId=(Find-DataValue $data @('ProcessId','PID','Pid')); UserName=(Find-DataValue $data @('UserName','User','AccountName')); SessionId=(Find-DataValue $data @('SessionId','Session'));
                                CpuPercent=(Find-DataValue $data @('CpuPercent','CPU','CpuUsage','PercentCpu')); OldPriority=(Find-DataValue $data @('OldPriority','OriginalPriority','PreviousPriority')); NewPriority=(Find-DataValue $data @('NewPriority','Priority','TargetPriority'));
                                Action=$action; Success=([int]$event.Id -ne 7004); Message=$event.Message; RawEventXml=$xmlText; ParseError=$parseError; QueryError=''
                            }
                        }
                    } catch {
                        $errorId = [string]$_.FullyQualifiedErrorId
                        $errorMessage = [string]$_.Exception.Message
                        if ($errorId -like 'NoMatchingEventsFound,*' -or $errorMessage -like 'No events were found that match the specified selection criteria*') {
                            $diagRows += [pscustomobject]@{ RunId=$RunId; Server=$env:COMPUTERNAME; CandidateLogName=$candidate; LogExists=$true; Selected=$selected; QuerySucceeded=$true; EventsReturned=0; ErrorType=''; ErrorMessage='' }
                        }
                        else {
                            $diagRows += [pscustomobject]@{ RunId=$RunId; Server=$env:COMPUTERNAME; CandidateLogName=$candidate; LogExists=$true; Selected=$selected; QuerySucceeded=$false; EventsReturned=0; ErrorType=$_.Exception.GetType().FullName; ErrorMessage=$errorMessage }
                        }
                    }
                }
                [pscustomobject]@{ Events=@($eventRows); Diagnostics=@($diagRows) }
            } -ArgumentList $RunId, $StartTime, $EndTime -ErrorAction Stop
            $rows += @($result.Events)
            $diagnostics += @($result.Diagnostics)
        } catch {
            $diagnostics += [pscustomobject]@{ RunId=$RunId; Server=$server; CandidateLogName=''; LogExists=$false; Selected=$false; QuerySucceeded=$false; EventsReturned=0; ErrorType=$_.Exception.GetType().FullName; ErrorMessage=$_.Exception.Message }
        }
    }
    $script:LastWemCpuSpikeProtectionQueryDiagnostics = @($diagnostics)
    @($rows | Group-Object { '{0}|{1}|{2}|{3}' -f $_.Server,$_.LogName,$_.ProviderName,$_.RecordId } | ForEach-Object { $_.Group | Select-Object -First 1 })
}

function Add-WemCorrelationToProcessRows {
    param([array]$ProcessRows, [array]$WemRows, [int]$PriorityLoweringSeconds)
    foreach ($p in $ProcessRows) {
        $sampleTime = [datetime]$p.Timestamp
        $sampleProcessId = [string]$p.ProcessId
        $matches = @($WemRows | Where-Object { $_.Server -eq $p.Server -and [string]$_.ProcessId -eq $sampleProcessId -and $_.TimeCreated } | Sort-Object TimeCreated)
        $recent = @($matches | Where-Object { ([datetime]$_.TimeCreated) -le $sampleTime -and ($sampleTime - [datetime]$_.TimeCreated).TotalSeconds -le $PriorityLoweringSeconds } | Select-Object -Last 1)
        if ($recent) {
            $p.WemSpikeTriggered = (@($matches | Where-Object { $_.EventId -in @(7001,7002) }).Count -gt 0)
            $p.WemPriorityChanged = (@($matches | Where-Object { $_.EventId -eq 7003 }).Count -gt 0)
            $p.WemNewPriority = $recent.NewPriority
            $p.WemTriggerTime = $recent.TimeCreated
            $p.SecondsSinceWemTrigger = [Math]::Round(($sampleTime - [datetime]$recent.TimeCreated).TotalSeconds, 1)
            $p.WemProtectionLikelyActive = $true
        }
    }
}

function New-WemCpuSpikeProtectionSummaryRows {
    param([array]$WemRows, [array]$ProcessRows, [string]$RunId, [int]$PriorityLoweringSeconds)
    $eventRows = @($WemRows | Where-Object { $_.EventId })
    foreach ($group in ($eventRows | Group-Object Server,ProcessName,UserName)) {
        $first = @($group.Group | Sort-Object TimeCreated | Select-Object -First 1)[0]
        $last = @($group.Group | Sort-Object TimeCreated | Select-Object -Last 1)[0]
        $samples = @($ProcessRows | Where-Object { $_.Server -eq $first.Server -and $_.ProcessName -eq $first.ProcessName })
        [pscustomobject]@{
            RunId=$RunId; Server=$first.Server; ProcessName=$first.ProcessName; UserName=$first.UserName;
            TriggerCount=@($group.Group | Where-Object { $_.EventId -in @(7001,7002) }).Count; PriorityChangeCount=@($group.Group | Where-Object { $_.EventId -eq 7003 }).Count; FailedPriorityChangeCount=@($group.Group | Where-Object { $_.EventId -eq 7004 }).Count;
            FirstTriggerTime=$first.TimeCreated; LastTriggerTime=$last.TimeCreated;
            MaximumProcessCpuPercent=if($samples){($samples | Measure-Object ProcessCpuServerPercent -Maximum).Maximum}else{''}; AverageProcessCpuPercent=if($samples){[Math]::Round(($samples | Measure-Object ProcessCpuServerPercent -Average).Average,2)}else{''};
            MaximumServerCpuAtTrigger=if($samples){($samples | Measure-Object CpuPercent -Maximum).Maximum}else{''}; AverageServerCpuAtTrigger=if($samples){[Math]::Round(($samples | Measure-Object CpuPercent -Average).Average,2)}else{''};
            MostFrequentNewPriority=(@($group.Group | Where-Object NewPriority | Group-Object NewPriority | Sort-Object Count -Descending | Select-Object -First 1).Name);
            SamplesDuringProtection=@($samples | Where-Object WemProtectionLikelyActive).Count; SamplesOutsideProtection=@($samples | Where-Object { -not $_.WemProtectionLikelyActive }).Count
        }
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
Ensure-SettingProperty -Settings $settings -Name 'DefenderPerfFinalWaitSeconds' -DefaultValue 120
Ensure-SettingProperty -Settings $settings -Name 'FinalizationTimeoutSeconds' -DefaultValue 300
Ensure-SettingProperty -Settings $settings -Name 'DefenderPerfLocalRoot' -DefaultValue 'C:\ProgramData\CitrixTSHealthCheck\DefenderPerf'
Ensure-SettingProperty -Settings $settings -Name 'DefenderPerfCopyToOutputPath' -DefaultValue $true
Ensure-SettingProperty -Settings $settings -Name 'IncludeWemEventContext' -DefaultValue $false
Ensure-SettingProperty -Settings $settings -Name 'IncludeWemCpuSpikeProtectionEvents' -DefaultValue $false
Ensure-SettingProperty -Settings $settings -Name 'WemTriggerServerCpuPercent' -DefaultValue 10
Ensure-SettingProperty -Settings $settings -Name 'WemEventWindowMinutes' -DefaultValue 10
Ensure-SettingProperty -Settings $settings -Name 'IncludeWemLogTail' -DefaultValue $false
Ensure-SettingProperty -Settings $settings -Name 'WemLogTailLines' -DefaultValue 200
Ensure-SettingProperty -Settings $settings -Name 'WemPriorityLoweringSeconds' -DefaultValue 180
Ensure-SettingProperty -Settings $settings -Name 'WemEventContextCooldownMinutes' -DefaultValue 10
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
$wemCpuSpikeProtectionOutputPath = Join-Path $rawPath "WemCpuSpikeProtectionEvents_$runId.csv"
$wemCpuSpikeProtectionSummaryOutputPath = Join-Path $rawPath "WemCpuSpikeProtectionSummary_$runId.csv"
$wemCpuSpikeProtectionQueryDiagnosticsOutputPath = Join-Path $rawPath "WemCpuSpikeProtectionQueryDiagnostics_$runId.csv"
$eventContextOutputPath = Join-Path $rawPath "EventContext_$runId.csv"
$allServerRows = @()
$allProcessRows = @()
$defenderJobs = @()
$wemJobs = @()
$defenderLastTriggerByServer = @{}
$defenderTriggeredServers = @{}
$defenderTriggered = $false
$wemTriggered = $false
$wemEventContextLastTriggerByServer = @{}
$round = 0
$completedRounds = 0
$measurementStartTime = $runStart
$plannedDurationMinutes = [int]$settings.DurationMinutes
$intervalSeconds = [Math]::Max(1, [int]$settings.IntervalSeconds)
$hardEndTime = $runStart.AddMinutes($plannedDurationMinutes)
$maxRounds = if ($plannedDurationMinutes -lt 1) { 1 } else { ([int][Math]::Floor(($plannedDurationMinutes * 60.0) / $intervalSeconds)) + 1 }
$endReason = 'Completed'
$runError = $null

try {
    Write-RunLog -Path $runLogFile -Message "Run gestartet. Server=$($servers.Count), DurationMinutes=$($settings.DurationMinutes), IntervalSeconds=$($settings.IntervalSeconds), MaxParallel=$($settings.MaxParallel), StartTime=$($runStart.ToString('s')), HardEndTime=$($hardEndTime.ToString('s')), MaxRounds=$maxRounds, DefenderPerfLocalRoot=$($settings.DefenderPerfLocalRoot), DefenderPerfTriggerServerCpuPercent=$($settings.DefenderPerfTriggerServerCpuPercent), DefenderPerfRecordingSeconds=$($settings.DefenderPerfRecordingSeconds), DefenderPerfCooldownMinutes=$($settings.DefenderPerfCooldownMinutes), MaxConcurrentDefenderPerfRecordings=$($settings.MaxConcurrentDefenderPerfRecordings), DefenderPerfFinalWaitSeconds=$($settings.DefenderPerfFinalWaitSeconds), FinalizationTimeoutSeconds=$($settings.FinalizationTimeoutSeconds), WemTriggerServerCpuPercent=$($settings.WemTriggerServerCpuPercent), WemEventWindowMinutes=$($settings.WemEventWindowMinutes), IncludeEventlogContext=$($settings.IncludeEventLogContext), IncludeWemCpuSpikeProtectionEvents=$($settings.IncludeWemCpuSpikeProtectionEvents), WemPriorityLoweringSeconds=$($settings.WemPriorityLoweringSeconds)"
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
    $measurementStartTime = Get-Date
    $hardEndTime = $measurementStartTime.AddMinutes($plannedDurationMinutes)
    Write-RunLog -Path $runLogFile -Message "MeasurementStartTime=$($measurementStartTime.ToString('s')), HardEndTime=$($hardEndTime.ToString('s')), MaxRounds=$maxRounds"
    Write-EffectiveConfiguration -Settings $settings -RunLogFile $runLogFile -HardEndTime $hardEndTime -MaxRounds $maxRounds
    if ($settings.IncludeEventLogContext) { Export-EmptyCsv -Path $eventContextOutputPath -Columns @('RunId','Timestamp','Server','TargetServer','ComputerName','LogName','EventTime','EventId','ProviderName','Level','Message') -Delimiter $settings.OutputDelimiter; Export-EmptyCsv -Path (Join-Path $runOutputPath 'EventContext.csv') -Columns @('RunId','Timestamp','Server','TargetServer','ComputerName','LogName','EventTime','EventId','ProviderName','Level','Message') -Delimiter $settings.OutputDelimiter }
    if ($settings.IncludeWemCpuSpikeProtectionEvents) { Export-EmptyCsv -Path $wemCpuSpikeProtectionOutputPath -Columns (Get-WemCpuSpikeProtectionColumns) -Delimiter $settings.OutputDelimiter; Export-EmptyCsv -Path (Join-Path $runOutputPath 'WemCpuSpikeProtectionEvents.csv') -Columns (Get-WemCpuSpikeProtectionColumns) -Delimiter $settings.OutputDelimiter; Export-EmptyCsv -Path $wemCpuSpikeProtectionSummaryOutputPath -Columns (Get-WemCpuSpikeProtectionSummaryColumns) -Delimiter $settings.OutputDelimiter; Export-EmptyCsv -Path (Join-Path $runOutputPath 'WemCpuSpikeProtectionSummary.csv') -Columns (Get-WemCpuSpikeProtectionSummaryColumns) -Delimiter $settings.OutputDelimiter; Export-EmptyCsv -Path $wemCpuSpikeProtectionQueryDiagnosticsOutputPath -Columns (Get-WemCpuSpikeProtectionQueryDiagnosticsColumns) -Delimiter $settings.OutputDelimiter; Export-EmptyCsv -Path (Join-Path $runOutputPath 'WemCpuSpikeProtectionQueryDiagnostics.csv') -Columns (Get-WemCpuSpikeProtectionQueryDiagnosticsColumns) -Delimiter $settings.OutputDelimiter }
    while ($true) {
        $nextRound = $round + 1
        if ($nextRound -gt $maxRounds) { $endReason = 'MaxRoundsReached'; Write-RunLog -Path $runLogFile -Message "Keine neue Runde: MaxRounds erreicht ($maxRounds)."; break }
        $plannedRoundStart = $measurementStartTime.AddSeconds(($nextRound - 1) * $intervalSeconds)
        if ($plannedRoundStart -gt $hardEndTime) { $endReason = 'HardEndTimeReached'; Write-RunLog -Path $runLogFile -Message "Keine neue Runde: geplanter Start $($plannedRoundStart.ToString('s')) liegt nach HardEndTime $($hardEndTime.ToString('s'))."; break }
        $now = Get-Date
        if ($now -lt $plannedRoundStart) {
            $sleepSeconds = [int][Math]::Ceiling(($plannedRoundStart - $now).TotalSeconds)
            if ($sleepSeconds -gt 0) { Start-Sleep -Seconds ([Math]::Min($sleepSeconds, [Math]::Max(0, [int][Math]::Floor(($hardEndTime - $now).TotalSeconds)))) }
        }
        $round = $nextRound
        $roundStart = Get-Date
        $message = "Runde $round gestartet. Geplant=$($plannedRoundStart.ToString('s')), Server=$($servers.Count), CpuSampleSeconds=$($settings.CpuSampleSeconds), Parallel=$($settings.MaxParallel)"
        Write-Host ("[{0}] {1}" -f $roundStart.ToString('HH:mm:ss'), $message)
        Write-RunLog -Path $runLogFile -Message $message
        $roundResults = @(Invoke-ServerCollectionRound -Servers $servers -Settings $settings -RoundTimestamp $roundStart)
        $day = (Get-Date).ToString('yyyy-MM-dd')
        $serverRows = @(); $processRows = @(); $alertRows = @(); $categoryRows = @(); $eventRows = @()

        foreach ($item in $roundResults) {
            if ($item.Status -ne 'OK') {
                $errorRow = New-ErrorServerSample -Server $item.TargetServer -Message $item.ErrorMessage -Timestamp $roundStart -RunId $runId
                $serverRows += $errorRow
                Write-RunLog -Path $runLogFile -Level 'ERROR' -Message ("Server={0}; Round={1}; ProcessName={2}; PID={3}; ObjectType={4}; Properties={5}; ExceptionType={6}; Exception={7}; ScriptStackTrace={8}" -f $item.TargetServer, $round, (Get-SafePropertyValue -Object $item -PropertyName 'ProcessName' -DefaultValue ''), (Get-SafePropertyValue -Object $item -PropertyName 'ProcessId' -DefaultValue ''), (Get-SafePropertyValue -Object $item -PropertyName 'ObjectType' -DefaultValue ''), (Get-SafePropertyValue -Object $item -PropertyName 'PropertyNames' -DefaultValue ''), (Get-SafePropertyValue -Object $item -PropertyName 'ErrorType' -DefaultValue ''), $item.ErrorMessage, (Get-SafePropertyValue -Object $item -PropertyName 'ScriptStackTrace' -DefaultValue ''))
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
                $remainingForDefenderSeconds = ($hardEndTime - (Get-Date)).TotalSeconds
                $requiredDefenderSeconds = [int]$settings.DefenderPerfRecordingSeconds + 60
                if ((Get-Date) -ge $hardEndTime -or $remainingForDefenderSeconds -lt $requiredDefenderSeconds) {
                    Write-RunLog -Path $runLogFile -Level 'WARN' -Message "Defender Recording skipped: insufficient remaining time. Server=$($item.TargetServer), RemainingSeconds=$([Math]::Round($remainingForDefenderSeconds,1)), RequiredSeconds=$requiredDefenderSeconds"
                }
                else {
                    $runningDefenderJobs = @($defenderJobs | Where-Object { $_.State -eq 'Running' }).Count
                    $cooldownOk = (-not $defenderLastTriggerByServer.ContainsKey($item.TargetServer)) -or (($sampleTime - $defenderLastTriggerByServer[$item.TargetServer]).TotalMinutes -ge [int]$settings.DefenderPerfCooldownMinutes)
                    if (-not $defenderTriggeredServers.ContainsKey($item.TargetServer) -and $cooldownOk -and $runningDefenderJobs -lt [int]$settings.MaxConcurrentDefenderPerfRecordings) {
                        $defenderLastTriggerByServer[$item.TargetServer] = $sampleTime
                        $defenderTriggeredServers[$item.TargetServer] = $true
                        $defenderTriggered = $true
                        $recordingStartTime = Get-Date
                        $defenderJobs += Start-DefenderPerfRecordingJob -Server $item.TargetServer -RunId $runId -OutputPath $OutputPath -LocalRoot $settings.DefenderPerfLocalRoot -CopyToOutputPath ([bool]$settings.DefenderPerfCopyToOutputPath) -Seconds ([int]$settings.DefenderPerfRecordingSeconds) -TriggerProcess 'MsMpEng' -TriggerCpu ([double]$msMpEng.ProcessCpuServerPercent) -TriggerThreshold ([double]$settings.DefenderPerfTriggerServerCpuPercent) -TriggerTime $sampleTime
                        $pendingRow = New-DefenderPerfRecordingRow -Values @{ RunId=$runId; Server=$item.TargetServer; TriggerTimestamp=$sampleTime.ToString('s'); TriggerProcessName='MsMpEng'; TriggerProcessCpuServerPercent=[double]$msMpEng.ProcessCpuServerPercent; TriggerThreshold=[double]$settings.DefenderPerfTriggerServerCpuPercent; TriggerReason='MsMpEng ProcessCpuServerPercent >= DefenderPerfTriggerServerCpuPercent'; Status='Pending'; StartTime=$recordingStartTime.ToString('s'); TriggerToRecordingStartSeconds=[Math]::Round(($recordingStartTime - $sampleTime).TotalSeconds, 1); ReportGenerationStatus='Pending'; CopyStatus='Pending'; ErrorMessage='' }
                        Export-DefenderPerfRecordingRows -Rows @($pendingRow) -RawPath $rawPath -RunOutputPath $runOutputPath -RunId $runId -Delimiter $settings.OutputDelimiter -RunLogFile $runLogFile
                        Write-RunLog -Path $runLogFile -Message "Defender Performance Recording getriggert: $($item.TargetServer), MsMpEng=$($msMpEng.ProcessCpuServerPercent)%"
                    }
                }
            }
            $wemCpu = [double]$cat.WemCpuServerPercent
            $wemProc = @($item.Result.ProcessSamples | Where-Object { $_.ProcessName -in @('Citrix.Wem.Agent.Service','VUEMUIAgent') } | Sort-Object ProcessCpuServerPercent -Descending | Select-Object -First 1)
            if ($settings.IncludeWemEventContext -and (($wemCpu -ge [double]$settings.WemTriggerServerCpuPercent) -or ($wemProc -and [double]$wemProc.ProcessCpuServerPercent -ge [double]$settings.WemTriggerServerCpuPercent))) {
                $cooldownOk = (-not $wemEventContextLastTriggerByServer.ContainsKey($item.TargetServer)) -or (($sampleTime - $wemEventContextLastTriggerByServer[$item.TargetServer]).TotalMinutes -ge [int]$settings.WemEventContextCooldownMinutes)
                if ($cooldownOk) {
                    $wemEventContextLastTriggerByServer[$item.TargetServer] = $sampleTime
                    $wemTriggered = $true
                    $wemAlertId = '{0}_{1}_{2}_WEM' -f $runId, ($item.TargetServer -replace '[^A-Za-z0-9_.-]', '_'), ($serverSample.Timestamp -replace '[:]', '')
                    $wemTriggerName = 'Category:Wem'
                    $wemTriggerCpu = $wemCpu
                    if ($wemProc -and [double]$wemProc.ProcessCpuServerPercent -ge [double]$settings.WemTriggerServerCpuPercent) { $wemTriggerName = $wemProc.ProcessName; $wemTriggerCpu = [double]$wemProc.ProcessCpuServerPercent }
                    $wemJobs += Start-WemEventContextJob -Server $item.TargetServer -RunId $runId -AlertId $wemAlertId -TriggerTime $sampleTime -WindowMinutes ([int]$settings.WemEventWindowMinutes) -IncludeLogTail ([bool]$settings.IncludeWemLogTail) -TailLines ([int]$settings.WemLogTailLines) -MaxEventsPerAlert ([int]$settings.MaxEventsPerAlert) -TriggerProcessName $wemTriggerName -TriggerProcessCpuServerPercent $wemTriggerCpu -TriggerThreshold ([double]$settings.WemTriggerServerCpuPercent)
                    Write-RunLog -Path $runLogFile -Message "WEM Event Context getriggert: $($item.TargetServer), Trigger=$wemTriggerName, CPU=$wemTriggerCpu%"
                } else { Write-RunLog -Path $runLogFile -Message "WEM Event Context skipped wegen Cooldown: $($item.TargetServer)" }
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
                    ProcessCpuCorePercent=$p.ProcessCpuCorePercent; ProcessCpuServerPercent=$p.ProcessCpuServerPercent; ProcessCpuSecondsDelta=$p.ProcessCpuSecondsDelta; ProcessWorkingSetMB=$p.ProcessWorkingSetMB; ProcessPrivateMemoryMB=$p.ProcessPrivateMemoryMB; ProcessPath=$p.ProcessPath; ProcessCommandLine=$p.ProcessCommandLine; ProcessStartTime=$p.ProcessStartTime; PriorityClass=(Get-SafePropertyValue -Object $p -PropertyName 'PriorityClass' -DefaultValue $null); BasePriority=(Get-SafePropertyValue -Object $p -PropertyName 'BasePriority' -DefaultValue $null); SessionId=$p.ProcessSessionId;
                    WemSpikeTriggered=$false; WemPriorityChanged=$false; WemNewPriority=''; WemTriggerTime=''; SecondsSinceWemTrigger=''; WemProtectionLikelyActive=$false;
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
        Receive-DetailDiagnosticJobs -DefenderJobs $defenderJobs -WemJobs $wemJobs -RawPath $rawPath -RunOutputPath $runOutputPath -RunId $runId -Delimiter $settings.OutputDelimiter -RunLogFile $runLogFile
        $defenderJobs = @($defenderJobs | Where-Object { $_.State -eq 'Running' })
        $wemJobs = @($wemJobs | Where-Object { $_.State -eq 'Running' })
        $allServerRows += $serverRows; $allProcessRows += $processRows
        $completedRounds = $round
        $doneMessage = "Runde $round beendet. OK=$(@($serverRows | Where-Object Status -eq 'OK').Count), Fehler=$(@($serverRows | Where-Object Status -eq 'ERROR').Count)"
        Write-Host ("[{0}] {1}" -f (Get-Date).ToString('HH:mm:ss'), $doneMessage)
        Write-RunLog -Path $runLogFile -Message $doneMessage
        if ($round -ge $maxRounds) { $endReason = 'MaxRoundsReached'; Write-RunLog -Path $runLogFile -Message "MaxRounds erreicht: $completedRounds/$maxRounds."; break }
        if ((Get-Date) -ge $hardEndTime) { $endReason = 'HardEndTimeReached'; Write-RunLog -Path $runLogFile -Message "HardEndTime erreicht: $($hardEndTime.ToString('s'))."; break }
    }
}
catch {
    $runError = $_.Exception.Message
    if ($endReason -eq 'Completed') { $endReason = 'Error' }
    Write-RunLog -Path $runLogFile -Level 'ERROR' -Message "Runfehler: $runError"
    Write-Warning $runError
}
finally {
    $runEnd = $null
    try {
        $categoryAggregateRows = @(New-CategoryAggregateRows -ProcessRows $allProcessRows -RunId $runId)
        Export-Rows -Rows $categoryAggregateRows -Path (Join-Path $rawPath "CategorySummaryAggregated_$runId.csv") -Delimiter $settings.OutputDelimiter
        Export-Rows -Rows $categoryAggregateRows -Path (Join-Path $runOutputPath 'CategorySummaryAggregated.csv') -Delimiter $settings.OutputDelimiter
        Complete-DetailDiagnosticJobs -DefenderJobs $defenderJobs -WemJobs $wemJobs -RawPath $rawPath -RunOutputPath $runOutputPath -RunId $runId -Delimiter $settings.OutputDelimiter -DefenderFinalWaitSeconds ([int]$settings.DefenderPerfFinalWaitSeconds) -FinalizationTimeoutSeconds ([int]$settings.FinalizationTimeoutSeconds) -RunLogFile $runLogFile
        $runEnd = Get-Date
        $actualDurationMinutes = [Math]::Round(($runEnd - $runStart).TotalMinutes, 2)
        Write-RunLog -Path $runLogFile -Message "Run beendet. EndReason=$endReason, CompletedRounds=$completedRounds, ActualDurationMinutes=$actualDurationMinutes"
        $wemSpikeRows = @(); $wemSpikeSummaryRows = @(); $wemSpikeQueryDiagnosticsRows = @()
        if ($settings.IncludeWemCpuSpikeProtectionEvents) {
            try {
                Write-RunLog -Path $runLogFile -Message 'WEM CPU Spike Protection Eventsammlung gestartet.'
                $wemSpikeRows = @(Invoke-WemCpuSpikeProtectionEventCollection -Servers $servers -RunId $runId -StartTime $measurementStartTime -EndTime $runEnd)
                $wemSpikeQueryDiagnosticsRows = @($script:LastWemCpuSpikeProtectionQueryDiagnostics)
                if ($wemSpikeQueryDiagnosticsRows.Count -eq 0) {
                    foreach ($server in $servers) {
                        $wemSpikeQueryDiagnosticsRows += [pscustomobject]@{ RunId=$runId; Server=$server; CandidateLogName=''; LogExists=$false; Selected=$false; QuerySucceeded=$false; EventsReturned=0; ErrorType='NoDiagnosticsReturned'; ErrorMessage='Keine WEM Query Diagnostics vom Server erhalten.' }
                    }
                }
                Add-WemCorrelationToProcessRows -ProcessRows $allProcessRows -WemRows $wemSpikeRows -PriorityLoweringSeconds ([int]$settings.WemPriorityLoweringSeconds)
                $wemSpikeSummaryRows = @(New-WemCpuSpikeProtectionSummaryRows -WemRows $wemSpikeRows -ProcessRows $allProcessRows -RunId $runId -PriorityLoweringSeconds ([int]$settings.WemPriorityLoweringSeconds))
            }
            catch {
                Write-RunLog -Path $runLogFile -Level 'ERROR' -Message ("WEM CPU Spike Protection Finalisierung fehlgeschlagen: Type={0}; Message={1}; Stack={2}" -f $_.Exception.GetType().FullName, $_.Exception.Message, $_.ScriptStackTrace)
                if ($wemSpikeQueryDiagnosticsRows.Count -eq 0) {
                    foreach ($server in $servers) {
                        $wemSpikeQueryDiagnosticsRows += [pscustomobject]@{ RunId=$runId; Server=$server; CandidateLogName=''; LogExists=$false; Selected=$false; QuerySucceeded=$false; EventsReturned=0; ErrorType=$_.Exception.GetType().FullName; ErrorMessage=$_.Exception.Message }
                    }
                }
            }
            $eventRowsOnly = @($wemSpikeRows | Where-Object { $_.EventId })
            if ($eventRowsOnly.Count -gt 0) { $eventRowsOnly | ConvertTo-InvariantObject | Export-Csv -LiteralPath $wemCpuSpikeProtectionOutputPath -Delimiter $settings.OutputDelimiter -NoTypeInformation -Encoding UTF8; $eventRowsOnly | ConvertTo-InvariantObject | Export-Csv -LiteralPath (Join-Path $runOutputPath 'WemCpuSpikeProtectionEvents.csv') -Delimiter $settings.OutputDelimiter -NoTypeInformation -Encoding UTF8 }
            if ($wemSpikeSummaryRows.Count -gt 0) { $wemSpikeSummaryRows | ConvertTo-InvariantObject | Export-Csv -LiteralPath $wemCpuSpikeProtectionSummaryOutputPath -Delimiter $settings.OutputDelimiter -NoTypeInformation -Encoding UTF8; $wemSpikeSummaryRows | ConvertTo-InvariantObject | Export-Csv -LiteralPath (Join-Path $runOutputPath 'WemCpuSpikeProtectionSummary.csv') -Delimiter $settings.OutputDelimiter -NoTypeInformation -Encoding UTF8 }
            if ($wemSpikeQueryDiagnosticsRows.Count -gt 0) { $wemSpikeQueryDiagnosticsRows | ConvertTo-InvariantObject | Export-Csv -LiteralPath $wemCpuSpikeProtectionQueryDiagnosticsOutputPath -Delimiter $settings.OutputDelimiter -NoTypeInformation -Encoding UTF8; $wemSpikeQueryDiagnosticsRows | ConvertTo-InvariantObject | Export-Csv -LiteralPath (Join-Path $runOutputPath 'WemCpuSpikeProtectionQueryDiagnostics.csv') -Delimiter $settings.OutputDelimiter -NoTypeInformation -Encoding UTF8 }
            if ($allProcessRows.Count -gt 0) {
                $allProcessRows | ConvertTo-InvariantObject | Export-Csv -LiteralPath (Join-Path $rawPath "Raw_ProcessSamples_$runId.csv") -Delimiter $settings.OutputDelimiter -NoTypeInformation -Encoding UTF8
                $allProcessRows | ConvertTo-InvariantObject | Export-Csv -LiteralPath (Join-Path $runOutputPath 'Raw_ProcessSamples.csv') -Delimiter $settings.OutputDelimiter -NoTypeInformation -Encoding UTF8
            }
            Write-RunLog -Path $runLogFile -Message ("WEM CPU Spike Protection Events beendet. Rows={0}, QueryErrors={1}, ParseErrors={2}, DiagnosticsRows={3}" -f $eventRowsOnly.Count, @($wemSpikeQueryDiagnosticsRows | Where-Object { $_.Selected -and -not $_.QuerySucceeded }).Count, @($wemSpikeRows | Where-Object ParseError).Count, $wemSpikeQueryDiagnosticsRows.Count)
        }
        $eventContextRowsForLog = @(Import-CsvIfExists -Path $eventContextOutputPath -Delimiter $settings.OutputDelimiter)
        $eventContextTriggerCount = @($allServerRows | Where-Object { $_.IsCpuWarning -eq $true -or $_.IsCpuWarning -eq 'True' }).Count
        $eventContextRowCount = @($eventContextRowsForLog | Where-Object { $_.EventId }).Count
        Write-RunLog -Path $runLogFile -Message ("EventContext: IncludeEventlogContext={0}, EventContextTriggerCount={1}, EventContextRows={2}, EventContextQueryErrors=0, EventContextOutputPath={3}" -f $settings.IncludeEventLogContext, $eventContextTriggerCount, $eventContextRowCount, $eventContextOutputPath)
        $wemSpikeEventRows = @($wemSpikeRows | Where-Object { $_.EventId })
        $wemSpikeEventTotal = $wemSpikeEventRows.Count
        $wemSpike7001Count = @($wemSpikeEventRows | Where-Object { [int]$_.EventId -eq 7001 }).Count
        $wemSpike7002Count = @($wemSpikeEventRows | Where-Object { [int]$_.EventId -eq 7002 }).Count
        $wemSpike7003Count = @($wemSpikeEventRows | Where-Object { [int]$_.EventId -eq 7003 }).Count
        $wemSpike7004Count = @($wemSpikeEventRows | Where-Object { [int]$_.EventId -eq 7004 }).Count
        $wemSpikeParseErrors = @($wemSpikeRows | Where-Object { $_.ParseError }).Count
        $wemSpikeQueryErrors = @($wemSpikeQueryDiagnosticsRows | Where-Object { $_.Selected -and -not $_.QuerySucceeded }).Count
        $wemSpikeServersWithEvents = @($wemSpikeEventRows | Select-Object -ExpandProperty Server -Unique).Count
        if ($settings.IncludeWemCpuSpikeProtectionEvents) { Write-RunLog -Path $runLogFile -Message ("WemCpuSpikeProtectionEventsTotal={0}, WemCpuSpikeProtectionEvent7001Count={1}, WemCpuSpikeProtectionEvent7002Count={2}, WemCpuSpikeProtectionEvent7003Count={3}, WemCpuSpikeProtectionEvent7004Count={4}, WemCpuSpikeProtectionParseErrors={5}, WemCpuSpikeProtectionQueryErrors={6}, WemCpuSpikeProtectionServersWithEvents={7}, WemCpuSpikeProtectionOutputPath={8}, WemCpuSpikeProtectionSummaryOutputPath={9}, WemCpuSpikeProtectionQueryDiagnosticsOutputPath={10}" -f $wemSpikeEventTotal, $wemSpike7001Count, $wemSpike7002Count, $wemSpike7003Count, $wemSpike7004Count, $wemSpikeParseErrors, $wemSpikeQueryErrors, $wemSpikeServersWithEvents, $wemCpuSpikeProtectionOutputPath, $wemCpuSpikeProtectionSummaryOutputPath, $wemCpuSpikeProtectionQueryDiagnosticsOutputPath) }
        $defenderRecordingRows = @(Import-CsvIfExists -Path (Join-Path $rawPath "DefenderPerfRecordings_$runId.csv") -Delimiter $settings.OutputDelimiter)
        $wemEventRows = @(Import-CsvIfExists -Path (Join-Path $rawPath "WemEventContext_$runId.csv") -Delimiter $settings.OutputDelimiter)
        $defenderStartedCount = @($defenderRecordingRows | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.TriggerTimestamp) -and -not [string]::IsNullOrWhiteSpace([string]$_.TriggerProcessName) -and $_.Status -in @('Pending','Running','Completed','Failed','TimedOut','ReportFailed','CopyFailed') }).Count
        $defenderReportOkCount = @($defenderRecordingRows | Where-Object { $_.ReportGenerationStatus -eq 'Completed' }).Count
        $defenderCopyOkCount = @($defenderRecordingRows | Where-Object { $_.CopyStatus -eq 'Completed' }).Count
        $defenderFailedCount = @($defenderRecordingRows | Where-Object { $_.Status -in @('Failed','TimedOut','ReportFailed','CopyFailed') }).Count
        Write-RunLog -Path $runLogFile -Message "Detaildiagnosen: DefenderRecordingsStarted=$defenderStartedCount, DefenderReportsOk=$defenderReportOkCount, DefenderCopiesOk=$defenderCopyOkCount, DefenderFailedOrTimedOut=$defenderFailedCount, WemEventContextRows=$($wemEventRows.Count)"
        $runSummaryRows = @(New-RunSummaryRows -ServerRows $allServerRows -ProcessRows $allProcessRows -Start $runStart -End $runEnd -ServerCount $servers.Count -EndReason $endReason -PlannedDurationMinutes $plannedDurationMinutes -ActualDurationMinutes $actualDurationMinutes -PlannedRounds $maxRounds -CompletedRounds $completedRounds)
        foreach ($summaryRow in $runSummaryRows) {
            $summaryRow | Add-Member -MemberType NoteProperty -Name DefenderPerfRecordingTriggered -Value $defenderTriggered -Force
            $summaryRow | Add-Member -MemberType NoteProperty -Name WemEventContextTriggered -Value $wemTriggered -Force
            $summaryRow | Add-Member -MemberType NoteProperty -Name IncludeWemCpuSpikeProtectionEvents -Value ([bool]$settings.IncludeWemCpuSpikeProtectionEvents) -Force
            $summaryRow | Add-Member -MemberType NoteProperty -Name WemPriorityLoweringSeconds -Value ([int]$settings.WemPriorityLoweringSeconds) -Force
            $summaryRow | Add-Member -MemberType NoteProperty -Name WemCpuSpikeProtectionEventsTotal -Value $wemSpikeEventTotal -Force
            $summaryRow | Add-Member -MemberType NoteProperty -Name WemCpuSpikeProtectionEvent7001Count -Value $wemSpike7001Count -Force
            $summaryRow | Add-Member -MemberType NoteProperty -Name WemCpuSpikeProtectionEvent7002Count -Value $wemSpike7002Count -Force
            $summaryRow | Add-Member -MemberType NoteProperty -Name WemCpuSpikeProtectionEvent7003Count -Value $wemSpike7003Count -Force
            $summaryRow | Add-Member -MemberType NoteProperty -Name WemCpuSpikeProtectionEvent7004Count -Value $wemSpike7004Count -Force
            $summaryRow | Add-Member -MemberType NoteProperty -Name WemCpuSpikeProtectionParseErrors -Value $wemSpikeParseErrors -Force
            $summaryRow | Add-Member -MemberType NoteProperty -Name WemCpuSpikeProtectionQueryErrors -Value $wemSpikeQueryErrors -Force
            $summaryRow | Add-Member -MemberType NoteProperty -Name WemCpuSpikeProtectionServersWithEvents -Value $wemSpikeServersWithEvents -Force
            $summaryRow | Add-Member -MemberType NoteProperty -Name WemCpuSpikeProtectionOutputPath -Value $wemCpuSpikeProtectionOutputPath -Force
            $summaryRow | Add-Member -MemberType NoteProperty -Name WemCpuSpikeProtectionSummaryOutputPath -Value $wemCpuSpikeProtectionSummaryOutputPath -Force
            $summaryRow | Add-Member -MemberType NoteProperty -Name WemCpuSpikeProtectionQueryDiagnosticsOutputPath -Value $wemCpuSpikeProtectionQueryDiagnosticsOutputPath -Force
        }
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
