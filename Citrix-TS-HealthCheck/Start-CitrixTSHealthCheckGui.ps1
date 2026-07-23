#requires -version 3.0
<#
.SYNOPSIS
    Starts a Windows Forms GUI for Citrix-TS-HealthCheck.
.DESCRIPTION
    Provides a small Windows PowerShell 5.1-compatible GUI to edit server/settings
    files, start the health check script and open generated summary, raw and log files.
#>
[CmdletBinding()]
param(
    [string]$ProjectRoot
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:HealthCheckProcess = $null
$script:StdOutFile = $null
$script:StdErrFile = $null
$script:InvocationPath = $MyInvocation.MyCommand.Path
$script:GuiClosing = $false

function Resolve-ProjectRoot {
    param([string]$ConfiguredProjectRoot)

    if (-not [string]::IsNullOrWhiteSpace($ConfiguredProjectRoot)) {
        return (Resolve-Path -LiteralPath $ConfiguredProjectRoot).ProviderPath
    }

    if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        return $PSScriptRoot
    }

    if (-not [string]::IsNullOrWhiteSpace($script:InvocationPath)) {
        return (Split-Path -Parent $script:InvocationPath)
    }

    return (Get-Location).ProviderPath
}

$ProjectRoot = Resolve-ProjectRoot -ConfiguredProjectRoot $ProjectRoot

function Join-ProjectPath {
    param([Parameter(Mandatory=$true)][string[]]$ChildPath)
    $path = $ProjectRoot
    if ([string]::IsNullOrWhiteSpace($path)) { throw 'ProjectRoot ist leer. Starten Sie die GUI aus dem Projektverzeichnis oder uebergeben Sie -ProjectRoot.' }
    foreach ($child in $ChildPath) { $path = Join-Path -Path $path -ChildPath $child }
    return $path
}

function New-Label {
    param([string]$Text, [int]$X, [int]$Y, [int]$Width = 120)
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Location = New-Object System.Drawing.Point($X, $Y)
    $label.Size = New-Object System.Drawing.Size($Width, 20)
    return $label
}

function New-Button {
    param([string]$Text, [int]$X, [int]$Y, [int]$Width = 110, [int]$Height = 28)
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Location = New-Object System.Drawing.Point($X, $Y)
    $button.Size = New-Object System.Drawing.Size($Width, $Height)
    return $button
}

function Read-SettingsFile {
    $settingsPath = Join-ProjectPath -ChildPath @('config','settings.json')
    if (-not (Test-Path -LiteralPath $settingsPath)) {
        return [pscustomobject]@{
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
            WemTriggerServerCpuPercent = 10
            WemEventWindowMinutes = 10
            IncludeWemLogTail = $false
            WemLogTailLines = 200
            IncludeEventLogContext = $false
            IncludeEventContext = $false
            ImageVersion = ''
            Notes = ''
            RunId = ''
            IncludeScheduledTaskInventory = $false
            IncludeCylanceHealth = $false
            TaskNamesToCheck = @('nWizard_{B2FE1952-0186-46C3-BAEC-A80AA35AC5B8}','Adobe Acrobat Update Task','MicrosoftEdgeUpdateTaskMachineUA','Launch Adobe CCXProcess','LexwareAppSysOpt','Office Automatic Updates 2.0','Office Feature Updates','Office Feature Updates Logon','BackgroundDownload')
            AnonymizeUsers = $false
            MaxEventsPerAlert = 50
            OutputDelimiter = ';'
            WinRMTimeoutSeconds = 5
            IncludeDisconnectedSessions = $true
        }
    }
    return Get-Content -LiteralPath $settingsPath -Raw -Encoding UTF8 | ConvertFrom-Json
}


function Get-SettingValue {
    param($Settings, [string]$Name, $DefaultValue)
    if ($null -eq $Settings) { return $DefaultValue }
    if ($Settings.PSObject.Properties.Name -contains $Name) {
        $value = $Settings.$Name
        if ($null -ne $value) { return $value }
    }
    return $DefaultValue
}

function Save-SettingsFile {
    param(
        [int]$CpuSampleSeconds,
        [int]$TopProcessCount,
        [int]$WinRMTimeoutSeconds,
        [int]$DurationMinutes = 480,
        [int]$IntervalSeconds = 300,
        [int]$AlertTopProcessCount = 25,
        [double]$CpuWarningThreshold = 70,
        [double]$CpuCriticalThreshold = 90,
        [int]$MaxParallel = 4,
        [int]$MaxForcedProcessesPerCategory = 10,
        [bool]$AutoDefenderPerfRecording = $false,
        [double]$DefenderPerfTriggerServerCpuPercent = 10,
        [int]$DefenderPerfRecordingSeconds = 900,
        [int]$DefenderPerfCooldownMinutes = 120,
        [int]$MaxConcurrentDefenderPerfRecordings = 2,
        [bool]$IncludeWemEventContext = $false,
        [double]$WemTriggerServerCpuPercent = 10,
        [int]$WemEventWindowMinutes = 10,
        [bool]$IncludeWemLogTail = $false,
        [int]$WemLogTailLines = 200,
        [string]$OutputDelimiter,
        [bool]$IncludeDisconnectedSessions,
        [string]$ImageVersion = '',
        [string]$Notes = '',
        [string]$RunId = '',
        [bool]$IncludeScheduledTaskInventory = $false,
        [bool]$IncludeCylanceHealth = $false,
        [bool]$IncludeEventContext = $false,
        [bool]$AnonymizeUsers = $false,
        [int]$MaxEventsPerAlert = 50,
        [string[]]$TaskNamesToCheck = @()
    )
    $configFolder = Join-ProjectPath -ChildPath @('config')
    if (-not (Test-Path -LiteralPath $configFolder)) { New-Item -ItemType Directory -Path $configFolder -Force | Out-Null }
    $settingsPath = Join-ProjectPath -ChildPath @('config','settings.json')
    $existing = $null
    if (Test-Path -LiteralPath $settingsPath) { try { $existing = Get-Content -LiteralPath $settingsPath -Raw -Encoding UTF8 | ConvertFrom-Json } catch { $existing = $null } }
    $defaultTaskNames = @('nWizard_{B2FE1952-0186-46C3-BAEC-A80AA35AC5B8}','Adobe Acrobat Update Task','MicrosoftEdgeUpdateTaskMachineUA','Launch Adobe CCXProcess','LexwareAppSysOpt','Office Automatic Updates 2.0','Office Feature Updates','Office Feature Updates Logon','BackgroundDownload')
    $existingTaskNames = @(Get-SettingValue -Settings $existing -Name 'TaskNamesToCheck' -DefaultValue $defaultTaskNames)
    $taskNames = if ($TaskNamesToCheck -and $TaskNamesToCheck.Count -gt 0) { @($TaskNamesToCheck) } elseif ($existingTaskNames -and $existingTaskNames.Count -gt 0) { @($existingTaskNames) } else { $defaultTaskNames }
    $defenderPerfFinalWaitSeconds = [int](Get-SettingValue -Settings $existing -Name 'DefenderPerfFinalWaitSeconds' -DefaultValue 120)
    $finalizationTimeoutSeconds = [int](Get-SettingValue -Settings $existing -Name 'FinalizationTimeoutSeconds' -DefaultValue 300)
    $defenderPerfLocalRoot = [string](Get-SettingValue -Settings $existing -Name 'DefenderPerfLocalRoot' -DefaultValue 'C:\ProgramData\CitrixTSHealthCheck\DefenderPerf')
    $defenderPerfCopyToOutputPath = [bool](Get-SettingValue -Settings $existing -Name 'DefenderPerfCopyToOutputPath' -DefaultValue $true)
    $settings = [ordered]@{
        DurationMinutes = $DurationMinutes
        IntervalSeconds = $IntervalSeconds
        CpuSampleSeconds = $CpuSampleSeconds
        TopProcessCount = $TopProcessCount
        AlertTopProcessCount = $AlertTopProcessCount
        CpuWarningThreshold = $CpuWarningThreshold
        CpuCriticalThreshold = $CpuCriticalThreshold
        MaxParallel = $MaxParallel
        MaxForcedProcessesPerCategory = $MaxForcedProcessesPerCategory
        AutoDefenderPerfRecording = $AutoDefenderPerfRecording
        DefenderPerfTriggerServerCpuPercent = $DefenderPerfTriggerServerCpuPercent
        DefenderPerfRecordingSeconds = $DefenderPerfRecordingSeconds
        DefenderPerfCooldownMinutes = $DefenderPerfCooldownMinutes
        MaxConcurrentDefenderPerfRecordings = $MaxConcurrentDefenderPerfRecordings
        DefenderPerfFinalWaitSeconds = $defenderPerfFinalWaitSeconds
        FinalizationTimeoutSeconds = $finalizationTimeoutSeconds
        DefenderPerfLocalRoot = $defenderPerfLocalRoot
        DefenderPerfCopyToOutputPath = $defenderPerfCopyToOutputPath
        IncludeWemEventContext = $IncludeWemEventContext
        WemTriggerServerCpuPercent = $WemTriggerServerCpuPercent
        WemEventWindowMinutes = $WemEventWindowMinutes
        IncludeWemLogTail = $IncludeWemLogTail
        WemLogTailLines = $WemLogTailLines
        IncludeEventLogContext = $IncludeEventContext
        IncludeEventContext = $IncludeEventContext
        ImageVersion = $ImageVersion
        Notes = $Notes
        RunId = $RunId
        IncludeScheduledTaskInventory = $IncludeScheduledTaskInventory
        IncludeCylanceHealth = $IncludeCylanceHealth
        TaskNamesToCheck = $taskNames
        AnonymizeUsers = $AnonymizeUsers
        MaxEventsPerAlert = $MaxEventsPerAlert
        OutputDelimiter = $OutputDelimiter
        WinRMTimeoutSeconds = $WinRMTimeoutSeconds
        IncludeDisconnectedSessions = $IncludeDisconnectedSessions
    }
    ($settings | ConvertTo-Json -Depth 3) | Set-Content -LiteralPath $settingsPath -Encoding UTF8
}

function Get-LatestFile {
    param([Parameter(Mandatory=$true)][string]$Folder, [Parameter(Mandatory=$true)][string]$Filter)
    if (-not (Test-Path -LiteralPath $Folder)) { return $null }
    return Get-ChildItem -LiteralPath $Folder -Filter $Filter -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
}

function Quote-Argument {
    param([Parameter(Mandatory=$true)][string]$Value)
    return '"{0}"' -f ($Value -replace '"', '\"')
}


function Quote-ArrayArgument {
    param([string[]]$Values)
    return (@($Values) | ForEach-Object { Quote-Argument -Value $_ }) -join ','
}

function Get-LinesFromTextBox {
    param([System.Windows.Forms.TextBox]$TextBox)
    @($TextBox.Lines | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function Open-PathWithShell {
    param([Parameter(Mandatory=$true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { throw "Pfad nicht gefunden: $Path" }
    Start-Process -FilePath $Path | Out-Null
}

function Add-StatusLine {
    param([Parameter(Mandatory=$true)][string]$Message)
    if ($script:GuiClosing) { return }
    if ($null -eq $statusBox -or $statusBox.IsDisposed -or -not $statusBox.IsHandleCreated) { return }
    try {
        $line = '{0}  {1}{2}' -f (Get-Date).ToString('HH:mm:ss'), $Message, [Environment]::NewLine
        if ($statusBox.InvokeRequired) {
            [void]$statusBox.BeginInvoke([System.Action[string]]{
                param($text)
                try {
                    if ($null -ne $statusBox -and -not $statusBox.IsDisposed -and $statusBox.IsHandleCreated) { $statusBox.AppendText($text) }
                }
                catch [ObjectDisposedException] { }
                catch [InvalidOperationException] { }
            }, $line)
        }
        else {
            if (-not $statusBox.IsDisposed -and $statusBox.IsHandleCreated) { $statusBox.AppendText($line) }
        }
    }
    catch [ObjectDisposedException] { }
    catch [InvalidOperationException] { }
}

function Load-GuiData {
    $serversPath = Join-ProjectPath -ChildPath @('config','servers.txt')
    if (Test-Path -LiteralPath $serversPath) {
        $serversBox.Text = [string]::Join([Environment]::NewLine, (Get-Content -LiteralPath $serversPath -Encoding UTF8))
    }
    $settings = Read-SettingsFile
    $cpuSampleBox.Value = [decimal](Get-SettingValue -Settings $settings -Name 'CpuSampleSeconds' -DefaultValue 5)
    $topProcessBox.Value = [decimal](Get-SettingValue -Settings $settings -Name 'TopProcessCount' -DefaultValue 10)
    $winRmTimeoutBox.Value = [decimal](Get-SettingValue -Settings $settings -Name 'WinRMTimeoutSeconds' -DefaultValue 5)
    $delimiterBox.Text = [string](Get-SettingValue -Settings $settings -Name 'OutputDelimiter' -DefaultValue ';')
    $includeDisconnectedBox.Checked = [bool](Get-SettingValue -Settings $settings -Name 'IncludeDisconnectedSessions' -DefaultValue $true)
    if ($outputPathBox -and [string]::IsNullOrWhiteSpace($outputPathBox.Text)) { $outputPathBox.Text = [IO.Path]::Combine($ProjectRoot, 'output') }
    if ($manualDurationBox) { $manualDurationBox.Value = [decimal](Get-SettingValue -Settings $settings -Name 'DurationMinutes' -DefaultValue 480) }
    if ($manualIntervalBox) { $manualIntervalBox.Value = [decimal](Get-SettingValue -Settings $settings -Name 'IntervalSeconds' -DefaultValue 300) }
    if ($runCpuSampleBox) { $runCpuSampleBox.Value = [decimal](Get-SettingValue -Settings $settings -Name 'CpuSampleSeconds' -DefaultValue 5) }
    if ($runTopProcessBox) { $runTopProcessBox.Value = [decimal](Get-SettingValue -Settings $settings -Name 'TopProcessCount' -DefaultValue 10) }
    if ($runAlertTopProcessBox) { $runAlertTopProcessBox.Value = [decimal](Get-SettingValue -Settings $settings -Name 'AlertTopProcessCount' -DefaultValue 25) }
    if ($runWarningBox) { $runWarningBox.Value = [decimal](Get-SettingValue -Settings $settings -Name 'CpuWarningThreshold' -DefaultValue 70) }
    if ($runCriticalBox) { $runCriticalBox.Value = [decimal](Get-SettingValue -Settings $settings -Name 'CpuCriticalThreshold' -DefaultValue 90) }
    if ($runMaxParallelBox) { $runMaxParallelBox.Value = [decimal](Get-SettingValue -Settings $settings -Name 'MaxParallel' -DefaultValue 4) }
    if ($runMaxForcedBox) { $runMaxForcedBox.Value = [decimal](Get-SettingValue -Settings $settings -Name 'MaxForcedProcessesPerCategory' -DefaultValue 10) }
    if ($runAutoDefenderBox) { $runAutoDefenderBox.Checked = [bool](Get-SettingValue -Settings $settings -Name 'AutoDefenderPerfRecording' -DefaultValue $false) }
    if ($runDefenderTriggerBox) { $runDefenderTriggerBox.Value = [decimal](Get-SettingValue -Settings $settings -Name 'DefenderPerfTriggerServerCpuPercent' -DefaultValue 10) }
    if ($runDefenderSecondsBox) { $runDefenderSecondsBox.Value = [decimal](Get-SettingValue -Settings $settings -Name 'DefenderPerfRecordingSeconds' -DefaultValue 900) }
    if ($runDefenderCooldownBox) { $runDefenderCooldownBox.Value = [decimal](Get-SettingValue -Settings $settings -Name 'DefenderPerfCooldownMinutes' -DefaultValue 120) }
    if ($runDefenderConcurrentBox) { $runDefenderConcurrentBox.Value = [decimal](Get-SettingValue -Settings $settings -Name 'MaxConcurrentDefenderPerfRecordings' -DefaultValue 2) }
    if ($runIncludeWemBox) { $runIncludeWemBox.Checked = [bool](Get-SettingValue -Settings $settings -Name 'IncludeWemEventContext' -DefaultValue $false) }
    if ($runWemTriggerBox) { $runWemTriggerBox.Value = [decimal](Get-SettingValue -Settings $settings -Name 'WemTriggerServerCpuPercent' -DefaultValue 10) }
    if ($runWemWindowBox) { $runWemWindowBox.Value = [decimal](Get-SettingValue -Settings $settings -Name 'WemEventWindowMinutes' -DefaultValue 10) }
    if ($runIncludeWemTailBox) { $runIncludeWemTailBox.Checked = [bool](Get-SettingValue -Settings $settings -Name 'IncludeWemLogTail' -DefaultValue $false) }
    if ($runWemTailLinesBox) { $runWemTailLinesBox.Value = [decimal](Get-SettingValue -Settings $settings -Name 'WemLogTailLines' -DefaultValue 200) }
    if ($runIncludeEventsBox) { $runIncludeEventsBox.Checked = [bool](Get-SettingValue -Settings $settings -Name 'IncludeEventLogContext' -DefaultValue (Get-SettingValue -Settings $settings -Name 'IncludeEventContext' -DefaultValue $false)) }
    if ($runMaxEventsBox) { $runMaxEventsBox.Value = [decimal](Get-SettingValue -Settings $settings -Name 'MaxEventsPerAlert' -DefaultValue 50) }
    if ($runAnonymizeBox) { $runAnonymizeBox.Checked = [bool](Get-SettingValue -Settings $settings -Name 'AnonymizeUsers' -DefaultValue $false) }
    if ($runImageVersionBox) { $runImageVersionBox.Text = [string](Get-SettingValue -Settings $settings -Name 'ImageVersion' -DefaultValue '') }
    if ($runNotesBox) { $runNotesBox.Text = [string](Get-SettingValue -Settings $settings -Name 'Notes' -DefaultValue '') }
    if ($runRunIdBox) { $runRunIdBox.Text = [string](Get-SettingValue -Settings $settings -Name 'RunId' -DefaultValue '') }
    if ($runTaskInventoryBox) { $runTaskInventoryBox.Checked = [bool](Get-SettingValue -Settings $settings -Name 'IncludeScheduledTaskInventory' -DefaultValue $false) }
    if ($runCylanceBox) { $runCylanceBox.Checked = [bool](Get-SettingValue -Settings $settings -Name 'IncludeCylanceHealth' -DefaultValue $false) }
    if ($runTaskNamesBox) { $runTaskNamesBox.Text = [string]::Join([Environment]::NewLine, @(Get-SettingValue -Settings $settings -Name 'TaskNamesToCheck' -DefaultValue @('nWizard_{B2FE1952-0186-46C3-BAEC-A80AA35AC5B8}','Adobe Acrobat Update Task','MicrosoftEdgeUpdateTaskMachineUA','Launch Adobe CCXProcess','LexwareAppSysOpt','Office Automatic Updates 2.0','Office Feature Updates','Office Feature Updates Logon','BackgroundDownload'))) }
    if ($taskImageVersionBox) { $taskImageVersionBox.Text = [string](Get-SettingValue -Settings $settings -Name 'ImageVersion' -DefaultValue '') }
    if ($taskNotesBox) { $taskNotesBox.Text = [string](Get-SettingValue -Settings $settings -Name 'Notes' -DefaultValue '') }
    if ($taskRunIdBox) { $taskRunIdBox.Text = [string](Get-SettingValue -Settings $settings -Name 'RunId' -DefaultValue '') }
    if ($taskMaxForcedBox) { $taskMaxForcedBox.Value = [decimal](Get-SettingValue -Settings $settings -Name 'MaxForcedProcessesPerCategory' -DefaultValue 10) }
    if ($taskAutoDefenderBox) { $taskAutoDefenderBox.Checked = [bool](Get-SettingValue -Settings $settings -Name 'AutoDefenderPerfRecording' -DefaultValue $false) }
    if ($taskDefenderTriggerBox) { $taskDefenderTriggerBox.Value = [decimal](Get-SettingValue -Settings $settings -Name 'DefenderPerfTriggerServerCpuPercent' -DefaultValue 10) }
    if ($taskDefenderSecondsBox) { $taskDefenderSecondsBox.Value = [decimal](Get-SettingValue -Settings $settings -Name 'DefenderPerfRecordingSeconds' -DefaultValue 900) }
    if ($taskDefenderCooldownBox) { $taskDefenderCooldownBox.Value = [decimal](Get-SettingValue -Settings $settings -Name 'DefenderPerfCooldownMinutes' -DefaultValue 120) }
    if ($taskDefenderConcurrentBox) { $taskDefenderConcurrentBox.Value = [decimal](Get-SettingValue -Settings $settings -Name 'MaxConcurrentDefenderPerfRecordings' -DefaultValue 2) }
    if ($taskIncludeWemBox) { $taskIncludeWemBox.Checked = [bool](Get-SettingValue -Settings $settings -Name 'IncludeWemEventContext' -DefaultValue $false) }
    if ($taskWemTriggerBox) { $taskWemTriggerBox.Value = [decimal](Get-SettingValue -Settings $settings -Name 'WemTriggerServerCpuPercent' -DefaultValue 10) }
    if ($taskWemWindowBox) { $taskWemWindowBox.Value = [decimal](Get-SettingValue -Settings $settings -Name 'WemEventWindowMinutes' -DefaultValue 10) }
    if ($taskIncludeWemTailBox) { $taskIncludeWemTailBox.Checked = [bool](Get-SettingValue -Settings $settings -Name 'IncludeWemLogTail' -DefaultValue $false) }
    if ($taskWemTailLinesBox) { $taskWemTailLinesBox.Value = [decimal](Get-SettingValue -Settings $settings -Name 'WemLogTailLines' -DefaultValue 200) }
    if ($taskIncludeEventsBox) { $taskIncludeEventsBox.Checked = [bool](Get-SettingValue -Settings $settings -Name 'IncludeEventLogContext' -DefaultValue (Get-SettingValue -Settings $settings -Name 'IncludeEventContext' -DefaultValue $false)) }
    if ($taskMaxEventsBox) { $taskMaxEventsBox.Value = [decimal](Get-SettingValue -Settings $settings -Name 'MaxEventsPerAlert' -DefaultValue 50) }
    if ($taskAnonymizeBox) { $taskAnonymizeBox.Checked = [bool](Get-SettingValue -Settings $settings -Name 'AnonymizeUsers' -DefaultValue $false) }
    if ($taskTaskInventoryBox) { $taskTaskInventoryBox.Checked = [bool](Get-SettingValue -Settings $settings -Name 'IncludeScheduledTaskInventory' -DefaultValue $false) }
    if ($taskCylanceBox) { $taskCylanceBox.Checked = [bool](Get-SettingValue -Settings $settings -Name 'IncludeCylanceHealth' -DefaultValue $false) }
    if ($taskTaskNamesBox) { $taskTaskNamesBox.Text = [string]::Join([Environment]::NewLine, @(Get-SettingValue -Settings $settings -Name 'TaskNamesToCheck' -DefaultValue @('nWizard_{B2FE1952-0186-46C3-BAEC-A80AA35AC5B8}','Adobe Acrobat Update Task','MicrosoftEdgeUpdateTaskMachineUA','Launch Adobe CCXProcess','LexwareAppSysOpt','Office Automatic Updates 2.0','Office Feature Updates','Office Feature Updates Logon','BackgroundDownload'))) }
    Add-StatusLine 'Konfiguration geladen.'
}

function Save-GuiData {
    $configFolder = Join-ProjectPath -ChildPath @('config')
    if (-not (Test-Path -LiteralPath $configFolder)) { New-Item -ItemType Directory -Path $configFolder -Force | Out-Null }
    $serversPath = Join-ProjectPath -ChildPath @('config','servers.txt')
    $serversBox.Lines | Set-Content -LiteralPath $serversPath -Encoding UTF8
    Save-SettingsFile -CpuSampleSeconds ([int]$runCpuSampleBox.Value) `
        -TopProcessCount ([int]$runTopProcessBox.Value) `
        -WinRMTimeoutSeconds ([int]$winRmTimeoutBox.Value) `
        -DurationMinutes ([int]$manualDurationBox.Value) `
        -IntervalSeconds ([int]$manualIntervalBox.Value) `
        -AlertTopProcessCount ([int]$runAlertTopProcessBox.Value) `
        -CpuWarningThreshold ([double]$runWarningBox.Value) `
        -CpuCriticalThreshold ([double]$runCriticalBox.Value) `
        -MaxParallel ([int]$runMaxParallelBox.Value) `
        -MaxForcedProcessesPerCategory ([int]$runMaxForcedBox.Value) `
        -AutoDefenderPerfRecording $runAutoDefenderBox.Checked `
        -DefenderPerfTriggerServerCpuPercent ([double]$runDefenderTriggerBox.Value) `
        -DefenderPerfRecordingSeconds ([int]$runDefenderSecondsBox.Value) `
        -DefenderPerfCooldownMinutes ([int]$runDefenderCooldownBox.Value) `
        -MaxConcurrentDefenderPerfRecordings ([int]$runDefenderConcurrentBox.Value) `
        -IncludeWemEventContext $runIncludeWemBox.Checked `
        -WemTriggerServerCpuPercent ([double]$runWemTriggerBox.Value) `
        -WemEventWindowMinutes ([int]$runWemWindowBox.Value) `
        -IncludeWemLogTail $runIncludeWemTailBox.Checked `
        -WemLogTailLines ([int]$runWemTailLinesBox.Value) `
        -OutputDelimiter $delimiterBox.Text `
        -IncludeDisconnectedSessions $includeDisconnectedBox.Checked `
        -ImageVersion $runImageVersionBox.Text `
        -Notes $runNotesBox.Text `
        -RunId $runRunIdBox.Text `
        -IncludeScheduledTaskInventory $runTaskInventoryBox.Checked `
        -IncludeCylanceHealth $runCylanceBox.Checked `
        -IncludeEventContext $runIncludeEventsBox.Checked `
        -AnonymizeUsers $runAnonymizeBox.Checked `
        -MaxEventsPerAlert ([int]$runMaxEventsBox.Value) `
        -TaskNamesToCheck (Get-LinesFromTextBox -TextBox $runTaskNamesBox)
    Add-StatusLine 'Konfiguration gespeichert.'
}

function Start-HealthCheckRun {
    if ($script:HealthCheckProcess -and -not $script:HealthCheckProcess.HasExited) {
        [System.Windows.Forms.MessageBox]::Show('Es laeuft bereits ein HealthCheck.', 'Citrix-TS-HealthCheck', 'OK', 'Information') | Out-Null
        return
    }

    Save-GuiData
    $scriptPath = Join-ProjectPath -ChildPath @('Invoke-CitrixTSHealthCheck.ps1')
    if (-not (Test-Path -LiteralPath $scriptPath)) { throw "HealthCheck-Script nicht gefunden: $scriptPath" }

    $tempRoot = [System.IO.Path]::GetTempPath()
    $runId = Get-Date -Format 'yyyyMMdd-HHmmss'
    $script:StdOutFile = Join-Path $tempRoot "CitrixTSHealthCheck-$runId.out.txt"
    $script:StdErrFile = Join-Path $tempRoot "CitrixTSHealthCheck-$runId.err.txt"

    $arguments = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', ('"{0}"' -f $scriptPath),
        '-ServerListPath', (Quote-Argument -Value (Join-ProjectPath -ChildPath @('config','servers.txt'))),
        '-ConfigPath', ('"{0}"' -f (Join-ProjectPath -ChildPath @('config'))),
        '-OutputPath', (Quote-Argument -Value $outputPathBox.Text),
        '-DurationMinutes', ([int]$manualDurationBox.Value),
        '-IntervalSeconds', ([int]$manualIntervalBox.Value),
        '-CpuSampleSeconds', ([int]$runCpuSampleBox.Value),
        '-TopProcessCount', ([int]$runTopProcessBox.Value),
        '-AlertTopProcessCount', ([int]$runAlertTopProcessBox.Value),
        '-CpuWarningThreshold', ([double]$runWarningBox.Value),
        '-CpuCriticalThreshold', ([double]$runCriticalBox.Value),
        '-MaxParallel', ([int]$runMaxParallelBox.Value),
        '-MaxForcedProcessesPerCategory', ([int]$runMaxForcedBox.Value),
        '-MaxEventsPerAlert', ([int]$runMaxEventsBox.Value),
        '-DefenderPerfTriggerServerCpuPercent', ([double]$runDefenderTriggerBox.Value),
        '-DefenderPerfRecordingSeconds', ([int]$runDefenderSecondsBox.Value),
        '-DefenderPerfCooldownMinutes', ([int]$runDefenderCooldownBox.Value),
        '-MaxConcurrentDefenderPerfRecordings', ([int]$runDefenderConcurrentBox.Value),
        '-WemTriggerServerCpuPercent', ([double]$runWemTriggerBox.Value),
        '-WemEventWindowMinutes', ([int]$runWemWindowBox.Value),
        '-WemLogTailLines', ([int]$runWemTailLinesBox.Value)
    )
    if (-not [string]::IsNullOrWhiteSpace($runImageVersionBox.Text)) { $arguments += @('-ImageVersion', (Quote-Argument -Value $runImageVersionBox.Text)) }
    if (-not [string]::IsNullOrWhiteSpace($runNotesBox.Text)) { $arguments += @('-Notes', (Quote-Argument -Value $runNotesBox.Text)) }
    if (-not [string]::IsNullOrWhiteSpace($runRunIdBox.Text)) { $arguments += @('-RunId', (Quote-Argument -Value $runRunIdBox.Text)) }
    if ($runTaskInventoryBox.Checked) { $arguments += '-IncludeScheduledTaskInventory' }
    if ($runCylanceBox.Checked) { $arguments += '-IncludeCylanceHealth' }
    $runTaskNames = @(Get-LinesFromTextBox -TextBox $runTaskNamesBox)
    if ($runTaskNames.Count -gt 0) { $arguments += @('-TaskNamesToCheck', (Quote-ArrayArgument -Values $runTaskNames)) }
    if ($runIncludeEventsBox.Checked) { $arguments += '-IncludeEventContext' }
    if ($runAutoDefenderBox.Checked) { $arguments += '-AutoDefenderPerfRecording' }
    if ($runIncludeWemBox.Checked) { $arguments += '-IncludeWemEventContext' }
    if ($runIncludeWemTailBox.Checked) { $arguments += '-IncludeWemLogTail' }
    if ($runAnonymizeBox.Checked) { $arguments += '-AnonymizeUsers' }
    $arguments = $arguments -join ' '
    Add-StatusLine ("PowerShell-Aufruf: powershell.exe $arguments")

    $script:HealthCheckProcess = Start-Process -FilePath 'powershell.exe' `
        -ArgumentList $arguments `
        -WorkingDirectory $ProjectRoot `
        -RedirectStandardOutput $script:StdOutFile `
        -RedirectStandardError $script:StdErrFile `
        -WindowStyle Hidden `
        -PassThru

    $runButton.Enabled = $false
    $stopButton.Enabled = $true
    $progressBar.Style = 'Marquee'
    Add-StatusLine "HealthCheck gestartet. PID: $($script:HealthCheckProcess.Id)"
    $timer.Start()
}

function Register-HealthCheckScheduledTask {
    Save-GuiData
    $scriptPath = Join-ProjectPath -ChildPath @('Invoke-CitrixTSHealthCheck.ps1')
    if (-not (Test-Path -LiteralPath $scriptPath)) { throw "HealthCheck-Script nicht gefunden: $scriptPath" }

    $taskNameBase = $taskNameBox.Text.Trim()
    if (-not $taskNameBase) { throw 'Bitte einen Tasknamen angeben.' }
    $taskNameSafe = $taskNameBase -replace '[\\/:*?"<>|]', '_'
    $taskTimestamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
    $taskName = '{0}_{1}' -f $taskNameSafe, $taskTimestamp
    $taskPath = '\TS Health Checks\'

    $actionArguments = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', (Quote-Argument -Value $scriptPath),
        '-ServerListPath', (Quote-Argument -Value (Join-ProjectPath -ChildPath @('config','servers.txt'))),
        '-ConfigPath', (Quote-Argument -Value (Join-ProjectPath -ChildPath @('config'))),
        '-OutputPath', (Quote-Argument -Value $outputPathBox.Text),
        '-DurationMinutes', ([int]$taskRunDurationBox.Value),
        '-IntervalSeconds', ([int]$taskIntervalBox.Value),
        '-CpuSampleSeconds', ([int]$taskCpuSampleBox.Value),
        '-TopProcessCount', ([int]$taskTopProcessBox.Value),
        '-AlertTopProcessCount', ([int]$taskAlertTopProcessBox.Value),
        '-CpuWarningThreshold', ([double]$taskWarningBox.Value),
        '-CpuCriticalThreshold', ([double]$taskCriticalBox.Value),
        '-MaxParallel', ([int]$taskMaxParallelBox.Value),
        '-MaxForcedProcessesPerCategory', ([int]$taskMaxForcedBox.Value),
        '-MaxEventsPerAlert', ([int]$taskMaxEventsBox.Value),
        '-DefenderPerfTriggerServerCpuPercent', ([double]$taskDefenderTriggerBox.Value),
        '-DefenderPerfRecordingSeconds', ([int]$taskDefenderSecondsBox.Value),
        '-DefenderPerfCooldownMinutes', ([int]$taskDefenderCooldownBox.Value),
        '-MaxConcurrentDefenderPerfRecordings', ([int]$taskDefenderConcurrentBox.Value),
        '-WemTriggerServerCpuPercent', ([double]$taskWemTriggerBox.Value),
        '-WemEventWindowMinutes', ([int]$taskWemWindowBox.Value),
        '-WemLogTailLines', ([int]$taskWemTailLinesBox.Value)
    )
    if (-not [string]::IsNullOrWhiteSpace($taskImageVersionBox.Text)) { $actionArguments += @('-ImageVersion', (Quote-Argument -Value $taskImageVersionBox.Text)) }
    if (-not [string]::IsNullOrWhiteSpace($taskNotesBox.Text)) { $actionArguments += @('-Notes', (Quote-Argument -Value $taskNotesBox.Text)) }
    if (-not [string]::IsNullOrWhiteSpace($taskRunIdBox.Text)) { $actionArguments += @('-RunId', (Quote-Argument -Value $taskRunIdBox.Text)) }
    if ($taskTaskInventoryBox.Checked) { $actionArguments += '-IncludeScheduledTaskInventory' }
    if ($taskCylanceBox.Checked) { $actionArguments += '-IncludeCylanceHealth' }
    $taskNamesToCheck = @(Get-LinesFromTextBox -TextBox $taskTaskNamesBox)
    if ($taskNamesToCheck.Count -gt 0) { $actionArguments += @('-TaskNamesToCheck', (Quote-ArrayArgument -Values $taskNamesToCheck)) }
    if ($taskIncludeEventsBox.Checked) { $actionArguments += '-IncludeEventContext' }
    if ($taskAutoDefenderBox.Checked) { $actionArguments += '-AutoDefenderPerfRecording' }
    if ($taskIncludeWemBox.Checked) { $actionArguments += '-IncludeWemEventContext' }
    if ($taskIncludeWemTailBox.Checked) { $actionArguments += '-IncludeWemLogTail' }
    if ($taskAnonymizeBox.Checked) { $actionArguments += '-AnonymizeUsers' }
    $actionArguments = $actionArguments -join ' '
    $taskCommandLine = "powershell.exe $actionArguments"
    Add-StatusLine ("Task PowerShell-Aufruf: $taskCommandLine")
    $taskCommandLinePath = $null
    try {
        $taskCommandLogPath = [IO.Path]::Combine($outputPathBox.Text, 'logs')
        if (-not (Test-Path -LiteralPath $taskCommandLogPath)) { [IO.Directory]::CreateDirectory($taskCommandLogPath) | Out-Null }
        $taskCommandLinePath = [IO.Path]::Combine($taskCommandLogPath, ("TaskCommandLine_{0}.txt" -f $taskTimestamp))
        @(
            "Created=$(Get-Date -Format s)",
            "TaskPath=$taskPath",
            "TaskName=$taskName",
            "LogonType=S4U",
            "RunWhetherUserIsLoggedOnOrNot=True",
            "StartTime=$($taskStartPicker.Value.ToString('s'))",
            "CommandLine=$taskCommandLine"
        ) | Set-Content -LiteralPath $taskCommandLinePath -Encoding UTF8
        Add-StatusLine ("TaskCommandLine gespeichert: $taskCommandLinePath")
    }
    catch {
        Add-StatusLine ("TaskCommandLine konnte nicht gespeichert werden: $($_.Exception.Message)")
    }

    $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $actionArguments -WorkingDirectory $ProjectRoot
    if ($taskRepeatEnabledBox.Checked) {
        $trigger = New-ScheduledTaskTrigger -Once -At $taskStartPicker.Value `
            -RepetitionInterval (New-TimeSpan -Minutes ([int]$taskRepeatMinutesBox.Value)) `
            -RepetitionDuration (New-TimeSpan -Days ([int]$taskRepeatDaysBox.Value))
    }
    else {
        $trigger = New-ScheduledTaskTrigger -Once -At $taskStartPicker.Value
    }
    $principalUser = if ($env:USERDOMAIN) { "$env:USERDOMAIN\$env:USERNAME" } else { $env:USERNAME }
    $principal = New-ScheduledTaskPrincipal -UserId $principalUser -LogonType S4U -RunLevel Highest
    $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Hours ([int]$taskExecutionLimitHoursBox.Value))
    Register-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
    $successMessage = "Scheduled Task eingerichtet: $taskPath$taskName`r`nAusfuehrung: unabhaengig von Benutzeranmeldung (Run whether user is logged on or not)"
    if ($taskCommandLinePath) { $successMessage = "$successMessage`r`nTaskCommandLine: $taskCommandLinePath" }
    Add-StatusLine $successMessage
    [System.Windows.Forms.MessageBox]::Show($successMessage, 'Task erfolgreich eingerichtet', 'OK', 'Information') | Out-Null
}


function Stop-HealthCheckRun {
    if (-not $script:HealthCheckProcess -or $script:HealthCheckProcess.HasExited) {
        Add-StatusLine 'Kein laufender HealthCheck zum Stoppen gefunden.'
        return
    }
    $answer = [System.Windows.Forms.MessageBox]::Show('Laufenden HealthCheck wirklich abbrechen?', 'Citrix-TS-HealthCheck', 'YesNo', 'Warning')
    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    try {
        if ($timer) { $timer.Stop() }
        Add-StatusLine "Stop angefordert. PID: $($script:HealthCheckProcess.Id)"
        Stop-Process -Id $script:HealthCheckProcess.Id -Force -ErrorAction Stop
        try { $script:HealthCheckProcess.WaitForExit(5000) | Out-Null } catch { }
        Add-StatusLine 'HealthCheck-Prozess wurde beendet.'
        $script:HealthCheckProcess = $null
    }
    catch {
        Add-StatusLine "Stop fehlgeschlagen: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Stop fehlgeschlagen', 'OK', 'Error') | Out-Null
    }
    finally {
        $stopButton.Enabled = $false
        $runButton.Enabled = $true
        $progressBar.Style = 'Blocks'
        $progressBar.Value = 0
    }
}

function Complete-HealthCheckRun {
    $timer.Stop()
    $progressBar.Style = 'Blocks'
    $progressBar.Value = 0
    $runButton.Enabled = $true
    $stopButton.Enabled = $false

    if (-not $script:HealthCheckProcess) { return }
    $script:HealthCheckProcess.Refresh()
    $exitCode = $script:HealthCheckProcess.ExitCode
    Add-StatusLine "HealthCheck beendet. ExitCode: $exitCode"

    if ($script:StdOutFile -and (Test-Path -LiteralPath $script:StdOutFile)) {
        $output = Get-Content -LiteralPath $script:StdOutFile -Raw -ErrorAction SilentlyContinue
        if ($output) { Add-StatusLine "Ausgabe: $($output.Trim())" }
    }
    if ($script:StdErrFile -and (Test-Path -LiteralPath $script:StdErrFile)) {
        $errorOutput = Get-Content -LiteralPath $script:StdErrFile -Raw -ErrorAction SilentlyContinue
        if ($errorOutput) { Add-StatusLine "Fehlerausgabe: $($errorOutput.Trim())" }
    }

    $latestSummary = Get-LatestFile -Folder ([IO.Path]::Combine($outputPathBox.Text, 'summary')) -Filter 'RunSummary_*.csv'
    if ($latestSummary) { Add-StatusLine "Letzte Zusammenfassung: $($latestSummary.FullName)" }
    Add-StatusLine "Output: $($outputPathBox.Text)"
    Add-StatusLine "Logs: $([IO.Path]::Combine($outputPathBox.Text, 'logs'))"
    $script:HealthCheckProcess = $null
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Citrix-TS-HealthCheck'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(1020, 800)
$form.MinimumSize = New-Object System.Drawing.Size(900, 700)

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = 'Fill'
$form.Controls.Add($tabs)

$configTab = New-Object System.Windows.Forms.TabPage
$configTab.Text = 'Konfiguration'
$runTab = New-Object System.Windows.Forms.TabPage
$runTab.Text = 'Ausfuehren'
$runTab.AutoScroll = $true
$outputTab = New-Object System.Windows.Forms.TabPage
$outputTab.Text = 'Ausgaben'
$taskTab = New-Object System.Windows.Forms.TabPage
$taskTab.Text = 'Taskplanung'
$taskTab.AutoScroll = $true
$tabs.TabPages.AddRange(@($configTab, $runTab, $taskTab, $outputTab))

$configTab.Controls.Add((New-Label -Text 'Serverliste' -X 15 -Y 15 -Width 200))
$serversBox = New-Object System.Windows.Forms.TextBox
$serversBox.Multiline = $true
$serversBox.ScrollBars = 'Vertical'
$serversBox.AcceptsReturn = $true
$serversBox.AcceptsTab = $false
$serversBox.Location = New-Object System.Drawing.Point(15, 40)
$serversBox.Size = New-Object System.Drawing.Size(390, 420)
$configTab.Controls.Add($serversBox)

$configTab.Controls.Add((New-Label -Text 'CPU Sample Sekunden' -X 430 -Y 45 -Width 180))
$cpuSampleBox = New-Object System.Windows.Forms.NumericUpDown
$cpuSampleBox.Minimum = 1
$cpuSampleBox.Maximum = 300
$cpuSampleBox.Location = New-Object System.Drawing.Point(620, 42)
$configTab.Controls.Add($cpuSampleBox)

$configTab.Controls.Add((New-Label -Text 'Top Prozesse' -X 430 -Y 85 -Width 180))
$topProcessBox = New-Object System.Windows.Forms.NumericUpDown
$topProcessBox.Minimum = 1
$topProcessBox.Maximum = 100
$topProcessBox.Location = New-Object System.Drawing.Point(620, 82)
$configTab.Controls.Add($topProcessBox)

$configTab.Controls.Add((New-Label -Text 'WinRM Timeout Sekunden' -X 430 -Y 125 -Width 180))
$winRmTimeoutBox = New-Object System.Windows.Forms.NumericUpDown
$winRmTimeoutBox.Minimum = 1
$winRmTimeoutBox.Maximum = 300
$winRmTimeoutBox.Location = New-Object System.Drawing.Point(620, 122)
$configTab.Controls.Add($winRmTimeoutBox)

$configTab.Controls.Add((New-Label -Text 'CSV Trennzeichen' -X 430 -Y 165 -Width 180))
$delimiterBox = New-Object System.Windows.Forms.TextBox
$delimiterBox.Location = New-Object System.Drawing.Point(620, 162)
$delimiterBox.Size = New-Object System.Drawing.Size(70, 22)
$configTab.Controls.Add($delimiterBox)


$configTab.Controls.Add((New-Label -Text 'OutputPath' -X 430 -Y 205 -Width 180))
$outputPathBox = New-Object System.Windows.Forms.TextBox
$outputPathBox.Location = New-Object System.Drawing.Point(620, 202)
$outputPathBox.Size = New-Object System.Drawing.Size(230, 22)
$outputPathBox.Text = [IO.Path]::Combine($ProjectRoot, 'output')
$configTab.Controls.Add($outputPathBox)

$includeDisconnectedBox = New-Object System.Windows.Forms.CheckBox
$includeDisconnectedBox.Text = 'Getrennte Sessions beruecksichtigen'
$includeDisconnectedBox.Location = New-Object System.Drawing.Point(430, 235)
$includeDisconnectedBox.Size = New-Object System.Drawing.Size(320, 24)
$configTab.Controls.Add($includeDisconnectedBox)

$saveButton = New-Button -Text 'Speichern' -X 430 -Y 290 -Width 130
$reloadButton = New-Button -Text 'Neu laden' -X 575 -Y 290 -Width 130
$configTab.Controls.AddRange(@($saveButton, $reloadButton))

$runInfo = New-Object System.Windows.Forms.Label
$runInfo.Text = 'Startet Invoke-CitrixTSHealthCheck.ps1 mit den aktuellen Einstellungen. Bei Dauer 0 wird genau ein Lauf ausgefuehrt.'
$runInfo.Location = New-Object System.Drawing.Point(15, 25)
$runInfo.Size = New-Object System.Drawing.Size(840, 40)
$runTab.Controls.Add($runInfo)

$runTab.Controls.Add((New-Label -Text 'Sammeldauer Minuten' -X 15 -Y 80 -Width 170))
$manualDurationBox = New-Object System.Windows.Forms.NumericUpDown
$manualDurationBox.Minimum = 0
$manualDurationBox.Maximum = 10080
$manualDurationBox.Value = 0
$manualDurationBox.Location = New-Object System.Drawing.Point(205, 77)
$runTab.Controls.Add($manualDurationBox)

$runTab.Controls.Add((New-Label -Text 'Intervall Sekunden' -X 15 -Y 120 -Width 170))
$manualIntervalBox = New-Object System.Windows.Forms.NumericUpDown
$manualIntervalBox.Minimum = 5
$manualIntervalBox.Maximum = 86400
$manualIntervalBox.Value = 60
$manualIntervalBox.Location = New-Object System.Drawing.Point(205, 117)
$runTab.Controls.Add($manualIntervalBox)


$runTab.Controls.Add((New-Label -Text 'CPU Delta Sekunden' -X 430 -Y 80 -Width 150))
$runCpuSampleBox = New-Object System.Windows.Forms.NumericUpDown
$runCpuSampleBox.Minimum = 1
$runCpuSampleBox.Maximum = 300
$runCpuSampleBox.Value = 5
$runCpuSampleBox.Location = New-Object System.Drawing.Point(620, 77)
$runTab.Controls.Add($runCpuSampleBox)

$runTab.Controls.Add((New-Label -Text 'Top Prozesse' -X 430 -Y 110 -Width 150))
$runTopProcessBox = New-Object System.Windows.Forms.NumericUpDown
$runTopProcessBox.Minimum = 1
$runTopProcessBox.Maximum = 100
$runTopProcessBox.Value = 10
$runTopProcessBox.Location = New-Object System.Drawing.Point(620, 107)
$runTab.Controls.Add($runTopProcessBox)

$runTab.Controls.Add((New-Label -Text 'Alert Top Prozesse' -X 430 -Y 140 -Width 150))
$runAlertTopProcessBox = New-Object System.Windows.Forms.NumericUpDown
$runAlertTopProcessBox.Minimum = 1
$runAlertTopProcessBox.Maximum = 200
$runAlertTopProcessBox.Value = 25
$runAlertTopProcessBox.Location = New-Object System.Drawing.Point(620, 137)
$runTab.Controls.Add($runAlertTopProcessBox)

$runTab.Controls.Add((New-Label -Text 'CPU Warn/Kritisch %' -X 430 -Y 170 -Width 150))
$runWarningBox = New-Object System.Windows.Forms.NumericUpDown
$runWarningBox.Minimum = 1
$runWarningBox.Maximum = 100
$runWarningBox.Value = 70
$runWarningBox.Location = New-Object System.Drawing.Point(620, 167)
$runTab.Controls.Add($runWarningBox)
$runCriticalBox = New-Object System.Windows.Forms.NumericUpDown
$runCriticalBox.Minimum = 1
$runCriticalBox.Maximum = 100
$runCriticalBox.Value = 90
$runCriticalBox.Location = New-Object System.Drawing.Point(735, 167)
$runTab.Controls.Add($runCriticalBox)

$runTab.Controls.Add((New-Label -Text 'MaxParallel' -X 430 -Y 200 -Width 150))
$runMaxParallelBox = New-Object System.Windows.Forms.NumericUpDown
$runMaxParallelBox.Minimum = 1
$runMaxParallelBox.Maximum = 64
$runMaxParallelBox.Value = 4
$runMaxParallelBox.Location = New-Object System.Drawing.Point(620, 197)
$runTab.Controls.Add($runMaxParallelBox)

$runIncludeEventsBox = New-Object System.Windows.Forms.CheckBox
$runIncludeEventsBox.Text = 'Eventlog-Kontext bei Critical'
$runIncludeEventsBox.Location = New-Object System.Drawing.Point(430, 230)
$runIncludeEventsBox.Size = New-Object System.Drawing.Size(230, 24)
$runTab.Controls.Add($runIncludeEventsBox)

$runTab.Controls.Add((New-Label -Text 'Max Events/Alert' -X 430 -Y 255 -Width 150))
$runMaxEventsBox = New-Object System.Windows.Forms.NumericUpDown
$runMaxEventsBox.Minimum = 0
$runMaxEventsBox.Maximum = 500
$runMaxEventsBox.Value = 50
$runMaxEventsBox.Location = New-Object System.Drawing.Point(620, 252)
$runTab.Controls.Add($runMaxEventsBox)

$runAnonymizeBox = New-Object System.Windows.Forms.CheckBox
$runAnonymizeBox.Text = 'Benutzer anonymisieren'
$runAnonymizeBox.Location = New-Object System.Drawing.Point(665, 230)
$runAnonymizeBox.Size = New-Object System.Drawing.Size(190, 24)
$runTab.Controls.Add($runAnonymizeBox)


$runTab.Controls.Add((New-Label -Text 'ImageVersion' -X 15 -Y 215 -Width 170))
$runImageVersionBox = New-Object System.Windows.Forms.TextBox
$runImageVersionBox.Location = New-Object System.Drawing.Point(205, 212)
$runImageVersionBox.Size = New-Object System.Drawing.Size(210, 22)
$runTab.Controls.Add($runImageVersionBox)

$runTab.Controls.Add((New-Label -Text 'RunId optional' -X 15 -Y 245 -Width 170))
$runRunIdBox = New-Object System.Windows.Forms.TextBox
$runRunIdBox.Location = New-Object System.Drawing.Point(205, 242)
$runRunIdBox.Size = New-Object System.Drawing.Size(210, 22)
$runTab.Controls.Add($runRunIdBox)

$runTab.Controls.Add((New-Label -Text 'Notes' -X 15 -Y 275 -Width 170))
$runNotesBox = New-Object System.Windows.Forms.TextBox
$runNotesBox.Multiline = $true
$runNotesBox.ScrollBars = 'Vertical'
$runNotesBox.Location = New-Object System.Drawing.Point(205, 272)
$runNotesBox.Size = New-Object System.Drawing.Size(650, 45)
$runTab.Controls.Add($runNotesBox)

$runTaskInventoryBox = New-Object System.Windows.Forms.CheckBox
$runTaskInventoryBox.Text = 'Scheduled Task Inventory erfassen'
$runTaskInventoryBox.Location = New-Object System.Drawing.Point(15, 330)
$runTaskInventoryBox.Size = New-Object System.Drawing.Size(260, 24)
$runTab.Controls.Add($runTaskInventoryBox)

$runCylanceBox = New-Object System.Windows.Forms.CheckBox
$runCylanceBox.Text = 'Cylance/Aurora Health erfassen'
$runCylanceBox.Location = New-Object System.Drawing.Point(300, 330)
$runCylanceBox.Size = New-Object System.Drawing.Size(260, 24)
$runTab.Controls.Add($runCylanceBox)

$runTab.Controls.Add((New-Label -Text 'TaskNamesToCheck' -X 15 -Y 365 -Width 170))
$runTaskNamesBox = New-Object System.Windows.Forms.TextBox
$runTaskNamesBox.Multiline = $true
$runTaskNamesBox.ScrollBars = 'Vertical'
$runTaskNamesBox.Location = New-Object System.Drawing.Point(205, 360)
$runTaskNamesBox.Size = New-Object System.Drawing.Size(650, 70)
$runTab.Controls.Add($runTaskNamesBox)

$runTab.Controls.Add((New-Label -Text 'Max Forced/Kategorie' -X 430 -Y 285 -Width 150))
$runMaxForcedBox = New-Object System.Windows.Forms.NumericUpDown
$runMaxForcedBox.Minimum = 0
$runMaxForcedBox.Maximum = 100
$runMaxForcedBox.Value = 10
$runMaxForcedBox.Location = New-Object System.Drawing.Point(620, 282)
$runTab.Controls.Add($runMaxForcedBox)

$runAutoDefenderBox = New-Object System.Windows.Forms.CheckBox
$runAutoDefenderBox.Text = 'Defender Recording automatisch'
$runAutoDefenderBox.Location = New-Object System.Drawing.Point(15, 445)
$runAutoDefenderBox.Size = New-Object System.Drawing.Size(250, 24)
$runTab.Controls.Add($runAutoDefenderBox)

$runTab.Controls.Add((New-Label -Text 'Defender Trigger %' -X 15 -Y 475 -Width 170))
$runDefenderTriggerBox = New-Object System.Windows.Forms.NumericUpDown
$runDefenderTriggerBox.Minimum = 1
$runDefenderTriggerBox.Maximum = 100
$runDefenderTriggerBox.Value = 10
$runDefenderTriggerBox.Location = New-Object System.Drawing.Point(205, 472)
$runTab.Controls.Add($runDefenderTriggerBox)

$runTab.Controls.Add((New-Label -Text 'Defender Sekunden' -X 15 -Y 505 -Width 170))
$runDefenderSecondsBox = New-Object System.Windows.Forms.NumericUpDown
$runDefenderSecondsBox.Minimum = 30
$runDefenderSecondsBox.Maximum = 7200
$runDefenderSecondsBox.Value = 900
$runDefenderSecondsBox.Location = New-Object System.Drawing.Point(205, 502)
$runTab.Controls.Add($runDefenderSecondsBox)

$runTab.Controls.Add((New-Label -Text 'Defender Cooldown Min.' -X 15 -Y 535 -Width 170))
$runDefenderCooldownBox = New-Object System.Windows.Forms.NumericUpDown
$runDefenderCooldownBox.Minimum = 0
$runDefenderCooldownBox.Maximum = 1440
$runDefenderCooldownBox.Value = 120
$runDefenderCooldownBox.Location = New-Object System.Drawing.Point(205, 532)
$runTab.Controls.Add($runDefenderCooldownBox)

$runTab.Controls.Add((New-Label -Text 'Defender max parallel' -X 15 -Y 565 -Width 170))
$runDefenderConcurrentBox = New-Object System.Windows.Forms.NumericUpDown
$runDefenderConcurrentBox.Minimum = 1
$runDefenderConcurrentBox.Maximum = 32
$runDefenderConcurrentBox.Value = 2
$runDefenderConcurrentBox.Location = New-Object System.Drawing.Point(205, 562)
$runTab.Controls.Add($runDefenderConcurrentBox)

$runIncludeWemBox = New-Object System.Windows.Forms.CheckBox
$runIncludeWemBox.Text = 'WEM Event-Kontext'
$runIncludeWemBox.Location = New-Object System.Drawing.Point(430, 445)
$runIncludeWemBox.Size = New-Object System.Drawing.Size(200, 24)
$runTab.Controls.Add($runIncludeWemBox)

$runTab.Controls.Add((New-Label -Text 'WEM Trigger %' -X 430 -Y 475 -Width 150))
$runWemTriggerBox = New-Object System.Windows.Forms.NumericUpDown
$runWemTriggerBox.Minimum = 1
$runWemTriggerBox.Maximum = 100
$runWemTriggerBox.Value = 10
$runWemTriggerBox.Location = New-Object System.Drawing.Point(620, 472)
$runTab.Controls.Add($runWemTriggerBox)

$runTab.Controls.Add((New-Label -Text 'WEM Fenster Min.' -X 430 -Y 505 -Width 150))
$runWemWindowBox = New-Object System.Windows.Forms.NumericUpDown
$runWemWindowBox.Minimum = 1
$runWemWindowBox.Maximum = 240
$runWemWindowBox.Value = 10
$runWemWindowBox.Location = New-Object System.Drawing.Point(620, 502)
$runTab.Controls.Add($runWemWindowBox)

$runIncludeWemTailBox = New-Object System.Windows.Forms.CheckBox
$runIncludeWemTailBox.Text = 'WEM Log-Tail'
$runIncludeWemTailBox.Location = New-Object System.Drawing.Point(430, 535)
$runIncludeWemTailBox.Size = New-Object System.Drawing.Size(150, 24)
$runTab.Controls.Add($runIncludeWemTailBox)

$runTab.Controls.Add((New-Label -Text 'WEM Tail Zeilen' -X 430 -Y 565 -Width 150))
$runWemTailLinesBox = New-Object System.Windows.Forms.NumericUpDown
$runWemTailLinesBox.Minimum = 1
$runWemTailLinesBox.Maximum = 5000
$runWemTailLinesBox.Value = 200
$runWemTailLinesBox.Location = New-Object System.Drawing.Point(620, 562)
$runTab.Controls.Add($runWemTailLinesBox)

$eightHourPresetButton = New-Button -Text '8h Preset' -X 205 -Y 165 -Width 100 -Height 34
$runTab.Controls.Add($eightHourPresetButton)

$runButton = New-Button -Text 'HealthCheck starten' -X 15 -Y 165 -Width 170 -Height 34
$runTab.Controls.Add($runButton)

$stopButton = New-Button -Text 'HealthCheck stoppen' -X 320 -Y 165 -Width 170 -Height 34
$stopButton.Enabled = $false
$runTab.Controls.Add($stopButton)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(15, 615)
$progressBar.Size = New-Object System.Drawing.Size(840, 24)
$runTab.Controls.Add($progressBar)

$statusBox = New-Object System.Windows.Forms.TextBox
$statusBox.Multiline = $true
$statusBox.ScrollBars = 'Vertical'
$statusBox.ReadOnly = $true
$statusBox.Location = New-Object System.Drawing.Point(15, 650)
$statusBox.Size = New-Object System.Drawing.Size(840, 225)
$runTab.Controls.Add($statusBox)

$openSummaryButton = New-Button -Text 'Letzte Summary' -X 25 -Y 35 -Width 150
$openRawButton = New-Button -Text 'Letzte Raw CSV' -X 190 -Y 35 -Width 150
$openLogButton = New-Button -Text 'Letztes Log' -X 355 -Y 35 -Width 150
$openOutputButton = New-Button -Text 'Output Ordner' -X 520 -Y 35 -Width 150
$outputTab.Controls.AddRange(@($openSummaryButton, $openRawButton, $openLogButton, $openOutputButton))


$taskTab.Controls.Add((New-Label -Text 'Taskname' -X 25 -Y 30 -Width 190))
$taskNameBox = New-Object System.Windows.Forms.TextBox
$taskNameBox.Text = 'Citrix TS HealthCheck'
$taskNameBox.Location = New-Object System.Drawing.Point(240, 27)
$taskNameBox.Size = New-Object System.Drawing.Size(260, 22)
$taskTab.Controls.Add($taskNameBox)

$taskTab.Controls.Add((New-Label -Text 'Startzeit' -X 25 -Y 70 -Width 190))
$taskStartPicker = New-Object System.Windows.Forms.DateTimePicker
$taskStartPicker.Format = 'Custom'
$taskStartPicker.CustomFormat = 'yyyy-MM-dd HH:mm'
$taskStartPicker.Value = (Get-Date).AddMinutes(5)
$taskStartPicker.Location = New-Object System.Drawing.Point(240, 67)
$taskStartPicker.Size = New-Object System.Drawing.Size(180, 22)
$taskTab.Controls.Add($taskStartPicker)

$taskTomorrow7Button = New-Button -Text 'Morgen 07:00' -X 520 -Y 27 -Width 110 -Height 28
$taskTab.Controls.Add($taskTomorrow7Button)
$taskTomorrow8Button = New-Button -Text 'Morgen 08:00' -X 645 -Y 27 -Width 110 -Height 28
$taskTab.Controls.Add($taskTomorrow8Button)

$taskRepeatEnabledBox = New-Object System.Windows.Forms.CheckBox
$taskRepeatEnabledBox.Text = 'Wiederholen aktivieren'
$taskRepeatEnabledBox.Location = New-Object System.Drawing.Point(520, 67)
$taskRepeatEnabledBox.Size = New-Object System.Drawing.Size(200, 24)
$taskRepeatEnabledBox.Checked = $false
$taskTab.Controls.Add($taskRepeatEnabledBox)

$taskTab.Controls.Add((New-Label -Text 'Wdh. alle Minuten' -X 25 -Y 110 -Width 190))
$taskRepeatMinutesBox = New-Object System.Windows.Forms.NumericUpDown
$taskRepeatMinutesBox.Minimum = 1
$taskRepeatMinutesBox.Maximum = 1440
$taskRepeatMinutesBox.Value = 15
$taskRepeatMinutesBox.Location = New-Object System.Drawing.Point(240, 107)
$taskTab.Controls.Add($taskRepeatMinutesBox)

$taskTab.Controls.Add((New-Label -Text 'Wiederholen fuer Tage' -X 25 -Y 150 -Width 190))
$taskRepeatDaysBox = New-Object System.Windows.Forms.NumericUpDown
$taskRepeatDaysBox.Minimum = 1
$taskRepeatDaysBox.Maximum = 3650
$taskRepeatDaysBox.Value = 365
$taskRepeatDaysBox.Location = New-Object System.Drawing.Point(240, 147)
$taskTab.Controls.Add($taskRepeatDaysBox)

$taskTab.Controls.Add((New-Label -Text 'Sammeldauer je Start Min.' -X 25 -Y 190 -Width 190))
$taskRunDurationBox = New-Object System.Windows.Forms.NumericUpDown
$taskRunDurationBox.Minimum = 0
$taskRunDurationBox.Maximum = 10080
$taskRunDurationBox.Value = 480
$taskRunDurationBox.Location = New-Object System.Drawing.Point(240, 187)
$taskTab.Controls.Add($taskRunDurationBox)

$taskTab.Controls.Add((New-Label -Text 'Messintervall Sekunden' -X 25 -Y 230 -Width 190))
$taskIntervalBox = New-Object System.Windows.Forms.NumericUpDown
$taskIntervalBox.Minimum = 5
$taskIntervalBox.Maximum = 86400
$taskIntervalBox.Value = 300
$taskIntervalBox.Location = New-Object System.Drawing.Point(240, 227)
$taskTab.Controls.Add($taskIntervalBox)

$taskTab.Controls.Add((New-Label -Text 'Max. Laufzeit Stunden' -X 25 -Y 270 -Width 190))
$taskExecutionLimitHoursBox = New-Object System.Windows.Forms.NumericUpDown
$taskExecutionLimitHoursBox.Minimum = 1
$taskExecutionLimitHoursBox.Maximum = 168
$taskExecutionLimitHoursBox.Value = 10
$taskExecutionLimitHoursBox.Location = New-Object System.Drawing.Point(240, 267)
$taskTab.Controls.Add($taskExecutionLimitHoursBox)


$taskTab.Controls.Add((New-Label -Text 'CPU Delta Sekunden' -X 430 -Y 110 -Width 170))
$taskCpuSampleBox = New-Object System.Windows.Forms.NumericUpDown
$taskCpuSampleBox.Minimum = 1
$taskCpuSampleBox.Maximum = 300
$taskCpuSampleBox.Value = 5
$taskCpuSampleBox.Location = New-Object System.Drawing.Point(620, 107)
$taskTab.Controls.Add($taskCpuSampleBox)

$taskTab.Controls.Add((New-Label -Text 'Top Prozesse' -X 430 -Y 150 -Width 170))
$taskTopProcessBox = New-Object System.Windows.Forms.NumericUpDown
$taskTopProcessBox.Minimum = 1
$taskTopProcessBox.Maximum = 100
$taskTopProcessBox.Value = 10
$taskTopProcessBox.Location = New-Object System.Drawing.Point(620, 147)
$taskTab.Controls.Add($taskTopProcessBox)

$taskTab.Controls.Add((New-Label -Text 'Alert Top Prozesse' -X 430 -Y 190 -Width 170))
$taskAlertTopProcessBox = New-Object System.Windows.Forms.NumericUpDown
$taskAlertTopProcessBox.Minimum = 1
$taskAlertTopProcessBox.Maximum = 200
$taskAlertTopProcessBox.Value = 25
$taskAlertTopProcessBox.Location = New-Object System.Drawing.Point(620, 187)
$taskTab.Controls.Add($taskAlertTopProcessBox)

$taskTab.Controls.Add((New-Label -Text 'CPU Warn/Kritisch %' -X 430 -Y 230 -Width 170))
$taskWarningBox = New-Object System.Windows.Forms.NumericUpDown
$taskWarningBox.Minimum = 1
$taskWarningBox.Maximum = 100
$taskWarningBox.Value = 70
$taskWarningBox.Location = New-Object System.Drawing.Point(620, 227)
$taskTab.Controls.Add($taskWarningBox)
$taskCriticalBox = New-Object System.Windows.Forms.NumericUpDown
$taskCriticalBox.Minimum = 1
$taskCriticalBox.Maximum = 100
$taskCriticalBox.Value = 90
$taskCriticalBox.Location = New-Object System.Drawing.Point(735, 227)
$taskTab.Controls.Add($taskCriticalBox)

$taskTab.Controls.Add((New-Label -Text 'MaxParallel' -X 430 -Y 270 -Width 170))
$taskMaxParallelBox = New-Object System.Windows.Forms.NumericUpDown
$taskMaxParallelBox.Minimum = 1
$taskMaxParallelBox.Maximum = 64
$taskMaxParallelBox.Value = 4
$taskMaxParallelBox.Location = New-Object System.Drawing.Point(620, 267)
$taskTab.Controls.Add($taskMaxParallelBox)

$taskIncludeEventsBox = New-Object System.Windows.Forms.CheckBox
$taskIncludeEventsBox.Text = 'Eventlog-Kontext bei Critical'
$taskIncludeEventsBox.Location = New-Object System.Drawing.Point(430, 310)
$taskIncludeEventsBox.Size = New-Object System.Drawing.Size(230, 24)
$taskTab.Controls.Add($taskIncludeEventsBox)

$taskTab.Controls.Add((New-Label -Text 'Max Events/Alert' -X 430 -Y 340 -Width 170))
$taskMaxEventsBox = New-Object System.Windows.Forms.NumericUpDown
$taskMaxEventsBox.Minimum = 0
$taskMaxEventsBox.Maximum = 500
$taskMaxEventsBox.Value = 50
$taskMaxEventsBox.Location = New-Object System.Drawing.Point(620, 337)
$taskTab.Controls.Add($taskMaxEventsBox)

$taskAnonymizeBox = New-Object System.Windows.Forms.CheckBox
$taskAnonymizeBox.Text = 'Benutzer anonymisieren'
$taskAnonymizeBox.Location = New-Object System.Drawing.Point(665, 310)
$taskAnonymizeBox.Size = New-Object System.Drawing.Size(190, 24)
$taskTab.Controls.Add($taskAnonymizeBox)


$taskTab.Controls.Add((New-Label -Text 'ImageVersion' -X 25 -Y 375 -Width 190))
$taskImageVersionBox = New-Object System.Windows.Forms.TextBox
$taskImageVersionBox.Location = New-Object System.Drawing.Point(240, 372)
$taskImageVersionBox.Size = New-Object System.Drawing.Size(260, 22)
$taskTab.Controls.Add($taskImageVersionBox)

$taskTab.Controls.Add((New-Label -Text 'RunId optional' -X 520 -Y 375 -Width 120))
$taskRunIdBox = New-Object System.Windows.Forms.TextBox
$taskRunIdBox.Location = New-Object System.Drawing.Point(650, 372)
$taskRunIdBox.Size = New-Object System.Drawing.Size(210, 22)
$taskTab.Controls.Add($taskRunIdBox)

$taskTab.Controls.Add((New-Label -Text 'Notes' -X 25 -Y 415 -Width 190))
$taskNotesBox = New-Object System.Windows.Forms.TextBox
$taskNotesBox.Multiline = $true
$taskNotesBox.ScrollBars = 'Vertical'
$taskNotesBox.Location = New-Object System.Drawing.Point(240, 412)
$taskNotesBox.Size = New-Object System.Drawing.Size(620, 45)
$taskTab.Controls.Add($taskNotesBox)

$taskTaskInventoryBox = New-Object System.Windows.Forms.CheckBox
$taskTaskInventoryBox.Text = 'Scheduled Task Inventory erfassen'
$taskTaskInventoryBox.Location = New-Object System.Drawing.Point(25, 470)
$taskTaskInventoryBox.Size = New-Object System.Drawing.Size(260, 24)
$taskTab.Controls.Add($taskTaskInventoryBox)

$taskCylanceBox = New-Object System.Windows.Forms.CheckBox
$taskCylanceBox.Text = 'Cylance/Aurora Health erfassen'
$taskCylanceBox.Location = New-Object System.Drawing.Point(320, 470)
$taskCylanceBox.Size = New-Object System.Drawing.Size(260, 24)
$taskTab.Controls.Add($taskCylanceBox)

$taskTab.Controls.Add((New-Label -Text 'TaskNamesToCheck' -X 25 -Y 505 -Width 190))
$taskTaskNamesBox = New-Object System.Windows.Forms.TextBox
$taskTaskNamesBox.Multiline = $true
$taskTaskNamesBox.ScrollBars = 'Vertical'
$taskTaskNamesBox.Location = New-Object System.Drawing.Point(240, 500)
$taskTaskNamesBox.Size = New-Object System.Drawing.Size(620, 75)
$taskTab.Controls.Add($taskTaskNamesBox)

$taskTab.Controls.Add((New-Label -Text 'Max Forced/Kategorie' -X 430 -Y 585 -Width 170))
$taskMaxForcedBox = New-Object System.Windows.Forms.NumericUpDown
$taskMaxForcedBox.Minimum = 0
$taskMaxForcedBox.Maximum = 100
$taskMaxForcedBox.Value = 10
$taskMaxForcedBox.Location = New-Object System.Drawing.Point(620, 582)
$taskTab.Controls.Add($taskMaxForcedBox)

$taskAutoDefenderBox = New-Object System.Windows.Forms.CheckBox
$taskAutoDefenderBox.Text = 'Defender Recording automatisch'
$taskAutoDefenderBox.Location = New-Object System.Drawing.Point(25, 620)
$taskAutoDefenderBox.Size = New-Object System.Drawing.Size(250, 24)
$taskTab.Controls.Add($taskAutoDefenderBox)

$taskTab.Controls.Add((New-Label -Text 'Defender Trigger %' -X 25 -Y 650 -Width 190))
$taskDefenderTriggerBox = New-Object System.Windows.Forms.NumericUpDown
$taskDefenderTriggerBox.Minimum = 1
$taskDefenderTriggerBox.Maximum = 100
$taskDefenderTriggerBox.Value = 10
$taskDefenderTriggerBox.Location = New-Object System.Drawing.Point(240, 647)
$taskTab.Controls.Add($taskDefenderTriggerBox)

$taskTab.Controls.Add((New-Label -Text 'Defender Sekunden' -X 25 -Y 680 -Width 190))
$taskDefenderSecondsBox = New-Object System.Windows.Forms.NumericUpDown
$taskDefenderSecondsBox.Minimum = 30
$taskDefenderSecondsBox.Maximum = 7200
$taskDefenderSecondsBox.Value = 900
$taskDefenderSecondsBox.Location = New-Object System.Drawing.Point(240, 677)
$taskTab.Controls.Add($taskDefenderSecondsBox)

$taskTab.Controls.Add((New-Label -Text 'Defender Cooldown Min.' -X 25 -Y 710 -Width 190))
$taskDefenderCooldownBox = New-Object System.Windows.Forms.NumericUpDown
$taskDefenderCooldownBox.Minimum = 0
$taskDefenderCooldownBox.Maximum = 1440
$taskDefenderCooldownBox.Value = 120
$taskDefenderCooldownBox.Location = New-Object System.Drawing.Point(240, 707)
$taskTab.Controls.Add($taskDefenderCooldownBox)

$taskTab.Controls.Add((New-Label -Text 'Defender max parallel' -X 25 -Y 740 -Width 190))
$taskDefenderConcurrentBox = New-Object System.Windows.Forms.NumericUpDown
$taskDefenderConcurrentBox.Minimum = 1
$taskDefenderConcurrentBox.Maximum = 32
$taskDefenderConcurrentBox.Value = 2
$taskDefenderConcurrentBox.Location = New-Object System.Drawing.Point(240, 737)
$taskTab.Controls.Add($taskDefenderConcurrentBox)

$taskIncludeWemBox = New-Object System.Windows.Forms.CheckBox
$taskIncludeWemBox.Text = 'WEM Event-Kontext'
$taskIncludeWemBox.Location = New-Object System.Drawing.Point(430, 620)
$taskIncludeWemBox.Size = New-Object System.Drawing.Size(200, 24)
$taskTab.Controls.Add($taskIncludeWemBox)

$taskTab.Controls.Add((New-Label -Text 'WEM Trigger %' -X 430 -Y 650 -Width 170))
$taskWemTriggerBox = New-Object System.Windows.Forms.NumericUpDown
$taskWemTriggerBox.Minimum = 1
$taskWemTriggerBox.Maximum = 100
$taskWemTriggerBox.Value = 10
$taskWemTriggerBox.Location = New-Object System.Drawing.Point(620, 647)
$taskTab.Controls.Add($taskWemTriggerBox)

$taskTab.Controls.Add((New-Label -Text 'WEM Fenster Min.' -X 430 -Y 680 -Width 170))
$taskWemWindowBox = New-Object System.Windows.Forms.NumericUpDown
$taskWemWindowBox.Minimum = 1
$taskWemWindowBox.Maximum = 240
$taskWemWindowBox.Value = 10
$taskWemWindowBox.Location = New-Object System.Drawing.Point(620, 677)
$taskTab.Controls.Add($taskWemWindowBox)

$taskIncludeWemTailBox = New-Object System.Windows.Forms.CheckBox
$taskIncludeWemTailBox.Text = 'WEM Log-Tail'
$taskIncludeWemTailBox.Location = New-Object System.Drawing.Point(430, 710)
$taskIncludeWemTailBox.Size = New-Object System.Drawing.Size(150, 24)
$taskTab.Controls.Add($taskIncludeWemTailBox)

$taskTab.Controls.Add((New-Label -Text 'WEM Tail Zeilen' -X 430 -Y 740 -Width 170))
$taskWemTailLinesBox = New-Object System.Windows.Forms.NumericUpDown
$taskWemTailLinesBox.Minimum = 1
$taskWemTailLinesBox.Maximum = 5000
$taskWemTailLinesBox.Value = 200
$taskWemTailLinesBox.Location = New-Object System.Drawing.Point(620, 737)
$taskTab.Controls.Add($taskWemTailLinesBox)

$task4hPresetButton = New-Button -Text '4h Preset' -X 190 -Y 785 -Width 110 -Height 34
$taskTab.Controls.Add($task4hPresetButton)
$task8hPresetButton = New-Button -Text '8h Preset' -X 315 -Y 785 -Width 110 -Height 34
$taskTab.Controls.Add($task8hPresetButton)

$createTaskButton = New-Button -Text 'Task einrichten' -X 25 -Y 785 -Width 150 -Height 34
$taskTab.Controls.Add($createTaskButton)

$taskHint = New-Object System.Windows.Forms.Label
$taskHint.Text = 'Der Task wird fuer den aktuellen Windows-Benutzer mit hoechsten Rechten eingerichtet. Die GUI muss dafuer ggf. als Administrator gestartet werden.'
$taskHint.Location = New-Object System.Drawing.Point(25, 840)
$taskHint.Size = New-Object System.Drawing.Size(820, 45)
$taskTab.Controls.Add($taskHint)

$outputHint = New-Object System.Windows.Forms.Label
$outputHint.Text = 'Die Schaltflaechen oeffnen die jeweils neueste erzeugte Datei beziehungsweise den Output-Ordner mit dem Windows-Standardprogramm.'
$outputHint.Location = New-Object System.Drawing.Point(25, 85)
$outputHint.Size = New-Object System.Drawing.Size(830, 40)
$outputTab.Controls.Add($outputHint)

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 1000
$timer.Add_Tick({
    try {
        if (-not $script:GuiClosing -and $script:HealthCheckProcess -and $script:HealthCheckProcess.HasExited) { Complete-HealthCheckRun }
    }
    catch [ObjectDisposedException] { if ($timer) { $timer.Stop() } }
    catch [InvalidOperationException] { if ($timer) { $timer.Stop() } }
    catch { Add-StatusLine "GUI-Timerfehler: $($_.Exception.Message)" }
})

$saveButton.Add_Click({
    try { Save-GuiData }
    catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Speichern fehlgeschlagen', 'OK', 'Error') | Out-Null }
})
$reloadButton.Add_Click({
    try { Load-GuiData }
    catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Laden fehlgeschlagen', 'OK', 'Error') | Out-Null }
})
$eightHourPresetButton.Add_Click({
    $manualDurationBox.Value = 480
    $manualIntervalBox.Value = 300
    $runCpuSampleBox.Value = 5
    $runTopProcessBox.Value = 10
    $runAlertTopProcessBox.Value = 25
    $runWarningBox.Value = 70
    $runCriticalBox.Value = 90
    $runMaxParallelBox.Value = 4
    $runMaxForcedBox.Value = 10
    $runAutoDefenderBox.Checked = $false
    $runDefenderTriggerBox.Value = 10
    $runDefenderSecondsBox.Value = 900
    $runDefenderCooldownBox.Value = 120
    $runDefenderConcurrentBox.Value = 2
    $runIncludeWemBox.Checked = $false
    $runWemTriggerBox.Value = 10
    $runWemWindowBox.Value = 10
    $runIncludeWemTailBox.Checked = $false
    $runWemTailLinesBox.Value = 200
    $runMaxEventsBox.Value = 50
    Add-StatusLine '8h Preset gesetzt.'
})
$runButton.Add_Click({
    try { Start-HealthCheckRun }
    catch {
        $runButton.Enabled = $true
        $stopButton.Enabled = $false
        $progressBar.Style = 'Blocks'
        Add-StatusLine "Start fehlgeschlagen: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Start fehlgeschlagen', 'OK', 'Error') | Out-Null
    }
})
$stopButton.Add_Click({
    Stop-HealthCheckRun
})
$taskTomorrow7Button.Add_Click({
    $taskStartPicker.Value = (Get-Date).Date.AddDays(1).AddHours(7)
    Add-StatusLine 'Taskstart auf morgen 07:00 gesetzt.'
})
$taskTomorrow8Button.Add_Click({
    $taskStartPicker.Value = (Get-Date).Date.AddDays(1).AddHours(8)
    Add-StatusLine 'Taskstart auf morgen 08:00 gesetzt.'
})
$task4hPresetButton.Add_Click({
    $taskRunDurationBox.Value = 240
    $taskIntervalBox.Value = 300
    $taskCpuSampleBox.Value = 5
    $taskTopProcessBox.Value = 10
    $taskAlertTopProcessBox.Value = 25
    $taskWarningBox.Value = 70
    $taskCriticalBox.Value = 90
    $taskMaxParallelBox.Value = 4
    $taskMaxForcedBox.Value = 10
    $taskAutoDefenderBox.Checked = $false
    $taskDefenderTriggerBox.Value = 10
    $taskDefenderSecondsBox.Value = 900
    $taskDefenderCooldownBox.Value = 120
    $taskDefenderConcurrentBox.Value = 2
    $taskIncludeWemBox.Checked = $false
    $taskWemTriggerBox.Value = 10
    $taskWemWindowBox.Value = 10
    $taskIncludeWemTailBox.Checked = $false
    $taskWemTailLinesBox.Value = 200
    $taskExecutionLimitHoursBox.Value = 6
    $taskMaxEventsBox.Value = 50
    Add-StatusLine 'Task 4h Preset gesetzt.'
})
$task8hPresetButton.Add_Click({
    $taskRunDurationBox.Value = 480
    $taskIntervalBox.Value = 300
    $taskCpuSampleBox.Value = 5
    $taskTopProcessBox.Value = 10
    $taskAlertTopProcessBox.Value = 25
    $taskWarningBox.Value = 70
    $taskCriticalBox.Value = 90
    $taskMaxParallelBox.Value = 4
    $taskMaxForcedBox.Value = 10
    $taskAutoDefenderBox.Checked = $false
    $taskDefenderTriggerBox.Value = 10
    $taskDefenderSecondsBox.Value = 900
    $taskDefenderCooldownBox.Value = 120
    $taskDefenderConcurrentBox.Value = 2
    $taskIncludeWemBox.Checked = $false
    $taskWemTriggerBox.Value = 10
    $taskWemWindowBox.Value = 10
    $taskIncludeWemTailBox.Checked = $false
    $taskWemTailLinesBox.Value = 200
    $taskExecutionLimitHoursBox.Value = 10
    $taskMaxEventsBox.Value = 50
    Add-StatusLine 'Task 8h Preset gesetzt.'
})
$createTaskButton.Add_Click({
    try { Register-HealthCheckScheduledTask }
    catch {
        Add-StatusLine "Task-Einrichtung fehlgeschlagen: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Task-Einrichtung fehlgeschlagen', 'OK', 'Error') | Out-Null
    }
})
$openSummaryButton.Add_Click({
    try {
        $file = Get-LatestFile -Folder ([IO.Path]::Combine($outputPathBox.Text, 'summary')) -Filter 'RunSummary_*.csv'
        if (-not $file) { throw 'Keine Summary-Datei gefunden.' }
        Open-PathWithShell -Path $file.FullName
    }
    catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Datei oeffnen', 'OK', 'Information') | Out-Null }
})
$openRawButton.Add_Click({
    try {
        $file = Get-LatestFile -Folder ([IO.Path]::Combine($outputPathBox.Text, 'raw')) -Filter 'Raw_ProcessSamples_*.csv'
        if (-not $file) { throw 'Keine Raw-CSV gefunden.' }
        Open-PathWithShell -Path $file.FullName
    }
    catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Datei oeffnen', 'OK', 'Information') | Out-Null }
})
$openLogButton.Add_Click({
    try {
        $file = Get-LatestFile -Folder ([IO.Path]::Combine($outputPathBox.Text, 'logs')) -Filter 'RunLog_*.log'
        if (-not $file) { throw 'Keine Logdatei gefunden.' }
        Open-PathWithShell -Path $file.FullName
    }
    catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Datei oeffnen', 'OK', 'Information') | Out-Null }
})
$openOutputButton.Add_Click({
    try { Open-PathWithShell -Path ($outputPathBox.Text) }
    catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Ordner oeffnen', 'OK', 'Information') | Out-Null }
})
$form.Add_Shown({
    try { Load-GuiData }
    catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Initialisierung fehlgeschlagen', 'OK', 'Error') | Out-Null }
})
$form.Add_FormClosing({
    $script:GuiClosing = $true
    if ($timer) { $timer.Stop() }
    if ($script:HealthCheckProcess -and -not $script:HealthCheckProcess.HasExited) {
        $answer = [System.Windows.Forms.MessageBox]::Show('Ein HealthCheck laeuft noch. GUI trotzdem schliessen?', 'Citrix-TS-HealthCheck', 'YesNo', 'Warning')
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { $_.Cancel = $true; $script:GuiClosing = $false; if ($timer) { $timer.Start() } }
    }
})
$form.Add_FormClosed({
    $script:GuiClosing = $true
    if ($timer) { $timer.Stop() }
})

[void][System.Windows.Forms.Application]::Run($form)
