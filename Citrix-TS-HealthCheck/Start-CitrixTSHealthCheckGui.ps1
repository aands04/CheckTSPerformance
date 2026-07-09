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
            CpuSampleSeconds = 5
            TopProcessCount = 10
            WinRMTimeoutSeconds = 5
            OutputDelimiter = ';'
            IncludeDisconnectedSessions = $true
        }
    }
    return Get-Content -LiteralPath $settingsPath -Raw -Encoding UTF8 | ConvertFrom-Json
}

function Save-SettingsFile {
    param(
        [int]$CpuSampleSeconds,
        [int]$TopProcessCount,
        [int]$WinRMTimeoutSeconds,
        [string]$OutputDelimiter,
        [bool]$IncludeDisconnectedSessions
    )
    $configFolder = Join-ProjectPath -ChildPath @('config')
    if (-not (Test-Path -LiteralPath $configFolder)) { New-Item -ItemType Directory -Path $configFolder -Force | Out-Null }
    $settingsPath = Join-ProjectPath -ChildPath @('config','settings.json')
    $settings = [ordered]@{
        CpuSampleSeconds = $CpuSampleSeconds
        TopProcessCount = $TopProcessCount
        WinRMTimeoutSeconds = $WinRMTimeoutSeconds
        OutputDelimiter = $OutputDelimiter
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

function Open-PathWithShell {
    param([Parameter(Mandatory=$true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { throw "Pfad nicht gefunden: $Path" }
    Start-Process -FilePath $Path | Out-Null
}

function Add-StatusLine {
    param([Parameter(Mandatory=$true)][string]$Message)
    $statusBox.AppendText(('{0}  {1}{2}' -f (Get-Date).ToString('HH:mm:ss'), $Message, [Environment]::NewLine))
}

function Load-GuiData {
    $serversPath = Join-ProjectPath -ChildPath @('config','servers.txt')
    if (Test-Path -LiteralPath $serversPath) {
        $serversBox.Text = [string]::Join([Environment]::NewLine, (Get-Content -LiteralPath $serversPath -Encoding UTF8))
    }
    $settings = Read-SettingsFile
    $cpuSampleBox.Value = [decimal]$settings.CpuSampleSeconds
    $topProcessBox.Value = [decimal]$settings.TopProcessCount
    $winRmTimeoutBox.Value = [decimal]$settings.WinRMTimeoutSeconds
    $delimiterBox.Text = [string]$settings.OutputDelimiter
    $includeDisconnectedBox.Checked = [bool]$settings.IncludeDisconnectedSessions
    Add-StatusLine 'Konfiguration geladen.'
}

function Save-GuiData {
    $configFolder = Join-ProjectPath -ChildPath @('config')
    if (-not (Test-Path -LiteralPath $configFolder)) { New-Item -ItemType Directory -Path $configFolder -Force | Out-Null }
    $serversPath = Join-ProjectPath -ChildPath @('config','servers.txt')
    $serversBox.Lines | Set-Content -LiteralPath $serversPath -Encoding UTF8
    Save-SettingsFile -CpuSampleSeconds ([int]$cpuSampleBox.Value) `
        -TopProcessCount ([int]$topProcessBox.Value) `
        -WinRMTimeoutSeconds ([int]$winRmTimeoutBox.Value) `
        -OutputDelimiter $delimiterBox.Text `
        -IncludeDisconnectedSessions $includeDisconnectedBox.Checked
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
        '-ConfigPath', ('"{0}"' -f (Join-ProjectPath -ChildPath @('config'))),
        '-OutputPath', ('"{0}"' -f (Join-ProjectPath -ChildPath @('output')))
    ) -join ' '

    $script:HealthCheckProcess = Start-Process -FilePath 'powershell.exe' `
        -ArgumentList $arguments `
        -WorkingDirectory $ProjectRoot `
        -RedirectStandardOutput $script:StdOutFile `
        -RedirectStandardError $script:StdErrFile `
        -WindowStyle Hidden `
        -PassThru

    $runButton.Enabled = $false
    $progressBar.Style = 'Marquee'
    Add-StatusLine "HealthCheck gestartet. PID: $($script:HealthCheckProcess.Id)"
    $timer.Start()
}

function Complete-HealthCheckRun {
    $timer.Stop()
    $progressBar.Style = 'Blocks'
    $progressBar.Value = 0
    $runButton.Enabled = $true

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

    $latestSummary = Get-LatestFile -Folder (Join-ProjectPath -ChildPath @('output','summary')) -Filter 'summary-*.csv'
    if ($latestSummary) { Add-StatusLine "Letzte Zusammenfassung: $($latestSummary.FullName)" }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Citrix-TS-HealthCheck'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(920, 680)
$form.MinimumSize = New-Object System.Drawing.Size(820, 600)

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = 'Fill'
$form.Controls.Add($tabs)

$configTab = New-Object System.Windows.Forms.TabPage
$configTab.Text = 'Konfiguration'
$runTab = New-Object System.Windows.Forms.TabPage
$runTab.Text = 'Ausfuehren'
$outputTab = New-Object System.Windows.Forms.TabPage
$outputTab.Text = 'Ausgaben'
$tabs.TabPages.AddRange(@($configTab, $runTab, $outputTab))

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

$includeDisconnectedBox = New-Object System.Windows.Forms.CheckBox
$includeDisconnectedBox.Text = 'Getrennte Sessions beruecksichtigen'
$includeDisconnectedBox.Location = New-Object System.Drawing.Point(430, 205)
$includeDisconnectedBox.Size = New-Object System.Drawing.Size(320, 24)
$configTab.Controls.Add($includeDisconnectedBox)

$saveButton = New-Button -Text 'Speichern' -X 430 -Y 260 -Width 130
$reloadButton = New-Button -Text 'Neu laden' -X 575 -Y 260 -Width 130
$configTab.Controls.AddRange(@($saveButton, $reloadButton))

$runInfo = New-Object System.Windows.Forms.Label
$runInfo.Text = 'Startet Invoke-CitrixTSHealthCheck.ps1 mit den aktuellen Einstellungen. Die GUI bleibt waehrend des Laufs bedienbar.'
$runInfo.Location = New-Object System.Drawing.Point(15, 25)
$runInfo.Size = New-Object System.Drawing.Size(840, 40)
$runTab.Controls.Add($runInfo)

$runButton = New-Button -Text 'HealthCheck starten' -X 15 -Y 80 -Width 170 -Height 34
$runTab.Controls.Add($runButton)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(205, 84)
$progressBar.Size = New-Object System.Drawing.Size(650, 24)
$runTab.Controls.Add($progressBar)

$statusBox = New-Object System.Windows.Forms.TextBox
$statusBox.Multiline = $true
$statusBox.ScrollBars = 'Vertical'
$statusBox.ReadOnly = $true
$statusBox.Location = New-Object System.Drawing.Point(15, 135)
$statusBox.Size = New-Object System.Drawing.Size(840, 410)
$runTab.Controls.Add($statusBox)

$openSummaryButton = New-Button -Text 'Letzte Summary' -X 25 -Y 35 -Width 150
$openRawButton = New-Button -Text 'Letzte Raw CSV' -X 190 -Y 35 -Width 150
$openLogButton = New-Button -Text 'Letztes Log' -X 355 -Y 35 -Width 150
$openOutputButton = New-Button -Text 'Output Ordner' -X 520 -Y 35 -Width 150
$outputTab.Controls.AddRange(@($openSummaryButton, $openRawButton, $openLogButton, $openOutputButton))

$outputHint = New-Object System.Windows.Forms.Label
$outputHint.Text = 'Die Schaltflaechen oeffnen die jeweils neueste erzeugte Datei beziehungsweise den Output-Ordner mit dem Windows-Standardprogramm.'
$outputHint.Location = New-Object System.Drawing.Point(25, 85)
$outputHint.Size = New-Object System.Drawing.Size(830, 40)
$outputTab.Controls.Add($outputHint)

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 1000
$timer.Add_Tick({
    if ($script:HealthCheckProcess -and $script:HealthCheckProcess.HasExited) { Complete-HealthCheckRun }
})

$saveButton.Add_Click({
    try { Save-GuiData }
    catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Speichern fehlgeschlagen', 'OK', 'Error') | Out-Null }
})
$reloadButton.Add_Click({
    try { Load-GuiData }
    catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Laden fehlgeschlagen', 'OK', 'Error') | Out-Null }
})
$runButton.Add_Click({
    try { Start-HealthCheckRun }
    catch {
        $runButton.Enabled = $true
        $progressBar.Style = 'Blocks'
        Add-StatusLine "Start fehlgeschlagen: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Start fehlgeschlagen', 'OK', 'Error') | Out-Null
    }
})
$openSummaryButton.Add_Click({
    try {
        $file = Get-LatestFile -Folder (Join-ProjectPath -ChildPath @('output','summary')) -Filter 'summary-*.csv'
        if (-not $file) { throw 'Keine Summary-Datei gefunden.' }
        Open-PathWithShell -Path $file.FullName
    }
    catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Datei oeffnen', 'OK', 'Information') | Out-Null }
})
$openRawButton.Add_Click({
    try {
        $file = Get-LatestFile -Folder (Join-ProjectPath -ChildPath @('output','raw')) -Filter 'healthcheck-*.csv'
        if (-not $file) { throw 'Keine Raw-CSV gefunden.' }
        Open-PathWithShell -Path $file.FullName
    }
    catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Datei oeffnen', 'OK', 'Information') | Out-Null }
})
$openLogButton.Add_Click({
    try {
        $file = Get-LatestFile -Folder (Join-ProjectPath -ChildPath @('output','logs')) -Filter 'healthcheck-*.log'
        if (-not $file) { throw 'Keine Logdatei gefunden.' }
        Open-PathWithShell -Path $file.FullName
    }
    catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Datei oeffnen', 'OK', 'Information') | Out-Null }
})
$openOutputButton.Add_Click({
    try { Open-PathWithShell -Path (Join-ProjectPath -ChildPath @('output')) }
    catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Ordner oeffnen', 'OK', 'Information') | Out-Null }
})
$form.Add_Shown({
    try { Load-GuiData }
    catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Initialisierung fehlgeschlagen', 'OK', 'Error') | Out-Null }
})
$form.Add_FormClosing({
    if ($script:HealthCheckProcess -and -not $script:HealthCheckProcess.HasExited) {
        $answer = [System.Windows.Forms.MessageBox]::Show('Ein HealthCheck laeuft noch. GUI trotzdem schliessen?', 'Citrix-TS-HealthCheck', 'YesNo', 'Warning')
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { $_.Cancel = $true }
    }
})

[void][System.Windows.Forms.Application]::Run($form)
