# Citrix-TS-HealthCheck

PowerShell-Projekt zur regelmaessigen Performance- und Ursachenanalyse von Windows/Citrix-Terminalservern von einem zentralen Managementserver aus. Die erste Version setzt keine externen PowerShell-Module voraus und ist fuer Windows PowerShell 5.1 ausgelegt.

## Projektstruktur

```text
Citrix-TS-HealthCheck/
├── Invoke-CitrixTSHealthCheck.ps1   # Hauptscript fuer CLI/Scheduled Task
├── Start-CitrixTSHealthCheckGui.ps1  # Windows-Forms-GUI
├── config/
│   ├── servers.txt                  # Zielserver, ein Host pro Zeile
│   └── settings.json                # Laufzeitparameter
├── output/
│   ├── raw/                         # Tages-CSV mit Detaildaten
│   ├── logs/                        # Tages-Logdateien
│   └── summary/                     # Zusammenfassung pro Lauf
└── src/                             # Reserviert fuer Erweiterungen
```

## Voraussetzungen

- Windows PowerShell 5.1 auf dem Managementserver.
- WinRM/PowerShell Remoting vom Managementserver zu allen Zielservern.
- Berechtigungen zum Ausfuehren von Remote-Befehlen und zum Lesen von Prozess-, Counter- und Sessioninformationen.
- Keine externen PowerShell-Module.

## Einrichtung

1. Projekt auf den Managementserver kopieren, z. B. nach `C:\Scripts\Citrix-TS-HealthCheck`.
2. Zielserver in `config\servers.txt` pflegen. Leerzeilen und Zeilen mit `#` werden ignoriert.
3. Parameter in `config\settings.json` anpassen.
4. WinRM grundsaetzlich testen:

```powershell
Test-WSMan -ComputerName TS-SERVER01
```

5. Optional die GUI starten:

```powershell
Set-Location C:\Scripts\Citrix-TS-HealthCheck
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Start-CitrixTSHealthCheckGui.ps1
```

6. Healthcheck manuell ohne GUI starten:

```powershell
Set-Location C:\Scripts\Citrix-TS-HealthCheck
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Invoke-CitrixTSHealthCheck.ps1 -Verbose
```

Laengerer manueller Sammellauf, z. B. 120 Minuten mit Messung alle 60 Sekunden:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Invoke-CitrixTSHealthCheck.ps1 -DurationMinutes 120 -IntervalSeconds 60
```

## Konfiguration

`config/settings.json` enthaelt folgende Werte:

| Einstellung | Bedeutung |
| --- | --- |
| `CpuSampleSeconds` | Messfenster fuer Prozess-CPU-Delta. Standard: 5 Sekunden. |
| `TopProcessCount` | Anzahl der Top-CPU-Prozesse pro Server. Standard: 10. |
| `WinRMTimeoutSeconds` | Timeout fuer `Test-WSMan` je Server. |
| `OutputDelimiter` | CSV-Trennzeichen, standardmaessig `;`. |
| `IncludeDisconnectedSessions` | Getrennte Sessions bei der Sessionzuordnung beruecksichtigen. |

## Ausgaben

- `output/raw/healthcheck-YYYY-MM-DD.csv`: Detaildaten je Server und Top-Prozess. Die Datei wird pro Tag fortgeschrieben.
- `output/summary/summary-YYYYMMDD-HHMMSS.csv`: Zusammenfassung pro Lauf und Server.
- `output/logs/healthcheck-YYYY-MM-DD.log`: Lauf- und Fehlermeldungen.

## GUI verwenden

Die Datei `Start-CitrixTSHealthCheckGui.ps1` stellt eine einfache Windows-Forms-Oberflaeche bereit und benoetigt keine externen Module. Sie ist fuer die Bedienung auf dem Managementserver gedacht.

Die GUI bietet vier Bereiche:

- **Konfiguration**: `config\servers.txt` direkt bearbeiten sowie `CpuSampleSeconds`, `TopProcessCount`, `WinRMTimeoutSeconds`, CSV-Trennzeichen und getrennte Sessions setzen.
- **Ausfuehren**: Konfiguration speichern, `Invoke-CitrixTSHealthCheck.ps1` in einem separaten `powershell.exe`-Prozess starten, die Sammeldauer in Minuten festlegen und das Intervall zwischen den einzelnen Messlaeufen setzen. `0` Minuten bedeutet weiterhin: genau ein Lauf.
- **Taskplanung**: Einen wiederkehrenden Windows Scheduled Task fuer den aktuellen Benutzer einrichten. Konfigurierbar sind Taskname, Startzeit, Wiederholintervall, Wiederholdauer, Sammeldauer je Taskstart, Messintervall und maximale Laufzeit.
- **Ausgaben**: Neueste Summary-, Raw-CSV- und Logdatei oder den gesamten Output-Ordner mit dem Windows-Standardprogramm oeffnen.

Startbefehl:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Start-CitrixTSHealthCheckGui.ps1
```

Falls die GUI ueber eine Verknuepfung, ISE oder ein anderes Startverzeichnis geoeffnet wird, kann der Projektpfad explizit angegeben werden:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\Scripts\Citrix-TS-HealthCheck\Start-CitrixTSHealthCheckGui.ps1 -ProjectRoot C:\Scripts\Citrix-TS-HealthCheck
```

## Scheduled Task Beispiel

Die GUI kann den Task im Tab **Taskplanung** einrichten. Alternativ kann er manuell per PowerShell erstellt werden:

```powershell
$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '-NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\Citrix-TS-HealthCheck\Invoke-CitrixTSHealthCheck.ps1" -DurationMinutes 10 -IntervalSeconds 60'
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).Date -RepetitionInterval (New-TimeSpan -Minutes 15) -RepetitionDuration (New-TimeSpan -Days 1)
$principal = New-ScheduledTaskPrincipal -UserId 'DOMAIN\svc-ts-healthcheck' -LogonType Password -RunLevel Highest
Register-ScheduledTask -TaskName 'Citrix TS HealthCheck' -Action $action -Trigger $trigger -Principal $principal -Description 'Prueft Citrix/Terminalserver Performance per PowerShell Remoting.'
```

## Beispielauswertung

Top belastete Server des Tages:

```powershell
Import-Csv .\output\summary\summary-20260709-120000.csv -Delimiter ';' |
    Where-Object Status -eq 'OK' |
    Sort-Object {[double]$_.CpuPercent} -Descending |
    Select-Object -First 10 TargetServer,CpuPercent,MemoryPercent,ActiveSessions,TopProcess,TopProcessCpuPercent
```

Top Prozesse aus der Tagesdatei:

```powershell
Import-Csv .\output\raw\healthcheck-2026-07-09.csv -Delimiter ';' |
    Sort-Object {[double]$_.ProcessCpuPercent} -Descending |
    Select-Object -First 20 Timestamp,TargetServer,ProcessName,ProcessId,ProcessSessionId,ProcessUserName,ProcessCpuPercent
```

Fehlerhafte Server eines Laufes:

```powershell
Import-Csv .\output\summary\summary-20260709-120000.csv -Delimiter ';' |
    Where-Object Status -eq 'ERROR' |
    Select-Object TargetServer,ErrorMessage
```

## Hinweise zur Messlogik

Die Gesamt-CPU wird ueber die CIM-Klasse `Win32_PerfFormattedData_PerfOS_Processor` gelesen, damit keine lokalisierten Performance-Counter-Pfade wie `\Processor(_Total)\% Processor Time` benoetigt werden. Die Prozess-CPU wird als Delta gemessen: Das Script liest `Get-Process` zu Beginn, wartet standardmaessig 5 Sekunden und liest die Prozesse erneut. Aus der Differenz der CPU-Sekunden je Prozess und der realen Messdauer wird `ProcessCpuPercent` berechnet. Dadurch werden aktuell CPU-lastige Prozesse sichtbar und nicht nur Prozesse mit hoher historisch kumulierter CPU-Zeit.

## Betriebsanleitung fuer 4-8h CPU-Peak-Analyse

Empfohlene Startwerte fuer belastbare Daten bei geringer Zusatzlast:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Invoke-CitrixTSHealthCheck.ps1 `
  -ServerListPath .\config\servers.txt `
  -OutputPath .\output `
  -DurationMinutes 480 `
  -IntervalSeconds 300 `
  -CpuSampleSeconds 5 `
  -TopProcessCount 10 `
  -AlertTopProcessCount 25 `
  -CpuWarningThreshold 70 `
  -CpuCriticalThreshold 90 `
  -MaxParallel 4
```

Kurztest fuer 30 Minuten:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Invoke-CitrixTSHealthCheck.ps1 -DurationMinutes 30 -IntervalSeconds 120 -CpuSampleSeconds 5 -MaxParallel 4
```

Optional mit Eventlog-Kontext und anonymisierten Benutzernamen:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Invoke-CitrixTSHealthCheck.ps1 -DurationMinutes 480 -IntervalSeconds 300 -IncludeEventLogContext -AnonymizeUsers
```

### CSV-Dateien

- `output/raw/Raw_ProcessSamples_YYYY-MM-DD.csv`: eine Zeile je geloggtem Prozess und Messpunkt, inklusive PID, SessionID, Benutzer/Domain, CPU-Delta, Speicher, Pfad, CommandLine, Parent-Prozess, Kategorie und Monitoring-Markierung.
- `output/raw/ServerSamples_YYYY-MM-DD.csv`: eine Zeile je Server und Messpunkt mit CPU, RAM, Sessions, Status, Fehlertext und Warning/Critical-Kennzeichnung.
- `output/raw/AlertSamples_YYYY-MM-DD.csv`: nur bei CPU-Warnung oder CPU-kritischem Zustand; enthaelt Serverdaten, Alert-Level, Kategorie-Summen und die erweiterten Top-Prozesse.
- `output/raw/CategorySummary_YYYY-MM-DD.csv`: je Server und Messpunkt aggregierte CPU-Anteile fuer Security, Nexus, Office, Browser, Citrix, Printing, Monitoring, Windows und Other.
- `output/raw/EventContext_YYYY-MM-DD.csv`: optional bei CPU >= Critical und `-IncludeEventLogContext`; begrenzter Kontext aus System, Application und Defender Operational.
- `output/summary/RunSummary_YYYY-MM-DD_HH-mm.csv`: Zusammenfassung des gesamten Laufs pro Server mit CPU Durchschnitt/Median/Maximum, Warning/Critical-Anzahl, RAM-/Session-Durchschnitt sowie Top-Prozessen und Top-Kategorien nach kumulierter CPU-Zeit.

### Kategorien und Interpretation

- **Security**: `MsSense`, `SenseNdr`, `MsMpEng`, `CylanceSvc`.
- **Citrix**: `BrokerAgent`, `CtxGfx`, `Citrix.Wem.Agent.Service`, `wfcrun32`, `wfica32`.
- **Nexus**: `nexus.framework.healthcare`, `Infoclient`, `anm_neu`.
- **Office**: `WINWORD`, `EXCEL`, `OUTLOOK`, `POWERPNT`.
- **Browser**: `msedge`, `msedgewebview2`, `chrome`.
- **Printing**: `spoolsv`, `splwow64`.
- **Monitoring**: `wsmprovhost`, `WmiPrvSE`, `powershell`, `pwsh`.
- **Windows**: `svchost`, `System`, `explorer`, `dwm`.

Interpretationshinweise:

- CPU hoch + Security hoch: Defender/MDE/Cylance Policies, Scans, Ausschluesse und Updatezeitpunkte pruefen.
- CPU hoch + Nexus hoch: Nexus/Fachanwendung, Updates, Add-ins oder konkrete Benutzeraktion pruefen.
- CPU hoch + Printing hoch: Druckertreiber, Spooler, haengende Druckjobs und Mapping pruefen.
- CPU hoch + wenige Sessions: einzelner Prozess oder Stoerfall ist wahrscheinlicher.
- CPU hoch + viele Sessions: Sizing, Kapazitaet oder Lastverteilung pruefen.
- `wsmprovhost` und `WmiPrvSE` koennen teilweise durch das Monitoring selbst entstehen; diese Prozesse werden als `IsMonitoringRelated=True` markiert.

### Datenschutz

Standardmaessig werden Prozessbenutzer im Klartext gespeichert, damit eine Ursachenanalyse moeglich ist. Schuetze die CSV-Dateien entsprechend. Mit `-AnonymizeUsers` werden Benutzernamen stabil gehasht; Domains bleiben fuer die Einordnung erhalten.

### CPU-Prozent verstehen

`CpuPercent` ist die Gesamt-CPU des Servers. `ProcessCpuPercent` wird aus CPU-Sekunden-Delta geteilt durch Messdauer berechnet. Ein einzelner mehrthreadiger Prozess kann auf Mehrkernsystemen rechnerisch ueber 100 erreichen; der Wert ist damit kernbezogen. Zur Plausibilisierung werden zusaetzlich `TopProcessCpuPercentSum` und Kategorie-Summen ausgegeben.

### 8h-Lauf aus der GUI starten

Im Tab **Ausfuehren** koennen jetzt auch die wichtigsten Langlaufparameter gesetzt werden:

- Sammeldauer Minuten, z. B. `480` fuer 8 Stunden.
- Intervall Sekunden, z. B. `300` fuer eine Messrunde alle 5 Minuten.
- CPU Delta Sekunden, Top Prozesse, Alert Top Prozesse, CPU Warn-/Kritisch-Schwelle und MaxParallel.
- Optional Eventlog-Kontext bei Critical und Benutzer-Anonymisierung.

Die Schaltflaeche **8h Preset** setzt empfohlene Startwerte fuer einen 8-Stunden-Lauf mit geringer Zusatzlast.

### 4h-/8h-Lauf per GUI im Taskplaner einrichten

Im Tab **Taskplanung** kann ein geplanter Lauf z. B. fuer morgen 07:00 oder 08:00 eingerichtet werden:

1. **Morgen 07:00** oder **Morgen 08:00** klicken oder Startzeit manuell setzen.
2. **4h Preset** oder **8h Preset** klicken; alternativ Sammeldauer, Intervall, CPU-Delta, Top-Prozess-Anzahlen, Schwellwerte und MaxParallel manuell setzen.
3. **Wiederholen aktivieren** nur setzen, wenn der Task regelmaessig wiederholt werden soll. Ohne Haken wird genau ein geplanter Lauf angelegt.
4. Optional Eventlog-Kontext oder Benutzer-Anonymisierung aktivieren.
5. **Task einrichten** klicken. Die GUI sollte dafuer bei Bedarf als Administrator gestartet werden.

## Aktuelle Auswertungshinweise

Die Prozess-CPU wird ab dieser Version mit zwei Feldern ausgegeben:

- `ProcessCpuCorePercent`: CPU-Delta bezogen auf einen logischen Prozessor. `100` entspricht ungefaehr einem voll ausgelasteten logischen Prozessor im Messfenster.
- `ProcessCpuServerPercent`: CPU-Delta bezogen auf den gesamten Server. Formel: `ProcessCpuCorePercent / LogicalProcessorCount`.

`ServerSamples_YYYY-MM-DD.csv` und `CategorySummary_YYYY-MM-DD.csv` enthalten zusaetzlich `LogicalProcessorCount`, Kategorie-Summen als Server-Prozentwerte und `TopCategoryByCpuServerPercent`. Am Laufende werden `RunSummary_YYYY-MM-DD_HH-mm.csv` und `RunSummary_YYYY-MM-DD_HH-mm.txt` in `output/summary` geschrieben. Das Laufprotokoll liegt in `output/logs/RunLog_YYYY-MM-DD_HH-mm.log`.

### RunId und Laufordner

Jeder Start erzeugt eine eindeutige `RunId` im Format `yyyyMMdd_HHmmss_GUIDkurz`. Diese `RunId` wird in die CSV-Zeilen geschrieben. Zusaetzlich werden die Daten des Laufs nach `output/runs/<RunId>/` geschrieben, damit mehrere Laeufe am gleichen Tag nicht nur ueber Tagesdateien unterschieden werden muessen.

Session-Felder:

- `ActiveSessions`: aktive Benutzersessions aus `quser`.
- `DisconnectedSessions`: getrennte Benutzersessions aus `quser`.
- `UserSessionsTotal`: `ActiveSessions + DisconnectedSessions`.
- `RawSessionCount`: alle sauber aus `quser` erkannten Sessions.

Alert-Felder:

- `ProcessRank`: Rang innerhalb der geloggten Alert-Prozesse nach CPU.
- `InclusionReason`: `TopCpu`, `ForcedSecurity`, `ForcedNexus`, `ForcedPrinting`, `ForcedCitrixWEM` oder `ForcedMonitoring`.

## Vergleichslaeufe nach Image-Aenderungen

Fuer Vergleiche zwischen Citrix-/MCS-Image-Versionen kann jeder Lauf mit Metadaten markiert werden:

```powershell
.\Invoke-CitrixTSHealthCheck.ps1 `
  -ServerListPath .\config\servers.txt `
  -DurationMinutes 240 `
  -IntervalSeconds 300 `
  -CpuSampleSeconds 5 `
  -ImageVersion "2026-07-16-CylanceFixed-AVExclusions" `
  -Notes "Test nach korrigierter Cylance-Installation und AV-Ausnahmen" `
  -IncludeScheduledTaskInventory `
  -IncludeCylanceHealth `
  -IncludeEventContext
```

Die automatisch erzeugte `RunId` hat das Format `yyyyMMdd_HHmmss_GUIDkurz`, kann bei Bedarf aber mit `-RunId` vorgegeben werden. Alle Laufdateien enthalten die `RunId` im Dateinamen, z. B. `ServerSamples_<RunId>.csv`, `Raw_ProcessSamples_<RunId>.csv`, `AlertSamples_<RunId>.csv`, `CategorySummary_<RunId>.csv`, `RunSummary_<RunId>.csv` und `RunLog_<RunId>.log`. Parallel wird ein isolierter Ordner `output/runs/<RunId>/` geschrieben.

Optionale Zusatzdateien:

- `ScheduledTaskInventory_<RunId>.csv`: inventarisiert die konfigurierten Update-/Autostart-Tasks pro Server. `TaskNamesToCheck` wird als Liste verarbeitet; kommaseparierte Eingaben werden auf einzelne Tasks aufgeteilt. Nicht vorhandene Tasks werden mit `State = NOT_FOUND` protokolliert; das Script deaktiviert keine Tasks.
- `CylanceHealth_<RunId>.csv`: sammelt Basisdaten zu `CylanceSvc`, `sc.exe qprotection CylanceSvc` und `C:\ProgramData\Cylance\Status\Status.json`.
- `CylanceHealth_Duplicates_<RunId>.csv`: meldet doppelte `SerialNumber` oder `StatusDeviceName`, damit geklonte Gold-Image-Identitaeten auffallen.
- `CategorySummaryAggregated_<RunId>.csv`: aggregierte Kategorieauswertung ueber den gesamten Lauf pro Server und Kategorie.
- `EventContext_<RunId>.csv`: optionaler Event-Kontext bei CPU-Warnungen aus TaskScheduler, System, Application, WMI-Activity, Defender, SENSE und, falls vorhanden, Citrix-WEM-Logs.

Die Kategorisierung beruecksichtigt neben dem Prozessnamen auch Pfad und CommandLine. Nexus-Prozesse unter `C:\Program Files (x86)\Nexus\Prog\`, Adobe-Prozesse unter `*\Adobe\*`, Citrix-/Workspace-/WEM-Prozesse unter `*\Citrix\*`, Edge/WebView2, Office sowie Monitoring-Prozesse wie `uberAgent`, `wsmprovhost`, `WmiPrvSE`, `powershell` und `pwsh` werden konsistent in Raw-, Alert-, Server- und Kategorieausgaben markiert.

### Parameter-Prioritaet und Effective Configuration

Beim Start gilt eine feste Reihenfolge: explizite CLI-/GUI-Parameter haben Vorrang vor `config\settings.json`; die Config ueberschreibt nur Werte, die nicht als Parameter uebergeben wurden; Script-Defaults gelten nur, wenn weder CLI/GUI noch Config einen Wert setzt. Das RunLog schreibt deshalb am Anfang einen Block `Effective Configuration` mit Wert und Quelle (`CLI`, `Config`, `Default` oder `Calculated`). Damit muss z. B. ein Start mit `-DurationMinutes 20 -IntervalSeconds 300 -DefenderPerfTriggerServerCpuPercent 1 -DefenderPerfRecordingSeconds 300 -DefenderPerfCooldownMinutes 60` im Log `DurationMinutes=20 Source=CLI`, `MaxRounds=5 Source=Calculated` und die Defender-Werte mit `Source=CLI` zeigen; die geplanten Messpunkte liegen inklusive Minute 0 und Minute 20.

### GUI: alle Collector-Optionen setzen

Die GUI stellt die wichtigsten Collector-Parameter sowohl fuer den manuellen Start als auch fuer geplante Tasks bereit: Laufdauer, Intervall, CPU-Delta, Top-Prozess-Anzahlen, Warn-/Kritisch-Schwellen, Parallelitaet, EventContext/MaxEvents, Anonymisierung, `ImageVersion`, `Notes`, optionale `RunId`, Scheduled-Task-Inventar, Cylance/Aurora-Health und die Liste `TaskNamesToCheck`. Der erzeugte PowerShell-Aufruf wird im Statusfenster protokolliert, damit nachvollziehbar ist, welche Optionen tatsaechlich gestartet oder im Taskplaner hinterlegt wurden. Bei der Taskplanung wird dieselbe vollstaendige Commandline zusaetzlich als `TaskCommandLine_<Zeitstempel>.txt` unter `output\logs` gespeichert.

### Lauf aus der GUI abbrechen

Ein laufender manueller HealthCheck kann im Tab **Ausfuehren** mit **HealthCheck stoppen** beendet werden. Die GUI fragt vorher nach einer Bestaetigung und beendet dann den gestarteten PowerShell-Prozess; bereits geschriebene CSV-/Logdaten bleiben erhalten.

### Getriggerte Detaildiagnosen fuer Defender und WEM

Optional kann der Collector bei auffaelligem `MsMpEng.exe` automatisch ein Defender Performance Recording starten (`-AutoDefenderPerfRecording`). Gesteuert wird dies ueber `DefenderPerfTriggerServerCpuPercent`, `DefenderPerfRecordingSeconds`, `DefenderPerfCooldownMinutes` und `MaxConcurrentDefenderPerfRecordings`. Das Recording wird auf dem Zielserver unter `C:\ProgramData\CitrixTSHealthCheck\DefenderPerf\<RunId>\` gespeichert; die Recording-Ausgabe landet in einer `.log`-Datei statt in der Hauptkonsole. Danach werden ETL, Log, Textreport und Raw-JSON nach `<OutputPath>\DefenderPerf\<RunId>\<Server>\` kopiert und in `DefenderPerfRecordings_<RunId>.csv` referenziert, sofern tatsaechlich ein Recording gestartet wurde.

Mit `-IncludeWemEventContext` sammelt der Collector bei auffaelliger WEM-CPU (`Citrix.Wem.Agent.Service`, `VUEMUIAgent` oder Kategorie `Wem`) asynchron WEM-Eventlogs im konfigurierten Zeitfenster. Optional kann mit `-IncludeWemLogTail` ein Tail bekannter WEM-Logdateien geschrieben werden. Fehler in diesen Detaildiagnosen werden protokolliert und brechen den Hauptlauf nicht ab.

### GUI: neue Detaildiagnose-Optionen

Die Windows-Forms-GUI zeigt die getriggerten Detaildiagnosen jetzt sowohl im Tab **Ausfuehren** als auch im Tab **Taskplanung** an. Dadurch koennen manuelle 4h/8h-Laeufe und geplante Laeufe dieselben Optionen setzen:

- **Defender Recording automatisch** aktiviert `-AutoDefenderPerfRecording`.
- **Defender Trigger %**, **Defender Sekunden**, **Defender Cooldown Min.** und **Defender max parallel** setzen die zugehoerigen Defender-Parameter.
- **WEM Event-Kontext**, **WEM Trigger %**, **WEM Fenster Min.**, **WEM Log-Tail** und **WEM Tail Zeilen** setzen die WEM-Diagnoseparameter.
- **CPU Spike Protection aktiv** setzt `-IncludeWemCpuSpikeProtectionEvents` und protokolliert WEM-CPU-Spike-Protection-Events 7001 bis 7004 unabhaengig vom CPU-Wert des WEM-Agent-Prozesses.
- **Max Forced/Kategorie** setzt `-MaxForcedProcessesPerCategory` und begrenzt forced Alert-Prozesse je Kategorie.

Der Button **HealthCheck stoppen** beendet den aktuell gestarteten PowerShell-Prozess, stoppt den GUI-Timer und setzt die GUI wieder in den Startzustand. Ein Abbruch kann je nach Zeitpunkt dazu fuehren, dass das CLI-Script nur die bis dahin gesammelten Daten und eine bestmoegliche Summary schreibt.


### Laufzeitgrenze und Finalisierung

Der Collector berechnet nach optionalem Scheduled-Task-Inventory und CylanceHealth den `MeasurementStartTime`. Die harte Endzeit ist `MeasurementStartTime + DurationMinutes`; die geplante Rundenzahl ist `Floor(DurationMinutes * 60 / IntervalSeconds) + 1`, damit die Messpunkte bei Minute 0 und am Laufzeitende enthalten sind. Rundenstarts werden absolut geplant (`MeasurementStartTime`, `MeasurementStartTime + IntervalSeconds`, ...). Wenn eine Runde laenger dauert, startet die naechste faellige Runde sofort ohne zusaetzliche Pause; nach `HardEndTime` oder `MaxRounds` wird keine neue Runde gestartet; ein geplanter Start exakt zur `HardEndTime` ist der letzte zulaessige Messpunkt. `RunSummary_<RunId>.csv` enthaelt `EndReason`, `PlannedDurationMinutes`, `ActualDurationMinutes`, `PlannedRounds` und `CompletedRounds`.

Defender Performance Recordings laufen als Background-Jobs. Neue Recordings werden nicht mehr gestartet, wenn die `HardEndTime` erreicht ist oder weniger als `DefenderPerfRecordingSeconds + 60` Sekunden Restzeit verfuegbar sind; das RunLog meldet dann `Defender Recording skipped: insufficient remaining time`. Am Ende wartet der Hauptlauf nur bis `DefenderPerfFinalWaitSeconds` (Default 120) beziehungsweise insgesamt maximal `FinalizationTimeoutSeconds` (Default 300) auf Detaildiagnosen; offene Defender-Jobs werden als `TimedOut` in `DefenderPerfRecordings_<RunId>.csv` protokolliert. `DefenderPerfRecordings_<RunId>.csv` wird pro `RunId + Server + TriggerTimestamp + TriggerProcessName` dedupliziert, sodass Pending-/Completed-Status nicht als doppelte logische Recordings stehen bleiben.


### Defender Trigger- und Reportdetails

`AutoDefenderPerfRecording` startet ausschliesslich bei `MsMpEng.exe`, wenn `ProcessCpuServerPercent` im aktuellen Sample groesser oder gleich `DefenderPerfTriggerServerCpuPercent` ist. `MsSense.exe`, `SenseNdr.exe`, `CylanceSvc.exe` und die Security-Gesamtkategorie loesen kein Defender Recording aus.

Defender-Dateien werden auf dem Zielserver unter `DefenderPerfLocalRoot\<RunId>\<Server>\` abgelegt und optional nach `<OutputPath>\DefenderPerf\<RunId>\<Server>\` kopiert. `DefenderPerfRecordings_<RunId>.csv` enthaelt Triggerprozess, Trigger-CPU, Schwellwert, lokale und zentrale Pfade sowie Report-/Copy-Status. Wenn `AutoDefenderPerfRecording` aktiv ist, aber kein `MsMpEng.exe`-Trigger ein Recording startet, werden DefenderPerfRecordings und DefenderReports im Output-Manifest als nicht erwartete optionale Ausgaben gewertet und erzeugen keinen PostProcessingError.


### WEM CPU Spike Protection

Mit `-IncludeWemCpuSpikeProtectionEvents` erfasst der Collector waehrend der Messzeit WEM-Ereignisse 7001 bis 7004 aus den dynamisch gefundenen WEM-Eventlogs (`Get-WinEvent -ListLog *WEM*`) sowie den bekannten WEM-/Norskale-Kanaelen. Die Option ist standardmaessig deaktiviert und bleibt getrennt von `-IncludeWemEventContext`, das weiterhin nur bei WEM-CPU-Triggern umfangreichen Kontext sammelt. `WemEventContextCooldownMinutes` (Default 10) begrenzt diese umfangreiche Kontextsammlung pro Server; die Spike-Protection-Events sind davon nicht betroffen.

Bei aktivierter Option entstehen pro Lauf zusaetzlich `WemCpuSpikeProtectionEvents_<RunId>.csv` im Raw-Ordner und `output\runs\<RunId>\WemCpuSpikeProtectionEvents.csv` im Run-Ordner. Die Datei wird auch ohne Treffer mit Kopfzeile erstellt. Die zugehoerige `WemCpuSpikeProtectionSummary_<RunId>.csv` fasst Trigger, Prioritaetswechsel und korrelierte Prozesssamples zusammen. `WemPriorityLoweringSeconds` (Default 180) beschreibt die konfigurierte Schutzdauer fuer die Korrelation: Prozesssamples innerhalb dieses Zeitfensters nach einem passenden WEM-Prioritaetsereignis werden mit `WemProtectionLikelyActive` markiert.

Die Event-XML-Felder koennen je nach WEM-Version unterschiedlich heissen. Das Script liest bekannte Feldvarianten bestmoeglich dynamisch aus; wenn ein Feld fehlt, bleibt es leer. Die vollstaendige Eventnachricht und das Raw-XML werden gespeichert, damit unvollstaendige oder versionsabhaengige Events spaeter manuell ausgewertet werden koennen. Zusaetzlich schreibt `WemCpuSpikeProtectionQueryDiagnostics_<RunId>.csv`, welche WEM-Kandidatenlogs pro Server gefunden, ausgewaehlt und erfolgreich abgefragt wurden; nicht vorhandene Alternativlogs zaehlen dabei nicht als QueryError.

Beispiel fuer einen geplanten Task/CLI-Aufruf mit aktiver WEM-CPU-Spike-Erfassung:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\Citrix-TS-HealthCheck\Invoke-CitrixTSHealthCheck.ps1" `
  -ServerListPath "C:\Scripts\Citrix-TS-HealthCheck\config\servers.txt" `
  -ConfigPath "C:\Scripts\Citrix-TS-HealthCheck\config" `
  -OutputPath "C:\Scripts\Citrix-TS-HealthCheck\output" `
  -DurationMinutes 240 `
  -IntervalSeconds 300 `
  -CpuSampleSeconds 5 `
  -IncludeWemCpuSpikeProtectionEvents `
  -WemPriorityLoweringSeconds 180
```

### Taskplanung in der GUI

Beim Erstellen eines geplanten HealthChecks legt die GUI den Task im Task-Scheduler-Unterordner `TS Health Checks` an und haengt dem eingegebenen Basisnamen automatisch einen Zeitstempel im Format `yyyyMMdd_HHmmss` an. Der Task wird mit `LogonType Password` registriert; die GUI fragt beim Erstellen nach dem Kennwort. Dadurch ist im Task Scheduler **Run whether user is logged on or not** ausgewaehlt und **Do not store password** bleibt deaktiviert. Nach erfolgreicher Registrierung erscheint eine Erfolgsmeldung mit vollstaendigem Taskpfad, Namen und Anmeldeoption.

### WEM-Nachverarbeitung, HealthCheck-Eigenlast und Laufstatus

Die WEM-CPU-Spike-Protection-Nachverarbeitung normalisiert Prozessnamen zentral (z. B. `WmiPrvSE.exe` -> `wmiprvse`) und wertet WEM-Events mehrstufig aus: benannte XML-Data-Felder, unbenannte XML-Data-Felder und bekannte Nachrichtenmuster. Die Spalten `ParseSucceeded`, `ParseMethod`, `ParseError`, `OriginalProcessName` und `NormalizedProcessName` helfen bei versionsabhaengigen WEM-Eventformaten; unbekannte Felder werden nicht erfunden, sondern leer gelassen, waehrend `Message` und `RawEventXml` erhalten bleiben.

`Raw_ProcessSamples_<RunId>.csv` enthaelt zusaetzliche WEM-Korrelationsfelder (`WemEventId`, `WemEventRecordId`, `WemCorrelationMethod`, `WemCorrelationConfidence`). Eine Schutzwirkung gilt nur als wahrscheinlich, wenn ein passendes Ereignis zeitlich innerhalb von `WemPriorityLoweringSeconds` liegt und die Zuordnung mindestens ueber PID/Zeit oder PID/Startzeit belastbar ist.

Monitoring-Prozesse wie `wsmprovhost`, `powershell`, `pwsh` und `WmiPrvSE` werden nur dann als `HealthCheck/WinRM` markiert, wenn CommandLine, Parent-Prozess, RunId oder Scriptpfad eine belastbare Zuordnung zum aktuellen HealthCheck liefern. Ohne diesen Nachweis bleibt `WmiPrvSE` als `Monitoring/WMI` auswertbar. Die RunSummary ergaenzt dafuer `CpuAverageIncludingHealthCheck`, `CpuAverageExcludingHealthCheck`, `HealthCheckCpuAverage`, `HealthCheckCpuMaximum`, `HealthCheckCpuP95` und `HealthCheckProcessSampleCount`.

Der Laufstatus unterscheidet kuenftig Messung und Nachverarbeitung ueber `RunStatus`, `MeasurementStatus` und `PostProcessingStatus`. Empfohlene Exitcodes: `0` fuer vollstaendigen Erfolg, `2` fuer erfolgreiche Messung mit fehlerhafter Teil-Nachverarbeitung, `1` fuer fehlgeschlagene Hauptmessung. Die GUI zeigt den numerischen Exitcode und liest den RunStatus aus der neuesten RunSummary aus.

### Regression 12-Minuten-Lauf: WEM-Parsing, HealthCheck-CPU und Output-Manifest

Der WEM-Parser erkennt zusaetzlich das Format `Initializing CPU spike protection for process <Name> (ID: <PID>), created by user <DOMAIN\\User>` und trennt `Process CPU` von `System CPU`. `CpuPercent` wird aus `Process CPU` gelesen, `WemReportedSystemCpuPercent` aus `System CPU`; Dezimalkomma und Dezimalpunkt werden invariant mit Punkt exportiert.

Das verwendete HealthCheck-Konto wird zu Laufbeginn ueber die Windows-Identitaet ermittelt und als `HealthCheckAccountName` in RunLog, Effective Configuration und RunSummary ausgegeben. `wsmprovhost.exe -Embedding` wird nur dann als `HealthCheck/WinRM` klassifiziert, wenn Owner, CommandLine und Messzeitpunkt zum aktuellen Lauf passen. `WmiPrvSE` bleibt ohne diese belastbaren Nachweise `Monitoring/WMI`.

Die CPU-Bereinigung erfolgt pro Server-Sample: `HealthCheckCpuAtSample = Sum(ProcessCpuServerPercent der IsHealthCheckProcess=True-Prozesse)`, `CpuExcludingHealthCheckAtSample = Max(0, Min(ServerCpuPercent, ServerCpuPercent - HealthCheckCpuAtSample))`. Danach werden Durchschnitt, Maximum, P95 und Summe ueber die Sample-Zeitpunkte aggregiert; ohne erkannte HealthCheck-Prozesse sind inklusive und exklusive CPU identisch.

Am Laufende wird `OutputFileManifest_<RunId>.csv` erzeugt. Es enthaelt fuer jede erwartete oder optionale Ausgabedatei LogicalName, Aktivierung, Pfad, Existenz, RowCount, Dateigroesse und Schreibstatus. Der Manifest-Eintrag fuer das Manifest selbst wird erst nach einem erfolgreichen temporären Schreibvorgang als erstellt bewertet, damit das Manifest nicht seinen eigenen Status faelschlich verschlechtert. RunSummary enthaelt daraus `OutputFilesExpectedCount`, `OutputFilesCreatedCount`, `OutputFilesFailedCount`, `OutputFilesExpected`, `OutputFilesCreated`, `OutputFilesFailed` und `OutputFileManifestPath`.

Bei aktivierter WEM-EventContext-Funktion erzeugt der Collector `WemEventContext_<RunId>.csv` immer mit stabilem Header. Ein Lauf ohne WEM-Trigger oder ohne passende WEM-Ereignisse ist ein erfolgreiches Nullergebnis: die Datei existiert header-only, `RowCount=0`, der Manifest-Eintrag bleibt erfolgreich und `RunStatus`/`ExitCode` werden dadurch nicht verschlechtert. Nur echte Export- oder Queryfehler werden als Nachverarbeitungsfehler bewertet.
