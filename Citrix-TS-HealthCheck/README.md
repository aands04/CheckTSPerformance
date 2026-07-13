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

Jeder Start erzeugt eine eindeutige `RunId` im Format `YYYY-MM-DD_HH-mm-ss`. Diese `RunId` wird in die CSV-Zeilen geschrieben. Zusaetzlich werden die Daten des Laufs nach `output/runs/<RunId>/` geschrieben, damit mehrere Laeufe am gleichen Tag nicht nur ueber Tagesdateien unterschieden werden muessen.

Session-Felder:

- `ActiveSessions`: aktive Benutzersessions aus `quser`.
- `DisconnectedSessions`: getrennte Benutzersessions aus `quser`.
- `UserSessionsTotal`: `ActiveSessions + DisconnectedSessions`.
- `RawSessionCount`: alle sauber aus `quser` erkannten Sessions.

Alert-Felder:

- `ProcessRank`: Rang innerhalb der geloggten Alert-Prozesse nach CPU.
- `InclusionReason`: `TopCpu`, `ForcedSecurity`, `ForcedNexus`, `ForcedPrinting`, `ForcedCitrixWEM` oder `ForcedMonitoring`.
