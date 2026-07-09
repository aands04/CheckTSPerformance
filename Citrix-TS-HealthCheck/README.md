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
