# Richtet die Outlook-Erinnerung als geplante Aufgabe ein: Start bei jeder Anmeldung,
# unbegrenzte Laufzeit, Neustart nach Absturz. Zusätzliche Optionen über -Argumente,
# z. B.: .\install_task.ps1 -Argumente "--minuten 10 --uebersicht 08:00"
param([string]$Argumente = "")

$ErrorActionPreference = "Stop"
$TaskName = "Outlook-Erinnerung"
$Skript = Join-Path $PSScriptRoot "outlook_reminder.py"
$Pythonw = (Get-Command pythonw.exe -ErrorAction SilentlyContinue).Source
if (-not $Pythonw) { throw "pythonw.exe nicht gefunden. Ist Python im PATH?" }
$Benutzer = "$env:USERDOMAIN\$env:USERNAME"

$Aktion = New-ScheduledTaskAction -Execute $Pythonw -Argument "`"$Skript`" $Argumente" -WorkingDirectory $PSScriptRoot
$Trigger = New-ScheduledTaskTrigger -AtLogOn -User $Benutzer
$Einstellungen = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit ([TimeSpan]::Zero) -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 1) `
    -MultipleInstances IgnoreNew -StartWhenAvailable
$Principal = New-ScheduledTaskPrincipal -UserId $Benutzer -LogonType Interactive -RunLevel Limited

Register-ScheduledTask -TaskName $TaskName -Action $Aktion -Trigger $Trigger -Settings $Einstellungen `
    -Principal $Principal -Description "Tagesübersicht, Wochenübersicht und Pop-up-Erinnerungen für Outlook-Termine" -Force | Out-Null
Start-ScheduledTask -TaskName $TaskName
Write-Host "Aufgabe '$TaskName' eingerichtet und gestartet. Log: $env:LOCALAPPDATA\OutlookReminder\outlook_reminder.log"
