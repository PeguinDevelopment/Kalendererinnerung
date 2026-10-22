# Entfernt die geplante Aufgabe und beendet die laufende Outlook-Erinnerung.
$TaskName = "Outlook-Erinnerung"
Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
Get-CimInstance Win32_Process -Filter "Name='pythonw.exe'" |
    Where-Object { $_.CommandLine -like "*outlook_reminder.py*" } |
    ForEach-Object { Stop-Process -Id $_.ProcessId -Force }
Write-Host "Aufgabe '$TaskName' entfernt."
