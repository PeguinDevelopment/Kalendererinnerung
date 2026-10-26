# Outlook-Erinnerung

Liest deinen Kalender aus dem klassischen Outlook (optional über Microsoft Graph) und zeigt

- **täglich** (um 08:55 bzw. beim ersten Start danach) eine Übersicht der heutigen Termine,
- **täglich dazu** die Wochenübersicht mit der Anzahl der Termine pro Tag (beide Fenster nebeneinander),
- **nachmittags** (ab 16:00) einen Ausblick auf morgen und die nächsten Werktage,
- **vor jedem Termin** ein bildschirmfüllendes Pop-up mit Countdown, Ton und einem Pinguin,
  der durch den Schnee watschelt (bei wichtigen Terminen mit „★ WICHTIG“-Schild).
  Es lässt sich nicht per X oder Alt+F4 schließen, holt sich jede Sekunde wieder nach
  vorne und die Knöpfe „Verstanden“ / „In 5 Min. nochmal“ werden erst nach 3 Sekunden aktiv.

Die Vorlaufzeit kommt aus der Erinnerung im Outlook-Termin, sonst 15 Minuten.
Abgesagte und abgelehnte Termine werden ignoriert, ganztägige bekommen kein Pop-up.

## Wichtige Personen

Termine, bei denen eine der Personen aus `config.json` Organisator oder Teilnehmer ist, werden
hervorgehoben: in den Übersichten rot und mit ★, im Pop-up mit rosa Hintergrund, Namen,
Pinguin mit Schild, lauterem Ton und 6 statt 3 Sekunden gesperrten Knöpfen.

```json
{ "wichtige_personen": ["Person 1", "Person 2"] }
```

Die Namen werden auch in der Form „Nachname, Vorname“ oder als E-Mail-Adresse erkannt.

## Kalenderquelle

**Standard: klassisches Outlook.** Ohne `client_id` in `config.json` liest das Skript den Standardkalender
des klassischen Outlook (`OUTLOOK.EXE`). Dort muss dein Konto eingerichtet sein. Ist Outlook nicht
geöffnet, startet das Skript es unsichtbar im Hintergrund. Das neue Outlook (`olk.exe`) kannst du
trotzdem weiter benutzen, es hat nur selbst keine Skript-Schnittstelle.

**Alternative: Microsoft Graph.** Braucht kein klassisches Outlook, dafür eine App-Registrierung.
Sobald `config.json` eine `client_id` enthält, wird Graph verwendet.

## 1. Nur für Graph: App-Registrierung

Im [Entra-Portal](https://entra.microsoft.com) → *Identität* → *Anwendungen* → *App-Registrierungen* → *Neue Registrierung*:

1. Name: `Outlook-Erinnerung`, Kontotyp: *Nur Konten in diesem Organisationsverzeichnis*
2. Umleitungs-URI: Plattform *Öffentlicher Client/nativ (mobil und Desktop)*, URI `http://localhost`
3. Nach dem Anlegen: *API-Berechtigungen* → *Microsoft Graph* → *Delegiert* → `Calendars.Read`
4. Von der Übersichtsseite **Anwendungs-ID (Client)** und **Verzeichnis-ID (Mandant)** kopieren

Darfst du keine Apps registrieren oder fehlt die Zustimmung für `Calendars.Read`, muss das eure IT erledigen.

Die beiden IDs als `client_id` und `tenant_id` in `config.json` eintragen (siehe `config.example.json`).

## 2. Installation

```powershell
pip install -r requirements.txt
python outlook_reminder.py --liste      # prüfen, ob die Termine stimmen
python outlook_reminder.py --anmelden   # nur bei Graph: Browser öffnet sich, einmal anmelden
```

Bei Graph wird die Anmeldung verschlüsselt (DPAPI) unter `%LOCALAPPDATA%\OutlookReminder` gespeichert
und automatisch erneuert. Läuft sie doch einmal ab, öffnet das Programm den Browser von selbst.

## 3. Ausprobieren

```powershell
python outlook_reminder.py --test-erinnerung   # Beispiel-Pop-up
python outlook_reminder.py --test-wichtig      # Beispiel-Pop-up für einen wichtigen Termin
python outlook_reminder.py --test-tag          # Tagesübersicht
python outlook_reminder.py --test-woche        # Wochenübersicht
python outlook_reminder.py --test-ausblick     # Ausblick auf die nächsten Tage
```

## 4. Dauerhaft im Hintergrund

Ein echter Windows-Dienst kann keine Fenster auf deinem Bildschirm anzeigen. Deshalb läuft das
Programm als geplante Aufgabe: Start bei jeder Anmeldung, ohne Konsolenfenster, Neustart nach Absturz.

```powershell
powershell -ExecutionPolicy Bypass -File install_task.ps1
# mit eigenen Optionen:
powershell -ExecutionPolicy Bypass -File install_task.ps1 -Argumente "--minuten 10 --uebersicht 08:00"
# entfernen:
powershell -ExecutionPolicy Bypass -File uninstall_task.ps1
```

Meldet `install_task.ps1` „Zugriff verweigert“, PowerShell als Administrator starten.

## Optionen

| Option | Standard | Bedeutung |
|---|---|---|
| `--minuten` | 15 | Vorlauf, falls im Termin keine Erinnerung gesetzt ist |
| `--schlummern` | 5 | Minuten bis zur erneuten Erinnerung |
| `--uebersicht` | 08:55 | ab dieser Uhrzeit Tages-/Wochenübersicht zeigen |
| `--nachholen` | 15 | bis so viele Minuten nach Terminbeginn noch erinnern (z. B. Laptop erst um 9:12 geöffnet) |
| `--vorschau` | 16:00 | ab dieser Uhrzeit den Ausblick zeigen |
| `--vorschau-tage` | 3 | Anzahl Werktage im Ausblick |
| `--intervall` | 60 | Prüfintervall in Sekunden |
| `--leise` | – | Pop-ups ohne Ton |

Log-Datei: `%LOCALAPPDATA%\OutlookReminder\outlook_reminder.log`
