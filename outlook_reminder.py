"""Outlook-Erinnerung: liest den Kalender aus dem klassischen Outlook (oder per Microsoft Graph) und zeigt

- täglich eine Übersicht der heutigen Termine,
- täglich dazu die Wochenübersicht mit der Anzahl der Termine pro Tag,
- nachmittags einen Ausblick auf morgen und die nächsten Tage,
- vor jedem Termin ein bildschirmfüllendes Pop-up mit Ton und einem watschelnden Pinguin.

Läuft dauerhaft im Hintergrund (Start über die Aufgabenplanung, siehe install_task.ps1).
Ohne config.json wird das klassische Outlook verwendet; enthält config.json Client- und
Tenant-ID einer App-Registrierung, wird Microsoft Graph verwendet (siehe README).

Aufruf:
    python outlook_reminder.py --liste     # Tages- und Wochenübersicht in der Konsole ausgeben
    python outlook_reminder.py --anmelden  # nur Graph: einmalig im Browser anmelden
    pythonw outlook_reminder.py            # Normalbetrieb (ohne Konsolenfenster)
    python outlook_reminder.py --test-erinnerung | --test-wichtig | --test-tag | --test-woche | --test-ausblick
"""
# hi
import argparse
import ctypes
import json
import logging
import math
import os
import random
import sys  
import winsound
from datetime import date, datetime, timedelta, timezone
import tkinter as tk
from tkinter import font as tkfont

import customtkinter as ctk

import msal
import requests
from msal_extensions import PersistedTokenCache, build_encrypted_persistence

PROJEKTVERZEICHNIS = os.path.dirname(os.path.abspath(__file__))
KONFIGDATEI = os.path.join(PROJEKTVERZEICHNIS, "config.json")
DATENVERZEICHNIS = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "OutlookReminder")
STATUSDATEI = os.path.join(DATENVERZEICHNIS, "status.json")
TOKENDATEI = os.path.join(DATENVERZEICHNIS, "token_cache.bin")  # per DPAPI verschlüsselt
LOGDATEI = os.path.join(DATENVERZEICHNIS, "outlook_reminder.log")
WOCHENTAGE = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag"]
WICHTIG = "★"  # Kennzeichen für Termine mit wichtigen Personen


def person_dabei(person: str, namen: list) -> bool:
    """True, wenn alle Namensteile der Person in einem der Namen vorkommen
    (findet "Vorname Nachname" auch als "Nachname, Vorname" oder "Vorname.Nachname@...")."""
    teile = person.lower().split()
    return any(all(teil in name.lower() for teil in teile) for name in namen)


class Kalender:
    """Gemeinsame Basis: markiert Termine, an denen wichtige Personen teilnehmen."""

    wichtige_personen: list = []

    def markieren(self, termine: list) -> list:
        for t in termine:
            t["wichtig_mit"] = [p for p in self.wichtige_personen if person_dabei(p, t["teilnehmer"])]
        return termine


# --- Microsoft Graph -------------------------------------------------------

class AnmeldungNoetig(Exception):
    pass


def graph_zeit(feld: dict, ganztaegig: bool) -> datetime:
    """Graph-Zeit (UTC) in naive Lokalzeit umwandeln. Ganztägige Termine sind zeitzonenlos."""
    dt = datetime.fromisoformat(feld["dateTime"][:26])  # Graph liefert 7 Nachkommastellen
    if ganztaegig:
        return dt
    return dt.replace(tzinfo=timezone.utc).astimezone().replace(tzinfo=None)


class GraphKalender(Kalender):
    SCOPES = ["Calendars.Read"]
    URL = "https://graph.microsoft.com/v1.0/me/calendarView"
    FELDER = ("id,subject,start,end,location,isAllDay,isCancelled,isReminderOn,"
              "reminderMinutesBeforeStart,responseStatus,organizer,attendees")

    def __init__(self, client_id: str, tenant_id: str):
        cache = PersistedTokenCache(build_encrypted_persistence(TOKENDATEI))
        self.app = msal.PublicClientApplication(
            client_id, authority=f"https://login.microsoftonline.com/{tenant_id}", token_cache=cache)
        self.session = requests.Session()
        self._naechster_login_versuch = datetime.min

    def token(self) -> str:
        konten = self.app.get_accounts()
        if konten:
            ergebnis = self.app.acquire_token_silent(self.SCOPES, account=konten[0])
            if ergebnis and "access_token" in ergebnis:
                return ergebnis["access_token"]

        # Browser-Anmeldung höchstens alle 30 Minuten anstoßen, sonst nervt es im Minutentakt
        if datetime.now() < self._naechster_login_versuch:
            raise AnmeldungNoetig("Anmeldung ausstehend")
        self._naechster_login_versuch = datetime.now() + timedelta(minutes=30)
        logging.info("Öffne Browser zur Anmeldung bei Microsoft 365")
        ergebnis = self.app.acquire_token_interactive(self.SCOPES, prompt="select_account", timeout=300)
        if "access_token" not in ergebnis:
            raise AnmeldungNoetig(ergebnis.get("error_description", "Anmeldung fehlgeschlagen"))
        return ergebnis["access_token"]

    def termine(self, von: datetime, bis: datetime) -> list:
        """Alle Termine, die sich mit dem Zeitraum [von, bis) überschneiden, nach Beginn sortiert."""
        utc = lambda dt: dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        url = self.URL
        params = {"startDateTime": utc(von), "endDateTime": utc(bis), "$select": self.FELDER,
                  "$orderby": "start/dateTime", "$top": "100"}
        kopf = {"Authorization": f"Bearer {self.token()}", "Prefer": 'outlook.timezone="UTC"'}

        termine = []
        while url:
            antwort = self.session.get(url, params=params, headers=kopf, timeout=30)
            antwort.raise_for_status()
            daten = antwort.json()
            for e in daten.get("value", []):
                if e.get("isCancelled") or e.get("responseStatus", {}).get("response") == "declined":
                    continue
                ganztaegig = bool(e.get("isAllDay"))
                start = graph_zeit(e["start"], ganztaegig)
                personen = [e.get("organizer") or {}] + (e.get("attendees") or [])
                termine.append({
                    "teilnehmer": [p.get("emailAddress", {}).get("name") or "" for p in personen],
                    "id": f"{e['id']}|{start.isoformat()}",
                    "betreff": e.get("subject") or "(ohne Betreff)",
                    "ort": (e.get("location") or {}).get("displayName") or "",
                    "start": start,
                    "ende": graph_zeit(e["end"], ganztaegig),
                    "ganztaegig": ganztaegig,
                    "vorlauf": e.get("reminderMinutesBeforeStart") if e.get("isReminderOn") else None,
                })
            url, params = daten.get("@odata.nextLink"), None  # nextLink enthält die Parameter schon
        termine.sort(key=lambda t: t["start"])
        return self.markieren(termine)


# --- Klassisches Outlook (COM) ---------------------------------------------

def outlook_datum(dt: datetime) -> str:
    """Datumsformat, das Outlooks Restrict-Filter versteht."""
    return dt.strftime("%m/%d/%Y %H:%M")


def lokale_zeit(com_datum) -> datetime:
    # pywin32 liefert Lokalzeit, markiert sie aber als UTC -> tzinfo verwerfen
    return datetime(com_datum.year, com_datum.month, com_datum.day,
                    com_datum.hour, com_datum.minute, com_datum.second)


class OutlookKalender(Kalender):
    """Liest den Standardkalender des klassischen Outlook (OUTLOOK.EXE, nicht das neue Outlook)."""

    OL_FOLDER_CALENDAR = 9
    OL_RESPONSE_DECLINED = 4

    def __init__(self):
        import pythoncom
        pythoncom.CoInitialize()

    def termine(self, von: datetime, bis: datetime) -> list:
        """Alle Termine, die sich mit dem Zeitraum [von, bis) überschneiden, nach Beginn sortiert."""
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        items = outlook.GetNamespace("MAPI").GetDefaultFolder(self.OL_FOLDER_CALENDAR).Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True  # Serientermine einzeln auflösen (muss nach Sort kommen)
        gefiltert = items.Restrict(
            f"[Start] < '{outlook_datum(bis)}' AND [End] > '{outlook_datum(von)}'"
        )

        termine = []
        item = gefiltert.GetFirst()
        while item is not None:
            start, ende = lokale_zeit(item.Start), lokale_zeit(item.End)
            if start >= bis:
                break
            if ende > von and item.ResponseStatus != self.OL_RESPONSE_DECLINED:
                termine.append({
                    "teilnehmer": [item.Organizer or ""] + [r.Name for r in item.Recipients],
                    "id": f"{item.EntryID}|{start.isoformat()}",
                    "betreff": item.Subject or "(ohne Betreff)",
                    "ort": item.Location or "",
                    "start": start,
                    "ende": ende,
                    "ganztaegig": bool(item.AllDayEvent),
                    "vorlauf": item.ReminderMinutesBeforeStart if item.ReminderSet else None,
                })
            item = gefiltert.GetNext()
        return self.markieren(termine)


def kalender_laden() -> Kalender:
    """Graph, wenn config.json eine App-Registrierung enthält, sonst klassisches Outlook."""
    try:
        with open(KONFIGDATEI, encoding="utf-8") as f:
            konfig = json.load(f)
    except OSError:
        konfig = {}
    except ValueError as fehler:
        raise SystemExit(f"config.json ist fehlerhaft ({fehler}).")

    if konfig.get("client_id"):
        if not konfig.get("tenant_id"):
            raise SystemExit("config.json: 'tenant_id' fehlt. Siehe README.")
        kalender = GraphKalender(konfig["client_id"], konfig["tenant_id"])
    else:
        kalender = OutlookKalender()
    kalender.wichtige_personen = konfig.get("wichtige_personen", [])
    return kalender


# --- Übersichten -----------------------------------------------------------

def tagesbeginn(tag: date) -> datetime:
    return datetime.combine(tag, datetime.min.time())


def tagesdaten(kalender, tag: date) -> list:
    von = tagesbeginn(tag)
    return kalender.termine(von, von + timedelta(days=1))


def wochendaten(kalender, tag: date) -> tuple:
    """(Montag, Termine pro Tag, ganztägige pro Tag, wichtige pro Tag) für die Woche von `tag`.
    Ganztägige Termine zählen extra und nicht bei den wichtigen."""
    montag = tag - timedelta(days=tag.weekday())
    von = tagesbeginn(montag)
    anzahl, ganztags, wichtig = [0] * 7, [0] * 7, [0] * 7
    for t in kalender.termine(von, von + timedelta(days=7)):
        index = (max(t["start"], von).date() - montag).days
        if t["ganztaegig"]:
            ganztags[index] += 1
        else:
            anzahl[index] += 1
            wichtig[index] += bool(t["wichtig_mit"])
    return montag, anzahl, ganztags, wichtig


def vorschaudaten(kalender, heute: date, werktage: int) -> list:
    """[(Tag, Termine)] ab morgen für die nächsten `werktage` Werktage.
    Wochenendtage nur, wenn dort Termine liegen."""
    tage, tag = [], heute
    while werktage > 0:
        tag += timedelta(days=1)
        termine = tagesdaten(kalender, tag)
        if tag.weekday() < 5:
            werktage -= 1
            tage.append((tag, termine))
        elif termine:
            tage.append((tag, termine))
    return tage


def relativer_tag(tag: date, heute: date) -> str:
    return {1: "Morgen", 2: "Übermorgen"}.get((tag - heute).days, WOCHENTAGE[tag.weekday()])


def anzahl_text(anzahl: int, wichtige: int) -> str:
    text = "Keine Termine" if anzahl == 0 else f"{anzahl} Termin" + ("" if anzahl == 1 else "e")
    return text + (f" · davon {wichtige} wichtig" if wichtige else "")


def tagesuebersicht(kalender, tag: date) -> tuple:
    """Tagesübersicht als Text (für --liste)."""
    termine = tagesdaten(kalender, tag)
    titel = f"Termine heute – {WOCHENTAGE[tag.weekday()]}, {tag:%d.%m.%Y}"
    if not termine:
        return titel, "Heute stehen keine Termine an."

    zeilen = [termin_zeile(t) for t in termine]
    zeilen += ["", anzahl_text(len(termine), sum(1 for t in termine if t["wichtig_mit"]))]
    return titel, "\n".join(zeilen)


def termin_zeile(t: dict) -> str:
    zeit = "ganztägig  " if t["ganztaegig"] else f"{t['start']:%H:%M}–{t['ende']:%H:%M}"
    ort = f"  ({t['ort']})" if t["ort"] else ""
    if t["wichtig_mit"]:
        return f"{WICHTIG} {zeit}  {t['betreff']}{ort}  – mit {', '.join(t['wichtig_mit'])}"
    return f"  {zeit}  {t['betreff']}{ort}"


def vorschau_uebersicht(kalender, heute: date, werktage: int) -> tuple:
    """Ausblick auf die nächsten Tage als Text (für --liste)."""
    bloecke = []
    for tag, termine in vorschaudaten(kalender, heute, werktage):
        zeilen = [termin_zeile(t) for t in termine] or ["  Keine Termine"]
        bloecke.append(f"{relativer_tag(tag, heute)}, {tag:%d.%m.}\n" + "\n".join(zeilen))
    return "Ausblick – die nächsten Tage", "\n\n".join(bloecke)


def wochenuebersicht(kalender, tag: date) -> tuple:
    """Wochenübersicht als Text (für --liste)."""
    montag, anzahl, ganztags, wichtig = wochendaten(kalender, tag)
    titel = (f"Wochenübersicht KW {montag.isocalendar()[1]} "
             f"({montag:%d.%m.} – {montag + timedelta(days=6):%d.%m.%Y})")
    zeilen = []
    for i in range(7):
        extra = f"  (+{ganztags[i]} ganztägig)" if ganztags[i] else ""
        if wichtig[i]:
            extra += f"  davon {wichtig[i]} wichtig"
        zeilen.append(f"{WICHTIG if wichtig[i] else ' '} {WOCHENTAGE[i][:2]} "
                      f"{montag + timedelta(days=i):%d.%m.}  {anzahl[i]:>2}  {'█' * anzahl[i]}{extra}")
    zeilen += ["", anzahl_text(sum(anzahl), sum(wichtig))]
    return titel, "\n".join(zeilen)


# --- Fenster ---------------------------------------------------------------

MONATE = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli",
          "August", "September", "Oktober", "November", "Dezember"]

# Farben als (hell, dunkel) – die Übersichten folgen dem Farbmodus von Windows
HINTERGRUND = ("#F3F4F6", "#18181B")
KARTE = ("#FFFFFF", "#27272A")
TEXT = ("#111827", "#F4F4F5")
GEDAEMPFT = ("#6B7280", "#A1A1AA")
SPUR = ("#E5E7EB", "#3F3F46")
AKZENT = ("#2563EB", "#60A5FA")
AKZENT_HOVER = ("#1D4ED8", "#3B82F6")
ROT = ("#DC2626", "#F87171")
ROT_FLAECHE = ("#FEF2F2", "#3B1F22")

SCHRIFTART = "Segoe UI"


def schriftart_waehlen(root) -> None:
    global SCHRIFTART
    vorhanden = set(tkfont.families(root))
    SCHRIFTART = next((s for s in ("Segoe UI Variable Display", "Segoe UI") if s in vorhanden), "TkDefaultFont")


def schrift(groesse: int, fett: bool = False) -> ctk.CTkFont:
    return ctk.CTkFont(family=SCHRIFTART, size=groesse, weight="bold" if fett else "normal")


class UebersichtsFenster(ctk.CTkToplevel):
    """Grundgerüst der Übersichten: Kopf, Inhalt, Knopf „Alles klar“; anfangs im Vordergrund."""

    BREITE = 480

    def __init__(self, master, fenstertitel, beim_schliessen=None):
        super().__init__(master, fg_color=HINTERGRUND)
        self._beim_schliessen = beim_schliessen
        self.title(fenstertitel)
        self.resizable(False, False)
        self.minsize(self.BREITE, 0)
        self.attributes("-topmost", True)
        self.protocol("WM_DELETE_WINDOW", self.schliessen)
        self.inhalt = ctk.CTkFrame(self, fg_color="transparent")
        self.inhalt.pack(fill="both", expand=True, padx=28, pady=(24, 22))

    def kopf(self, oben, titel, unten):
        ctk.CTkLabel(self.inhalt, text=oben.upper(), font=schrift(12, True), text_color=AKZENT,
                     anchor="w", height=18).pack(fill="x")
        ctk.CTkLabel(self.inhalt, text=titel, font=schrift(30, True), text_color=TEXT,
                     anchor="w", height=40).pack(fill="x")
        ctk.CTkLabel(self.inhalt, text=unten, font=schrift(14), text_color=GEDAEMPFT,
                     anchor="w", height=20).pack(fill="x", pady=(0, 18))

    def fertig(self):
        ctk.CTkButton(self.inhalt, text="Alles klar", font=schrift(14, True), width=140, height=40,
                      corner_radius=20, fg_color=AKZENT, hover_color=AKZENT_HOVER, text_color=("#FFFFFF", "#0B1220"),
                      command=self.schliessen).pack(anchor="e", pady=(20, 0))
        self.update_idletasks()
        x = (self.winfo_screenwidth() - self.winfo_reqwidth()) // 2
        y = (self.winfo_screenheight() - self.winfo_reqheight()) // 3
        tk.Toplevel.geometry(self, f"+{x}+{y}")  # an CTk vorbei: Werte sind bereits echte Pixel
        self.after(250, self.focus_force)

    def schliessen(self):
        self.destroy()
        if self._beim_schliessen:
            self._beim_schliessen()

    def liste(self, scrollen: bool):
        """Container für Termin-Karten, bei vielen Terminen mit Scrollleiste."""
        if scrollen:
            liste = ctk.CTkScrollableFrame(self.inhalt, fg_color="transparent", width=self.BREITE - 40, height=460)
        else:
            liste = ctk.CTkFrame(self.inhalt, fg_color="transparent")
        liste.pack(fill="both", expand=True)
        return liste

    def karte(self, liste, t):
        wichtig = bool(t["wichtig_mit"])
        karte = ctk.CTkFrame(liste, fg_color=ROT_FLAECHE if wichtig else KARTE, corner_radius=14)
        karte.pack(fill="x", pady=4)
        # height=1: sonst nimmt CTkFrame 200 px Standardhöhe an; fill="y" streckt den Balken auf Kartenhöhe
        ctk.CTkFrame(karte, width=4, height=1, corner_radius=2, fg_color=ROT if wichtig else AKZENT) \
            .pack(side="left", fill="y", padx=(12, 0), pady=14)

        zeit = ctk.CTkFrame(karte, fg_color="transparent")
        zeit.pack(side="left", anchor="n", padx=(14, 6), pady=12)
        if t["ganztaegig"]:
            ctk.CTkLabel(zeit, text="Ganztägig", font=schrift(12, True), text_color=GEDAEMPFT,
                         width=62, anchor="w", height=22).pack()
        else:
            ctk.CTkLabel(zeit, text=f"{t['start']:%H:%M}", font=schrift(16, True), text_color=TEXT,
                         width=62, anchor="w", height=22).pack()
            ctk.CTkLabel(zeit, text=f"{t['ende']:%H:%M}", font=schrift(12), text_color=GEDAEMPFT,
                         width=62, anchor="w", height=16).pack()

        text = ctk.CTkFrame(karte, fg_color="transparent")
        text.pack(side="left", fill="x", expand=True, padx=(0, 16), pady=12)
        ctk.CTkLabel(text, text=t["betreff"], font=schrift(15, True), text_color=TEXT, anchor="w",
                     justify="left", wraplength=320, height=22).pack(fill="x")
        if t["ort"]:
            ctk.CTkLabel(text, text=t["ort"], font=schrift(12), text_color=GEDAEMPFT, anchor="w",
                         justify="left", wraplength=320, height=16).pack(fill="x")
        if wichtig:
            ctk.CTkLabel(text, text=f"{WICHTIG}  mit {', '.join(t['wichtig_mit'])}", font=schrift(12, True),
                         text_color=ROT, anchor="w", height=18).pack(fill="x", pady=(4, 0))


def nebeneinander(*fenster, abstand=24):
    """Mehrere Übersichten mittig nebeneinander statt deckungsgleich übereinander anordnen."""
    gesamt = sum(f.winfo_reqwidth() for f in fenster) + abstand * (len(fenster) - 1)
    x = max((fenster[0].winfo_screenwidth() - gesamt) // 2, 0)
    for f in fenster:
        y = max((f.winfo_screenheight() - f.winfo_reqheight()) // 3, 0)
        tk.Toplevel.geometry(f, f"+{x}+{y}")
        x += f.winfo_reqwidth() + abstand


class TagesFenster(UebersichtsFenster):
    def __init__(self, master, tag, termine, beim_schliessen=None):
        super().__init__(master, "Tagesübersicht", beim_schliessen)
        wichtige = sum(1 for t in termine if t["wichtig_mit"])
        self.kopf(f"{WOCHENTAGE[tag.weekday()]} · Heute", f"{tag.day}. {MONATE[tag.month - 1]}",
                  anzahl_text(len(termine), wichtige))

        if not termine:
            leer = ctk.CTkFrame(self.inhalt, fg_color=KARTE, corner_radius=16)
            leer.pack(fill="x")
            ctk.CTkLabel(leer, text="Heute stehen keine Termine an.", font=schrift(16, True),
                         text_color=TEXT).pack(pady=(28, 0))
            ctk.CTkLabel(leer, text="Zeit für konzentriertes Arbeiten.", font=schrift(13),
                         text_color=GEDAEMPFT).pack(pady=(0, 28))
        else:
            liste = self.liste(scrollen=len(termine) > 6)
            for t in termine:
                self.karte(liste, t)
        self.fertig()


class VorschauFenster(UebersichtsFenster):
    """Ausblick am Nachmittag: Termine von morgen und den nächsten Tagen."""

    def __init__(self, master, heute, tage, beim_schliessen=None):
        super().__init__(master, "Ausblick", beim_schliessen)
        alle = [t for _, termine in tage for t in termine]
        self.kopf("Ausblick", "Die nächsten Tage",
                  anzahl_text(len(alle), sum(1 for t in alle if t["wichtig_mit"])))

        liste = self.liste(scrollen=len(alle) + len(tage) > 8)
        for i, (tag, termine) in enumerate(tage):
            zeile = ctk.CTkFrame(liste, fg_color="transparent")
            zeile.pack(fill="x", pady=(0 if i == 0 else 16, 4))
            relativ = relativer_tag(tag, heute)
            datum = f"{tag.day}. {MONATE[tag.month - 1]}"
            if relativ != WOCHENTAGE[tag.weekday()]:
                datum = f"{WOCHENTAGE[tag.weekday()]}, {datum}"
            ctk.CTkLabel(zeile, text=relativ, font=schrift(16, True), text_color=TEXT, height=24).pack(side="left")
            ctk.CTkLabel(zeile, text=f"  {datum}", font=schrift(13), text_color=GEDAEMPFT, height=24).pack(side="left")
            wichtige = sum(1 for t in termine if t["wichtig_mit"])
            ctk.CTkLabel(zeile, text=anzahl_text(len(termine), wichtige).replace("davon ", ""), font=schrift(12),
                         text_color=ROT if wichtige else GEDAEMPFT, height=24).pack(side="right")
            if not termine:
                leer = ctk.CTkFrame(liste, fg_color=KARTE, corner_radius=14)
                leer.pack(fill="x", pady=4)
                ctk.CTkLabel(leer, text="Frei – keine Termine", font=schrift(13), text_color=GEDAEMPFT,
                             height=20).pack(anchor="w", padx=20, pady=12)
            for t in termine:
                self.karte(liste, t)
        self.fertig()


class WochenFenster(UebersichtsFenster):
    def __init__(self, master, montag, anzahl, ganztags, wichtig, beim_schliessen=None):
        super().__init__(master, "Wochenübersicht", beim_schliessen)
        sonntag = montag + timedelta(days=6)
        if montag.month == sonntag.month:
            zeitraum = f"{montag.day}.–{sonntag.day}. {MONATE[sonntag.month - 1]}"
        else:
            zeitraum = f"{montag.day}. {MONATE[montag.month - 1]} – {sonntag.day}. {MONATE[sonntag.month - 1]}"
        self.kopf(f"Kalenderwoche {montag.isocalendar()[1]}", zeitraum, anzahl_text(sum(anzahl), sum(wichtig)))

        karte = ctk.CTkFrame(self.inhalt, fg_color=KARTE, corner_radius=16)
        karte.pack(fill="x")
        maximum = max(max(anzahl), 1)
        for i in range(7):
            tag = montag + timedelta(days=i)
            heute = tag == date.today()
            zeile = ctk.CTkFrame(karte, fg_color="transparent")
            zeile.pack(fill="x", padx=20, pady=(16 if i == 0 else 4, 16 if i == 6 else 4))

            ctk.CTkLabel(zeile, text=WOCHENTAGE[i][:2], font=schrift(14, True),
                         text_color=AKZENT if heute else TEXT, width=30, anchor="w").pack(side="left")
            ctk.CTkLabel(zeile, text=f"{tag:%d.%m.}", font=schrift(12, heute),
                         text_color=AKZENT if heute else GEDAEMPFT, width=48, anchor="w").pack(side="left")

            spur = ctk.CTkFrame(zeile, fg_color=SPUR, width=210, height=12, corner_radius=6)
            spur.pack(side="left", padx=(4, 14))
            normal = (anzahl[i] - wichtig[i]) / maximum
            if normal:
                ctk.CTkFrame(spur, fg_color=AKZENT, corner_radius=6, height=12) \
                    .place(relx=0, rely=0, relwidth=normal, relheight=1)
            if wichtig[i]:
                ctk.CTkFrame(spur, fg_color=ROT, corner_radius=6, height=12) \
                    .place(relx=normal, rely=0, relwidth=wichtig[i] / maximum, relheight=1)

            ctk.CTkLabel(zeile, text=str(anzahl[i]), font=schrift(15, True),
                         text_color=TEXT if anzahl[i] else GEDAEMPFT, width=22, anchor="e").pack(side="left")
            if wichtig[i]:
                ctk.CTkLabel(zeile, text=f"  {WICHTIG} {wichtig[i]}", font=schrift(12, True),
                             text_color=ROT, anchor="w").pack(side="left")
            if ganztags[i]:
                ctk.CTkLabel(zeile, text=f"  +{ganztags[i]} ganztägig", font=schrift(11),
                             text_color=GEDAEMPFT, anchor="w").pack(side="left")

        legende = ctk.CTkFrame(self.inhalt, fg_color="transparent")
        legende.pack(fill="x", pady=(12, 0))
        for farbe, text in ((AKZENT, "Termine"), (ROT, "mit wichtigen Personen")):
            ctk.CTkFrame(legende, fg_color=farbe, width=10, height=10, corner_radius=3).pack(side="left", padx=(4, 6))
            ctk.CTkLabel(legende, text=text, font=schrift(12), text_color=GEDAEMPFT, height=16) \
                .pack(side="left", padx=(0, 14))
        self.fertig()


class ErinnerungsFenster(ctk.CTkToplevel):
    """Bildschirmfüllend, immer im Vordergrund, mit Ton – und einem Pinguin, der durchs Bild watschelt.
    Lässt sich nicht per X oder Alt+F4 schließen, nur über die Knöpfe."""

    TEXTFARBE = "#0F172A"
    # (Hintergrund, Hover für „später“, Akzent, Pinguin-Tempo px/Bild, Tonabstand ms, Ton, Sperrzeit s)
    STIL_NORMAL = ("#E0F2FE", "#BAE6FD", "#0F172A", 7, 2000, "SystemExclamation", 3)
    STIL_WICHTIG = ("#FFE4E6", "#FECDD3", "#DC2626", 11, 1000, "SystemHand", 6)
    BILDTAKT = 40  # ms pro Animationsbild (25 Bilder/s)

    def __init__(self, master, termin, schlummer_minuten, mit_ton, beim_schliessen):
        super().__init__(master)
        self.termin = termin
        self._beim_schliessen = beim_schliessen
        self._aktiv = True
        self._mit_ton = mit_ton
        self.wichtig_mit = termin.get("wichtig_mit")
        hintergrund, hover, self.akzent, self.tempo, self.tonabstand, self.tonname, self._sperre = \
            self.STIL_WICHTIG if self.wichtig_mit else self.STIL_NORMAL

        self.title(f"Termin: {termin['betreff']}")
        self.configure(fg_color=hintergrund)
        self.attributes("-fullscreen", True)
        self.attributes("-topmost", True)
        self.protocol("WM_DELETE_WINDOW", lambda: None)
        try:
            skalierung = ctk.ScalingTracker.get_window_scaling(self)
        except Exception:
            skalierung = 1
        self.breite, self.hoehe = self.winfo_screenwidth(), self.winfo_screenheight()
        textbreite = int(self.breite / skalierung * 0.7)

        # Leinwand für Schnee und Pinguin liegt unter dem Text
        self.leinwand = tk.Canvas(self, bg=hintergrund, highlightthickness=0)
        self.leinwand.place(x=0, y=0, relwidth=1, relheight=1)
        self.groesse = self.hoehe * 0.16          # Körperhöhe des Pinguins
        self.boden = self.hoehe * 0.97
        self.px, self.richtung, self.phase = -self.groesse, 1, 0.0
        self.flocken = []
        for _ in range(70):
            r = random.uniform(2, 5)
            x, y = random.uniform(0, self.breite), random.uniform(0, self.hoehe)
            flocke = self.leinwand.create_oval(x - r, y - r, x + r, y + r, fill="white", outline="")
            self.flocken.append((flocke, random.uniform(0.6, 2.2), random.uniform(-0.4, 0.4)))

        mitte = ctk.CTkFrame(self, fg_color=hintergrund, corner_radius=0)
        mitte.place(relx=0.5, rely=0.38, anchor="center")
        marke = f"  {WICHTIG}  WICHTIGER TERMIN  " if self.wichtig_mit else "  ERINNERUNG  "
        ctk.CTkLabel(mitte, text=marke, font=schrift(18, True), text_color="#FFFFFF", fg_color=self.akzent,
                     corner_radius=19, height=38).pack(pady=(0, 20))
        self.countdown = ctk.CTkLabel(mitte, font=schrift(110, True), text_color=self.TEXTFARBE)
        self.countdown.pack()
        ctk.CTkLabel(mitte, text=termin["betreff"], font=schrift(52, True), text_color=self.TEXTFARBE,
                     wraplength=textbreite).pack(pady=(4, 10))
        details = " · ".join(filter(None, [f"{termin['start']:%H:%M} – {termin['ende']:%H:%M} Uhr", termin["ort"]]))
        ctk.CTkLabel(mitte, text=details, font=schrift(26), text_color="#475569", wraplength=textbreite).pack()
        if self.wichtig_mit:
            ctk.CTkLabel(mitte, text=f"mit {', '.join(self.wichtig_mit)}", font=schrift(30, True),
                         text_color=self.akzent, wraplength=textbreite).pack(pady=(14, 0))

        leiste = ctk.CTkFrame(mitte, fg_color=hintergrund, corner_radius=0)
        leiste.pack(pady=(44, 0))
        knopf = dict(font=schrift(20, True), width=270, height=62, corner_radius=31, state="disabled")
        self.ok = ctk.CTkButton(leiste, text="Verstanden", fg_color=self.TEXTFARBE, hover_color="#1E293B",
                                text_color="#FFFFFF", text_color_disabled="#94A3B8",
                                command=lambda: self.schliessen("ok"), **knopf)
        self.spaeter = ctk.CTkButton(leiste, text=f"In {schlummer_minuten} Min. nochmal", fg_color=hintergrund,
                                     border_width=2, border_color=self.TEXTFARBE, hover_color=hover,
                                     text_color=self.TEXTFARBE, text_color_disabled="#94A3B8",
                                     command=lambda: self.schliessen("schlummern"), **knopf)
        self.ok.pack(side="left", padx=14)
        self.spaeter.pack(side="left", padx=14)

        self.animieren()
        self.sekundentakt()
        self.ton()

    # --- Animation ---

    def animieren(self):
        if not self._aktiv:
            return
        self.phase += 0.25 + self.tempo * 0.02
        self.px += self.richtung * self.tempo
        if self.px > self.breite + self.groesse:   # am Rand umdrehen und zurücklaufen
            self.richtung = -1
        elif self.px < -self.groesse:
            self.richtung = 1

        for flocke, fall, wind in self.flocken:
            self.leinwand.move(flocke, wind, fall)
            if self.leinwand.coords(flocke)[1] > self.hoehe:
                self.leinwand.move(flocke, 0, -self.hoehe - 10)

        self.pinguin_zeichnen()
        self.after(self.BILDTAKT, self.animieren)

    def _drehen(self, punkte):
        """Punkte um die Füße kippen – das ergibt das Watscheln."""
        cos, sin = math.cos(self._kipp), math.sin(self._kipp)
        ergebnis = []
        for x, y in punkte:
            dx, dy = x - self._fx, y - self._fy
            ergebnis += [self._fx + dx * cos - dy * sin, self._fy + dx * sin + dy * cos]
        return ergebnis

    def _oval(self, cx, cy, rx, ry, farbe):
        punkte = [(cx + rx * math.cos(a), cy + ry * math.sin(a)) for a in (i * math.pi / 14 for i in range(28))]
        self.leinwand.create_polygon(self._drehen(punkte), fill=farbe, outline="", smooth=True, tags="pinguin")

    def _form(self, punkte, farbe):
        self.leinwand.create_polygon(self._drehen(punkte), fill=farbe, outline="", tags="pinguin")

    def pinguin_zeichnen(self):
        c, s, d, x = self.leinwand, self.groesse, self.richtung, self.px
        c.delete("pinguin")
        schritt = math.sin(self.phase)
        y = self.boden - abs(schritt) * s * 0.04        # kleines Hüpfen bei jedem Schritt
        self._fx, self._fy, self._kipp = x, y, schritt * 0.09

        # Füße (nicht gekippt), abwechselnd angehoben
        for seite, hub in ((-1, max(0, schritt)), (1, max(0, -schritt))):
            fx, fy = x + seite * s * 0.12, y - hub * s * 0.06
            c.create_oval(fx - s * 0.11 + d * s * 0.03, fy - s * 0.035, fx + s * 0.11 + d * s * 0.03, fy + s * 0.035,
                          fill="#F59E0B", outline="", tags="pinguin")

        flatter = schritt * s * 0.05
        for seite in (-1, 1):                            # Flügel
            self._form([(x + seite * s * 0.22, y - s * 0.62), (x + seite * s * 0.37, y - s * 0.30 + seite * flatter),
                        (x + seite * s * 0.25, y - s * 0.26)], "#1E293B")
        self._oval(x, y - s * 0.46, s * 0.30, s * 0.46, "#1E293B")               # Körper
        self._oval(x + d * s * 0.05, y - s * 0.38, s * 0.21, s * 0.34, "#F8FAFC")  # Bauch
        self._oval(x + d * s * 0.20, y - s * 0.60, s * 0.035, s * 0.025, "#FDA4AF")  # Bäckchen
        for ex in (0.03, 0.17):                          # Augen mit Glanzpunkt
            self._oval(x + d * s * ex, y - s * 0.72, s * 0.055, s * 0.06, "#FFFFFF")
            self._oval(x + d * s * (ex + 0.015), y - s * 0.715, s * 0.03, s * 0.035, "#0F172A")
            self._oval(x + d * s * (ex + 0.025), y - s * 0.73, s * 0.009, s * 0.009, "#FFFFFF")
        self._form([(x + d * s * 0.25, y - s * 0.67), (x + d * s * 0.40, y - s * 0.63),
                    (x + d * s * 0.25, y - s * 0.59)], "#F59E0B")                 # Schnabel

        if self.wichtig_mit:
            self.schild(x - d * s * 0.33, y)
        else:
            self.sprechblase(x + d * s * 0.35, y - s * 1.12, "Psst … gleich ist Termin!")

    def schild(self, sx, y):
        c, s = self.leinwand, self.groesse
        c.create_line(sx, y - s * 0.30, sx, y - s * 1.10, width=max(3, s * 0.03), fill="#92400E", tags="pinguin")
        b, h = s * 0.95, s * 0.30
        c.create_rectangle(sx - b / 2, y - s * 1.10 - h, sx + b / 2, y - s * 1.10,
                           fill="#DC2626", outline="#FFFFFF", width=3, tags="pinguin")
        c.create_text(sx, y - s * 1.10 - h / 2, text=f"{WICHTIG} WICHTIG", fill="#FFFFFF",
                      font=(SCHRIFTART, -int(s * 0.14), "bold"), tags="pinguin")

    def sprechblase(self, bx, by, text):
        c, s = self.leinwand, self.groesse
        schriftgroesse = -int(s * 0.12)
        probe = c.create_text(0, 0, text=text, font=(SCHRIFTART, schriftgroesse, "bold"))
        x1, y1, x2, y2 = c.bbox(probe)
        c.delete(probe)
        b, h, r = (x2 - x1) + s * 0.25, (y2 - y1) + s * 0.16, s * 0.10
        l, o, re, u = bx - b / 2, by - h, bx + b / 2, by
        # abgerundetes Rechteck: doppelte Eckpunkte + smooth ergeben runde Ecken
        c.create_polygon(l + r, o, re - r, o, re, o, re, o + r, re, u - r, re, u, re - r, u, l + r, u,
                         l, u, l, u - r, l, o + r, l, o, fill="#FFFFFF", outline="", smooth=True, tags="pinguin")
        spitze = bx - self.richtung * s * 0.25
        c.create_polygon(spitze - s * 0.06, by - 2, spitze + s * 0.06, by - 2, spitze - self.richtung * s * 0.02, by + s * 0.12,
                         fill="#FFFFFF", outline="", tags="pinguin")
        c.create_text(bx, by - h / 2, text=text, fill=self.TEXTFARBE, font=(SCHRIFTART, schriftgroesse, "bold"),
                      tags="pinguin")

    # --- Takt, Ton, Schließen ---

    def sekundentakt(self):
        if not self._aktiv:
            return
        sekunden = (self.termin["start"] - datetime.now()).total_seconds()
        if sekunden > 60:
            text = f"in {int(-(-sekunden // 60))} Min."
        elif sekunden > 0:
            text = "gleich!"
        elif sekunden > -60:
            text = "jetzt!"
        else:
            text = f"seit {int(-sekunden // 60)} Min."
        self.countdown.configure(text=text)

        # Knöpfe erst nach der Sperrzeit klickbar, damit man nicht versehentlich wegklickt
        if self._sperre > 0:
            self.ok.configure(text=f"Verstanden ({self._sperre})")
            self._sperre -= 1
        elif self.ok.cget("state") == "disabled":
            self.ok.configure(text="Verstanden", state="normal")
            self.spaeter.configure(state="normal")

        self.lift()
        self.attributes("-topmost", True)
        self.focus_force()
        self.after(1000, self.sekundentakt)

    def ton(self):
        if not (self._aktiv and self._mit_ton):
            return
        winsound.PlaySound(self.tonname, winsound.SND_ALIAS | winsound.SND_ASYNC)
        self.after(self.tonabstand, self.ton)

    def schliessen(self, aktion):
        self._aktiv = False
        winsound.PlaySound(None, 0)
        self.destroy()
        self._beim_schliessen(self.termin, aktion)


# --- Hintergrundbetrieb ----------------------------------------------------

def status_laden() -> dict:
    try:
        with open(STATUSDATEI, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def status_speichern(status: dict) -> None:
    with open(STATUSDATEI, "w", encoding="utf-8") as f:
        json.dump(status, f, indent=2)


class ErinnerungsApp:
    def __init__(self, root, kalender, args):
        self.root = root
        self.kalender = kalender
        self.args = args
        self.status = status_laden()
        self.erledigt = set()   # quittierte Erinnerungen
        self.offen = set()      # Erinnerungen mit offenem Pop-up
        self.schlummern = {}    # Termin-ID -> nächste Erinnerung

    def tick(self):
        try:
            self.uebersichten_pruefen()
            self.ausblick_pruefen()
            self.erinnerungen_pruefen()
        except AnmeldungNoetig as fehler:
            logging.warning("Nicht angemeldet: %s", fehler)
        except requests.RequestException as fehler:  # z. B. noch kein Netz nach dem Hochfahren
            logging.warning("Kalender nicht erreichbar: %s", fehler)
        except Exception:
            logging.exception("Fehler beim Lesen des Kalenders")
        self.root.after(self.args.intervall * 1000, self.tick)

    def uebersichten_pruefen(self):
        jetzt = datetime.now()
        if jetzt.time() < self.args.uebersicht:
            return
        heute = jetzt.date()

        fenster = []
        if self.status.get("wochenuebersicht") != heute.isoformat():
            fenster.append(WochenFenster(self.root, *wochendaten(self.kalender, heute)))
            self.status["wochenuebersicht"] = heute.isoformat()
            status_speichern(self.status)
            logging.info("Wochenübersicht angezeigt")

        if self.status.get("tagesuebersicht") != heute.isoformat():
            fenster.append(TagesFenster(self.root, heute, tagesdaten(self.kalender, heute)))
            self.status["tagesuebersicht"] = heute.isoformat()
            status_speichern(self.status)
            logging.info("Tagesübersicht angezeigt")

        if len(fenster) > 1:
            nebeneinander(*fenster)

    def ausblick_pruefen(self):
        jetzt = datetime.now()
        heute = jetzt.date().isoformat()
        if jetzt.time() < self.args.vorschau or self.status.get("ausblick") == heute:
            return
        VorschauFenster(self.root, jetzt.date(), vorschaudaten(self.kalender, jetzt.date(), self.args.vorschau_tage))
        self.status["ausblick"] = heute
        status_speichern(self.status)
        logging.info("Ausblick angezeigt")

    def erinnerungen_pruefen(self):
        jetzt = datetime.now()
        for t in self.kalender.termine(jetzt, jetzt + timedelta(hours=24)):
            if t["ganztaegig"] or t["id"] in self.erledigt or t["id"] in self.offen:
                continue
            if t["id"] in self.schlummern:
                faellig, grenze = self.schlummern[t["id"]], t["ende"]
            else:
                vorlauf = t["vorlauf"] if t["vorlauf"] is not None else self.args.minuten
                # auch kurz nach Beginn noch erinnern, z. B. wenn der Laptop erst um 9:12 aufgeklappt wird
                nachholen = min(t["start"] + timedelta(minutes=self.args.nachholen), t["ende"])
                faellig, grenze = t["start"] - timedelta(minutes=vorlauf), nachholen
            if faellig <= jetzt < grenze:
                self.offen.add(t["id"])
                logging.info("Erinnerung: %s (%s)", t["betreff"], f"{t['start']:%d.%m. %H:%M}")
                ErinnerungsFenster(self.root, t, self.args.schlummern, not self.args.leise,
                                   self.erinnerung_beendet)

    def erinnerung_beendet(self, termin, aktion):
        self.offen.discard(termin["id"])
        if aktion == "schlummern":
            self.schlummern[termin["id"]] = datetime.now() + timedelta(minutes=self.args.schlummern)
        else:
            self.erledigt.add(termin["id"])
            self.schlummern.pop(termin["id"], None)
        logging.info("Erinnerung %s: %s", aktion, termin["betreff"])


def schon_gestartet() -> bool:
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    global _mutex  # Handle muss bis Programmende offen bleiben
    _mutex = kernel32.CreateMutexW(None, False, "Local\\OutlookReminderEinzelinstanz")
    return ctypes.get_last_error() == 183  # ERROR_ALREADY_EXISTS


def main() -> None:
    parser = argparse.ArgumentParser(description="Erinnerungen und Übersichten für Outlook-Termine")
    parser.add_argument("--minuten", type=int, default=15,
                        help="Vorlauf, falls im Termin keine Erinnerung gesetzt ist (Standard: 15)")
    parser.add_argument("--schlummern", type=int, default=5, help="Schlummerzeit in Minuten (Standard: 5)")
    parser.add_argument("--nachholen", type=int, default=15,
                        help="Bis so viele Minuten nach Terminbeginn noch erinnern (Standard: 15)")
    parser.add_argument("--uebersicht", default="08:55",
                        help="Ab dieser Uhrzeit Tages-/Wochenübersicht zeigen (Standard: 08:55)")
    parser.add_argument("--vorschau", default="16:00",
                        help="Ab dieser Uhrzeit Ausblick auf die nächsten Tage zeigen (Standard: 16:00)")
    parser.add_argument("--vorschau-tage", type=int, default=3,
                        help="Anzahl Werktage im Ausblick (Standard: 3)")
    parser.add_argument("--intervall", type=int, default=60, help="Prüfintervall in Sekunden (Standard: 60)")
    parser.add_argument("--leise", action="store_true", help="Erinnerungen ohne Ton")
    parser.add_argument("--anmelden", action="store_true", help="Im Browser bei Microsoft 365 anmelden")
    parser.add_argument("--liste", action="store_true", help="Übersichten in der Konsole ausgeben")
    parser.add_argument("--test-erinnerung", action="store_true", help="Beispiel-Pop-up anzeigen")
    parser.add_argument("--test-wichtig", action="store_true", help="Beispiel-Pop-up für einen wichtigen Termin")
    parser.add_argument("--test-tag", action="store_true", help="Tagesübersicht sofort anzeigen")
    parser.add_argument("--test-woche", action="store_true", help="Wochenübersicht sofort anzeigen")
    parser.add_argument("--test-ausblick", action="store_true", help="Ausblick auf die nächsten Tage sofort anzeigen")
    args = parser.parse_args()
    args.uebersicht = datetime.strptime(args.uebersicht, "%H:%M").time()
    args.vorschau = datetime.strptime(args.vorschau, "%H:%M").time()

    os.makedirs(DATENVERZEICHNIS, exist_ok=True)
    handler = [logging.FileHandler(LOGDATEI, encoding="utf-8")]
    if sys.stderr:  # unter pythonw gibt es keine Konsole
        handler.append(logging.StreamHandler())
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s", handlers=handler)

    heute = date.today()
    if args.test_erinnerung or args.test_wichtig:
        kalender = None  # Beispiel-Pop-up braucht keinen Kalender
    else:
        kalender = kalender_laden()
    if args.anmelden:
        if not isinstance(kalender, GraphKalender):
            raise SystemExit("--anmelden ist nur für Microsoft Graph nötig (config.json).")
        kalender.token()
        print(f"Angemeldet als {kalender.app.get_accounts()[0]['username']}")
        return
    if args.liste:
        for titel, text in (tagesuebersicht(kalender, heute), wochenuebersicht(kalender, heute),
                            vorschau_uebersicht(kalender, heute, args.vorschau_tage)):
            print(f"{titel}\n{'-' * len(titel)}\n{text}\n")
        return

    ctk.set_appearance_mode("system")  # hell/dunkel wie Windows; DPI-Skalierung übernimmt CustomTkinter
    root = ctk.CTk()
    root.withdraw()
    schriftart_waehlen(root)

    if args.test_erinnerung or args.test_wichtig:
        start = datetime.now().replace(second=0, microsecond=0) + timedelta(minutes=10)
        termin = {"id": "test", "betreff": "Testtermin", "ort": "Besprechungsraum 1",
                  "start": start, "ende": start + timedelta(hours=1),
                  "wichtig_mit": ["Wichtige Person"] if args.test_wichtig else []}
        ErinnerungsFenster(root, termin, args.schlummern, not args.leise, lambda *_: root.destroy())
    elif args.test_tag:
        TagesFenster(root, heute, tagesdaten(kalender, heute), beim_schliessen=root.destroy)
    elif args.test_woche:
        WochenFenster(root, *wochendaten(kalender, heute), beim_schliessen=root.destroy)
    elif args.test_ausblick:
        VorschauFenster(root, heute, vorschaudaten(kalender, heute, args.vorschau_tage), beim_schliessen=root.destroy)
    else:
        if schon_gestartet():
            logging.info("Läuft bereits – beende diese Instanz")
            return
        logging.info("Outlook-Erinnerung gestartet")
        root.after(0, ErinnerungsApp(root, kalender, args).tick)
    root.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        logging.exception("Unerwarteter Fehler")
        raise
