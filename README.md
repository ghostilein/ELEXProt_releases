# ELEXProt – Elektrische Prüfprotokolle aus PADF/PADFX

Windows-Desktop-Tool zum automatischen Erstellen von Prüfprotokollen (Excel `.xlsx`) aus PADF/PADFX-Messdateien von **BENNING** und **Metrel**-Prüfgeräten.

Die neueste Version steht unter [Releases](../../releases) zum Download bereit.

> Quellcode im privaten Repository.

---

## Was macht ELEXProt?

ELEXProt liest eine PADF- oder PADFX-Messdatei ein und erzeugt pro Unterverteiler (UV) eine fertige Excel-Datei mit zwei bis drei Arbeitsblättern:

- **Deckblatt** – standardisierter Prüfbericht (DIN VDE, mit Checklisten, Prüferunterschrift, Auftraggeber-Daten etc.)
- **Prüfprotokoll** – tabellarische Übersicht aller Stromkreise mit Messwerten, Ampel-Bewertung und Statistik
- **Bemerkungen** *(bei Mängeln)* – Mängelliste im Klartext plus Foto-Platzhalter für Mangel-Dokumentation

---

## Features

### Einlesen & Parsen
- Unterstützt `.padf` (XML) und `.padfx` (ZIP-Archiv)
- Erkennt alle gängigen Sicherungscharakteristiken: **B, C, D, G, K, Z, L, U, MA, gG, gL, NV, NH**
- Liest BENNING-eigene Custom-Charakteristiken korrekt aus (z. B. „Custom G" → Charakteristik `G` via Eigenschaft P430)
- Unterstützt FI-Gruppen mit verschachtelten Unterkreisen
- Einspeisung wird separat als eigene Zeile ausgewiesen

### Prüfprotokoll
- 19 Hauptspalten: Pos., Beschreibung, Querschnitt, Charakteristik, Nennstrom, ZS/Ik, ZL/Ik, RPE, RISO, FI-Auslösestrom, FI-Auslösezeit, Berührspannung, Phasen-/Leiterzahl, Messdaten-Zähler u.a.
- **Live-Ampel-Formeln** direkt in Excel: Korrekturen in der ZS/Ik- oder ZL/Ik-Zelle kippen die Ampel sofort, ohne Re-Export
- **Statistik-Block** mit COUNTIF-Formeln: Anzahl Kreise je Leiterzahl, Messabdeckung in %
- **Einzelmessungs-Block** für ERROR-Zeilen: listet jeden Einzelmesswert mit ZS/Ik und ZL/Ik auf, damit Ausreißer (schlechter Kontakt) leicht erkennbar sind
- Grenzwertprüfung nach VDE:
  - Abschaltbedingung (Ik-basiert): Ik gemessen ≥ Ik_min für die jeweilige Charakteristik + Nennstrom
  - RPE > 2 Ω = Fehler, > 1 Ω = Warnung
  - RISO < 0,5 MΩ = Fehler, < 1 MΩ = Warnung
  - FI: Auslösestrom und -zeit nach VDE 0664 (Typen AC/F/A/B)
  - Berührspannung Uc > 50 V = Fehler, > 25 V = Warnung
- Druckbereich auf die 19 Hauptspalten begrenzt; Statistik/Ampel-Blöcke sind nicht-druckend

### Deckblatt
- Vollständig ausgefüllter Prüfbericht: Auftraggeber, Auftragnehmer, Anlagenadresse, Messdatum, Prüfgeräte, Prüfernamen
- 4-Tab-Formular im Programm zum Ausfüllen (Allgemein, Checklisten, Messgeräte, Bilder)
- Echtzeit-Vorschau (QPainter-Zoom 0,75× bis 3,0×) beim Ausfüllen
- Unterschriften als Bilddatei einbettbar
- **Pro UV individuell überschreibbar** oder gemeinsamer Standard für alle UVs
- Automatisches Setzen der Mängel-Checkbox bei erkannten Grenzwertverletzungen

### Bemerkungsblatt
- Wird automatisch erzeugt, sobald Grenzwertverletzungen vorliegen (unabhängig von der Deckblatt-Checkbox)
- Mängeltext im Klartext: Messwert, Grenzwert und Konsequenz ausgeschrieben (z. B. *„Kurzschlussstrom Ik (L-PE) = 68 A < 80 A (Mindest-Ik für B16 A) – Abschaltbedingung nicht erfüllt"*)
- Langer Text bricht auf gleichhohe Folgezeilen um (kein Quetschen)
- Foto-Boxen starten erst nach dem Mängeltext-Block (kein visueller Zusammenhang zwischen Auto-Text und Fotos)
- Foto-Platzhalter für manuell hinzugefügte Mängelfotos (werden automatisch eingepasst)

### Arbeitsstand & Datenbank
- **SQLite-Datenbank** (`%APPDATA%\ELEXProt\padf_archiv.db`) speichert alle UVs und Stromkreise
- **Lebender Arbeitsstand** (`work_circuits`): ISO-Regeln, Testdaten, RPE-Ergänzungen und manuelle Korrekturen werden pro UV in der DB gespeichert – der Export liest daraus, nicht aus der PADF
- **UV-Reset**: löscht den Arbeitsstand der UV und liest circuits + work_circuits frisch aus der PADF ein (Parser läuft neu → kein Stale-Cache); Fallback auf DB-Cache wenn PADF nicht mehr erreichbar
- Gesperrte UVs (als „fertig" markiert) bleiben vom Reset unberührt

### Schnell-Werkzeuge (Ctrl+I)
- **ISO-Schnellregeln**: Regex-basierte Massenzuweisung von RISO-Werten (z. B. „Bel\*→>200", „Steckd\*→>999") – schreibt nur in leere Felder; Reset stellt exakt die berührten Felder auf PADF-Baseline zurück
- **Testdaten**: füllt leere ZS/Ik- und ZL/Ik-Felder mit plausiblen Testwerten
- **RPE-Ergänzen**: berechnet fehlende RPE-Werte aus ZL/2 (Untergrenze: Einspeisung RPE)

### Excel-Rückimport
- Manuell bearbeitete `.xlsx`-Dateien können zurückgelesen werden (Button „Excel zurücklesen")
- Liest Prüfprotokoll-Daten *und* Deckblatt-Felder zurück in den DB-Arbeitsstand
- Ermöglicht z. B. Korrektur eines Ausreißer-Messwertes direkt in Excel, dann Re-Import

### Protokoll-Archiv
- Jeder Export wird als unveränderlicher Snapshot gespeichert
- Archiv-Fenster: dreispaltige Übersicht (Dateien | UVs | Stromkreise + Export-Verlauf)
- **Re-Export aus Snapshot**: vergangene Exporte können 1:1 reproduziert werden, auch wenn sich der Arbeitsstand oder die Konfiguration inzwischen geändert hat
- SHA-256-Deduplizierung der PADF-Quelldateien

### PDF-Export
- Direkt aus dem Programm: Excel-COM (PowerShell) → PDF
- Alle erzeugten `.xlsx` auf einmal, druckoptimiert (A4, Seitenränder, Drucktitel in Zeile 1–4)

### Manuelle UVs
- UVs ohne PADF-Quelldatei manuell anlegen und Stromkreise erfassen
- ZS/Ik ↔ ZL/Ik werden aus 230 V und dem eingegebenen Widerstandswert automatisch berechnet

### Update-Check
- Prüft beim Start im Hintergrund ob eine neuere Version auf GitHub verfügbar ist

---

## Systemvoraussetzungen

- Windows 10/11
- Keine Installation nötig – portable `.exe`
- Microsoft Excel optional (nur für PDF-Export via COM benötigt)

---

## Changelog (Auswahl)

| Version | Änderungen |
|---------|-----------|
| **2.4.0** | UV-Reset-Button, Archiv-Highlighting, Fensterbreite optimiert |
| **2.3.x** | BENNING Custom-Charakteristik (P430), automatische Mängelwarnungen im UV-Listing, Bemerkungsblatt bei Grenzwertverletzungen, Klartext-Fehlermeldungen, UV-Reset parst frisch aus PADF |
| **2.2.x** | Protokoll-Archiv (SQLite-Snapshots), Re-Export, Excel-Rückimport |
| **2.1.x** | DB-Arbeitsstand, ISO-Schnellregeln, Testdaten, RPE-Ergänzen |
| **2.0.x** | PySide6-GUI, Deckblatt-Vorschau, Foto-Integration, Bemerkungsblatt |
