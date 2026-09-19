# Hans Glanzmann Website-Dokumentation & Extraktion

Dieses Projekt dient der automatisierten Analyse, Extraktion und Katalogisierung eines vollständigen lokalen Backups der Weebly-Webpräsenz von **Hans Glanzmann**. Es durchläuft alle Seiten, Unterseiten und Medienbestände, um eine strukturierte und formatierte Dokumentation sowohl als **Excel-Arbeitsmappe (`.xlsx`)** als auch als **interaktive HTML-Dokumentation (`.html`)** im identischen visuellen Design zu generieren.

---

## Zweck & Nutzen

- **Vollständige Archivierung**: Sicherung und strukturierte Erfassung aller Inhalte, Werktitel, Abbildungen und Verlinkungen.
- **Katalogisierung des Werkbestands**: Automatisches Parsen von Werkangaben (Titel, Technik, Maße, Entstehungsjahr, Katalognummern).
- **Interaktive Bild- & Werkgalerie**: Sofortige visuelle Vorschau aller Kunstwerke im Browser mit Lightbox-Zoom, Filterung nach Maltechniken und Wechsel zwischen Tabellen- und Galerie-Kartenansicht.
- **Prüfung der Datenintegrität**: Abgleich referenzierter Bilder und Medien mit den tatsächlich lokal vorhandenen Dateien (inkl. Größenangaben und Fehlererkennung).
- **Effiziente Weiterverarbeitung**: Bereitstellung der Daten in 5 strukturierten Tabellenblättern / Tabs für Forschung, Archivierung oder Web-Relaunch.

---

## Projektstruktur

```text
DocumentWebsite/
├── main.py                                      # Hauptskript für Parsing, Excel- & HTML-Generierung
├── README.md                                    # Projektdokumentation
├── Website_Dokumentation_Hans_Glanzmann.xlsx    # Generierte Excel-Dokumentation (lokale Kopie)
├── Website_Dokumentation_Hans_Glanzmann.html    # Generierte interaktive HTML-Dokumentation (lokale Kopie)
└── .venv/                                       # Python Virtual Environment
```

---

## Dokumentationsformate

### 1. Interaktive HTML-Dokumentation (`Website_Dokumentation_Hans_Glanzmann.html`)

Die HTML-Dokumentation spiegelt das Design der Excel-Arbeitsmappe (`#1F497D` Header, Segoe UI Typografie, strukturierte Zebrastreifen-Tabellen) wider und erweitert diese um webbasierte Interaktivität:

- **Tabs-Navigation**: Direktes Umschalten zwischen allen 5 Dokumentationsbereichen.
- **Echtzeit-Suche & Filter**: Dynamisches Filtern von Tabellenzeilen nach Begriffen, Hierarchie-Ebenen, Texttypen und Maltechniken ohne Neuladen der Seite.
- **Spaltensortierung**: Klickbare Tabellenköpfe zum auf- und absteigenden Sortieren aller Spalten.
- **Bildvorschau & Lightbox**: Miniaturvorschauen für alle 1.618 Kunstwerke mit Klick-Vergrößerung im Modal-Dialog inklusive vollständiger Werksdaten.
- **Galerie-Kartenansicht**: Umschalten zwischen tabellarischer Listenansicht und responsiver Bildkarten-Galerie.
- **Direktverlinkung**: Anklickbare Verweise zum direkten Öffnen der lokalen Original-HTML-Webseiten und Vollbilddateien.

### 2. Strukturierte Excel-Mappe (`Website_Dokumentation_Hans_Glanzmann.xlsx`)

Enthält 5 spezialisierte Arbeitsblätter mit Auto-Filtern und fixierten Kopfzeilen:

| Blattname / Tab | Beschreibung & Inhalt |
| :--- | :--- |
| **`Seitenübersicht & Struktur`** | Hierarchische Erfassung aller 35 Seiten (Ebenen 0 bis 2 sowie Direktlinks), Menütitel, Pfade (`Breadcrumbs`), Dateinamen, Meta-Angaben und Inhaltszusammenfassungen. |
| **`Inhalte & Texte`** | Vollständige Extraktion aller 189 redaktionellen Text- und Inhaltsblöcke inkl. Typisierung (Überschrift, Absatz, Zitat) und Wortanzahl. |
| **`Kunstwerke & Bildkatalog`** | Detaillierter Katalog aller 1.618 Kunstwerke und Abbildungen mit automatisiert extrahierten Metadaten (Titel, Technik, Maße in cm, Jahr, Nummer) und lokalem Dateistatus. |
| **`Medien & Verlinkungen`** | Inventar aller 1.704 internen Verlinkungen, externen Links, Kontakt-E-Mails und Mediendateien inkl. Prüfung der lokalen Erreichbarkeit. |
| **`Zusammenfassung & Statistik`** | Kompaktes Dashboard mit Kennzahlen zur Website, Verteilung der Menüebenen, Medienverfügbarkeit und Rubrikenübersichten. |

---

## Technische Voraussetzungen

- **Python**: Version 3.10 oder neuer
- **Abhängigkeiten**:
  - `beautifulsoup4` (inkl. `lxml`-Parser)
  - `openpyxl`

Installation der benötigten Pakete:

```powershell
pip install beautifulsoup4 lxml openpyxl
```

---

## Konfiguration & Ausführung

### Pfadeinstellungen in `main.py`

In `main.py` sind die Quell- und Zielpfade definiert:

```python
BASE_PATH = r'D:\Hans_Glanzmann\436167332711627763-1789638831\7897076666aabb50cb36a1'
INDEX_FILE = os.path.join(BASE_PATH, 'index.html')

OUTPUT_EXCEL_PRIMARY = r'D:\Hans_Glanzmann\Website_Dokumentation_Hans_Glanzmann.xlsx'
OUTPUT_EXCEL_LOCAL = r'C:\Users\edlei\PycharmProjects\DocumentWebsite\Website_Dokumentation_Hans_Glanzmann.xlsx'

OUTPUT_HTML_PRIMARY = r'D:\Hans_Glanzmann\Website_Dokumentation_Hans_Glanzmann.html'
OUTPUT_HTML_BASEDIR = os.path.join(BASE_PATH, 'Website_Dokumentation_Hans_Glanzmann.html')
OUTPUT_HTML_LOCAL = r'C:\Users\edlei\PycharmProjects\DocumentWebsite\Website_Dokumentation_Hans_Glanzmann.html'
```

### Ausführung

Das Skript kann direkt gestartet werden:

```powershell
python main.py
```

Nach der Ausführung werden sowohl die Excel-Arbeitsmappe als auch die HTML-Dokumentation in `D:\Hans_Glanzmann` sowie lokal im Projektverzeichnis gespeichert.
