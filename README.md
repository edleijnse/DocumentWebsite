# Hans Glanzmann Website-Dokumentation & Extraktion

Dieses Projekt dient der automatisierten Analyse, Extraktion und Katalogisierung eines vollständigen lokalen Backups der Weebly-Webpräsenz von **Hans Glanzmann**. Es durchläuft alle Seiten, Unterseiten und Medienbestände, um eine strukturierte und formatierte Dokumentation als Excel-Arbeitsmappe (`.xlsx`) zu generieren.

---

## Zweck & Nutzen

- **Vollständige Archivierung**: Sicherung und strukturierte Erfassung aller Inhalte, Werktitel, Abbildungen und Verlinkungen.
- **Katalogisierung des Werkbestands**: Automatisches Parsen von Werkangaben (Titel, Technik, Maße, Entstehungsjahr, Katalognummern).
- **Prüfung der Datenintegrität**: Abgleich referenzierter Bilder und Medien mit den tatsächlich lokal vorhandenen Dateien (inkl. Größenangaben und Fehlererkennung).
- **Effiziente Weiterverarbeitung**: Bereitstellung der Daten in 5 strukturierten Tabellenblättern mit Auto-Filtern, fixierten Kopfzeilen und optimierten Spaltenbreiten für Forschung, Archivierung oder Web-Relaunch.

---

## Projektstruktur

```text
DocumentWebsite/
├── main.py                                      # Hauptskript für Parsing & Excel-Generierung
├── README.md                                    # Projektdokumentation
├── Website_Dokumentation_Hans_Glanzmann.xlsx    # Generierte Excel-Dokumentation (lokale Kopie)
└── .venv/                                       # Python Virtual Environment
```

---

## Struktur der generierten Excel-Dokumentation

Die generierte Datei `Website_Dokumentation_Hans_Glanzmann.xlsx` enthält 5 spezialisierte Arbeitsblätter:

| Blattname | Beschreibung & Inhalt |
| :--- | :--- |
| **`Seitenübersicht & Struktur`** | Hierarchische Erfassung aller 35 Seiten (Ebenen 0 bis 2 sowie Direktlinks), Menütitel, Pfade (`Breadcrumbs`), Dateinamen, Meta-Angaben und Inhaltszusammenfassungen. |
| **`Inhalte & Texte`** | Vollständige Extraktion aller 189 redaktionellen Text- und Inhaltsblöcke inkl. Typisierung (Überschrift, Absatz, Zitat) und Wortanzahl. |
| **`Kunstwerke & Bildkatalog`** | Detaillierter Katalog aller 1.618 Kunstwerke und Abbildungen mit automatisiert extrahierten Metadaten (Titel, Technik, Maße in cm, Jahr, Nummer) und lokalem Dateistatus. |
| **`Medien & Verlinkungen`** | Inventar aller 1.704 internen Verlinkungen, externen Links, Kontakt-E-Mails und Mediendateien inkl. Prüfung der lokalen Erreichbarkeit. |
| **`Zusammenfassung & Statistik`** | Kompaktes Dashboard mit Kennzahlen zur Website, Verteilung der Menüebenen, Medienverfügbarkeit und Größenverteilungen. |

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
```

### Ausführung

Das Skript kann direkt über Python gestartet werden:

```powershell
python main.py
```

Nach der Ausführung wird die Excel-Dokumentation sowohl im primären Sicherungsverzeichnis als auch im Projektordner abgelegt.
