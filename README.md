# philsoritos_textabenteuer

Kleine statische Verteilseite für eine schulische Formatierungsübung.

## Für die Schüler

**Direktdownload der Übungsdatei:**  
https://raw.githubusercontent.com/GeRe87/philsoritos_textabenteuer/main/material/mikroplastik_formatierungsuebung.docx

Die Word-Datei enthält absichtlich Formatierungsfehler. Inhaltlich soll nichts neu geschrieben werden.

## Inhalt
- `index.html` – Schülerseite mit Download
- `material/mikroplastik_formatierungsuebung.docx` – absichtlich schlecht formatierte Rohfassung
- `LEHRKRAFT.md` – Hinweise, Fehlerliste und Bewertungsvorschlag
- `tools/build_exercise.py` – reproduzierbarer Generator für die Word-Datei
- `.github/workflows/build-material.yml` – baut die Datei automatisch bei Änderungen am Generator

## GitHub Pages
Die Dateien sind für GitHub Pages vorbereitet. In **Settings → Pages** als Quelle `Deploy from a branch`, Branch `main`, Ordner `/ (root)` auswählen.

Danach ist die Schülerseite unter

`https://gere87.github.io/philsoritos_textabenteuer/`

erreichbar.

## Fachliche Quellen
Die Übung verwendet 12 reale institutionelle bzw. wissenschaftliche Quellen. Die Rohfassung mischt ihre Formatierung absichtlich; Details stehen im Dokument und in `LEHRKRAFT.md`.
