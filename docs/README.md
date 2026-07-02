# WebApp — Heidelberg Haushaltsdaten-Browser

Klickbare Single-HTML-Sicht auf den Excel-Master `Daten\01_HD_Haushaltsdaten.xlsx` (Stufe 1).

## Öffnen

Doppelklick auf `index.html` (oder im Browser-Tab öffnen).
Beim ersten Aufruf wird Chart.js vom CDN geladen — danach Browser-Cache.

Falls der Browser `file://` blockt (selten), lokalen Server starten:

```bash
cd WebApp
python -m http.server 8765
# http://localhost:8765/index.html
```

## Inhalt

**Hauptsicht (Dashboard):** 3 Kernaussagen aus Stufe 1 als Cards plus 3 Charts:

1. Ord. Ergebnis 2017–2029 (Line)
2. Schuldenstand 2018–2029 (Line)
3. Top-10 TH nach Defizit 2025 (Horizontal-Bar, TH FW gefiltert)

**7 Sheet-Sichten** (alle voll befüllten Sheets aus dem Excel-Master):

| Sheet | Inhalt |
|---|---|
| 00 · README | Lesehinweise (Key-Value) |
| 01 · Eckdaten | Gesamthaushalt 2017–2029, Filter (Jahr / Kennzahl / Plan_Typ) + Line-Chart |
| 02 · Teilhaushalte | 41 TH × 7 Kennzahlen × 2 Jahre, Filter (Jahr / Kennzahl) + TH-Bar-Chart |
| 08 · Schulden | Schulden-Trend 2018–2029 + Line-Chart |
| 09 · Kennzahlen | Übergreifende Kennzahlen + Hinweise (PDF-Rotation) |
| 13 · Mapping TH⇄PB | Akkordeon pro Produktbereich → aufklappbare PG-Listen |
| 14 · Stammdaten TH | Brüche im TH-Schnitt 2019–2025, Freitext-Suche |

## TH-FW-Toggle

Header rechts oben. TH FW (Allgemeine Finanzwirtschaft) ist standardmäßig **ausgeblendet** — zentrale Steuer-/Umlagenposition verzerrt sonst alle TH-Rankings (FW = Überschuss, alle anderen Defizit).

## Regenerieren

Nach Änderung am Excel-Master:

```bash
python WebApp/build_webapp.py
```

Liest `Daten\01_HD_Haushaltsdaten.xlsx`, transformiert 7 Sheets zu JSON und schreibt das komplette `index.html` neu (~165 KB).

## Tech

- **Single File:** Daten als JSON inline, kein Server, keine Build-Pipeline.
- **Chart.js v4** über CDN.
- **Vanilla JS ES6**, keine Frameworks.
- **Quellen:** Q1 = Haushaltsplan 2019/20 · Q2 = 2021/22 · Q3 = 2023/24 · Q4 = 2025/26 (Hauptquelle) · Q5 = Nachtragshaushalt 2026.

## Was bewusst NICHT in der App ist

- Skelett-Sheets (03 Produktbereiche, 04 Produktgruppen, 05 Produkte, 06 Investitionen, 07 Personal, 10 Eigenbetriebe, 11 Vergleichsstädte, 12 Hebel-Hypothesen) — kommt in Stufe 2/3.
- Edit-Funktion — App ist Read-only.
- Backend / Datenbank — vollständig statisch.
