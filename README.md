# PV-Mieterstrom: Modellierung und Optimierung wirtschaftlicher Systemkonfigurationen

## Projektüberblick

Dieses Repository enthält ein Modell zur Analyse und Optimierung von Photovoltaik-Mieterstromprojekten im urbanen Raum.

Kernthese: Die Wirtschaftlichkeit von Mieterstromprojekten hängt stärker von der Verbraucherstruktur als von der eingesetzten Anlagentechnik ab.

## Projektstruktur

Das Projekt besteht aus zwei zentralen Skripten:

* **PV-Simulation**: Simulation der PV-Stromerzeugung (15-Minuten-Auflösung, PVGIS-Daten)
* **Optimierungsmodell**: Wirtschaftliche Bewertung und Optimierung der Teilnehmerkonfiguration

## Voraussetzungen

### Julia-Packages

```julia
using Pkg
Pkg.add([
    "PythonCall",
    "Dates",
    "DataFrames",
    "XLSX",
    "Statistics",
    "Plots",
    "StatsPlots",
    "CSV",
    "Printf",
    "FinanceCore"
])
```

### Python-Abhängigkeiten

```bash
pip install pvlib pandas numpy
```

## Nutzung

### 1. PV-Simulation ausführen

Zuerst wird der PV-Stromertrag simuliert.

Wichtige Parameter:

```julia
projekt = "Leipzig"        # Projektname bzw. Ortsbezeichnung
lat, lon = 51.340, 12.374  # Standortkoordinaten
tz = "Europe/Berlin"       # Zeitzone
sunYear = 2020             # Jahr der solaren Einstrahlung

tilt = 30.0                # Neigung der PV-Anlage
azim = 180.0               # Ausrichtung der PV-Anlage
P_dc_W = 20_000.0          # Nennleistung der PV-Anlage in Wpeak
```

Output: PV-Ertragszeitreihe (Input für Optimierung)

---

### 2. Optimierungsmodell ausführen

Auf Basis der PV-Daten wird die optimale Teilnehmerkonfiguration bestimmt.

Wichtige Parameter:

```julia
projekt = "Leipzig"      # Projektname bzw. Ortsbezeichnung
sunYear = "2020"         # Jahr der solaren Einstrahlung
kWp = 20                 # Nennleistung der PV-Anlage in kWpeak
laufzeit = 20            # Laufzeit des Projekts
```

Das Modell:

* variiert Teilnehmerstrukturen (MIE, ALLG, GEW, EMOB)
* berechnet Energieflüsse (Eigenverbrauch, Einspeisung, Netzbezug)
* bewertet Wirtschaftlichkeit
* bestimmt das Gewinnmaximum

---

## Output

* Optimale Teilnehmerkonfiguration
* Jahresgewinn
* Amortisationszeit
* IRR
* LCOE
* Vergleich zur Volleinspeisung
* Excel-Export und Visualisierungen

---

## Kontext

Das Modell wurde im Rahmen einer Masterarbeit zur Wirtschaftlichkeit von PV-Mieterstromprojekten entwickelt.
