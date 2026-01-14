"""
@file       astrology_compatibility_calculator.py
@brief      Berechnet astrologische Transit-Scores basierend auf Geburtsdaten.
@details    Nutzt die NASA JPL Horizons API für präzise Planetenpositionen. 
            Erstellt eine Excel-Auswertung mit einem kombinierten Score-Diagramm 
            und einer mehrstufigen X-Achse für Tierkreisübergänge.
@author     A. KHOUK
@date       2024-06-04
@version    1.2
@copyright  Copyright (c) 2024
"""

# =============================================================================
# IMPORTS
# =============================================================================
import pandas as pd
import numpy as np
from astroquery.jplhorizons import Horizons
from astropy.time import Time
import warnings
from openpyxl import load_workbook
from openpyxl.chart import LineChart, Reference

# Ignoriere Warnungen der astroquery-Bibliothek für eine saubere Ausgabe
warnings.filterwarnings("ignore", category=UserWarning)

# =============================================================================
# KONFIGURATION
# =============================================================================

## @var GEBURTSDATUM
# Referenzzeitpunkt für das Radix-Horoskop (Format: YYYY-MM-DD HH:MM)
GEBURTSDATUM = '1979-06-04 12:00'

## @var JAHR
# Das Kalenderjahr, für welches die Transite berechnet werden
JAHR = 1993

## @var PLANET_IDS
# Mapping der Planeten auf die offiziellen JPL Horizons IDs
PLANET_IDS = {
    'Sonne': '10', 'Mond': '301', 'Merkur': '199', 'Venus': '299',
    'Mars': '499', 'Jupiter': '599', 'Saturn': '699'
}

## @var PRIORITY_PLANETS
# Liste von Planeten, die bei der Aspektberechnung stärker gewichtet werden
PRIORITY_PLANETS = ['Sonne', 'Venus', 'Mars']

## @var ASPECTS
# Definition der berücksichtigten Aspekte: (Winkel, Orbis, Basis-Score)
ASPECTS = [(0, 8, 100), (60, 6, 70), (90, 8, 25), (120, 8, 90), (180, 8, 40)]

## @var ZODIAC_START
# Statische Daten für den Beginn der Tierkreiszeichen (MM-DD) für die X-Achse
ZODIAC_START = {
    "Wassermann": "01-20", "Fische": "02-19", "Widder": "03-21",
    "Stier": "04-20", "Zwillinge": "05-21", "Krebs": "06-21",
    "Löwe": "07-23", "Jungfrau": "08-23", "Waage": "09-23",
    "Skorpion": "10-23", "Schütze": "11-22", "Steinbock": "12-22"
}

# =============================================================================
# FUNKTIONEN
# =============================================================================

def fetch_birth_positions(date_str):
    """
    @brief      Fragt die Planetenpositionen für das Geburtsdatum ab.
    @details    Nutzt das Julianische Datum, um über JPL Horizons die ekliptikale 
                Länge (ObsEclLon) für alle definierten Planeten zu erhalten.
    @param      date_str Das Geburtsdatum als String (YYYY-MM-DD HH:MM).
    @return     dict Ein Dictionary mit Planeten-Namen und ihren Gradzahlen (0-360).
    """
    jd = Time(date_str).jd
    pos = {}
    for name, pid in PLANET_IDS.items():
        eph = Horizons(id=pid, location='399', epochs=jd).ephemerides(quantities='31')
        pos[name] = float(eph['ObsEclLon'][0])
    return pos

def fetch_year_positions(year):
    """
    @brief      Lädt die täglichen Ephemeriden für ein ganzes Jahr.
    @details    Erzeugt einen Datumsbereich vom 01.01. bis 31.12. des Zieljahres 
                und fragt die täglichen Positionen für alle Planeten ab.
    @param      year Das Zieljahr als Integer (z. B. 1993).
    @return     tuple (dates, data): 'dates' ist ein Index von Zeitstempeln, 
                'data' ist ein Dictionary mit Listen von Gradzahlen pro Planet.
    """
    epochs = {'start': f"{year}-01-01", 'stop': f"{year}-12-31", 'step': '1d'}
    base = Horizons(id='10', location='399', epochs=epochs).ephemerides()
    dates = pd.to_datetime(base['datetime_str'])
    data = {}
    for name, pid in PLANET_IDS.items():
        eph = Horizons(id=pid, location='399', epochs=epochs).ephemerides(quantities='31')
        data[name] = eph['ObsEclLon'].filled(np.nan).astype(float)
    return dates, data

def calculate_score(birth, transit):
    """
    @brief      Berechnet die astrologische Resonanz (Score) für einen Zeitpunkt.
    @details    Vergleicht alle Geburtsplaneten mit den aktuellen Transiten. 
                Werden Aspekte innerhalb des Orbis gefunden, wird der Score 
                unter Berücksichtigung der Planeten-Priorität gewichtet.
    @param      birth Dictionary der Geburts-Planetenpositionen.
    @param      transit Dictionary der aktuellen Transit-Positionen.
    @return     float Der berechnete Score, gerundet auf zwei Dezimalstellen.
    """
    score_sum = 0
    hits = 0
    for b_name, b_lon in birth.items():
        for t_name, t_lon in transit.items():
            diff = abs(b_lon - t_lon) % 360
            if diff > 180: diff = 360 - diff
            for ang, orb, val in ASPECTS:
                if abs(diff - ang) <= orb:
                    # Gewichtung erhöhen, wenn wichtige Planeten involviert sind
                    weight = 1.0 + (0.25 if b_name in PRIORITY_PLANETS else 0) + (0.25 if t_name in PRIORITY_PLANETS else 0)
                    score_sum += val * weight
                    hits += 1
                    break
    # Logarithmische Skalierung der Trefferanzahl zur Score-Glättung
    return round((score_sum / hits) * np.log1p(hits), 2) if hits > 0 else 50.0

def get_zodiac(degree):
    """
    @brief      Bestimmt das Tierkreiszeichen für eine Gradzahl.
    @details    Mappt die ekliptikale Länge (0-360°) auf die 12 klassischen 
                Tierkreiszeichen à 30 Grad.
    @param      degree Die Position des Planeten in Grad (float).
    @return     str Der Name des entsprechenden Tierkreiszeichens.
    """
    signs = [(0, "Widder"), (30, "Stier"), (60, "Zwillinge"), (90, "Krebs"), (120, "Löwe"), (150, "Jungfrau"), (180, "Waage"), (210, "Skorpion"), (240, "Schütze"), (270, "Steinbock"), (300, "Wassermann"), (330, "Fische")]
    degree %= 360
    for i, (start, name) in enumerate(signs):
        if degree < start: return signs[i-1][1] if i > 0 else signs[0][1]
    return signs[-1][1]

# =============================================================================
# HAUPTLAUF
# =============================================================================
if __name__ == "__main__":
    """
    Steuerung des gesamten Analyseprozesses:
    1. Datenabruf
    2. Score-Berechnung pro Tag
    3. Datenaufbereitung in Pandas
    4. Excel-Export mit Diagramm-Generierung
    """
    # Daten für Geburt und Transitjahr laden
    birth_pos = fetch_birth_positions(GEBURTSDATUM)
    dates, year_data = fetch_year_positions(JAHR)

    scores, mars_signs, venus_signs = [], [], []
    for i in range(len(dates)):
        transit_pos = {p: year_data[p][i] for p in PLANET_IDS}
        scores.append(calculate_score(birth_pos, transit_pos))
        mars_signs.append(f"{get_zodiac(year_data['Mars'][i])}")
        venus_signs.append(f"{get_zodiac(year_data['Venus'][i])}")

    # DataFrame Erstellung
    df = pd.DataFrame({
        "Datum": [d.date() for d in dates],
        "Zodiac_Name": "",  
        "Score": scores,
        "Mars": mars_signs,
        "Venus": venus_signs
    })

    # Score normalisieren auf einen Prozentbereich
    s_min, s_max = df['Score'].min(), df['Score'].max()
    df['Score_Prozent'] = ((df['Score'] - s_min) / (s_max - s_min)) * 100

    # Logik für die vertikalen Trennlinien und Beschriftung der Tierkreisübergänge
    df['Zodiac_Start'] = 0
    for sign, start_md in ZODIAC_START.items():
        current_start_date = pd.to_datetime(f"{JAHR}-{start_md}").date()
        mask = df['Datum'] == current_start_date
        if mask.any():
            idx = df.index[mask][0]
            df.at[idx, 'Zodiac_Start'] = 100
            df.at[idx, 'Zodiac_Name'] = sign 

    # Excel-Export und Diagrammkonfiguration
    excel_path = f"Astrologie_Analyse_{JAHR}_DoppelAchse.xlsx"
    df.to_excel(excel_path, index=False)

    wb = load_workbook(excel_path)
    ws = wb.active

    # Diagramm-Typ: Linien-Diagramm
    chart = LineChart()
    chart.title = f"Astrologischer Score {JAHR} (mit Tierkreis-Achse)"
    chart.y_axis.title = "Score in %"
    chart.x_axis.title = "Zeitverlauf"

    # Daten: Score (Spalte F) und Marker-Nadeln (Spalte G)
    score_ref = Reference(ws, min_col=6, min_row=1, max_row=len(df)+1)
    chart.add_data(score_ref, titles_from_data=True)

    marker_ref = Reference(ws, min_col=7, min_row=1, max_row=len(df)+1)
    chart.add_data(marker_ref, titles_from_data=True)

    # Mehrstufige X-Achse (Kombination aus Datum und Tierkreis-Name)
    categories = Reference(ws, min_col=1, max_col=2, min_row=2, max_row=len(df)+1)
    chart.set_categories(categories)

    # Optische Gestaltung der vertikalen Linien (Serie 2)
    chart.series[1].graphicalProperties.line.solidFill = "FF0000" # Rot
    chart.series[1].graphicalProperties.line.dashStyle = "sysDash" # Gestrichelt

    # Diagramm im Excel-Sheet positionieren
    ws.add_chart(chart, "I5")
    wb.save(excel_path)

    print(f"Fertig! Die Datei '{excel_path}' wurde erfolgreich erstellt.")

# =============================================================================
# ENDE
# =============================================================================
