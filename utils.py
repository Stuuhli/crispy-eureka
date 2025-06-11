# utils.py

import pandas as pd
from openpyxl.utils import column_index_from_string

def spalten_index(spalte):
    """Konvertiert einen Excel-Spaltenbuchstaben in einen 0-basierten Index."""
    try:
        return column_index_from_string(spalte) - 1
    except ValueError:
        return -1 # Fallback, wenn die Spalte ungültig ist

def durchschnitt_ohne_null(werte):
    """Berechnet den Durchschnitt einer Liste, ignoriert dabei Nullen, leere Strings und Fehler."""
    bereinigte = []
    for v in werte:
        try:
            # Konvertiere zu float, falls möglich
            f = float(v)
            if f != 0:
                bereinigte.append(f)
        except (ValueError, TypeError):
            # Ignoriere Werte, die nicht in Zahlen umgewandelt werden können
            continue
    return sum(bereinigte) / len(bereinigte) if bereinigte else 0

def werte_aus_excel(df, rows, cols=None, col=None):
    """Extrahiert Werte aus einem DataFrame-Bereich und berechnet den Durchschnitt."""
    werte = []
    if cols and len(cols) == 2:
        col_start = spalten_index(cols[0])
        col_end = spalten_index(cols[1]) + 1
        if col_start != -1 and col_end != -1:
            werte = df.iloc[rows[0]:rows[1]+1, col_start:col_end].values.flatten().tolist()
    elif col:
        col_index = spalten_index(col)
        if col_index != -1:
            werte = df.iloc[rows[0]:rows[1]+1, col_index].values.flatten().tolist()
            
    return durchschnitt_ohne_null(werte)