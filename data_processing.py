# data_processing.py

import os
import re
import pandas as pd
import streamlit as st
from io import BytesIO

# Import aus unserer neuen utils.py Datei
from utils import spalten_index, werte_aus_excel

@st.cache_data(show_spinner=False)
def analyse_risiken(mapping, dateien):
    """Analysiert hochgeladene Excel-Dateien auf Risikowerte."""
    ergebnisse = []
    for file in dateien:
        try:
            xls = pd.ExcelFile(file, engine="openpyxl")
            dateiname = file.name
            for bedrohung, daten in mapping.items():
                sheet = daten["sheet"]
                if sheet not in xls.sheet_names:
                    continue

                df_sheet = xls.parse(sheet, header=None)
                schadenswirkung = werte_aus_excel(df_sheet, **daten["schadenswirkung"])
                wahrscheinlichkeit = werte_aus_excel(df_sheet, **daten["wahrscheinlichkeit"])
                risikowert = werte_aus_excel(df_sheet, **daten["risikowert"])

                ergebnisse.append({
                    "Datei": dateiname,
                    "Bedrohung": bedrohung,
                    "Blatt": sheet,
                    "∅ Schadenswirkung": schadenswirkung,
                    "∅ Eintrittswahrscheinlichkeit": wahrscheinlichkeit,
                    "∅ Risikowert": risikowert
                })
        except Exception as e:
            st.warning(f"Konnte Datei '{file.name}' für die Risikoanalyse nicht vollständig verarbeiten: {e}")
            
    return pd.DataFrame(ergebnisse)

@st.cache_data(show_spinner=False)
def get_asset_data(_folder_path, mapping_assets):
    """Durchsucht Ordner nach Asset-Dateien und extrahiert die Daten."""
    all_extracted_data = []
    
    if not os.path.isdir(_folder_path):
        return []

    for root, _, files in os.walk(_folder_path):
        for file_name in files:
            if re.match(r"OCTAVE_S1\.2-S2\.2 .*\.xlsx", file_name):
                file_path = os.path.join(root, file_name)
                try:
                    xls = pd.ExcelFile(file_path, engine="openpyxl")
                    folder_name = os.path.basename(root)
                    extracted_data_for_file = {"Ordnername": folder_name, "Dateiname": file_name}

                    for asset_type, asset_mapping in mapping_assets.items():
                        sheet_name = asset_type
                        if sheet_name in xls.sheet_names:
                            df_sheet = xls.parse(sheet_name, header=None)
                            start_row_excel = asset_mapping[next(iter(asset_mapping))]["row"]
                            start_row_pandas = start_row_excel - 1
                            
                            relevant_cols_end_idx = spalten_index("E") if asset_type == "Relationen" else spalten_index("J")

                            last_data_row_pandas = start_row_pandas - 1
                            for r in range(start_row_pandas, df_sheet.shape[0]):
                                # Prüfe, ob in den relevanten Spalten irgendwelche Daten stehen
                                row_slice = df_sheet.iloc[r, spalten_index("A") : relevant_cols_end_idx + 1]
                                if row_slice.notna().any() and not all(str(val).strip() == "" for val in row_slice):
                                    last_data_row_pandas = r
                                else:
                                    break # Erste leere Zeile gefunden
                            
                            if last_data_row_pandas >= start_row_pandas:
                                if asset_type == "Relationen":
                                    cols_to_extract = ["Prozess", "Information", "Anwendung / Dienst", "Systemname I", "Systemname II"]
                                    extracted_df = df_sheet.iloc[start_row_pandas : last_data_row_pandas + 1, spalten_index("A") : spalten_index("E") + 1].copy()
                                else:
                                    cols_to_extract = ["ID", "Name", "Kurzbeschreibung", "Vertraulichkeit", "Integrität", "Verfügbarkeit", "Sonstiges", "Begründung", "Kommentar", "Risk Owner"]
                                    extracted_df = df_sheet.iloc[start_row_pandas : last_data_row_pandas + 1, spalten_index("A") : spalten_index("J") + 1].copy()
                                
                                extracted_df.columns = cols_to_extract
                                extracted_data_for_file[asset_type] = extracted_df
                            else:
                                extracted_data_for_file[asset_type] = pd.DataFrame() # Leer, wenn keine Daten gefunden
                        else:
                            extracted_data_for_file[asset_type] = pd.DataFrame() # Leer, wenn Blatt nicht existiert

                    all_extracted_data.append(extracted_data_for_file)

                except Exception as e:
                    st.warning(f"Konnte Datei '{file_name}' im Ordner '{root}' nicht verarbeiten: {e}")

    return all_extracted_data

def consolidate_and_map_relations(all_extracted_data, unique_asset_dfs_map):
    """Konsolidiert alle Relationen und ersetzt Asset-Namen durch neue IDs."""
    all_relations_dfs = []
    for file_data in all_extracted_data:
        if "Relationen" in file_data and not file_data["Relationen"].empty:
            all_relations_dfs.append(file_data["Relationen"])

    if not all_relations_dfs:
        return pd.DataFrame(), pd.DataFrame()

    combined_relations_df = pd.concat(all_relations_dfs, ignore_index=True).drop_duplicates().reset_index(drop=True)

    mapping_dict = {}
    for asset_type, df in unique_asset_dfs_map.items():
        if 'Name' in df.columns and 'ID' in df.columns:
            # Ignoriere leere Namen beim Erstellen des Mappings
            temp_map = df.dropna(subset=['Name'])
            mapping_dict.update(pd.Series(temp_map['ID'].values, index=temp_map['Name']).to_dict())
        else:
            st.warning(f"Spalten 'Name' oder 'ID' fehlen im einzigartigen Asset-DataFrame für '{asset_type}'.")

    mapped_relations_df = combined_relations_df.copy()
    
    id_cols = []
    original_cols = ["Prozess", "Information", "Anwendung / Dienst", "Systemname I", "Systemname II"]

    for col in original_cols:
        id_col = f"{col.replace(' / ', '_').replace(' ', '_')}_ID"
        mapped_relations_df[id_col] = mapped_relations_df[col].map(mapping_dict).fillna(mapped_relations_df[col])
        id_cols.append(id_col)
    
    excel_export_cols = id_cols + original_cols
    df_for_excel = mapped_relations_df[excel_export_cols]
    
    df_for_display = mapped_relations_df[original_cols].sort_values(by=original_cols, ascending=True).reset_index(drop=True)
    
    return df_for_excel, df_for_display