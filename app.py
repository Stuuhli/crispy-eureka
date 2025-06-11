# app.py

import json
import os
import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

# Importieren der ausgelagerten Verarbeitungs- und Hilfsfunktionen
from data_processing import analyse_risiken, get_asset_data, consolidate_and_map_relations

# === Hilfsfunktion zum Laden der Konfigurationsdateien ===
@st.cache_data
def lade_mapping(datei_pfad):
    """Lädt eine JSON-Mapping-Datei."""
    try:
        with open(datei_pfad, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        st.error(f"❌ Kritischer Fehler: Mapping-Datei '{datei_pfad}' nicht gefunden. Die Anwendung kann nicht starten.")
        st.stop()
    except json.JSONDecodeError:
        st.error(f"❌ Kritischer Fehler: Fehler beim Lesen der JSON-Datei '{datei_pfad}'. Bitte Syntax prüfen.")
        st.stop()

# === UI-Bereich: Risikoanalyse ===
def zeige_risiko_analyse_bereich(mapping_risikoanalyse):
    st.subheader("1. Risikoanalyse")
    hochgeladene_dateien = st.file_uploader(
        "Excel-Dateien für Risikoanalyse auswählen",
        type="xlsx",
        accept_multiple_files=True,
        key="risk_files"
    )

    if not hochgeladene_dateien:
        st.info("⬆️ Bitte eine oder mehrere Excel-Dateien für die Risikoanalyse hochladen.")
        return

    df_risiken = analyse_risiken(mapping_risikoanalyse, hochgeladene_dateien)

    if df_risiken.empty:
        st.warning("Keine passenden Daten für die Risikoanalyse in den hochgeladenen Dateien gefunden. Bitte prüfen Sie die Dateiinhalte und das `mapping_risikoanalyse.json`.")
        return

    df_risiken["Datei_kurz"] = df_risiken["Datei"].str.extract(r"KiCloud-\d+-(.*)\.xlsx")[0].fillna(df_risiken["Datei"])

    farbmodus = st.radio(
        "🎨 Farbliche Darstellung der Risikomatrix",
        options=["Nach Bedrohung", "Nach Risikowert"],
        horizontal=True,
        key="farbmodus_risiko"
    )

    st.subheader("🎯 Risikomatrix (Bubble-Darstellung)")
    auswahl_bedrohungen = st.multiselect(
        "🔍 Bedrohungen filtern",
        options=sorted(df_risiken["Bedrohung"].unique()),
        default=sorted(df_risiken["Bedrohung"].unique()),
        key="bedrohungen_risiko"
    )

    df_gefiltert = df_risiken[df_risiken["Bedrohung"].isin(auswahl_bedrohungen)]

    if df_gefiltert.empty:
        st.info("Keine Daten für die ausgewählten Filter vorhanden.")
    else:
        color_arg = "Bedrohung" if farbmodus == "Nach Bedrohung" else "∅ Risikowert"
        fig = px.scatter(
            df_gefiltert,
            x="∅ Eintrittswahrscheinlichkeit",
            y="∅ Schadenswirkung",
            size="∅ Risikowert",
            color=color_arg,
            color_continuous_scale="RdYlGn_r",
            hover_name="Datei_kurz",
            text=df_gefiltert["∅ Risikowert"].round(2).astype(str),
            size_max=40,
            title="Risikomatrix: Schadenswirkung vs. Eintrittswahrscheinlichkeit"
        )
        fig.update_traces(textposition='middle center')
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("🔎 Top Risiken")
    top_x = st.slider("Anzahl der Top-Risiken", min_value=3, max_value=50, value=10, key="top_risiken_slider")

    top_risiken = df_risiken.nlargest(top_x, "∅ Risikowert").reset_index(drop=True)
    st.dataframe(top_risiken, use_container_width=True)

    export_buffer = BytesIO()
    df_risiken.to_excel(export_buffer, index=False, sheet_name="Risikoanalyse")
    st.download_button(
        "📥 Gesamte Risikoauswertung als Excel herunterladen",
        data=export_buffer.getvalue(),
        file_name="Risikoauswertung_Ergebnisse.xlsx",
        key="download_risiko"
    )

# === UI-Bereich: Asset-Auflistung ===
def zeige_asset_auflistung_bereich(mapping_assets):
    st.subheader("2. Asset-Auflistung & Konsolidierung")
    st.markdown("⚠️ **Hinweis:** Streamlit unterstützt keine direkte Ordnerauswahl im Browser. Bitte geben Sie den vollständigen Pfad zum Ordner auf Ihrem lokalen System ein.")
    folder_path = st.text_input("Pfad zum Root-Ordner (z.B. `C:\\Users\\DeinName\\Dokumente\\Assets`)", key="asset_folder_path")

    if not folder_path:
        st.info("Bitte geben Sie einen Ordnerpfad ein, um die Asset-Analyse zu starten.")
        return

    if not os.path.isdir(folder_path):
        st.error("❌ Der angegebene Pfad ist kein gültiger Ordner oder existiert nicht.")
        return

    asset_data = get_asset_data(folder_path, mapping_assets)

    if not asset_data:
        st.warning("Keine passenden Asset-Dateien (`OCTAVE_S1.2-S2.2 *.xlsx`) im angegebenen Ordner oder dessen Unterordnern gefunden.")
        return

    # -- Einzigartige Assets konsolidieren und anzeigen --
    st.markdown("---")
    st.subheader("Einzigartige Assets pro Kategorie")
    st.markdown("Hier werden alle Assets aus allen Dateien zusammengeführt und Duplikate (basierend auf dem Asset-Namen) entfernt. Jedem einzigartigen Asset wird eine neue, fortlaufende ID zugewiesen.")

    asset_types_for_unique = ["Prozesse", "Informationen", "Anwend. & Dienste", "Systeme"]
    all_asset_dfs_by_type = {atype: [] for atype in asset_types_for_unique}

    for folder_entry in asset_data:
        for asset_type, df_asset in folder_entry.items():
            if asset_type in all_asset_dfs_by_type and not df_asset.empty:
                all_asset_dfs_by_type[asset_type].append(df_asset)

    unique_asset_dfs_for_mapping = {}
    excel_buffer_unique = BytesIO()
    with pd.ExcelWriter(excel_buffer_unique, engine='xlsxwriter') as writer:
        for asset_type, list_of_dfs in all_asset_dfs_by_type.items():
            if not list_of_dfs:
                with st.expander(f"Kategorie: {asset_type} (0 Einträge)"):
                    st.info(f"Keine Daten für '{asset_type}' gefunden.")
                continue

            combined_df = pd.concat(list_of_dfs, ignore_index=True)
            if 'Name' in combined_df.columns:
                unique_df = combined_df.drop_duplicates(subset=['Name'], keep='first').reset_index(drop=True)
            else:
                unique_df = combined_df

            prefix_map = {"Prozesse": "P", "Informationen": "I", "Anwend. & Dienste": "A", "Systeme": "S"}
            prefix = prefix_map.get(asset_type, "U")

            unique_df['Neue_ID'] = [f"{prefix}{i+1}" for i in range(len(unique_df))]
            unique_df = unique_df.rename(columns={'ID': 'Original_ID'})
            cols = ['Neue_ID'] + [col for col in unique_df.columns if col != 'Neue_ID']
            unique_df = unique_df[cols]
            unique_df = unique_df.rename(columns={'Neue_ID': 'ID'})

            unique_asset_dfs_for_mapping[asset_type] = unique_df.copy()

            with st.expander(f"Kategorie: {asset_type} ({len(unique_df)} einzigartige Einträge)"):
                st.dataframe(unique_df, use_container_width=True)

            # --- NEU: Schreibe DataFrame in Excel und füge Formatierung hinzu ---
            sheet_name = asset_type.replace(" & ", "_")[:31]
            unique_df.to_excel(writer, sheet_name=sheet_name, index=False)

            # Hole die xlsxwriter-Objekte für die Formatierung
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]

            # Definiere die Zellformate für die Farben
            green_format = workbook.add_format({'bg_color': '#C6EFCE', 'fg_color': '#006100'})
            yellow_format = workbook.add_format({'bg_color': '#FFEB9C', 'fg_color': '#9C6500'})
            orange_format = workbook.add_format({'bg_color': '#FFDDC1', 'fg_color': '#C55A11'})


            # Spalten, auf die die Formatierung angewendet werden soll
            cols_to_format = ['Vertraulichkeit', 'Integrität', 'Verfügbarkeit']
            header = unique_df.columns.tolist()
            num_rows = len(unique_df)

            for col_name in cols_to_format:
                if col_name in header:
                    col_idx = header.index(col_name)
                    # Wende die bedingte Formatierung auf die gesamte Spalte an (ohne Header)
                    worksheet.conditional_format(1, col_idx, num_rows, col_idx,
                                                 {'type': 'cell', 'criteria': '==', 'value': '"niedrig"', 'format': green_format})
                    worksheet.conditional_format(1, col_idx, num_rows, col_idx,
                                                 {'type': 'cell', 'criteria': '==', 'value': '"mittel"', 'format': yellow_format})
                    worksheet.conditional_format(1, col_idx, num_rows, col_idx,
                                                 {'type': 'cell', 'criteria': '==', 'value': '"hoch"', 'format': orange_format})

            # Passe die Spaltenbreiten für eine bessere Lesbarkeit automatisch an
            for idx, col in enumerate(unique_df.columns):
                series = unique_df[col]
                # Die +2 ist für etwas zusätzlichen Abstand
                max_len = max((series.astype(str).map(len).max(), len(str(series.name)))) + 2
                worksheet.set_column(idx, idx, max_len)
            # --- Ende des neuen Code-Abschnitts ---


    st.download_button(
        "📥 Einzigartige Assets (mit Farbmarkierung) als Excel herunterladen",
        data=excel_buffer_unique.getvalue(),
        file_name="Assets_Einzigartig_Formatiert.xlsx",
        key="download_assets_unique"
    )

    # -- Konsolidierte Relationen anzeigen --
    st.markdown("---")
    st.subheader("Konsolidierte Relationen")
    st.markdown("Hier sind alle Relationen aus allen Dokumenten, zugeordnet mit den neuen, einzigartigen Asset-IDs.")

    df_relations_excel, df_relations_display = consolidate_and_map_relations(asset_data, unique_asset_dfs_for_mapping)

    if not df_relations_display.empty:
        st.dataframe(df_relations_display, use_container_width=True)

        export_relations_buffer = BytesIO()
        df_relations_excel.to_excel(export_relations_buffer, index=False, sheet_name="Konsolidierte_Relationen")
        st.download_button(
            "📥 Konsolidierte Relationen als Excel herunterladen",
            data=export_relations_buffer.getvalue(),
            file_name="Relationen_Konsolidiert.xlsx",
            key="download_relations"
        )
    else:
        st.info("Keine Relationen in den verarbeiteten Dokumenten gefunden.")

# === Hauptanwendung ===
def main():
    """Definiert das Layout der Streamlit-Anwendung."""
    st.set_page_config(layout="wide", page_title="Risiko-Dashboard")
    st.title("📊 Risiko- & Asset-Dashboard")

    # Mappings laden, bevor die UI-Elemente angezeigt werden
    mapping_risikoanalyse = lade_mapping("mapping_risikoanalyse.json")
    mapping_assets = lade_mapping("mapping_assets.json")

    # Bereiche der Anwendung anzeigen
    zeige_risiko_analyse_bereich(mapping_risikoanalyse)
    st.markdown("---")
    zeige_asset_auflistung_bereich(mapping_assets)

if __name__ == "__main__":
    main()
