import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Trasferimento Excel Iscritti per GolfBox")

st.title("Trasferimento Excel Iscritti per GolfBox")
st.markdown("Carica il file sorgente e il file di template per ottenere il file finale con i dati formattati correttamente.")

file_sorgente = st.file_uploader("Carica file sorgente (es. GOLF ELITE uff 2.xlsx)", type=["xlsx"])
file_template = st.file_uploader("Carica file template (Import_Template.xlsx)", type=["xlsx"])

if file_sorgente and file_template:
    try:
        df_sorgente = pd.read_excel(file_sorgente, skiprows=5)
        df_sorgente = df_sorgente[df_sorgente.iloc[:, 0].notna() & df_sorgente.iloc[:, 5].notna()].copy()

        df_sorgente["Last Name"] = df_sorgente.iloc[:, 0].astype(str).str.title()
        df_sorgente["First Name"] = df_sorgente.iloc[:, 5].astype(str).str.title()
        df_sorgente["Handicap"] = df_sorgente.iloc[:, 8].astype(str).str.replace(".", ",", regex=False)
        df_sorgente["Gender"] = df_sorgente.iloc[:, 10]
        df_sorgente["Club Name"] = df_sorgente.iloc[:, 12]

        df_template = pd.read_excel(file_template)
        df_output = pd.DataFrame(columns=df_template.columns)
        df_output["Last Name"] = df_sorgente["Last Name"]
        df_output["First Name"] = df_sorgente["First Name"]
        df_output["Handicap"] = df_sorgente["Handicap"]
        df_output["Gender"] = df_sorgente["Gender"]
        df_output["Club Name"] = df_sorgente["Club Name"]

        buffer = io.BytesIO()
        df_output.to_excel(buffer, index=False)
        buffer.seek(0)

        st.success("File elaborato correttamente. Clicca qui sotto per scaricarlo.")
        st.download_button(label="📥 Scarica file Excel", data=buffer, file_name="Import_Completato_GolfBox.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Errore durante l'elaborazione: {e}")
