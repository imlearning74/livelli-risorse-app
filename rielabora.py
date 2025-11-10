import streamlit as st
import pandas as pd
from collections import defaultdict
from io import BytesIO

st.set_page_config(page_title="Trasformazione Livelli Risorse", layout="centered")
st.title("📊 Trasformazione Livelli Risorse")

uploaded_file = st.file_uploader("Carica un file Excel", type=["xlsx"])

if uploaded_file:
    try:
        # Leggi i fogli disponibili
        xls = pd.ExcelFile(uploaded_file)
        sheet_name = st.selectbox("Seleziona il foglio da elaborare", xls.sheet_names)

        # Carica il foglio selezionato
        df = pd.read_excel(xls, sheet_name=sheet_name)

        # Verifica che le colonne necessarie siano presenti
        required_columns = ['Persona', 'Cognome', 'Nome', 'Divisione', 'Unità organizzativa']
        missing_columns = [col for col in required_columns if col not in df.columns]

        if missing_columns:
            st.error(f"❌ Il foglio selezionato non contiene le colonne richieste: {', '.join(missing_columns)}")
        else:
            # Trasformazione dei dati
            person_data = defaultdict(list)
            for _, row in df.iterrows():
                person_id = row['Persona']
                unita_org = row['Unità organizzativa']
                person_data[person_id].append({
                    'Cognome': row['Cognome'],
                    'Nome': row['Nome'],
                    'Divisione': row['Divisione'],
                    'Unita_Organizzative': [unita_org]
                })

            new_rows = []
            for person_id, records in person_data.items():
                base_record = records[0]
                all_units = []
                for record in records:
                    all_units.extend(record['Unita_Organizzative'])

                unique_units = []
                for unit in all_units:
                    if unit not in unique_units:
                        unique_units.append(unit)

                new_record = {
                    'Persona': person_id,
                    'Cognome': base_record['Cognome'],
                    'Nome': base_record['Nome'],
                    'Divisione': base_record['Divisione'],
                }

                for i in range(7):
                    new_record[f'Livello_{i+1}'] = unique_units[i] if i < len(unique_units) else ''

                new_rows.append(new_record)

            new_df = pd.DataFrame(new_rows)

            st.success("✅ Trasformazione completata!")
            st.dataframe(new_df)

            # Prepara il file per il download
            output = BytesIO()
            new_df.to_excel(output, index=False)
            st.download_button(
                label="📥 Scarica il file trasformato",
                data=output.getvalue(),
                file_name="Report_Livelli_Risorse_Riorganizzato.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"❌ Errore durante l'elaborazione del file: {e}")
