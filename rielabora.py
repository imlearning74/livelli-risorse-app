import pandas as pd
from collections import defaultdict

# Carica il file Excel
df = pd.read_excel('Report Livelli risorse tot_102025.xlsx', sheet_name='ZRU_TEMPLATE')

# Raggruppa per persona e raccogli tutte le unità organizzative
person_data = defaultdict(list)

for _, row in df.iterrows():
    person_id = row['Persona']
    unita_org = row['Unità organizzativa']
    
    # Aggiungi l'unità organizzativa alla lista della persona
    person_data[person_id].append({
        'Cognome': row['Cognome'],
        'Nome': row['Nome'],
        'Divisione': row['Divisione'],
        'Unita_Organizzative': [unita_org]  # Inizializza la lista
    })

# Prepara i dati per il nuovo DataFrame
new_rows = []
for person_id, records in person_data.items():
    # Prendi il primo record come base
    base_record = records[0]
    
    # Raccogli tutte le unità organizzative uniche
    all_units = []
    for record in records:
        all_units.extend(record['Unita_Organizzative'])
    
    # Rimuovi duplicati mantenendo l'ordine
    unique_units = []
    for unit in all_units:
        if unit not in unique_units:
            unique_units.append(unit)
    
    # Crea il nuovo record
    new_record = {
        'Persona': person_id,
        'Cognome': base_record['Cognome'],
        'Nome': base_record['Nome'],
        'Divisione': base_record['Divisione'],
        'Livello_1': unique_units[0] if len(unique_units) > 0 else '',
        'Livello_2': unique_units[1] if len(unique_units) > 1 else '',
        'Livello_3': unique_units[2] if len(unique_units) > 2 else '',
        'Livello_4': unique_units[3] if len(unique_units) > 3 else '',
        'Livello_5': unique_units[4] if len(unique_units) > 4 else '',
        'Livello_6': unique_units[5] if len(unique_units) > 5 else '',
        'Livello_7': unique_units[6] if len(unique_units) > 6 else ''
    }
    
    new_rows.append(new_record)

# Crea il nuovo DataFrame
new_df = pd.DataFrame(new_rows)

# Salva il risultato in un nuovo file Excel
new_df.to_excel('Report_Livelli_Risorse_Riorganizzato tot_102025.xlsx', index=False)

print("Trasformazione completata! File salvato come 'Report_Livelli_Risorse_Riorganizzato tot_102025.xlsx'")
print(f"Numero originale di righe: {len(df)}")
print(f"Numero di righe dopo trasformazione: {len(new_df)}")