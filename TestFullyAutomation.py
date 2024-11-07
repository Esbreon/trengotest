def process_excel_file(filepath):
    """Verwerkt Excel bestand en stuurt berichten."""
    try:
        print(f"\nVerwerken Excel bestand: {filepath}")
        df = pd.read_excel(filepath)
        
        if df.empty:
            print("Geen data gevonden in Excel bestand")
            return
        
        print(f"Aantal rijen gevonden: {len(df)}")
        print(f"Kolommen in bestand: {', '.join(df.columns)}")
        
        column_mapping = {
            'Naam bewoner': 'fields.Naam bewoner',
            'Datum bezoek': 'fields.Datum bezoek',
            'Tijdvak': 'fields.Tijdvak',
            'Reparatieduur': 'fields.Reparatieduur',
            'Mobielnummer': 'fields.Mobielnummer'
        }
        
        # Verify all required columns exist
        missing_columns = [col for col in column_mapping.keys() if col not in df.columns]
        if missing_columns:
            raise ValueError(f"Missende kolommen in Excel: {', '.join(missing_columns)}")
        
        df = df.rename(columns=column_mapping)
        
        # Remove duplicates based on name and visit date
        df['fields.Datum bezoek'] = pd.to_datetime(df['fields.Datum bezoek'])
        df_unique = df.drop_duplicates(subset=['fields.Naam bewoner', 'fields.Datum bezoek'])
        
        if len(df_unique) < len(df):
            print(f"Let op: {len(df) - len(df_unique)} dubbele afspraken verwijderd")
        
        for index, row in df_unique.iterrows():
            try:
                print(f"\nVerwerken rij {index + 1}/{len(df_unique)}")
                mobielnummer = format_phone_number(row['fields.Mobielnummer'])
                if not mobielnummer:
                    print(f"Geen geldig telefoonnummer voor {row['fields.Naam bewoner']}, deze overslaan")
                    continue
                
                send_whatsapp_message(
                    naam_bewoner=row['fields.Naam bewoner'],
                    datum=row['fields.Datum bezoek'],
                    tijdvak=row['fields.Tijdvak'],
                    reparatieduur=row['fields.Reparatieduur'],
                    mobielnummer=mobielnummer
                )
                
                print(f"Bericht verstuurd voor {row['fields.Naam bewoner']}")
                
            except Exception as e:
                print(f"Fout bij verwerken rij {index}: {str(e)}")
                continue
                
    except Exception as e:
        print(f"Fout bij verwerken Excel bestand: {str(e)}")
        raise
