import pandas as pd
import re
from datetime import datetime
import locale
import sys
import os

def parse_german_date(date_str):
    """Parst deutsche Datumsangaben wie 'Dienstag, 2. September 2025'"""
    try:
        # Entfernt Wochentag und Komma
        clean_date = re.sub(r'^[^,]+,\s*', '', date_str.strip())
        
        # Deutsche Monatsnamen zu Zahlen
        month_map = {
            'Januar': '01', 'Februar': '02', 'März': '03', 'April': '04',
            'Mai': '05', 'Juni': '06', 'Juli': '07', 'August': '08',
            'September': '09', 'Oktober': '10', 'November': '11', 'Dezember': '12'
        }
        
        # Parst "2. September 2025"
        parts = clean_date.split()
        if len(parts) == 3:
            day = parts[0].rstrip('.')
            month_name = parts[1]
            year = parts[2]
            
            if month_name in month_map:
                month = month_map[month_name]
                # Formatiert zu MM/DD/YYYY für Google Calendar
                return f"{month}/{day.zfill(2)}/{year}"
    except Exception as e:
        print(f"Fehler beim Parsen des Datums '{date_str}': {e}")
    
    return None

def parse_time_range(time_str):
    """Parst Zeitbereiche wie '13:15-16:45' und gibt Start- und Endzeit zurück"""
    try:
        if '-' in time_str:
            start_time, end_time = time_str.split('-')
            start_time = start_time.strip()
            end_time = end_time.strip()
            
            # Konvertiert zu 12-Stunden-Format mit AM/PM
            def convert_to_12h(time_24h):
                try:
                    time_obj = datetime.strptime(time_24h, '%H:%M')
                    return time_obj.strftime('%I:%M %p')
                except:
                    return time_24h
            
            return convert_to_12h(start_time), convert_to_12h(end_time)
    except Exception as e:
        print(f"Fehler beim Parsen der Zeit '{time_str}': {e}")
    
    return None, None

def convert_excel_to_google_calendar():
    """Konvertiert die Excel-Datei zu Google Calendar CSV"""
    
    # Liest die Excel-Datei
    try:
        df = pd.read_excel('termine.xlsx')
        print(f"Datei erfolgreich gelesen. {len(df)} Zeilen gefunden.")
    except FileNotFoundError:
        print("Fehler: termine.xlsx nicht gefunden!")
        return
    except Exception as e:
        print(f"Fehler beim Lesen der Excel-Datei: {e}")
        return
    
    # Initialisiert leere Listen für Google Calendar Format
    google_calendar_data = []
    
    for index, row in df.iterrows():
        try:
            # Parst deutsches Datum
            start_date = parse_german_date(str(row['Datum']))
            if not start_date:
                print(f"Zeile {index+1}: Datum konnte nicht geparst werden: {row['Datum']}")
                continue
            
            # Parst Zeitbereich
            start_time, end_time = parse_time_range(str(row['Zeit']))
            if not start_time:
                print(f"Zeile {index+1}: Zeit konnte nicht geparst werden: {row['Zeit']}")
                continue
            
            # Erstellt Subject aus Modul und Dozierender
            subject = str(row['Modul'])
            if pd.notna(row['Dozierender']) and str(row['Dozierender']).strip():
                subject += f" - {row['Dozierender']}"
            
            # Fügt Beschreibung hinzu falls vorhanden
            description = ""
            if pd.notna(row.get('Unnamed: 5')) and str(row['Unnamed: 5']).strip():
                description = str(row['Unnamed: 5'])
            
            # Erstellt Google Calendar Eintrag
            calendar_entry = {
                'Subject': subject,
                'Start Date': start_date,
                'Start Time': start_time,
                'End Date': start_date,  # Gleicher Tag
                'End Time': end_time,
                'All Day Event': 'False',
                'Description': description,
                'Location': '',
                'Private': 'False'
            }
            
            google_calendar_data.append(calendar_entry)
            
        except Exception as e:
            print(f"Fehler bei Zeile {index+1}: {e}")
            continue
    
    # Erstellt DataFrame für Google Calendar
    if google_calendar_data:
        google_df = pd.DataFrame(google_calendar_data)
        
        # Speichert als CSV
        output_file = 'kalender.csv'
        google_df.to_csv(output_file, index=False, encoding='utf-8')
        print(f"\nKonvertierung erfolgreich!")
        print(f"Ausgabedatei: {output_file}")
        print(f"Anzahl konvertierte Termine: {len(google_calendar_data)}")
        
        # Zeigt erste paar Zeilen zur Kontrolle
        print("\nErste 3 konvertierte Einträge:")
        print(google_df.head(3).to_string())
        
    else:
        print("Keine Termine konnten konvertiert werden!")

if __name__ == "__main__":
    convert_excel_to_google_calendar()
