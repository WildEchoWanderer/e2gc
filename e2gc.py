import pandas as pd
import re
from datetime import datetime, date, time
import sys
import os # Importiert das os Modul, um Dateipfade zu manipulieren
from icalendar import Calendar, Event # Importiert das icalendar Modul

print(r"""

        ____          
    ___|___ \ __ _  ___ 
   / _ \ __) / _` |/ __|
  |  __// __/ (_| | (__ 
   \___|_____\__, |\___|
               |___/    
""")

def parse_german_date(date_str):
    """
    Parst deutsche Datumsangaben wie 'Dienstag, 2. September 2025' und gibt ein datetime.date Objekt zurück.
    Behandelt die Konvertierung von deutschen Monatsnamen zu Zahlen.
    """
    try:
        # Entfernt Wochentag und Komma (z.B. "Dienstag, ")
        clean_date = re.sub(r'^[^,]+,\s*', '', date_str.strip())
        
        # Mapping von deutschen Monatsnamen zu ihren numerischen Werten
        month_map = {
            'Januar': 1, 'Februar': 2, 'März': 3, 'April': 4,
            'Mai': 5, 'Juni': 6, 'Juli': 7, 'August': 8,
            'September': 9, 'Oktober': 10, 'November': 11, 'Dezember': 12
        }
        
        # Zerlegt den Datumsstring (z.B. "2. September 2025") in Tag, Monat und Jahr
        parts = clean_date.split()
        if len(parts) == 3:
            # Extrahiert den Tag und entfernt den Punkt (z.B. "2.")
            day = int(parts[0].rstrip('.'))
            month_name = parts[1]
            year = int(parts[2])
            
            # Überprüft, ob der Monatsname im Mapping existiert
            if month_name in month_map:
                month = month_map[month_name]
                return date(year, month, day) # Gibt ein datetime.date Objekt zurück
    except Exception as e:
        print(f"Fehler beim Parsen des Datums '{date_str}': {e}")
    
    return None

def parse_time_range(time_str):
    """
    Parst Zeitbereiche wie '13:15-16:45' und gibt Start- und Endzeit als datetime.time Objekte zurück.
    Verwendet das 24-Stunden-Format für das Parsen.
    """
    try:
        if '-' in time_str:
            start_time_str, end_time_str = time_str.split('-')
            start_time_str = start_time_str.strip()
            end_time_str = end_time_str.strip()
            
            # Konvertiert die Zeitstrings in datetime.time Objekte
            start_t_obj = datetime.strptime(start_time_str, '%H:%M').time()
            end_t_obj = datetime.strptime(end_time_str, '%H:%M').time()
            
            return start_t_obj, end_t_obj
    except Exception as e:
        print(f"Fehler beim Parsen der Zeit '{time_str}': {e}")
    
    return None, None

def process_excel_to_events(input_file): # Parameter für den Dateinamen hinzugefügt
    """
    Liest die angegebene Excel-Datei und verarbeitet jede Zeile zu einem Event-Objekt (als Dictionary).
    Gibt eine Liste dieser Event-Objekte zurück.
    """
    
    # Versucht, die Excel-Datei zu lesen
    try:
        df = pd.read_excel(input_file) # Verwendet den übergebenen Dateinamen
        print(f"Datei '{input_file}' erfolgreich gelesen. {len(df)} Zeilen gefunden.")
    except FileNotFoundError:
        print(f"Fehler: '{input_file}' nicht gefunden! Bitte stellen Sie sicher, dass die Datei im selben Verzeichnis liegt.")
        return [] # Gibt eine leere Liste zurück, wenn die Datei nicht gefunden wird
    except Exception as e:
        print(f"Fehler beim Lesen der Excel-Datei '{input_file}': {e}")
        return [] # Gibt eine leere Liste zurück bei anderen Fehlern
    
    # Initialisiert eine leere Liste, um die verarbeiteten Ereignisse zu speichern
    processed_events = []
    
    # Iteriert über jede Zeile im DataFrame
    for index, row in df.iterrows():
        try:
            # Parst das Datum aus der 'Datum'-Spalte
            event_date = parse_german_date(str(row['Datum']))
            if not event_date:
                print(f"Zeile {index+1}: Datum konnte nicht geparst werden: '{row['Datum']}'. Überspringe Zeile.")
                continue # Springt zur nächsten Zeile, wenn das Datum nicht geparst werden kann
            
            # Parst den Zeitbereich aus der 'Zeit'-Spalte
            start_time_obj, end_time_obj = parse_time_range(str(row['Zeit']))
            if not start_time_obj or not end_time_obj:
                print(f"Zeile {index+1}: Zeit konnte nicht geparst werden: '{row['Zeit']}'. Überspringe Zeile.")
                continue # Springt zur nächsten Zeile, wenn die Zeit nicht geparst werden kann
            
            # Kombiniert das Datum-Objekt und die Zeit-Objekte zu vollständigen datetime-Objekten
            start_datetime = datetime.combine(event_date, start_time_obj)
            end_datetime = datetime.combine(event_date, end_time_obj)
            
            # Erstellt den Betreff (Subject) des Kalendereintrags aus 'Modul' und 'Dozierender'
            subject = str(row['Modul'])
            # Überprüft, ob die 'Dozierender'-Spalte existiert und nicht leer ist
            if pd.notna(row.get('Dozierender')) and str(row['Dozierender']).strip():
                subject += f" - {row['Dozierender']}"
            
            # Fügt eine Beschreibung hinzu, falls die 'Unnamed: 5'-Spalte existiert und Inhalt hat
            description = ""
            if 'Unnamed: 5' in row and pd.notna(row['Unnamed: 5']) and str(row['Unnamed: 5']).strip():
                description = str(row['Unnamed: 5'])
            
            # Erstellt ein Dictionary für das aktuelle Ereignis mit allen benötigten Informationen
            event_data = {
                'Subject': subject,
                'StartDateTime': start_datetime,
                'EndDateTime': end_datetime,
                'Description': description,
                'Location': '', # Platzhalter für den Ort, kann bei Bedarf aus einer Spalte gelesen werden
                'Private': False # Standardwert für privaten Status, kann angepasst werden
            }
            
            processed_events.append(event_data) # Fügt das Ereignis zur Liste hinzu
            
        except KeyError as ke:
            print(f"Fehler bei Zeile {index+1}: Spalte '{ke}' nicht gefunden. Bitte stellen Sie sicher, dass die Spalten 'Datum', 'Zeit', 'Modul', 'Dozierender' und ggf. 'Unnamed: 5' vorhanden sind.")
            continue # Springt zur nächsten Zeile bei fehlender Spalte
        except Exception as e:
            print(f"Unbekannter Fehler bei Zeile {index+1}: {e}. Überspringe Zeile.")
            continue # Springt zur nächsten Zeile bei anderen Ausnahmen
            
    # Gibt eine Zusammenfassung der Verarbeitung aus
    if not processed_events:
        print("Es wurden keine Termine erfolgreich verarbeitet!")
    else:
        print(f"Erfolgreich {len(processed_events)} Termine verarbeitet.")
    return processed_events

def export_to_csv(events_data, output_base_name): # Parameter für den Basis-Ausgabedateinamen hinzugefügt
    """
    Exportiert die verarbeiteten Event-Daten in eine Google Calendar kompatible CSV-Datei.
    Formatiert Datum und Uhrzeit gemäß den Anforderungen von Google Calendar.
    Der Dateiname basiert auf dem übergebenen output_base_name.
    """
    if not events_data:
        print("Keine Daten zum Exportieren als CSV.")
        return
        
    google_calendar_data = []
    for event in events_data:
        # Formatiert Start- und Enddatum im MM/DD/YYYY-Format für Google Calendar CSV
        start_date_str = event['StartDateTime'].strftime('%m/%d/%Y')
        end_date_str = event['EndDateTime'].strftime('%m/%d/%Y')
        # Formatiert Start- und Endzeit im 12-Stunden-Format mit AM/PM für Google Calendar CSV
        start_time_str = event['StartDateTime'].strftime('%I:%M %p')
        end_time_str = event['EndDateTime'].strftime('%I:%M %p')

        # Erstellt das Dictionary im Google Calendar CSV-Format
        calendar_entry = {
            'Subject': event['Subject'],
            'Start Date': start_date_str,
            'Start Time': start_time_str,
            'End Date': end_date_str,
            'End Time': end_time_str,
            'All Day Event': 'False', # Standardmäßig keine Ganztagesveranstaltung
            'Description': event['Description'],
            'Location': event['Location'],
            'Private': 'False' # Standardmäßig nicht privat
        }
        google_calendar_data.append(calendar_entry)

    # Erstellt einen Pandas DataFrame aus den vorbereiteten Daten
    google_df = pd.DataFrame(google_calendar_data)
    output_file = f"{output_base_name}.csv" # Erstellt den CSV-Dateinamen
    # Speichert den DataFrame als CSV-Datei
    google_df.to_csv(output_file, index=False, encoding='utf-8')
    print(f"\nKonvertierung erfolgreich! Ausgabedatei: {output_file}")
    print(f"Anzahl konvertierte Termine: {len(google_calendar_data)}")
    
    # Zeigt die ersten 3 konvertierten Einträge zur Überprüfung an
    print("\nErste 3 konvertierte CSV-Einträge:")
    print(google_df.head(3).to_string())

def export_to_ics(events_data, output_base_name): # Parameter für den Basis-Ausgabedateinamen hinzugefügt
    """
    Exportiert die verarbeiteten Event-Daten in eine ICS-Datei (iCalendar-Format).
    Der Dateiname basiert auf dem übergebenen output_base_name.
    """
    if not events_data:
        print("Keine Daten zum Exportieren als ICS.")
        return
        
    cal = Calendar()
    # Fügt grundlegende Kalenderinformationen hinzu
    cal.add('prodid', '-//My Calendar Product//mxm.dk//') # Produkt-ID
    cal.add('version', '2.0') # iCalendar-Version

    # Iteriert über jedes Event-Dictionary und erstellt einen VEVENT-Eintrag
    for event_data in events_data:
        event = Event()
        event.add('summary', event_data['Subject']) # Betreff/Zusammenfassung des Ereignisses
        event.add('dtstart', event_data['StartDateTime']) # Startzeitpunkt (datetime-Objekt)
        event.add('dtend', event_data['EndDateTime']) # Endzeitpunkt (datetime-Objekt)
        
        # Fügt Beschreibung und Ort hinzu, falls vorhanden
        if event_data['Description']:
            event.add('description', event_data['Description'])
        if event_data['Location']:
            event.add('location', event_data['Location'])
        
        cal.add_component(event) # Fügt das Ereignis dem Kalender hinzu

    output_file = f"{output_base_name}.ics" # Erstellt den ICS-Dateinamen
    # Schreibt den Kalenderinhalt in eine .ics-Datei im Binärmodus
    with open(output_file, 'wb') as f:
        f.write(cal.to_ical())
    
    print(f"\nKonvertierung erfolgreich! Ausgabedatei: {output_file}")
    print(f"Anzahl konvertierte Termine: {len(events_data)}")
    print("\nICS-Datei erfolgreich erstellt.")

def main():
    """
    Hauptfunktion des Programms, die den Ablauf steuert:
    - Begrüßung und Hinweis
    - Verarbeitung der Excel-Daten
    - Abfrage des gewünschten Ausgabeformats vom Benutzer
    - Aufruf der entsprechenden Exportfunktion
    """
    print("Willkommen beim Excel zu Kalender Konverter!")
    
    # Überprüft, ob ein Dateiname als Kommandozeilenargument übergeben wurde
    if len(sys.argv) > 1:
        input_file = sys.argv[1] # Verwendet das erste Argument als Eingabedatei
        print(f"Verwende die angegebene Eingabedatei: '{input_file}'")
    else:
        input_file = 'termine.xlsx' # Standard-Dateiname, wenn kein Argument übergeben wird
        print(f"Keine Eingabedatei angegeben. Verwende Standarddatei: '{input_file}'")

    # Extrahiert den Basisnamen der Datei ohne Erweiterung für die Ausgabedateien
    output_base_name = os.path.splitext(input_file)[0]
    
    # Verarbeitet die Excel-Datei und holt die Event-Daten
    events = process_excel_to_events(input_file) # Übergibt den Eingabedateinamen
    
    # Beendet das Programm, wenn keine Termine verarbeitet werden konnten
    if not events:
        print("Es wurden keine Termine erfolgreich verarbeitet. Beende Programm.")
        return

    # Schleife zur Abfrage des Ausgabeformats, bis eine gültige Eingabe erfolgt
    while True:
        print("\nBitte wählen Sie das gewünschte Ausgabeformat:")
        print("1: CSV (für Google Calendar Import)")
        print("2: ICS (Standard Kalenderformat, kompatibel mit den meisten Kalendern)")
        choice = input("Ihre Wahl (1 oder 2): ").strip()

        if choice == '1':
            export_to_csv(events, output_base_name) # Übergibt den Basisnamen für die Ausgabe
            break # Beendet die Schleife nach dem Export
        elif choice == '2':
            export_to_ics(events, output_base_name) # Übergibt den Basisnamen für die Ausgabe
            break # Beendet die Schleife nach dem Export
        else:
            print("Ungültige Eingabe. Bitte geben Sie '1' oder '2' ein.")

if __name__ == "__main__":
    # Überprüft, ob das 'icalendar'-Modul installiert ist, bevor das Programm startet
    try:
        import icalendar
    except ImportError:
        print("Fehler: Das 'icalendar' Modul ist nicht installiert.")
        print("Bitte installieren Sie es mit: pip install icalendar")
        sys.exit(1) # Beendet das Programm mit einem Fehlercode
        
    main() # Startet die Hauptfunktion des Programms
