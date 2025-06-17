
```
       ____            
   ___|___ \ __ _  ___ 
  / _ \ __) / _` |/ __|
 |  __// __/ (_| | (__ 
  \___|_____\__, |\___|
            |___/      
```

# Excel-to-Google-Calendar (now with .ics support)

This file converts an Excel sheet containing ZHAW school dates in a specific format into a CSV file for import into Google Calendar or into a .ics file. The Excel File needs to be in the same format as the example file termine.xlsx .

## Prerequisite:

```bash
pip install pandas icalendar
```

## Instructions:

1. Place `e2gc.py` and the Excel file to be converted in the same folder
2. Open terminal in this folder
3. Run the desired command
4. Import the CSV into Google Calendar

## Available Commands:

### Basic Conversion (uses termine.xlsx)

```bash
python e2gc.py 
```

### With Specific input File

```bash
python e2gc.py appointments.xlsx
```
