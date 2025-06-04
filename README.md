
```
       ____            
   ___|___ \ __ _  ___ 
  / _ \ __) / _` |/ __|
 |  __// __/ (_| | (__ 
  \___|_____\__, |\___|
            |___/      
```

# **E**xcel-**to**-**G**oogle-**C**alendar

This file converts an Excel sheet containing ZHAW school dates in a specific format into a CSV file for import into Google Calendar.

## Prerequisite:

```bash
pip install pandas openpyxl
```

## Instructions:

1. Place `e2gc.py` and the Excel file to be converted in the same folder
2. Run the command
3. Import the CSV into Google Calendar

## Available Commands:

### Basic Conversion

```bash
python google_calendar_converter.py your_appointments.csv
```

### With Specific Output File

```bash
python google_calendar_converter.py appointments.xlsx -o calendar.csv
```

### Excel with Specific Worksheet

```bash
python google_calendar_converter.py workbook.xlsx -s "Appointments"
```
