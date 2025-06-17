
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
2. Open terminal in this folder
3. Run the desired command
4. Import the CSV into Google Calendar

## Available Commands:

### Basic Conversion

```bash
python e2gc.py your_appointments.csv
```

### With Specific Output File

```bash
python e2gc.py appointments.xlsx -o calendar.csv
```

### Excel with Specific Worksheet

```bash
python e2gc.py workbook.xlsx -s "Appointments"
```
