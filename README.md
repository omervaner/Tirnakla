# Tırnakla

A desktop utility app I built to help with common SQL and data tasks. The name "Tırnakla" is Turkish for "Quote it" (and also means "fingernail" - hence the icon).

![Tırnakla](assets/icon.png)

## Features

### SQL Quoter
I use this tab constantly when I need to convert a list of values into SQL-friendly format:
- Paste a list of values (one per line)
- Automatically wraps them in single quotes
- Adds commas between items
- Supports custom templates like `SELECT %s FROM table`
- Has a "File Mode" for processing large files

### SQL Formatter
Formats SQL queries in "river style" where keywords right-align:
```sql
    SELECT column1
         , column2
         , column3
      FROM table1
      JOIN table2
        ON table1.id = table2.id
     WHERE condition1 = 'value'
       AND condition2 = 'value'
```

### Report Converter
Converts text reports (tab-delimited, pipe-separated, fixed-width) to Excel:
- Auto-detects delimiters
- Handles files with millions of rows
- Shows preview before converting
- Splits into multiple sheets if needed (Excel's 1M row limit)

### Clipboard History
Keeps track of the last 50 clipboard items:
- Click any item to copy it back
- Persists between sessions
- Search through history

## Installation

### Download Pre-built Apps
Check the [Releases](../../releases) page for:
- **macOS**: `Tirnakla-macOS.zip`
- **Windows**: `Tirnakla-Windows.zip`

### Run from Source
```bash
# Clone the repo
git clone https://github.com/omervaner/Tirnakla.git
cd Tirnakla

# Install dependencies
pip install -r requirements.txt

# Run
python main.py
```

## Building from Source

### macOS
```bash
# Install Homebrew Python (recommended)
brew install python@3.11 python-tk@3.11

# Install dependencies
pip3.11 install customtkinter pillow openpyxl pyinstaller

# Build
python3.11 -m PyInstaller --name "Tırnakla" --onefile --windowed \
  --icon=assets/icon.png --add-data "assets:assets" \
  --collect-all customtkinter main.py
```

### Windows
```bash
pip install customtkinter pillow openpyxl pyinstaller
pyinstaller --name "Tirnakla" --onefile --windowed ^
  --icon=assets/icon.png --add-data "assets;assets" ^
  --collect-all customtkinter main.py
```

## Tech Stack
- Python 3.9+
- CustomTkinter (modern UI)
- Pillow (image handling)
- openpyxl (Excel export)

## License
MIT

## Contributors
- [omervaner](https://github.com/omervaner)
- Claude (AI pair programmer)
