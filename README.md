<img width="1920" height="1058" alt="pic" src="https://github.com/user-attachments/assets/2a265c95-4a23-4501-88a1-c243de0312e7" />

# Warehouse Tracker

A lightweight desktop app for a storage/warehouse department to track product
locations over time.

Every day the department exports an inventory Excel file. This app imports
that file into a local database, automatically detects when a product's
storage location has changed, and keeps a full history of every location a
product has ever been stored in — so staff can always search a product by
name or code and see both its current location and its full location
history.

## Features

- **Drag & drop or browse** to upload the daily Excel file, all on one screen
- **Automatic change detection** — only creates a new history entry when a
  product's location actually changes, so the history stays meaningful
  instead of duplicating the same location every day
- **Case-insensitive search** by product name or part/technical code
- **Full location history** per product, newest first, with the date and
  source file of each change
- **Upload log** showing every import that's been run, with counts of new
  products and location changes
- **Persian UI** with a modern navy/blue and white design
- **Runs as a single `.exe`** — no Python installation needed on the
  department's computer
- All data stored locally in a single SQLite file (`warehouse.db`) — no
  internet connection or external server required

## Tech Stack

- Python 3
- Tkinter (GUI, built into Python)
- tkinterdnd2 (drag-and-drop file support)
- SQLite (local storage, built into Python)
- openpyxl (reading `.xlsx` files)
- PyInstaller (packaging into a standalone `.exe`)

## Project Structure

```
warehouse_app/
├── app.py            # Main GUI application (single-page layout)
├── db.py             # Database layer (SQLite)
├── importer.py       # Excel import + change-detection logic
├── fonts.py          # Font resolution helper (IRANSansX with fallback)
├── requirements.txt  # Python dependencies
├── build.bat          # Windows script to build the .exe
├── sample_data/       # Small synthetic sample files for demo purposes
└── README.md
```

## Font Note

The UI is designed for **IRANSansX**. Since it's a licensed commercial font,
it isn't bundled in this repo. Install it on the machine running the app to
get the intended look (Windows: right-click the font file → Install). If it's
not installed, the app automatically falls back to Tahoma, which also
supports Persian text well and comes pre-installed on Windows.

## Expected Excel Format

The importer looks for these column headers (matching the department's
export format):

| Column              | Meaning                          |
|---------------------|-----------------------------------|
| `نام_کالا`          | Product name                     |
| `شماره_فني`         | Part/technical code (unique key) |
| `مدل_به_کاررفته`    | Model(s) the part is used in     |
| `آدرس_در_انبار`     | Warehouse location                |

The part code is used as the unique key. If a code already exists in the
database and its location in the new file differs from the stored location,
a new history entry is created automatically.

## Running from Source

```bash
pip install -r requirements.txt
python app.py
```

## Building the Standalone .exe

On a Windows machine, inside the project folder:

```bash
build.bat
```

This installs dependencies and runs PyInstaller. The finished app will be at
`dist/WarehouseTracker.exe` — a single file that can be copied anywhere and
run without installing Python.

The database file (`warehouse.db`) is created automatically next to the
`.exe` the first time it's run, and persists between runs.

## Trying it out with sample data

Since real warehouse data isn't included in this repo, two small synthetic
sample files are provided in `sample_data/`:

1. Run the app and upload `sample_data/sample_day1.xlsx`
2. Upload `sample_data/sample_day2.xlsx` (a couple of items have moved)
3. Search for e.g. "Brake Pad" and see the location history update

## Notes

- The daily Excel files and the `warehouse.db` database are excluded from
  version control (see `.gitignore`) since they contain real company data.
