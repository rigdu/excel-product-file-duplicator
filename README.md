# Excel Product File Duplicator

A Tkinter GUI app to batch duplicate Excel files based on a lookup sheet.  
It copies (and renames) master Excel files for each product code in a lookup list.

## Features

- Select a product lookup Excel file (must have columns: **Product Code**, **Product Name**)
- Select a folder with master files (`Product Name.xlsx`)
- Select a destination folder
- One-click duplication and renaming with a log of all actions

## Requirements

- Python 3.8 or higher
- pandas
- openpyxl

Install requirements:
```bash
pip install -r requirements.txt
```

## Usage

```bash
python excel_duplicator_app.py
```

---

**The app will open a GUI window for you to select files and folders.**
