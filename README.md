# Material Master Data (MSD) Table App

A Python-based desktop application for parsing and managing SAP Material Master Data from text files.

## Overview

The **MSD Table App** is a GUI application built with Tkinter that processes SAP Material Master text files, extracts structured data, and exports it to Excel format. It's designed to handle material information including material codes, plant assignments, material groups, and control keys.

## Features

- **Text File Parser**: Reads and parses SAP Material Master data from `.txt` files
- **BOM Detection**: Automatically detects and handles different text encodings (UTF-8, UTF-16, Latin-1)
- **Data Extraction**: Extracts the following information per material:
  - Material Code
  - Plant (Plnt)
  - Material Group (Matl grp)
  - Control Type (CTyp)
  - Control Key (Ctrl key)
  - Material Description
- **Excel Export**: Exports parsed data directly to `.xlsx` files
- **User-Friendly GUI**: Clean interface with loading indicators

## Requirements

- Python 3.7+
- pandas
- openpyxl (automatically installed with pandas)

## Installation

1. Clone or download the repository
2. Install required dependencies:
   ```bash
   pip install pandas openpyxl
   ```

## Usage

### Running the Application

```bash
python MSD_Table_App.py
```

### How to Use

1. **Select File**: Click the "Seleccionar archivo TXT" (Select TXT file) button to choose a Material Master `.txt` file
2. **Processing**: The application will parse the file and display the extracted records
3. **Export**: Click the "Exportar a Excel" (Export to Excel) button to save the parsed data as an Excel file

## File Structure

```
MSD_Table_App/
├── MSD_Table_App.py          # Main application file
├── MSD_Table_App.spec        # PyInstaller specification file
├── README.md                 # This file
├── Reporte MSL.TXT          # Sample material master data file
└── build/                   # Build artifacts (PyInstaller)
```

## Technical Details

### Parsing Logic

The parser processes text files line by line:
- **Headers**: Skips header rows (Material, Material description, dates)
- **Material Records**: Identifies material lines with tab or multiple space separators
- **Descriptions**: Captures the following line as the material description

### Supported Encodings

The application automatically detects and handles:
- UTF-16 (with BOM)
- UTF-8 (with BOM)
- Latin-1 (fallback)

## Application Architecture

- **Threading**: File operations run in background threads to keep the UI responsive
- **Loading Dialog**: Shows progress indication during file processing and export
- **Error Handling**: Comprehensive error messages for debugging

## Dependencies

| Library | Purpose |
|---------|---------|
| tkinter | GUI framework |
| pandas | Data manipulation and Excel export |
| threading | Asynchronous file operations |
| re | Regular expression parsing |
| codecs | Text encoding detection |

## Building Executable

To create a standalone executable using PyInstaller:

```bash
pyinstaller MSD_Table_App.spec
```

The compiled executable will be available in the `dist/` folder.

## Notes

- The application is designed for processing SAP Material Master exports
- Large files may take longer to process; a loading dialog will appear
- Ensure the input text file follows the expected Material Master format

## License

Internal use - Jabil Corporation

## Support

For issues or feature requests, please contact the development team.
