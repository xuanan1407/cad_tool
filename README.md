
# CAD Column Inspector Pro - Manual Polygon Selection Tool

## Overview

CAD Column Inspector Pro is a lightweight AutoLISP-based tool for manually selecting and analyzing polygons in AutoCAD. Users click points to define polygon boundaries, and the tool automatically calculates area, centroid, and exports data to Excel via Python.

## System Architecture

This project consists of two main components:

### 1. **ColumnInspector.lsp** (AutoLISP Plugin)
The AutoCAD plugin that runs inside AutoCAD and handles:
- Manual point selection by user clicks
- Real-time polygon preview with PLINE
- Area calculation using Shoelace formula
- Centroid calculation
- JSON data export with multiple polygon support
- Automatic Python processor invocation

### 2. **python_processor.py** (Python Backend)
The Python script that processes exported JSON data:
- Loads polygon data from JSON (supports both single and array formats)
- Displays polygon statistics in console
- Exports all polygons to formatted Excel file with:
  - Individual polygon sections with headers
  - Property tables (type, area, centroid, etc.)
  - Complete point coordinate listings

## Key Features

### Multi-Polygon Support
- Select unlimited polygons in one session
- Each new polygon appends to existing data
- JSON array format preserves all selections
- Clear all data with `COLCLEAR` command

### Accurate Calculations
- **Shoelace formula** for polygon area
- **Centroid** calculation from vertices
- Handles arbitrary polygon shapes (3+ points)
- Preserves 3D coordinates (X, Y, Z)

### Automated Workflow
- AutoLISP → JSON → Python → Excel (fully automated)
- Auto-detects Python installation
- Fallback to manual processing if Python not found

## File Structure

```
cad_tool/
├── ColumnInspector.lsp      # AutoCAD plugin (main)
├── python_processor.py       # Python backend processor
├── column_data.json          # Polygon data storage (auto-generated)
├── requirements.txt          # Python dependencies
├── README.md                 # This file
└── PLUGIN_INSTALLATION.md    # Installation guide
```

## Installation

### Prerequisites
- AutoCAD (any version with AutoLISP support)
- Python 3.10+ (recommended: 3.12.10)

### Python Dependencies

Install required packages:

```bash
pip install pandas openpyxl
```

Or use requirements.txt:

```bash
pip install -r requirements.txt
```

### AutoLISP Plugin Setup

1. Copy `ColumnInspector.lsp` to a location accessible by AutoCAD
2. In AutoCAD, load the plugin:
   ```lisp
   (load "C:/path/to/ColumnInspector.lsp")
   ```
3. Or set up auto-loading in `acad.lsp` or `acaddoc.lsp`

See [PLUGIN_INSTALLATION.md](PLUGIN_INSTALLATION.md) for detailed instructions.

## Usage

### Step 1: Start Polygon Selection

In AutoCAD command line, type:
```
COLINSPECT
```

### Step 2: Select Points

Click points to define polygon boundary:
- Click at each corner/vertex of the polygon
- Press **Enter** when finished selecting points
- Minimum 3 points required

Example output:
```
=== CAD Column Inspector Pro ===
Select points to create polygon (Press Enter to finish)
Select first point: 
Point 1 selected: (100.0 50.0 0.0)
Point 2 selected: (200.0 50.0 0.0)
Point 3 selected: (200.0 150.0 0.0)
Point 4 selected: (100.0 150.0 0.0)

Polygon created with 4 points
Area: 10000.00 sq units
Centroid: (150.0 100.0)
```

### Step 3: Confirm Export

Choose Yes to export data:
```
Export this polygon data? [Yes/No] <Yes>: Yes
✓ Polygon #1 saved to: C:\path\to\column_data.json
Total polygons: 1
```

### Step 4: Select More Polygons (Optional)

Run `COLINSPECT` again to add more polygons:
- Each new polygon appends to the JSON file
- All polygons saved in single array format
- Track total count after each selection

### Step 5: Export to Excel

The Python processor runs automatically and creates an Excel file:
```
Column_Inspector_YYYYMMDD_HHMMSS.xlsx
```

Located in the same folder as `column_data.json`.

### Clear All Data

To start fresh:
```
COLCLEAR
```

Confirms deletion and removes all saved polygon data.

## Commands Reference

| Command | Description |
|---------|-------------|
| `COLINSPECT` | Start polygon selection mode |
| `COLCLEAR` | Clear all saved polygons and start fresh |

## JSON Data Format

The tool exports polygon data in JSON array format:

```json
[
  {
    "type": "polygon",
    "point_count": 4,
    "area": 10000.0,
    "centroid": [150.0, 100.0],
    "points": [
      [100.0, 50.0, 0.0],
      [200.0, 50.0, 0.0],
      [200.0, 150.0, 0.0],
      [100.0, 150.0, 0.0]
    ],
    "timestamp": "20260518",
    "drawing": "Drawing1.dwg"
  },
  {
    "type": "polygon",
    "point_count": 5,
    ...
  }
]
```

## Excel Output Format

The generated Excel file contains:

### Polygon Sections
Each polygon has its own section with:
- **Property table**: Type, point count, area, centroid coordinates, drawing info
- **Points table**: Complete list of X, Y, Z coordinates for all vertices

### Example:
```
POLYGON #1
Property    | Value
------------|--------
Type        | polygon
Point Count | 4
Area        | 10000.00
Centroid X  | 150.00
Centroid Y  | 100.00
Drawing     | Drawing1.dwg

Points:
Point # | X      | Y      | Z
--------|--------|--------|----
1       | 100.00 | 50.00  | 0.00
2       | 200.00 | 50.00  | 0.00
...
```

## Technical Details

### Area Calculation
Uses the **Shoelace formula** (Surveyor's formula):
```
Area = |Σ(x_i * y_{i+1} - x_{i+1} * y_i)| / 2
```

### Centroid Calculation
Simple arithmetic mean:
```
Centroid_X = Σ(x_i) / n
Centroid_Y = Σ(y_i) / n
```

### Python Auto-Detection
Searches for Python in:
1. `C:\Python312\python.exe`
2. `C:\Python311\python.exe`
3. `C:\Python310\python.exe`
4. `C:\Program Files\Python312\python.exe`
5. `C:\Program Files\Python311\python.exe`
6. System PATH

## Troubleshooting

### "Python not found"
- Install Python 3.10+
- Add Python to system PATH
- Or manually run: `python python_processor.py column_data.json`

### "No polygon data file found"
- Run `COLINSPECT` at least once to create data
- Check drawing folder for `column_data.json`

### "Need at least 3 points"
- Polygons require minimum 3 vertices
- Press Enter only after selecting 3+ points

### Python processor doesn't start
- Manually run: `python python_processor.py "path\to\column_data.json"`
- Check `python_processor.py` is in same folder or add to PATH

## License

Copyright (c) 2026 Tran Xuan An  
Licensed under the MIT License.

## Version

Current Version: 2.0 (Manual Polygon Selection System)  
Date: May 18, 2026