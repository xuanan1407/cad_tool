
# CAD Column Inspector Pro - Auto-Scan Tool

## Overview

CAD Column Inspector Pro is a powerful AutoLISP-based tool for analyzing shapes in AutoCAD. It offers two modes:
1. **Manual Mode** - Click points to define custom polygons
2. **Auto-Scan Mode** - Select an area and automatically detect all shapes

The tool automatically calculates area, centroid, and exports data to Excel via Python.

## System Architecture

This project consists of two main components:

### 1. **ColumnInspector.lsp** (AutoLISP Plugin)
The AutoCAD plugin that runs inside AutoCAD and handles:
- **Manual point selection** - User clicks points to create custom polygons
- **Auto-scan mode** - Automatically detects LWPOLYLINE, POLYLINE, CIRCLE shapes in selected area
- Real-time polygon preview with PLINE (manual mode)
- Area calculation (Shoelace formula for polygons, πr² for circles)
- Centroid calculation
- Automatic nearest text detection for labeling
- JSON data export with multiple shape support
- Automatic Python processor invocation

### 2. **python_processor.py** (Python Backend)
The Python script that processes exported JSON data:
- Loads shape data from JSON (supports both single and array formats)
- **Calculates perimeter** for all shapes (polygons and circles)
- Displays shape statistics in console
- Exports to **2-sheet Excel file**:
  - **Summary sheet**: Statistics grouped by name (count, total area, avg area, total perimeter, avg perimeter)
  - **Detail sheet**: Individual shape sections with full property tables and point coordinates
  - Professional formatting with colors, borders, and proper alignment

## Key Features

### Auto-Scan Mode (NEW!)
- **Select rectangular area** with 2 clicks
- **Customizable tolerance**:
  - Set minimum centroid-to-centroid distance (default 1.0mm)
  - Any shapes closer than tolerance will be removed
  - Keeps first shape, removes subsequent shapes in cluster
  - **Simple distance-only check** - no area or size comparison
- **Automatic shape detection**:
  - Closed LWPOLYLINE (lightweight polylines)
  - Closed POLYLINE (heavy polylines)
  - CIRCLE (with radius information)
- **Batch processing** - All shapes detected and saved at once
- **Smart labeling** - Finds nearest text for each shape automatically

### Manual Mode
- Custom polygon creation by clicking points
- Minimum 3 points required
- Real-time preview with PLINE
- User confirmation before saving

### Multi-Shape Support
- Process unlimited shapes in one session
- **Both COLSCAN and COLINSPECT support cumulative appending**
- Each new scan/selection adds to existing data (no overwrite)
- JSON array format preserves all selections
- Clear all data with `COLCLEAR` command to start fresh

### Accurate Calculations
- **Shoelace formula** for polygon area
- **Circle area formula** (πr²) for circles
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
├── column_data.json          # Shape data storage (auto-generated)
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

### Mode 1: Auto-Scan Area (Recommended)

**Command:** `COLSCAN`

1. In AutoCAD command line, type:
   ```
   COLSCAN
   ```

2. Enter minimum distance tolerance:
   ```
   Enter minimum distance between shapes (default 1.0mm): [Type value or press Enter]
   ```
   - Default: 1.0mm
   - Any shapes with centroids closer than this distance will be considered too close
   - The first shape is kept, subsequent shapes within tolerance are removed
   - Press Enter to use default value

3. Select the first corner of scan area:
   ```
   Select first corner:
   ```

4. Select the opposite corner to define rectangular area:
   ```
   Select opposite corner:
   ```

5. Tool will automatically:
   - Detect all closed shapes (polylines, circles) in the area
   - **Calculate centroid-to-centroid distance** for all shape pairs
   - **Remove shapes that are too close** (within your specified tolerance)
   - Keep the first shape in each cluster, remove subsequent ones
   - Calculate area and centroid for each unique shape
   - Find nearest text label for each
6. Confirm export:
   ```
   Export all shapes to JSON? [Yes/No] <Yes>:
   ```
   - Type `Y` or press Enter to export
   - Excel file will be created automatically

7  - Excel file will be created automatically

8. **Repeat for multiple areas** - COLSCAN supports cumulative scanning:
   - Run COLSCAN again to scan another area
   - New shapes will be appended to existing data
   - All shapes from previous scans are preserved
   - Use `COLCLEAR` to start fresh

**Example multi-area workflow:**
```
1. COLSCAN → Select area A (5 shapes detected)
2. COLSCAN → Select area B (3 shapes detected)
3. COLSCAN → Select area C (7 shapes detected)
→ Total: 15 shapes in one Excel file!
```

**What shapes are detected:**
- ✓ Closed LWPOLYLINE
- ✓ Closed POLYLINE  
- ✓ CIRCLE
- ✗ Open polylines (ignored)
- ✗ Lines, arcs (ignored)

### Mode 2: Manual Polygon Selection

**Command:** `COLINSPECT`

1. In AutoCAD command line, type:
   ```
   COLINSPECT
   ```

2. Click points to define polygon vertices:
   ```
   Select first point:
   Select next point (or press Enter to finish):
   ```

3. Press Enter when finished (minimum 3 points required)

4. Preview polyline will be drawn automatically

5. Confirm to export:
   ```

   Export this polygon data? [Yes/No] <Yes>:
   ```
   - Type `Y` or press Enter to export
   - Excel file will be created automatically

### Clear All Data

To start fresh and delete all saved shape data:
```
COLCLEAR
```

Confirms deletion and removes all saved data from JSON file.

---

## Commands Reference

| Command | Description |
|---------|-------------|
| `COLSCAN` | **Auto-scan area** - Select region and detect all shapes automatically |
| `COLINSPECT` | **Manual mode** - Click points to create custom polygon |
| `COLCLEAR` | Clear all saved shapes and start fresh |

---

## JSON Data Format

The tool exports shape data in JSON array format:

### Polygon/Polyline Format:
```json
[
  {
    "type": "lwpolyline",
    "name": "Column-A1",
    "point_count": 4,
    "area": 10000.0,
    "centroid": [150.0, 100.0],
    "points": [
      [100.0, 50.0, 0.0],
      [200.0, 50.0, 0.0],
      [200.0, 150.0, 0.0],
      [100.0, 150.0, 0.0]
    ],
    "timestamp": "20260521",
    "drawing": "FloorPlan.dwg"
  }
]
```

### Circle Format:
```json
[
  {
    "type": "circle",
    "name": "Column-B2",
    "point_count": 1,
    "area": 314.159,
    "centroid": [200.0, 200.0],
    "radius": 10.0,
    "points": [
      [200.0, 200.0, 0.0]
    ],
    "timestamp": "20260521",
    "drawing": "FloorPlan.dwg"
  }
]
```

---

## Excel Output Format

The generated Excel file contains **2 sheets**:

### Sheet 1: Summary (Statistics Table)

A summary table grouped by shape name with:
- **Name**: Shape identifier from nearest text
- **Count**: Number of shapes with this name
- **Total Area**: Sum of all areas for this name
- **Avg Area**: Average area per shape
- **Total Perimeter**: Sum of all perimeters for this name
- **Avg Perimeter**: Average perimeter per shape
- **TOTAL row**: Grand totals across all shapes

#### Example Summary Sheet:
```
SHAPE STATISTICS SUMMARY

Name        | Count | Total Area | Avg Area | Total Perimeter | Avg Perimeter
------------|-------|------------|----------|-----------------|---------------
Column-A1   | 3     | 7500.00    | 2500.00  | 600.00          | 200.00
Column-B1   | 2     | 628.32     | 314.16   | 62.83           | 31.42
Column-C1   | 5     | 12500.00   | 2500.00  | 1000.00         | 200.00
TOTAL       | 10    | 20628.32   | 2062.83  | 1662.83         | 166.28
```

**Use cases:**
- Quick overview of all shapes by name
- Compare quantities and sizes
- Calculate total materials needed
- Identify patterns or outliers

---

### Sheet 2: Shape Details

Detailed information for each individual shape:
- **Property table**: Type, name, point count, area, **perimeter**, radius (for circles), centroid coordinates, drawing info
- **Points table**: Complete list of X, Y, Z coordinates for all vertices (or center for circles)

#### Example Detail Sheet:
```
SHAPE #1
Property          | Value
------------------|--------
Name              | Column-A1
Type              | lwpolyline
Point Count       | 4
Area (sq units)   | 10000.00
Perimeter (units) | 400.00
Centroid X        | 150.00
Centroid Y        | 100.00
Drawing           | FloorPlan.dwg

Points:
Point # | X      | Y      | Z
--------|--------|--------|----
1       | 100.00 | 50.00  | 0.00
2       | 200.00 | 50.00  | 0.00
3       | 200.00 | 150.00 | 0.00
4       | 100.00 | 150.00 | 0.00

SHAPE #2
Property          | Value
------------------|--------
Name              | Column-B2
Type              | circle
Point Count       | 1
Area (sq units)   | 314.16
Perimeter (units) | 62.83
Radius            | 10.00
Centroid X        | 200.00
Centroid Y        | 200.00
...
```

**Perimeter Calculation:**
- **Polygons**: Sum of distances between consecutive vertices (closed loop)
- **Circles**: 2πr (circumference)

---

## Technical Details

### Shape Detection (Auto-Scan Mode)
- Uses AutoCAD selection set (`ssget`) with window selection
- Filters for closed shapes only:
  - Configurable duplicate detection**:
  - **Tolerance**: User-defined minimum distance (default 2.0mm)
  - **Loose mode** (default): Remove shapes within tolerance distance
  - **Strict mode**: Remove shapes within tolerance distance AND similar area
  - Keeps first shape, removes subsequent shapes within tolerance
  - CIRCLE (always closed)
- **Duplicate detection**: Compares centroid position and area with 0.01 tolerance
- Automatically removes overlapping duplicates before saving
- Extracts vertices using VLA objects
- Calculates properties using native AutoCAD functions

### Area Calculation
- **Polygons**: Uses AutoCAD VLA `Area` property (precise)
- **Circles**: Formula πr²
- **Manual polygons**: Shoelace formula fallback

### Centroid Calculation
Simple arithmetic mean:
```
Centroid_X = Σ(x_i) / n
Centroid_Y = Σ(y_i) / n
```

### Text Detection
- Searches all TEXT and MTEXT entities in drawing
- Calculates 2D distance from shape centroid to each text
- Selects nearest text as shape name/label

### Python Auto-Detection
Searches for Python in:
1. `C:\Python312\python.exe`
2. `C:\Python311\python.exe`
3. `C:\Python310\python.exe`
4. `C:\Program Files\Python312\python.exe`
5. `C:\Program Files\Python311\python.exe`
6. System PATH

---

## Troubleshooting

### "No shapes found in selected area"
- Make sure shapes are **closed** (polylines must be closed, open polylines are ignored)
- Check that your selection window includes the shapes
- Try selecting a larger area
- Verify shapes are LWPOLYLINE, POLYLINE, or CIRCLE (lines and arcs not supported)

### "Python not found"
- Install Python 3.10+
- Add Python to system PATH
- Or manually run: `python python_processor.py column_data.json`

### "No polygon data file found" (COLCLEAR command)
- Run `COLSCAN` or `COLINSPECT` at least once to create data
- Check drawing folder for `column_data.json`

### "Need at least 3 points" (Manual mode)
- Polygons require minimum 3 vertices
- Press Enter only after selecting 3+ points

### Python processor doesn't start
- Manually run: `python python_processor.py "path\to\column_data.json"`
- Check `python_processor.py` is in same folder or add to PATH
- Install dependencies: `pip install pandas openpyxl`
shapes too close together
**Cause:**
- Multiple layers overlapping
- Copy/paste at nearby locations
- Imported DWGs from multiple sources
- Shapes naturally spaced closer than tolerance

**Solution:**
- ✅ Tool prompts for tolerance distance at start
- ✅ Default 1.0mm - adjust based on your drawing scale
- ✅ Only centroid-to-centroid distance is checked (no area/size comparison)
- Check output: "⚠ Removed X shape(s) within Y.Ymm from each other"
- First shape in each cluster is kept, rest are removed

**How it works:**
- Distance measured from **centroid to centroid** (not edge to edge)
- If 2 shapes have centroids < tolerance apart → second shape removed
- Different sizes, different areas → doesn't matter, only distance counts

**Examples:**
- Small scale drawings (details): Use 0.5mm - 1.0mm tolerance
- Medium scale (floor plans): Use 1.0mm - 5.0mm tolerance  
- Large scale (site plans): Use 10.0mm - 50.0mm tolerance

**Solution:**
- ✅ Tool automatically removes duplicates!
- Check output: "⚠ Removed X duplicate shape(s)"
- Only unique shapes are saved to JSON/Excel
- Duplicates are detected by comparing position (centroid) and area

---

## Example Workflow

### Scenario: Extract all column data from floor plan

1. Open your AutoCAD drawing with columns
2. Load the plugin: `(load "ColumnInspector.lsp")`
3. Run auto-scan: `COLSCAN`
4. Select area covering all columns (2 clicks)
5. Review detected shapes in command line
6. Confirm export: Press Enter
7. Open generated Excel file in output folder
8. All column data ready for analysis!

**Time saved:** From hours of manual measurement to seconds of automated extraction.

---

## License2** (Customizable Duplicate Detection
2** (2026-05-21): Added customizable tolerance for duplicate detection
- **v3.
Copyright (c) 2026 Tran Xuan An  
Licensed under the MIT License.

## Version

Current Version: **3.4** (Summary Statistics + Perimeter Calculation)  
Date: May 21, 2026

### Changelog
- **v3.4** (2026-05-21): Added Summary sheet with statistics by name, perimeter calculation for all shapes
- **v3.3** (2026-05-21): Simplified to distance-only check (centroid-to-centroid), removed area comparison
- **v3.1** (2026-05-21): Added automatic duplicate detection and removal
- **v3.0** (2026-05-21): Added COLSCAN command for automatic shape detection
- **v2.0** (2026-05-18): Multi-polygon support with JSON array format
- **v1.0** (2026-05-15): Initial release with manual polygon selection