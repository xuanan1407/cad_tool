# AutoCAD Plugin Installation Guide
## CAD Column Inspector Pro - Hybrid Version

---

## 🎯 Architecture Overview

This is a **Hybrid Plugin** combining:
- **AutoLISP** - Interactive point selection in AutoCAD
- **Python** - Data processing and Excel export

```
[AutoCAD] → [AutoLISP Plugin] → [JSON Data] → [Python Processor] → [Excel Output]
```

---

## 📋 Prerequisites

### 1. AutoCAD
- AutoCAD 2018 or later
- Any version that supports AutoLISP

### 2. Python
- Python 3.8+ installed
- Required packages: `pandas`, `openpyxl`

---

## 🚀 Installation

### Step 1: Install Python Dependencies

```powershell
pip install pandas openpyxl
```

### Step 2: Copy Plugin Files

Copy these files to your working directory:
```
your_folder/
├── ColumnInspector.lsp      # AutoLISP plugin
├── python_processor.py        # Python processor
└── (your drawing files)
```

### Step 3: Load Plugin in AutoCAD

**Method 1: Manual Load (for testing)**
1. Open AutoCAD
2. Type `APPLOAD` and press Enter
3. Browse and select `ColumnInspector.lsp`
4. Click "Load"
5. Plugin is now loaded!

**Method 2: Auto-Load (recommended)**
1. Find your AutoCAD support path:
   - Type `OPTIONS` → Files tab → Support File Search Path
2. Copy `ColumnInspector.lsp` to one of these folders
3. Create/edit `acaddoc.lsp` in the same folder:
   ```lisp
   (load "ColumnInspector.lsp")
   ```
4. Plugin will auto-load with every drawing

**Method 3: Startup Suite (best for deployment)**
1. Type `APPLOAD`
2. Click "Contents" button in Startup Suite section
3. Click "Add" → Select `ColumnInspector.lsp`
4. Click "Close"
5. Plugin will load automatically with AutoCAD

---

## 🎮 How to Use

### 1. Start the Tool

In AutoCAD command line, type:
```
COLINSPECT
```

Press Enter.

### 2. Select Points

```
=== CAD Column Inspector Pro ===
Select points to create polygon (Press Enter to finish)
Select first point:
```

- Click on drawing to select points
- Select at least 3 points to create a polygon
- Press **Enter** when done

### 3. Preview

The plugin will:
- Draw temporary polyline connecting your points
- Calculate area and centroid
- Display information

### 4. Confirm Export

```
Polygon created with 5 points
Area: 123.45 sq units
Centroid: (10.0, 20.0)

Export this polygon data? [Yes/No] <Yes>:
```

- Type **Y** or press Enter → Export and process
- Type **N** → Cancel (polyline will be deleted)

### 5. Check Output

If successful:
```
✓ Data exported to: C:\path\to\column_data.json
✓ Python processor started successfully!
✓ Check the output folder for Excel file.
```

Excel file will be created: `Column_Inspector_YYYYMMDD_HHMMSS.xlsx`

---

## 📊 Output Format

### Excel File Contains:

**Sheet: "Polygon Data"**
- Type: polygon
- Point Count: 5
- Area: 123.45 sq units
- Centroid X, Y
- Drawing name
- Timestamp
- Detailed points table (X, Y, Z coordinates)

---

## 🔧 Configuration

### Modify Python Path

If Python is not auto-detected, edit `ColumnInspector.lsp`:

```lisp
(defun find-python-executable (/ python_paths test_path)
  (setq python_paths '(
    "C:\\Your\\Custom\\Path\\python.exe"  ; Add your path here
    "C:\\Python312\\python.exe"
    ...
  ))
  ...
)
```

### Change Output Location

By default, files are saved in the same folder as the drawing.

To change, edit `python_processor.py`:

```python
self.output_folder = "C:\\Your\\Custom\\Output\\Folder"
```

---

## 🐛 Troubleshooting

### Issue: "Python not found"

**Solution:**
1. Check Python is installed: `python --version`
2. Add custom Python path in `ColumnInspector.lsp`
3. Or manually run: `python python_processor.py column_data.json`

### Issue: "python_processor.py not found"

**Solution:**
- Ensure `python_processor.py` is in the same folder as your drawing
- Or use absolute path in AutoLISP code

### Issue: Command not recognized

**Solution:**
- Reload plugin: `(load "ColumnInspector.lsp")`
- Check plugin loaded successfully
- Try `APPLOAD` and load manually

### Issue: Excel file not created

**Solution:**
- Check Python dependencies: `pip list | findstr pandas`
- Check `column_data.json` was created
- Run Python manually to see errors:
  ```
  python python_processor.py column_data.json
  ```

---

## 🎨 Advanced Usage

### Process Multiple Polygons

Run `COLINSPECT` multiple times. Each execution creates a new polygon and Excel file.

### Batch Processing

Collect multiple JSON files and process together:

```python
# batch_process.py
import glob
from python_processor import PolygonProcessor

for json_file in glob.glob("column_data_*.json"):
    processor = PolygonProcessor(json_file)
    processor.load_data()
    processor.process()
    processor.export_to_excel()
```

### Integrate with Existing Code

The Python processor can be integrated with existing modules:

```python
from python_processor import PolygonProcessor
from excel_exporter import ExcelExporter  # Your existing exporter

# Process polygon
processor = PolygonProcessor("column_data.json")
processor.load_data()

# Use your existing export logic
# ...
```

---

## 📝 Comparison: Standalone vs Plugin

| Feature | Standalone App | Plugin Version |
|---------|---------------|----------------|
| **User Experience** | Export DXF → Load in app | Direct selection in CAD |
| **Workflow** | Multi-step | Single command |
| **Flexibility** | Auto-detect shapes | Manual selection |
| **Accuracy** | Based on DXF data | Based on user intent |
| **Integration** | External tool | Native CAD tool |

**Plugin version** is better for:
- Interactive workflows
- Custom shape definition
- Direct CAD integration

**Standalone version** is better for:
- Batch processing
- Automated analysis
- No CAD required

---

## 🚀 Next Steps

### Enhance the Plugin

1. **Add more shape types:**
   - Rectangles
   - Circles
   - Custom shapes

2. **Multi-selection:**
   - Select multiple polygons in one session
   - Compile into single Excel file

3. **Layer integration:**
   - Auto-detect layer
   - Filter by layer

4. **Database integration:**
   - Save to database instead of Excel
   - Query and analysis features

5. **UI Improvements:**
   - Add dialog box for options
   - Preview before export
   - Undo/redo functionality

---

## 📞 Support

For issues or questions:
- Check log files in drawing folder
- Review `column_data.json` for data accuracy
- Test Python processor independently

---

**Built with ❤️ by Tran Xuan An © 2026**
