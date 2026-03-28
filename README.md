
# CAD Column Inspector Pro - Automation Tool

## Source Code Design

This project is a modular Python application for inspecting and processing CAD (DXF) files, connecting to AutoCAD, and exporting analysis results to Excel. The application uses a Tkinter-based GUI and is organized into the following main modules:

- **main.py**: Entry point and main application logic. Manages the GUI, user interactions, and coordinates between modules.
- **ui_components.py**: Contains reusable Tkinter UI components (button panel, filter frame, data table, status bar) for a clean and modular interface.
- **cad_processor.py**: Handles the extraction and analysis of geometric objects (rectangles, circles, etc.) from DXF files using ezdxf.
- **cad_connector.py**: Manages connection to a running AutoCAD instance and provides functions to zoom and highlight objects in AutoCAD via COM automation.
- **excel_exporter.py**: Exports processed and filtered data to Excel files, including summary, detailed, and statistics sheets using pandas.
- **config.py**: Centralizes configuration constants for geometry processing, UI layout, and color codes.

## Detailed Implementation

### main.py
- Initializes the main Tkinter window and all UI components.
- Handles user actions: loading DXF files, connecting to AutoCAD, filtering data, exporting to Excel.
- Uses `CadProcessor` to extract rectangles and circles from DXF files.
- Uses `CadConnector` to zoom/highlight objects in AutoCAD.
- Uses `ExcelExporter` to export the currently displayed data to Excel.

### ui_components.py
- Provides modular UI widgets:
	- `ButtonPanel`: Button row for loading DXF, connecting AutoCAD, exporting Excel.
	- `FilterFrame`: Checkboxes to filter rectangles/circles.
	- `DataTable`: Treeview for displaying detected objects and their properties.
	- `StatusBar`: Displays current status and messages.

### cad_processor.py
- Extracts rectangles from LWPOLYLINE and POLYLINE entities by checking geometric properties.
- Extracts circles from CIRCLE, ARC (full circle), and ELLIPSE (with equal axes) entities.
- Provides utility methods for geometric checks and property extraction.

### cad_connector.py
- Connects to a running AutoCAD instance using pyautocad and pywin32.
- Provides `zoom_to_point(x, y)` to focus and highlight a location in AutoCAD, drawing temporary markers and using color blinking for emphasis.
- Handles COM automation and error management for robust operation.

### excel_exporter.py
- Exports data to Excel with three sheets:
	- **Summary_by_Type_Size**: Grouped by object type, size, and layer, with counts and area stats.
	- **Detailed_Data**: Full list of all detected objects and their properties.
	- **Statistics**: Overall and per-type/size statistics for quick analysis.
- Uses pandas and openpyxl for Excel file creation.

### config.py
- Stores constants for geometry tolerance, highlight colors, UI layout, and window size.

## Environment Setup

### Requirements
- Python 3.12.10

### Installation

Install required dependencies using pip:

```bash
pip install ezdxf pandas openpyxl pyautocad pywin32 xlsxwriter