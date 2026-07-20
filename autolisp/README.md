# AutoLISP Modular Architecture

## 📁 Cấu trúc thư mục

```
cad_tool/
├── ColumnInspector_new.lsp       # Main loader (load file này)
├── ColumnInspector.lsp            # Old monolithic version (backup)
│
├── autolisp/                      # Modules folder
│   ├── config.lsp                 # Configuration & constants
│   ├── utils.lsp                  # Utility functions
│   ├── geometry.lsp               # Geometric calculations
│   ├── detection.lsp              # Shape detection & analysis
│   ├── io.lsp                     # File I/O (JSON export/import)
│   └── commands.lsp               # Command definitions
│
├── main.py                        # Python processor
├── data_loader.py
├── shape_calculator.py
├── excel_exporter.py
├── find_shape.py
└── requirements.txt
```

---

## 🎯 Kiến trúc Module

### **Dependency Graph**

```
ColumnInspector_new.lsp (loader)
    │
    ├─► config.lsp          (no dependencies)
    ├─► utils.lsp           (no dependencies)
    ├─► geometry.lsp        (depends on: utils, config)
    ├─► detection.lsp       (depends on: geometry, utils, config)
    ├─► io.lsp              (depends on: utils, config)
    └─► commands.lsp        (depends on: all above)
```

### **Load Order**
1. **config.lsp** - Constants và global variables
2. **utils.lsp** - Helper functions
3. **geometry.lsp** - Area, centroid, distance calculations
4. **detection.lsp** - Entity analysis, shape detection
5. **io.lsp** - JSON read/write
6. **commands.lsp** - COLINSPECT, COLSCAN, etc.

---

## 📦 Module Chi tiết

### **1. config.lsp**
**Purpose:** Centralized configuration

**Contents:**
- File paths (`*COL_DATA_FILENAME*`, `*COL_PYTHON_SCRIPT*`)
- Processing settings (`*COL_DEFAULT_TOLERANCE*`)
- Visual settings (colors, line widths)
- Python search paths
- Entity filters for auto-scan
- Global variables (`*COLFIND_SHAPEID*`)

**Why separate:** Easy to modify settings without touching code

---

### **2. utils.lsp**
**Purpose:** Reusable utility functions

**Functions:**
- `str-split` - String splitting
- `str-clean` - Clean shape names
- `str-pad-zeros` - Format IDs with leading zeros
- `get-result-file` - Get JSON file path
- `file-exists-p` - Check file existence
- `delete-entities` - Safe entity deletion
- `generate-shape-id` - Generate unique IDs
- `find-python-executable` - Locate Python
- `call-python-processor` - Invoke Python script
- `find-nearest-text` - Find text near point
- `print-header/success/error/warning` - UI utilities

**Why separate:** Single Responsibility, reusable across modules

---

### **3. geometry.lsp**
**Purpose:** Mathematical calculations

**Functions:**
- `calculate-polygon-area` - Shoelace formula
- `calculate-centroid` - Geometric center
- `distance-2d` - 2D distance
- `get-lwpolyline-points` - Extract LWPOLYLINE vertices
- `get-polyline-points` - Extract POLYLINE vertices
- `is-duplicate-shape` - Check if shapes too close
- `remove-duplicate-shapes` - Filter duplicates
- `check-area-consistency` - Validate area consistency

**Why separate:** Pure math logic, easy to test and debug

---

### **4. detection.lsp**
**Purpose:** CAD entity detection and analysis

**Functions:**
- `analyze-entity` - Analyze LWPOLYLINE/POLYLINE/CIRCLE
- `draw-preview-polyline` - Red preview lines
- `draw-final-polygon` - Final closed polyline

**Why separate:** CAD-specific logic isolated from calculations

---

### **5. io.lsp**
**Purpose:** File I/O operations

**Functions:**
- `append-polygon-data` - Append single shape to JSON
- `save-all-shapes` - Save multiple shapes to JSON
- `load-shapes-from-json` - Load shapes for consistency check
- `find-shape-by-id` - Find shape by ID
- `parse-json-shapes` - JSON parser (stub)

**Why separate:** File operations isolated, easier to add new formats

---

### **6. commands.lsp**
**Purpose:** User-facing commands

**Commands:**
- `C:COLINSPECT` - Manual polygon selection
- `C:COLSCAN` - Auto-scan area
- `C:COLCLEAR` - Clear saved data
- `C:COLFINDPOLY` - Reverse lookup (highlight shape)
- `C:COLLIST` - List all shape IDs

**Why separate:** User interface layer, orchestrates other modules

---

## 🚀 Cách sử dụng

### **Load vào AutoCAD**

**Method 1: Manual Load**
```
Command: APPLOAD
Browse to: ColumnInspector_new.lsp
Click: Load
```

**Method 2: Auto-load (recommended)**

Thêm vào `acaddoc.lsp`:
```lisp
(if (findfile "C:/path/to/ColumnInspector_new.lsp")
  (load "C:/path/to/ColumnInspector_new.lsp")
)
```

**Method 3: Startup Suite**
```
Command: APPLOAD
Click: Contents (in Startup Suite)
Add: ColumnInspector_new.lsp
```

### **Commands**

```
COLINSPECT  - Select polygon points manually
COLSCAN     - Auto-scan rectangular area
COLCLEAR    - Delete all saved data
COLFINDPOLY - Highlight shape by ID
COLLIST     - Show all shape IDs
```

---

## ✅ Ưu điểm của Modular Architecture

### **1. Maintainability**
- Mỗi module < 300 dòng (vs 970 dòng monolithic)
- Dễ tìm và sửa bugs
- Clear separation of concerns

### **2. Reusability**
- Functions có thể gọi từ modules khác
- Không duplicate code
- Dễ extend với tính năng mới

### **3. Testability**
- Test từng module độc lập
- Mock dependencies dễ dàng
- Debug từng layer riêng

### **4. Scalability**
- Thêm module mới không ảnh hưởng code cũ
- Có thể load/unload modules động
- Dễ tạo variants (e.g., ColumnInspector_Lite)

### **5. Collaboration**
- Nhiều người làm việc trên modules khác nhau
- Merge conflicts ít hơn
- Code review dễ hơn

### **6. Configuration Management**
- Tất cả settings ở 1 file (config.lsp)
- Không cần search trong 970 dòng code
- Environment-specific configs dễ dàng

---

## 🔧 Customization

### **Thay đổi tolerance mặc định**

Edit `autolisp/config.lsp`:
```lisp
(setq *COL_DEFAULT_TOLERANCE* 30.0)  ; Was 20.0
```

### **Thay đổi màu highlight**

Edit `autolisp/config.lsp`:
```lisp
(setq *COL_HIGHLIGHT_COLOR* 3)  ; Cyan instead of yellow
```

### **Thêm Python path mới**

Edit `autolisp/config.lsp`:
```lisp
(setq *COL_PYTHON_PATHS* '(
  "D:\\MyPython\\python.exe"  ; Add new path
  "C:\\Python312\\python.exe"
  ; ... existing paths
))
```

### **Thêm command mới**

Tạo function trong `autolisp/commands.lsp`:
```lisp
(defun C:COLMYCOMMAND (/ ...)
  (print-header "My Custom Command")
  ; Your logic here
  (princ)
)
```

---

## 🐛 Debugging

### **Check if modules loaded**

Trong AutoCAD command line:
```
Command: !*COL_BASE_PATH*
Should return: path to project folder

Command: !*COL_DEFAULT_TOLERANCE*
Should return: 20.0
```

### **Reload một module**

```
Command: (load "autolisp/geometry.lsp")
```

### **Test một function**

```
Command: (setq test_points '((0 0 0) (100 0 0) (100 100 0) (0 100 0)))
Command: (calculate-polygon-area test_points)
Should return: 10000.0
```

---

## 📊 So sánh Old vs New

| Aspect | Old (Monolithic) | New (Modular) |
|--------|-----------------|---------------|
| **Lines per file** | 970 | 50-250 |
| **Files** | 1 | 7 |
| **Find function** | Scroll 970 lines | Open specific module |
| **Modify config** | Search code | Edit config.lsp |
| **Add feature** | Insert anywhere | Add to specific module |
| **Test** | Hard (coupled) | Easy (isolated) |
| **Collaboration** | Merge conflicts | Independent modules |
| **Load time** | Same | Same (all loaded) |

---

## 🔄 Migration Path

### **Option A: Direct Switch**
1. Backup `ColumnInspector.lsp` → `ColumnInspector_old.lsp`
2. Load `ColumnInspector_new.lsp`
3. Test all commands
4. If works → delete old file

### **Option B: Gradual Migration**
1. Keep both versions
2. Test new version in test drawings
3. Report any issues
4. Switch when confident

### **Option C: Compile to FAS**
```lisp
; Compile all modules into one FAS file
(vlisp-compile 'st "ColumnInspector_new.lsp" "ColumnInspector.fas")
```

---

## 📝 Future Enhancements

### **Possible new modules:**

1. **validation.lsp** - Data validation rules
2. **export.lsp** - Multiple export formats (CSV, DXF)
3. **ui.lsp** - DCL dialog interfaces
4. **network.lsp** - HTTP API calls
5. **database.lsp** - Direct database connections
6. **reports.lsp** - Generate reports in AutoCAD

### **Performance optimizations:**

1. Lazy loading - Only load modules when needed
2. Caching - Cache frequently used calculations
3. Batch operations - Process multiple shapes in one transaction

---

## 🆘 Troubleshooting

### **Error: "Module not found"**

**Cause:** `autolisp/` folder not in correct location

**Solution:** 
- Ensure folder structure is correct
- Check `*COL_BASE_PATH*` variable
- Try absolute paths in loader

### **Error: "Unknown command"**

**Cause:** commands.lsp didn't load

**Solution:**
- Check if all dependencies loaded first
- Load commands.lsp manually: `(load "autolisp/commands.lsp")`

### **Error: "no function definition"**

**Cause:** Function called before module loaded

**Solution:**
- Check load order in ColumnInspector_new.lsp
- Ensure dependencies are correct

---

## 📞 Support

For issues or questions about the modular architecture:
1. Check this README
2. Review module source code (well-commented)
3. Test functions individually
4. Check AutoCAD command line for error messages

---

**Copyright (c) 2026 Tran Xuan An**
**Version:** 4.0 - Modular Architecture
**Last Updated:** 2026-07-13
