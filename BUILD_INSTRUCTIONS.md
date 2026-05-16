# Build Instructions for CAD Column Inspector Pro

## 🎯 Chosen Solution: **Static Application (Standalone)**

### Reasons:
- ✅ Users don't need to install Python
- ✅ Easy to distribute and install
- ✅ Avoid dependency conflicts
- ✅ Professional for commercial software

---

## 📋 Prerequisites

1. **Python installed** (version 3.8+)
2. **Dependencies from requirements.txt**

---

## 🚀 How to Build the Application

### Step 1: Install PyInstaller
```powershell
pip install pyinstaller
```

### Step 2: Install all dependencies
```powershell
pip install -r requirements.txt
```

### Step 3: Build executable
```powershell
pyinstaller build_installer.spec
```

### Step 4: Get the result
After building, the executable will be located in:
```
dist/CAD_Column_Inspector_Pro/CAD_Column_Inspector_Pro.exe
```

---

## 📦 Packaging Options

### Option 1: **One-Folder Mode** (Recommended) ✅
```
dist/CAD_Column_Inspector_Pro/
├── CAD_Column_Inspector_Pro.exe  (main executable)
├── python311.dll
├── _tkinter.pyd
└── ... (other DLLs and dependencies)
```

**Advantages:**
- Fast build
- Easy to debug errors
- Can add/remove files easily
- Suitable for internal distribution

**How to distribute:**
- Compress `CAD_Column_Inspector_Pro` folder to ZIP
- Users extract and run the .exe file

---

### Option 2: **One-File Mode** (Simpler)
Edit `build_installer.spec` file, change:

```python
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,      # ← Add this line
    a.zipfiles,      # ← Add this line
    a.datas,         # ← Add this line
    [],
    exclude_binaries=False,  # ← Change to False
    name='CAD_Column_Inspector_Pro',
    ...
)

# DELETE the COLLECT section below
```

Rebuild:
```powershell
python -m PyInstaller build_installer.spec --clean
```

Result: **Single .exe file**

**Advantages:**
- Very easy to distribute (only 1 file)
- No missing dependency files

**Disadvantages:**
- Slower startup (must extract temp files)
- Larger file size (usually 80-150 MB)
- May be blocked by antivirus

---

## 🔧 Advanced Customization

### Add Icon to the application
1. Create/download `icon.ico` file (256x256px size)
2. Place in project folder
3. Edit in `build_installer.spec`:
```python
icon='icon.ico'
```

### Enable Console Mode (for debugging)
```python
console=True  # Show console window
```

### Reduce file size
```python
excludes=[
    'matplotlib',
    'scipy',
    'numpy.distutils',
    'test',
    'unittest',
],
```

### Build for machines without Python
```powershell
# Ensure building on Windows with similar architecture as target machine
pyinstaller build_installer.spec --clean --noconfirm
```

---

## 🎨 Create Installer (Optional)

Use **Inno Setup** to create professional installer:

### Step 1: Download Inno Setup
https://jrsoftware.org/isdl.php

### Step 2: Create `installer_script.iss` file
```iss
[Setup]
AppName=CAD Column Inspector Pro
AppVersion=1.0.0
DefaultDirName={autopf}\CAD Column Inspector Pro
DefaultGroupName=CAD Column Inspector Pro
OutputDir=output
OutputBaseFilename=CAD_Column_Inspector_Pro_Setup
Compression=lzma2
SolidCompression=yes

[Files]
Source: "dist\CAD_Column_Inspector_Pro\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs

[Icons]
Name: "{group}\CAD Column Inspector Pro"; Filename: "{app}\CAD_Column_Inspector_Pro.exe"
Name: "{autodesktop}\CAD Column Inspector Pro"; Filename: "{app}\CAD_Column_Inspector_Pro.exe"
```

### Step 3: Compile
Open Inno Setup Compiler → Load `.iss` file → Compile

**Result:** `CAD_Column_Inspector_Pro_Setup.exe` file for installation

---

## ⚠️ Important Notes

### 1. AutoCAD Requirement
The application requires **AutoCAD installed** on the machine to:
- Connect via COM interface
- Highlight and zoom objects

### 2. Antivirus may block
- PyInstaller executables are often flagged by Windows Defender
- **Solution:** Code signing certificate (paid) or whitelist

### 3. Test on clean machine
After building, test on a machine **without Python** to ensure:
- ✅ Executable runs independently
- ✅ No missing dependencies
- ✅ UI displays correctly

### 4. File sizes
- **One-Folder:** ~80-120 MB
- **One-File:** ~100-150 MB
- Mainly due to pandas, openpyxl, and Python runtime

---

## 🏗️ Quick Build Commands

### Build Development (with console)
```powershell
pyinstaller --onedir --console --name="CAD_Column_Inspector_Pro" main.py
```

### Build Production (no console, clean)
```powershell
pyinstaller build_installer.spec --clean --noconfirm
```

### Build One-File Production
```powershell
pyinstaller --onefile --windowed --name="CAD_Column_Inspector_Pro" --icon=icon.ico main.py
```

---

## 📝 Checklist Before Distribution

- [ ] Test on clean machine (without Python)
- [ ] Test with AutoCAD running
- [ ] Test with sample DXF file
- [ ] Check Excel export
- [ ] Check logging (log file created?)
- [ ] Create User Manual (README for end-users)
- [ ] Create LICENSE file
- [ ] (Optional) Code signing to avoid antivirus

---

## 📚 Static vs Dynamic Comparison

| Criteria | Static (Standalone) | Dynamic (Python Required) |
|----------|---------------------|---------------------------|
| **Installation** | No Python needed | Need Python + deps |
| **Distribution** | ✅ Easy | ❌ Complex |
| **Size** | 100-150 MB | 500 KB (code only) |
| **Code Security** | ✅ Hard to reverse | ❌ Easy to read .py |
| **Updates** | Must rebuild | Just copy files |
| **Professional** | ✅ High | ❌ Low |
| **Recommendation** | ✅ **FOR THIS PROJECT** | ❌ Not suitable |

---

## 🎯 Final Recommendations

**For CAD Column Inspector Pro project:**

1. ✅ **Use PyInstaller with One-Folder Mode**
2. ✅ **Create Inno Setup installer** (professional)
3. ✅ **Add icon and User Manual**
4. ❌ **DO NOT use dynamic** (CAD users typically don't know Python)

---

## 🆘 Troubleshooting

### Error: "Failed to execute script main"
```powershell
# Build with console mode to see errors
pyinstaller build_installer.spec --console
```

### Error: "ImportError: DLL load failed"
```powershell
# Add to hiddenimports in .spec file
```

### Error: "win32com.client not found"
```powershell
pip install --upgrade pywin32
python -m pywin32_postinstall -install
```

### Antivirus blocking
- Submit false positive report to Microsoft
- Or use code signing certificate

---

**Built with ❤️ by Tran Xuan An © 2026**
