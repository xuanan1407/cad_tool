# build_release.py
# Script to build release package for CAD Column Inspector Pro
# This script creates executable files to protect source code

import os
import sys
import shutil
import subprocess
from pathlib import Path
from datetime import datetime

print("=" * 60)
print("CAD Column Inspector Pro - Release Builder")
print("=" * 60)

# Configuration
PROJECT_ROOT = Path(__file__).parent
DIST_FOLDER = PROJECT_ROOT / "dist"
BUILD_FOLDER = PROJECT_ROOT / "build"
RELEASE_FOLDER = PROJECT_ROOT / "release"

def clean_folders():
    """Remove old build artifacts"""
    print("\n[1/5] Cleaning old build folders...")
    for folder in [DIST_FOLDER, BUILD_FOLDER, RELEASE_FOLDER]:
        if folder.exists():
            shutil.rmtree(folder)
            print(f"  ✓ Removed {folder.name}/")
    
    RELEASE_FOLDER.mkdir(exist_ok=True)
    BUILD_FOLDER.mkdir(exist_ok=True)
    print("  ✓ Created fresh release/ and build/ folders")
    print("  ✓ Created fresh release/ folder")

def check_pyinstaller():
    """Check if PyInstaller is installed"""
    print("\n[2/5] Checking PyInstaller...")
    try:
        subprocess.run([sys.executable, "-m", "PyInstaller", "--version"], 
                      check=True, capture_output=True)
        print("  ✓ PyInstaller is installed")
        return True
    except:
        print("  ✗ PyInstaller not found")
        print("  → Installing PyInstaller...")
        try:
            subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"],
                          check=True)
            print("  ✓ PyInstaller installed successfully")
            return True
        except:
            print("  ✗ Failed to install PyInstaller")
            return False

def build_main_processor():
    """Build main.py into standalone executable"""
    print("\n[3/5] Building main processor (main.exe)...")
    
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",                    # Single executable file
        "--clean",                      # Clean cache
        "--noconfirm",                  # Overwrite without asking
        "--name", "CADColumnProcessor", # Output name
        "--icon", "NONE",               # No icon (or add your .ico file)
        "--console",                    # Console app
        "main.py"
    ]
    
    try:
        subprocess.run(cmd, check=True, cwd=PROJECT_ROOT)
        print("  ✓ Built CADColumnProcessor.exe")
        return True
    except:
        print("  ✗ Failed to build main processor")
        return False

def build_find_shape():
    """Build find_shape.py into standalone executable"""
    print("\n[4/5] Building shape finder (find_shape.exe)...")
    
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--clean",
        "--noconfirm",
        "--name", "FindShape",
        "--icon", "NONE",
        "--console",
        "find_shape.py"
    ]
    
    try:
        subprocess.run(cmd, check=True, cwd=PROJECT_ROOT)
        print("  ✓ Built FindShape.exe")
        return True
    except:
        print("  ✗ Failed to build shape finder")
        return False

def package_release():
    """Copy files to release folder"""
    print("\n[5/5] Packaging release files...")
    
    # Check for .fas file first (compiled), fallback to .lsp
    autolisp_file = None
    autolisp_name = None
    
    if (PROJECT_ROOT / "ColumnInspector.fas").exists():
        autolisp_file = PROJECT_ROOT / "ColumnInspector.fas"
        autolisp_name = "ColumnInspector.fas"
        print("  ✓ Found compiled .fas file (source code protected)")
    elif (PROJECT_ROOT / "ColumnInspector.lsp").exists():
        autolisp_file = PROJECT_ROOT / "ColumnInspector.lsp"
        autolisp_name = "ColumnInspector.lsp"
        print("  ⚠ Using .lsp file (source code visible)")
        print("  → Consider compiling to .fas for better protection")
    
    # Copy executables
    files_to_copy = [
        (DIST_FOLDER / "CADColumnProcessor.exe", "CADColumnProcessor.exe"),
        (DIST_FOLDER / "FindShape.exe", "FindShape.exe"),
        (PROJECT_ROOT / "PLUGIN_INSTALLATION.md", "INSTALLATION_GUIDE.md"),
        (PROJECT_ROOT / "README.md", "README.md"),
    ]
    
    # Add AutoLISP file if found
    if autolisp_file:
        files_to_copy.insert(2, (autolisp_file, autolisp_name))
    
    # NOTE: scripts/ folder is NOT included in release
    # It will be generated automatically when users run the tool and export Excel
    # Each scripts/ contains .vbs/.bat files specific to that drawing's shapes
    
    for src, dst_name in files_to_copy:
        if src.exists():
            shutil.copy(src, RELEASE_FOLDER / dst_name)
            print(f"  ✓ Copied {dst_name}")
        else:
            print(f"  ⚠ Missing {src.name}")
    
    # Create README for release
    create_release_readme()
    
    print(f"\n✓ Release package ready in: {RELEASE_FOLDER.absolute()}")

def create_release_readme():
    """Create simplified README for end users"""
    content = """# CAD Column Inspector Pro - Release v1.0

## Installation Guide

### Step 1: Install AutoLISP Plugin
1. Open AutoCAD
2. Type: `APPLOAD`
3. Browse and select: `ColumnInspector.fas` (or .lsp if .fas not available)
4. Click "Load" → Click "Close"
5. (Optional) Add to Startup Suite for auto-load

### Step 2: Setup Executables (No Python Required!)
1. Copy `CADColumnProcessor.exe` to a permanent location
   Example: `C:\\CAD_Tools\\CADColumnProcessor.exe`

2. Copy `FindShape.exe` to the same location
   Example: `C:\\CAD_Tools\\FindShape.exe`

3. These are standalone executables - Python installation is NOT required!

### Step 3: Test the Tool
1. In AutoCAD, type: `COLINSPECT` (Manual mode)
2. Or type: `COLSCAN` (Auto-scan mode)
3. Follow on-screen instructions
4. Excel file will be generated automatically

## Commands Available
- `COLINSPECT` - Manual polygon selection
- `COLSCAN` - Auto-scan area for shapes
- `COLCLEAR` - Clear all data
- `COLFINDPOLY` - Find shape by ID in CAD

## Reverse Lookup (Excel → CAD)
**NEW FEATURE:** Find shapes in AutoCAD directly from Excel!

1. Open exported Excel file (Column_Inspector_*.xlsx)
2. Go to "Shape Details" sheet
3. Click "🔍 Find in CAD" button in any row
4. The shape will highlight in YELLOW in AutoCAD automatically!

**How it works:**
- Clicking the button runs FindShape.exe (no console window)
- AutoCAD plugin receives the shape ID
- Shape highlights for 5 seconds then returns to normal

**Note:** Both AutoCAD and Excel must be open for this to work.

## System Requirements
- AutoCAD 2018 or later
- Windows 10/11 (64-bit)
- NO Python installation required!

## Troubleshooting

### Issue: Commands not working
**Solution:** Reload the plugin: `(load "ColumnInspector.fas")`

### Issue: Excel file not created
**Solution:** 
- Check that CADColumnProcessor.exe is accessible
- Verify column_data.json was created in drawing folder
- Run CADColumnProcessor.exe manually to see errors

### Issue: Reverse lookup not working
**Solution:**
- Ensure FindShape.exe is in the same folder as the Excel file
- Verify AutoCAD is open with the correct drawing
- Check scripts/ folder exists with .vbs files

## Support
For issues or questions, contact: [Your Contact Info]

## Copyright
© 2026 Tran Xuan An - All Rights Reserved

---
**Version:** 1.0  
**Build Date:** """ + datetime.now().strftime("%Y-%m-%d") + """  
**All executables are standalone - no dependencies required!**
"""
    
    with open(RELEASE_FOLDER / "USER_GUIDE.txt", "w", encoding="utf-8") as f:
        f.write(content)
    print("  ✓ Created USER_GUIDE.txt")

def move_spec_files():
    """Move .spec files generated by PyInstaller to build/ folder"""
    spec_files_moved = 0
    for spec_file in PROJECT_ROOT.glob("*.spec"):
        if spec_file.name not in ["build", "dist", "release"]:
            dest = BUILD_FOLDER / spec_file.name
            shutil.move(str(spec_file), str(dest))
            spec_files_moved += 1
    
    if spec_files_moved > 0:
        print(f"\n  ✓ Moved {spec_files_moved} .spec file(s) to build/ folder")

def main():
    """Main build process"""
    try:
        clean_folders()
        
        if not check_pyinstaller():
            print("\n✗ Build failed: PyInstaller not available")
            return False
        
        if not build_main_processor():
            print("\n✗ Build failed: Could not build main processor")
            return False
        
        if not build_find_shape():
            print("\n✗ Build failed: Could not build shape finder")
            return False
        
        # Move .spec files to build/ folder
        move_spec_files()
        
        package_release()
        
        print("\n" + "=" * 60)
        print("✓ BUILD SUCCESSFUL!")
        print("=" * 60)
        print(f"\nRelease package location:")
        print(f"  → {RELEASE_FOLDER.absolute()}")
        
        # Check which AutoLISP format was used
        if (PROJECT_ROOT / "ColumnInspector.fas").exists():
            print("\n🔒 Security Status:")
            print("  ✅ AutoLISP: Protected (.fas binary)")
            print("  ✅ Python: Protected (.exe binaries)")
            print("\n✓ All source code is protected!")
        else:
            print("\n⚠ Security Warning:")
            print("  ⚠ AutoLISP: Source code visible (.lsp file)")
            print("  ✅ Python: Protected (.exe binaries)")
            print("\nTo protect AutoLISP source code:")
            print("  1. Open AutoCAD")
            print("  2. Load compile_lisp.lsp: (load \"compile_lisp.lsp\")")
            print("  3. Type: COMPILECOLUMN")
            print("  4. Re-run build.bat")
        
        print("\nNext steps:")
        print("  1. Test the release package")
        print("  2. Distribute to users")
        print("\n" + "=" * 60)
        return True
        
    except Exception as e:
        print(f"\n✗ Build failed with error: {e}")
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
