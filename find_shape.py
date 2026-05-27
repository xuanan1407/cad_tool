# find_shape.py
# Python script to find and highlight shape in AutoCAD
# Copyright (c) 2026 Tran Xuan An

import sys
import os
import subprocess
import time

try:
    import win32com.client
    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False


def find_shape_via_com(shape_id):
    """
    Find shape using COM automation (direct AutoCAD control)
    
    Args:
        shape_id: Shape identifier
        
    Returns:
        True if successful, False otherwise
    """
    if not COM_AVAILABLE:
        return False
    
    try:
        print("\n⚡ Attempting direct AutoCAD connection...")
        
        # Connect to running AutoCAD instance
        acad = win32com.client.Dispatch("AutoCAD.Application")
        doc = acad.ActiveDocument
        
        print("✓ Connected to AutoCAD")
        
        # Set shape ID in AutoLISP global variable
        doc.SendCommand(f"(setq *COLFIND_SHAPEID* \"{shape_id}\")\n")
        
        # Small delay to ensure variable is set
        time.sleep(0.3)
        
        # Call COLFINDPOLY command (it will check the global variable)
        doc.SendCommand("COLFINDPOLY\n")
        
        print(f"✓ Command sent: COLFINDPOLY with shape_id={shape_id}")
        print("✓ Shape should be highlighted in AutoCAD!")
        print("✓ Shape should be highlighted in AutoCAD!")
        
        return True
        
    except Exception as e:
        print(f"⚠ COM connection failed: {e}")
        return False


def find_shape_in_autocad(shape_id):
    """
    Send FINDPOLY command to AutoCAD with shape ID
    
    Args:
        shape_id: Shape identifier (e.g., C2_001)
    """
    print(f"\n=== Find Shape in AutoCAD ===")
    print(f"Shape ID: {shape_id}")
    
    # Try COM automation first
    if COM_AVAILABLE:
        if find_shape_via_com(shape_id):
            return  # Success!
        print("\n⚠ Falling back to script file method...")
    
    # Fallback: Create script file
    script_content = f"(setq *COLFIND_SHAPEID* \"{shape_id}\")\nCOLFINDPOLY\n"
    
    # Get current directory
    current_dir = os.path.dirname(os.path.abspath(__file__))
    script_file = os.path.join(current_dir, "_find_shape_temp.scr")
    
    try:
        # Write script file
        with open(script_file, 'w') as f:
            f.write(script_content)
        
        print(f"✓ Script created: {script_file}")
        print("\nInstructions:")
        print("1. Switch to AutoCAD")
        print("2. Type: SCRIPT")
        print(f"3. Select file: {script_file}")
        print("\nOr manually type in AutoCAD:")
        print(f"   Command: (setq *COLFIND_SHAPEID* \"{shape_id}\")")
        print(f"   Command: COLFINDPOLY")
        
        # Try to open the script file location
        try:
            os.startfile(current_dir)
        except:
            pass
            
    except Exception as e:
        print(f"✗ Error creating script: {e}")
        print(f"\nManually run in AutoCAD:")
        print(f"   Command: (setq *COLFIND_SHAPEID* \"{shape_id}\")")
        print(f"   Command: COLFINDPOLY")


def main():
    """Main entry point"""
    print("=" * 60)
    print("   CAD Column Inspector - Shape Finder")
    print("=" * 60)
    
    # Check arguments
    if len(sys.argv) < 2:
        print("\nUsage: python find_shape.py <shape_id>")
        print("Example: python find_shape.py C2_001")
        if not COM_AVAILABLE:
            print("\n⚠ Note: pywin32 not installed. Install for auto-execution:")
            print("   pip install pywin32")
        sys.exit(1)
    
    shape_id = sys.argv[1]
    
    # Validate shape ID format
    if not shape_id or len(shape_id) < 3:
        print("✗ Invalid shape ID format")
        sys.exit(1)
    
    # Find shape
    find_shape_in_autocad(shape_id)
    
    print("\n" + "=" * 60)
    # Auto-close without waiting for user input
    # input("Press Enter to close...")


if __name__ == "__main__":
    main()
