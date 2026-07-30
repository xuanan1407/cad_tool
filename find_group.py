# find_group.py
# Python script to find and highlight group of shapes by prefix in AutoCAD
# Copyright (c) 2026 Tran Xuan An

import sys
import os
import time

try:
    import win32com.client
    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False


def find_group_via_com(prefix):
    """
    Find shape group using COM automation (direct AutoCAD control)
    
    Args:
        prefix: Shape prefix (e.g., C1, C2)
        
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
        
        # Set prefix in AutoLISP global variable
        doc.SendCommand(f"(setq *COLFIND_PREFIX* \"{prefix}\")\n")
        
        # Small delay to ensure variable is set
        time.sleep(0.3)
        
        # Call COLFINDGROUP command (it will check the global variable)
        doc.SendCommand("COLFINDGROUP\n")
        
        print(f"✓ Command sent: COLFINDGROUP with prefix={prefix}")
        print("✓ All shapes with this prefix will be highlighted in AutoCAD!")
        print("✓ Press Enter in AutoCAD to remove highlights")
        
        return True
        
    except Exception as e:
        print(f"⚠ COM connection failed: {e}")
        return False


def find_group_in_autocad(prefix):
    """
    Send COLFINDGROUP command to AutoCAD with shape prefix
    
    Args:
        prefix: Shape prefix (e.g., C1, C2)
    """
    print(f"\n=== Find Shape Group in AutoCAD ===")
    print(f"Shape Prefix: {prefix}")
    
    # Try COM automation first
    if COM_AVAILABLE:
        if find_group_via_com(prefix):
            return  # Success!
        print("\n⚠ Falling back to script file method...")
    
    # Fallback: Create script file
    script_content = f"(setq *COLFIND_PREFIX* \"{prefix}\")\nCOLFINDGROUP\n"
    
    # Get current directory
    current_dir = os.path.dirname(os.path.abspath(__file__))
    script_file = os.path.join(current_dir, "_find_group_temp.scr")
    
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
        print(f"   Command: (setq *COLFIND_PREFIX* \"{prefix}\")")
        print(f"   Command: COLFINDGROUP")
        
        # Try to open the script file location
        try:
            os.startfile(current_dir)
        except:
            pass
            
    except Exception as e:
        print(f"✗ Error creating script: {e}")
        print(f"\nManually run in AutoCAD:")
        print(f"   Command: (setq *COLFIND_PREFIX* \"{prefix}\")")
        print(f"   Command: COLFINDGROUP")


def main():
    """Main entry point"""
    print("=" * 60)
    print("   CAD Column Inspector - Group Finder")
    print("=" * 60)
    
    # Check arguments
    if len(sys.argv) < 2:
        print("\nUsage: python find_group.py <prefix>")
        print("Example: python find_group.py C1")
        if not COM_AVAILABLE:
            print("\n⚠ Note: pywin32 not installed. Install for auto-execution:")
            print("   pip install pywin32")
        sys.exit(1)
    
    prefix = sys.argv[1]
    
    # Validate prefix format
    if not prefix or len(prefix) < 1:
        print("✗ Invalid prefix format")
        sys.exit(1)
    
    # Find shape group
    find_group_in_autocad(prefix)
    
    print("\n" + "=" * 60)
    # Auto-close without waiting for user input


if __name__ == "__main__":
    main()
