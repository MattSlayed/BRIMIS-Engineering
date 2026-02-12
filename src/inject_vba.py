"""
inject_vba.py
=============
Injects all VBA modules into the BRIMIS IMS workbook using win32com Excel automation.

This script:
1. Opens BRIMIS_IMS.xlsx (or .xlsm if it already exists)
2. Imports 6 standard .bas modules into the VBA project
3. Writes ThisWorkbook.cls code into the existing ThisWorkbook code module
4. Saves the workbook as BRIMIS_IMS.xlsm (macro-enabled format)

Prerequisites:
- Python 3.x with pywin32 (pip install pywin32)
- Excel must be installed on this machine
- "Trust access to the VBA project object model" must be enabled in Excel:
    File > Options > Trust Center > Trust Center Settings > Macro Settings
    > check "Trust access to the VBA project object model"
- The workbook must be CLOSED in Excel before running this script

Usage:
    python src/inject_vba.py
"""

import os
import sys
import shutil
import time

try:
    import win32com.client
    from pywintypes import com_error
except ImportError:
    print("ERROR: pywin32 is not installed.")
    print("Install it with: pip install pywin32")
    sys.exit(1)


# ---------------------------------------------------------------------------
# Path Configuration
# ---------------------------------------------------------------------------
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
VBA_MODULES_DIR = os.path.join(PROJECT_ROOT, "src", "vba_modules")

XLSX_PATH = os.path.join(PROJECT_ROOT, "BRIMIS_IMS.xlsx")
XLSM_PATH = os.path.join(PROJECT_ROOT, "BRIMIS_IMS.xlsm")

# Standard .bas modules to import (order does not matter for import)
STANDARD_MODULES = [
    "modConstants.bas",
    "modErrorHandler.bas",
    "modUtilities.bas",
    "modFormatting.bas",
    "modDataAccess.bas",
    "modInitialize.bas",
]

# ThisWorkbook is handled specially (code written to existing component)
THISWORKBOOK_FILE = "ThisWorkbook.cls"

# Excel file format constants
XL_OPEN_XML_WORKBOOK_MACRO_ENABLED = 52  # .xlsm


def get_source_path():
    """Determine which source workbook to open (.xlsm preferred, .xlsx fallback)."""
    if os.path.exists(XLSM_PATH):
        return XLSM_PATH
    elif os.path.exists(XLSX_PATH):
        return XLSX_PATH
    else:
        print(f"ERROR: No workbook found at:")
        print(f"  {XLSM_PATH}")
        print(f"  {XLSX_PATH}")
        sys.exit(1)


def remove_existing_module(vb_project, module_name):
    """Remove a VBA module by name if it exists, so we can re-import cleanly."""
    try:
        existing = vb_project.VBComponents(module_name)
        vb_project.VBComponents.Remove(existing)
        print(f"  Removed existing module: {module_name}")
    except com_error:
        pass  # Module does not exist yet -- that's fine


def import_standard_module(vb_project, bas_file):
    """Import a standard .bas module into the VBA project using AddFromString.

    Uses AddFromString instead of Import to avoid VBA auto-renaming modules
    when Attribute VB_Name conflicts with an existing module name.
    """
    module_name = os.path.splitext(bas_file)[0]
    full_path = os.path.join(VBA_MODULES_DIR, bas_file)

    if not os.path.exists(full_path):
        print(f"  WARNING: File not found: {full_path}")
        return False

    # Remove existing module with same name (if any)
    remove_existing_module(vb_project, module_name)
    # Also remove any auto-renamed duplicates (e.g., modConstants1)
    remove_existing_module(vb_project, module_name + "1")

    # Read the .bas file and strip Attribute lines
    with open(full_path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    code_lines = [line for line in lines if not line.strip().startswith("Attribute ")]
    code_content = "".join(code_lines)

    # Create a new empty module with the correct name
    new_module = vb_project.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
    new_module.Name = module_name

    # Add code to the module
    new_module.CodeModule.AddFromString(code_content)
    print(f"  Created: {bas_file} -> {module_name}")
    return True


def inject_thisworkbook_code(vb_project):
    """
    Write ThisWorkbook.cls code into the existing ThisWorkbook component.

    ThisWorkbook is a pre-existing component in every workbook -- it cannot
    be imported as a new component. We read the .cls file, strip Attribute
    lines, and write the remaining code to the existing CodeModule.
    """
    cls_path = os.path.join(VBA_MODULES_DIR, THISWORKBOOK_FILE)

    if not os.path.exists(cls_path):
        print(f"  WARNING: File not found: {cls_path}")
        return False

    # Read the .cls file
    with open(cls_path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    # Strip Attribute VB_* lines -- keep everything from Option Explicit onward
    code_lines = []
    found_option_explicit = False
    for line in lines:
        if not found_option_explicit:
            if line.strip().startswith("Option Explicit"):
                found_option_explicit = True
                code_lines.append(line)
        else:
            code_lines.append(line)

    if not code_lines:
        print("  WARNING: No code found in ThisWorkbook.cls after stripping Attribute lines")
        return False

    code_content = "".join(code_lines)

    # Get the existing ThisWorkbook component
    try:
        this_wb = vb_project.VBComponents("ThisWorkbook")
    except com_error:
        print("  ERROR: Could not access ThisWorkbook component")
        return False

    # Clear existing code
    code_module = this_wb.CodeModule
    if code_module.CountOfLines > 0:
        code_module.DeleteLines(1, code_module.CountOfLines)

    # Add the new code
    code_module.AddFromString(code_content)
    print(f"  Injected: {THISWORKBOOK_FILE} -> ThisWorkbook.CodeModule")
    return True


def main():
    print("=" * 60)
    print("BRIMIS IMS - VBA Module Injection Script")
    print("=" * 60)
    print()

    # Determine source workbook
    source_path = get_source_path()
    print(f"Source workbook: {source_path}")
    print(f"Output workbook: {XLSM_PATH}")
    print(f"VBA modules dir: {VBA_MODULES_DIR}")
    print()

    # If source is .xlsx, copy to .xlsm first (we need to save as macro-enabled)
    if source_path == XLSX_PATH and source_path != XLSM_PATH:
        print("Converting .xlsx to .xlsm format...")
        # We'll handle the format conversion via SaveAs after opening

    # Start Excel COM automation
    print("Starting Excel...")
    xl = None
    wb = None

    try:
        xl = win32com.client.Dispatch("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False

        # Open the workbook
        print(f"Opening workbook: {os.path.basename(source_path)}")
        wb = xl.Workbooks.Open(source_path)

        # Test VBA project access
        print("Checking VBA project access...")
        try:
            _ = wb.VBProject.VBComponents.Count
            print(f"  VBA project accessible ({wb.VBProject.VBComponents.Count} existing components)")
        except com_error:
            print()
            print("ERROR: Cannot access VBA project object model.")
            print()
            print("Please enable 'Trust access to the VBA project object model' in Excel:")
            print("  1. Open Excel")
            print("  2. Go to: File > Options > Trust Center > Trust Center Settings")
            print("  3. Click 'Macro Settings'")
            print("  4. Check 'Trust access to the VBA project object model'")
            print("  5. Click OK, close Excel, and re-run this script")
            return False

        # Import standard .bas modules
        print()
        print("Importing standard modules...")
        imported_count = 0
        for bas_file in STANDARD_MODULES:
            if import_standard_module(wb.VBProject, bas_file):
                imported_count += 1

        # Inject ThisWorkbook code
        print()
        print("Injecting ThisWorkbook code...")
        thisworkbook_ok = inject_thisworkbook_code(wb.VBProject)

        # Save as .xlsm (macro-enabled format)
        print()
        if source_path == XLSX_PATH:
            print(f"Saving as macro-enabled workbook: {os.path.basename(XLSM_PATH)}")
            wb.SaveAs(XLSM_PATH, FileFormat=XL_OPEN_XML_WORKBOOK_MACRO_ENABLED)
        else:
            print(f"Saving workbook: {os.path.basename(XLSM_PATH)}")
            wb.Save()

        print()
        print("=" * 60)
        print("INJECTION SUMMARY")
        print("=" * 60)
        print(f"  Standard modules imported: {imported_count}/{len(STANDARD_MODULES)}")
        print(f"  ThisWorkbook code injected: {'Yes' if thisworkbook_ok else 'No'}")
        print(f"  Total components: {imported_count + (1 if thisworkbook_ok else 0)}/7")
        print(f"  Output file: {XLSM_PATH}")

        if os.path.exists(XLSM_PATH):
            size_kb = os.path.getsize(XLSM_PATH) / 1024
            print(f"  File size: {size_kb:.1f} KB")

        print()
        if imported_count == len(STANDARD_MODULES) and thisworkbook_ok:
            print("SUCCESS: All VBA modules injected successfully.")
        else:
            print("WARNING: Some modules may not have been injected. Check output above.")

        return imported_count == len(STANDARD_MODULES) and thisworkbook_ok

    except com_error as e:
        print(f"\nCOM ERROR: {e}")
        print("Make sure Excel is installed and the workbook is not open in another instance.")
        return False
    except Exception as e:
        print(f"\nERROR: {e}")
        return False
    finally:
        # Clean up COM objects
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if xl is not None:
                xl.Quit()
        except Exception:
            pass

        # Release COM references
        wb = None
        xl = None

        # Give Excel time to shut down
        time.sleep(1)


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
