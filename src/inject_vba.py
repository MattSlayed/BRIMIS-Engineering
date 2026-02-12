"""
inject_vba.py
=============
Injects all VBA modules into the BRIMIS IMS workbook using win32com Excel automation.

This script:
1. Opens BRIMIS_IMS.xlsx (or .xlsm if it already exists)
2. Imports 7 standard .bas modules into the VBA project
3. Writes ThisWorkbook.cls code into the existing ThisWorkbook code module
4. Creates the frmIncidentEntry UserForm with all controls programmatically
5. Adds a "Log New Incident" button shape on the Dashboard sheet
6. Saves the workbook as BRIMIS_IMS.xlsm (macro-enabled format)

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
    "modIncidentEntry.bas",
]

# ThisWorkbook is handled specially (code written to existing component)
THISWORKBOOK_FILE = "ThisWorkbook.cls"

# UserForm code file (no Attribute VB_Name line -- injected via CodeModule.AddFromString)
FORM_CODE_FILE = "frmIncidentEntry.bas"

# Excel file format constants
XL_OPEN_XML_WORKBOOK_MACRO_ENABLED = 52  # .xlsm

# VBE component type constants
VBEXT_CT_MSFORM = 3  # vbext_ct_MSForm - UserForm


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


def create_incident_entry_form(vb_project):
    """Create the frmIncidentEntry UserForm with all controls programmatically.

    Builds a 4-step wizard form with:
      - Step Indicator label
      - MultiPage with 4 pages (Category, Details, Priority, Review)
      - Navigation buttons (Back, Next, Submit, Cancel)
      - All input controls on each page
      - Form event handler code from frmIncidentEntry.bas
    """
    form_name = "frmIncidentEntry"

    # Remove existing form if present (for re-runs)
    try:
        existing = vb_project.VBComponents(form_name)
        vb_project.VBComponents.Remove(existing)
        print(f"  Removed existing form: {form_name}")
    except com_error:
        pass  # Form does not exist yet

    # Create the UserForm component
    print(f"  Creating UserForm: {form_name}")
    form_comp = vb_project.VBComponents.Add(VBEXT_CT_MSFORM)
    form_comp.Properties("Name").Value = form_name
    form_comp.Properties("Caption").Value = "BRIMIS - Log New Incident"
    form_comp.Properties("Width").Value = 520
    form_comp.Properties("Height").Value = 450
    form_comp.Properties("BackColor").Value = 16777215  # CLR_BRIMIS_WHITE

    # Get the Designer for adding controls
    designer = form_comp.Designer

    # =====================================================================
    # Form-level controls (outside MultiPage)
    # =====================================================================

    # Step Indicator Label
    print("  Adding step indicator label...")
    lbl = designer.Controls.Add("Forms.Label.1", "lblStepIndicator", True)
    lbl.Caption = "Step 1 of 4"
    lbl.Left = 12
    lbl.Top = 8
    lbl.Width = 200
    lbl.Height = 20
    lbl.Font.Size = 12
    lbl.Font.Bold = True
    lbl.ForeColor = 2372078  # CLR_BRIMIS_RED

    # MultiPage Control
    print("  Adding MultiPage wizard control...")
    mp = designer.Controls.Add("Forms.MultiPage.1", "mpWizard", True)
    mp.Left = 6
    mp.Top = 32
    mp.Width = 504
    mp.Height = 340
    mp.Style = 2  # fmTabStyleNone - hides tab headers

    # MultiPage starts with 2 pages. Add 2 more for 4 total.
    try:
        mp.Pages.Add("Page3", "Page3", 2)
        mp.Pages.Add("Page4", "Page4", 3)
        print(f"  MultiPage has {mp.Pages.Count} pages")
    except Exception as e:
        print(f"  WARNING: Could not add pages via Pages.Add: {e}")
        print("  Form may need manual page adjustment.")

    # Navigation Buttons (below MultiPage, at bottom of form)
    print("  Adding navigation buttons...")

    # Back button
    btn_back = designer.Controls.Add("Forms.CommandButton.1", "btnBack", True)
    btn_back.Caption = "< Back"
    btn_back.Left = 12
    btn_back.Top = 382
    btn_back.Width = 80
    btn_back.Height = 30
    btn_back.Enabled = False  # Disabled on page 0
    btn_back.BackColor = 3946290  # CLR_BRIMIS_GRAY
    btn_back.ForeColor = 16777215  # White
    btn_back.BackStyle = 1  # fmBackStyleOpaque

    # Next button
    btn_next = designer.Controls.Add("Forms.CommandButton.1", "btnNext", True)
    btn_next.Caption = "Next >"
    btn_next.Left = 310
    btn_next.Top = 382
    btn_next.Width = 90
    btn_next.Height = 30
    btn_next.BackColor = 2372078  # CLR_BRIMIS_RED
    btn_next.ForeColor = 16777215  # White
    btn_next.BackStyle = 1  # fmBackStyleOpaque

    # Submit button (hidden initially)
    btn_submit = designer.Controls.Add("Forms.CommandButton.1", "btnSubmit", True)
    btn_submit.Caption = "Submit"
    btn_submit.Left = 310
    btn_submit.Top = 382
    btn_submit.Width = 90
    btn_submit.Height = 30
    btn_submit.Visible = False
    btn_submit.BackColor = 2372078  # CLR_BRIMIS_RED
    btn_submit.ForeColor = 16777215  # White
    btn_submit.BackStyle = 1  # fmBackStyleOpaque

    # Cancel button
    btn_cancel = designer.Controls.Add("Forms.CommandButton.1", "btnCancel", True)
    btn_cancel.Caption = "Cancel"
    btn_cancel.Left = 416
    btn_cancel.Top = 382
    btn_cancel.Width = 90
    btn_cancel.Height = 30
    btn_cancel.BackColor = 3946290  # CLR_BRIMIS_GRAY
    btn_cancel.ForeColor = 16777215  # White
    btn_cancel.BackStyle = 1  # fmBackStyleOpaque

    # =====================================================================
    # Page 0: Category Selection
    # =====================================================================
    print("  Adding Page 0 controls (Category Selection)...")
    page0 = mp.Pages(0)

    # Header label
    lbl = page0.Controls.Add("Forms.Label.1", "lblCategoryHeader", True)
    lbl.Caption = "Step 1: Select Category and Reporter"
    lbl.Left = 12
    lbl.Top = 8
    lbl.Width = 460
    lbl.Height = 22
    lbl.Font.Size = 11
    lbl.Font.Bold = True
    lbl.ForeColor = 1052688  # CLR_BRIMIS_DARK

    # Category label
    lbl = page0.Controls.Add("Forms.Label.1", "lblCategory", True)
    lbl.Caption = "Category: *"
    lbl.Left = 12
    lbl.Top = 50
    lbl.Width = 120
    lbl.Height = 18
    lbl.ForeColor = 1052688

    # Category combo
    cbo = page0.Controls.Add("Forms.ComboBox.1", "cboCategory", True)
    cbo.Left = 140
    cbo.Top = 48
    cbo.Width = 300
    cbo.Height = 22
    cbo.Style = 2  # fmStyleDropDownList

    # Subcategory label
    lbl = page0.Controls.Add("Forms.Label.1", "lblSubcategory", True)
    lbl.Caption = "Subcategory: *"
    lbl.Left = 12
    lbl.Top = 90
    lbl.Width = 120
    lbl.Height = 18
    lbl.ForeColor = 1052688

    # Subcategory combo
    cbo = page0.Controls.Add("Forms.ComboBox.1", "cboSubcategory", True)
    cbo.Left = 140
    cbo.Top = 88
    cbo.Width = 300
    cbo.Height = 22
    cbo.Style = 2  # fmStyleDropDownList

    # Reporter label
    lbl = page0.Controls.Add("Forms.Label.1", "lblReporter", True)
    lbl.Caption = "Reported By: *"
    lbl.Left = 12
    lbl.Top = 130
    lbl.Width = 120
    lbl.Height = 18
    lbl.ForeColor = 1052688

    # Reporter combo
    cbo = page0.Controls.Add("Forms.ComboBox.1", "cboReporter", True)
    cbo.Left = 140
    cbo.Top = 128
    cbo.Width = 300
    cbo.Height = 22
    cbo.Style = 2  # fmStyleDropDownList

    # =====================================================================
    # Page 1: Incident Details
    # =====================================================================
    print("  Adding Page 1 controls (Incident Details)...")
    page1 = mp.Pages(1)

    # Header label
    lbl = page1.Controls.Add("Forms.Label.1", "lblDetailsHeader", True)
    lbl.Caption = "Step 2: Incident Details"
    lbl.Left = 12
    lbl.Top = 8
    lbl.Width = 460
    lbl.Height = 22
    lbl.Font.Size = 11
    lbl.Font.Bold = True
    lbl.ForeColor = 1052688

    # Title label
    lbl = page1.Controls.Add("Forms.Label.1", "lblTitle", True)
    lbl.Caption = "Title: *"
    lbl.Left = 12
    lbl.Top = 50
    lbl.Width = 120
    lbl.Height = 18
    lbl.ForeColor = 1052688

    # Title textbox
    txt = page1.Controls.Add("Forms.TextBox.1", "txtTitle", True)
    txt.Left = 140
    txt.Top = 48
    txt.Width = 340
    txt.Height = 22
    txt.MaxLength = 200

    # Description label
    lbl = page1.Controls.Add("Forms.Label.1", "lblDescription", True)
    lbl.Caption = "Description: *"
    lbl.Left = 12
    lbl.Top = 90
    lbl.Width = 120
    lbl.Height = 18
    lbl.ForeColor = 1052688

    # Description textbox (multiline)
    txt = page1.Controls.Add("Forms.TextBox.1", "txtDescription", True)
    txt.Left = 12
    txt.Top = 110
    txt.Width = 468
    txt.Height = 120
    txt.MultiLine = True
    txt.ScrollBars = 2  # fmScrollBarsVertical
    txt.WordWrap = True
    txt.EnterKeyBehavior = True  # Allow Enter for new lines

    # Attachment label
    lbl = page1.Controls.Add("Forms.Label.1", "lblAttachment", True)
    lbl.Caption = "Attachment Reference:"
    lbl.Left = 12
    lbl.Top = 250
    lbl.Width = 120
    lbl.Height = 18
    lbl.ForeColor = 1052688

    # Attachment textbox
    txt = page1.Controls.Add("Forms.TextBox.1", "txtAttachment", True)
    txt.Left = 140
    txt.Top = 248
    txt.Width = 260
    txt.Height = 22

    # Browse button
    btn = page1.Controls.Add("Forms.CommandButton.1", "btnBrowse", True)
    btn.Caption = "Browse..."
    btn.Left = 410
    btn.Top = 248
    btn.Width = 70
    btn.Height = 22

    # =====================================================================
    # Page 2: Priority Assessment
    # =====================================================================
    print("  Adding Page 2 controls (Priority Assessment)...")
    page2 = mp.Pages(2)

    # Header label
    lbl = page2.Controls.Add("Forms.Label.1", "lblPriorityHeader", True)
    lbl.Caption = "Step 3: Priority Assessment"
    lbl.Left = 12
    lbl.Top = 8
    lbl.Width = 460
    lbl.Height = 22
    lbl.Font.Size = 11
    lbl.Font.Bold = True
    lbl.ForeColor = 1052688

    # Explanation label
    lbl = page2.Controls.Add("Forms.Label.1", "lblPriorityExplain", True)
    lbl.Caption = "Select the impact and urgency levels. Priority will be calculated automatically."
    lbl.Left = 12
    lbl.Top = 40
    lbl.Width = 460
    lbl.Height = 18
    lbl.ForeColor = 1052688

    # Impact label
    lbl = page2.Controls.Add("Forms.Label.1", "lblImpact", True)
    lbl.Caption = "Impact Level: *"
    lbl.Left = 12
    lbl.Top = 80
    lbl.Width = 120
    lbl.Height = 18
    lbl.ForeColor = 1052688

    # Impact combo
    cbo = page2.Controls.Add("Forms.ComboBox.1", "cboImpact", True)
    cbo.Left = 140
    cbo.Top = 78
    cbo.Width = 200
    cbo.Height = 22
    cbo.Style = 2  # fmStyleDropDownList

    # Urgency label
    lbl = page2.Controls.Add("Forms.Label.1", "lblUrgency", True)
    lbl.Caption = "Urgency Level: *"
    lbl.Left = 12
    lbl.Top = 120
    lbl.Width = 120
    lbl.Height = 18
    lbl.ForeColor = 1052688

    # Urgency combo
    cbo = page2.Controls.Add("Forms.ComboBox.1", "cboUrgency", True)
    cbo.Left = 140
    cbo.Top = 118
    cbo.Width = 200
    cbo.Height = 22
    cbo.Style = 2  # fmStyleDropDownList

    # Priority result label
    lbl = page2.Controls.Add("Forms.Label.1", "lblPriorityLabel", True)
    lbl.Caption = "Calculated Priority:"
    lbl.Left = 12
    lbl.Top = 170
    lbl.Width = 120
    lbl.Height = 22
    lbl.Font.Bold = True
    lbl.ForeColor = 1052688

    # Priority value label (shows P1, P2, etc.)
    lbl = page2.Controls.Add("Forms.Label.1", "lblPriorityValue", True)
    lbl.Caption = "(select both Impact and Urgency)"
    lbl.Left = 140
    lbl.Top = 170
    lbl.Width = 300
    lbl.Height = 22
    lbl.Font.Size = 12
    lbl.Font.Bold = True
    lbl.ForeColor = 1052688

    # =====================================================================
    # Page 3: Review and Submit
    # =====================================================================
    print("  Adding Page 3 controls (Review and Submit)...")
    page3 = mp.Pages(3)

    # Header label
    lbl = page3.Controls.Add("Forms.Label.1", "lblReviewHeader", True)
    lbl.Caption = "Step 4: Review and Submit"
    lbl.Left = 12
    lbl.Top = 8
    lbl.Width = 460
    lbl.Height = 22
    lbl.Font.Size = 11
    lbl.Font.Bold = True
    lbl.ForeColor = 1052688

    # Review labels (all positioned vertically, showing field: value)
    review_labels = [
        ("lblReviewCategory",    "Category: ",    40),
        ("lblReviewSubcategory", "Subcategory: ", 62),
        ("lblReviewReporter",    "Reported By: ", 84),
        ("lblReviewTitle",       "Title: ",       116),
        ("lblReviewDescription", "Description: ", 138),
        ("lblReviewAttachment",  "Attachment: ",  190),
        ("lblReviewPriority",    "Priority: ",    222),
        ("lblReviewImpact",      "Impact: ",      244),
        ("lblReviewUrgency",     "Urgency: ",     266),
    ]
    for name, caption, top in review_labels:
        lbl = page3.Controls.Add("Forms.Label.1", name, True)
        lbl.Caption = caption
        lbl.Left = 12
        lbl.Top = top
        lbl.Width = 468
        lbl.Height = 18
        lbl.ForeColor = 1052688
        # For description, allow more height (it may wrap)
        if name == "lblReviewDescription":
            lbl.Height = 48
            lbl.WordWrap = True

    # =====================================================================
    # Inject form code from frmIncidentEntry.bas
    # =====================================================================
    print("  Injecting form event handler code...")
    form_code_path = os.path.join(VBA_MODULES_DIR, FORM_CODE_FILE)

    if not os.path.exists(form_code_path):
        print(f"  ERROR: Form code file not found: {form_code_path}")
        return False

    with open(form_code_path, "r", encoding="utf-8") as f:
        form_code = f.read()

    # Strip any Attribute lines (there should be none, but be safe)
    clean_lines = [
        line for line in form_code.split('\n')
        if not line.strip().startswith('Attribute ')
    ]
    clean_code = '\n'.join(clean_lines)

    # Inject into form's CodeModule
    code_module = form_comp.CodeModule
    if code_module.CountOfLines > 0:
        code_module.DeleteLines(1, code_module.CountOfLines)
    code_module.AddFromString(clean_code)

    print(f"  UserForm {form_name} created successfully with all controls and code")
    return True


def add_dashboard_button(wb):
    """Add a 'Log New Incident' button shape on the Dashboard sheet.

    Creates a BRIMIS-branded rounded rectangle shape that calls
    ShowIncidentEntryForm when clicked.
    """
    dashboard = wb.Sheets("Dashboard")

    # Remove existing button if present (for re-runs)
    try:
        dashboard.Shapes("btnLogIncident").Delete()
        print("  Removed existing Dashboard button")
    except Exception:
        pass

    # msoShapeRoundedRectangle = 5
    # Position: Left=30, Top=60 (below header bar), Width=220, Height=50
    shape = dashboard.Shapes.AddShape(5, 30, 60, 220, 50)
    shape.Name = "btnLogIncident"

    # Style the shape with BRIMIS branding
    # CLR_BRIMIS_RED = 2372078 (already in Long/BGR format for OLE_COLOR)
    shape.Fill.ForeColor.RGB = 2372078
    shape.Line.Visible = False

    # Text formatting
    tf = shape.TextFrame2
    tf.TextRange.Text = "Log New Incident"
    tf.TextRange.Font.Size = 14
    tf.TextRange.Font.Bold = True
    tf.TextRange.Font.Fill.ForeColor.RGB = 16777215  # White
    tf.VerticalAnchor = 3  # msoAnchorMiddle
    tf.TextRange.ParagraphFormat.Alignment = 2  # msoAlignCenter

    # Assign the macro
    shape.OnAction = "ShowIncidentEntryForm"

    print("  Created Dashboard button: btnLogIncident -> ShowIncidentEntryForm")
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

        # Create the UserForm
        print()
        print("Creating incident entry UserForm...")
        form_ok = False
        try:
            form_ok = create_incident_entry_form(wb.VBProject)
        except Exception as e:
            print(f"  ERROR creating UserForm: {e}")

        # Add Dashboard button
        print()
        print("Adding Dashboard launch button...")
        button_ok = False
        try:
            button_ok = add_dashboard_button(wb)
        except Exception as e:
            print(f"  ERROR adding Dashboard button: {e}")

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
        print(f"  UserForm created: {'Yes' if form_ok else 'No'}")
        print(f"  Dashboard button added: {'Yes' if button_ok else 'No'}")
        total_ok = imported_count + (1 if thisworkbook_ok else 0) + (1 if form_ok else 0)
        total_expected = len(STANDARD_MODULES) + 2  # modules + ThisWorkbook + UserForm
        print(f"  Total VBA components: {total_ok}/{total_expected}")
        print(f"  Output file: {XLSM_PATH}")

        if os.path.exists(XLSM_PATH):
            size_kb = os.path.getsize(XLSM_PATH) / 1024
            print(f"  File size: {size_kb:.1f} KB")

        all_ok = (imported_count == len(STANDARD_MODULES) and thisworkbook_ok
                  and form_ok and button_ok)
        print()
        if all_ok:
            print("SUCCESS: All VBA modules, UserForm, and Dashboard button injected.")
        else:
            print("WARNING: Some components may not have been created. Check output above.")

        return all_ok

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
