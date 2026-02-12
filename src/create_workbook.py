"""
BRIMIS Incident Management System - Workbook Generator
=======================================================
Generates BRIMIS_IMS.xlsm with all sheets, Excel Tables (ListObjects),
column schemas, sample data, BRIMIS branding, and formatting.

Usage:
    python src/create_workbook.py

Output:
    BRIMIS_IMS.xlsm in the project root directory

Dependencies:
    pip install openpyxl
"""

import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

# =============================================================================
# BRIMIS Brand Colors (hex codes for openpyxl PatternFill)
# =============================================================================
BRIMIS_RED = "EE3124"
BRIMIS_DARK = "101010"
BRIMIS_WHITE = "FFFFFF"
BRIMIS_GRAY = "32373C"

# =============================================================================
# Reusable Style Objects
# =============================================================================
FILL_DARK = PatternFill(start_color=BRIMIS_DARK, end_color=BRIMIS_DARK, fill_type="solid")
FILL_RED = PatternFill(start_color=BRIMIS_RED, end_color=BRIMIS_RED, fill_type="solid")
FILL_WHITE = PatternFill(start_color=BRIMIS_WHITE, end_color=BRIMIS_WHITE, fill_type="solid")

FONT_HEADER_LARGE = Font(name="Calibri", size=16, bold=True, color=BRIMIS_WHITE)
FONT_HEADER_MEDIUM = Font(name="Calibri", size=14, bold=True, color=BRIMIS_WHITE)
FONT_TABLE_HEADER = Font(name="Calibri", size=11, bold=True, color=BRIMIS_WHITE)
FONT_INSTRUCTION = Font(name="Calibri", size=11, italic=True, color="808080")
FONT_LABEL_BOLD = Font(name="Calibri", size=11, bold=True, color=BRIMIS_DARK)

# Standard table style used as a base (row stripes)
TABLE_STYLE = TableStyleInfo(
    name="TableStyleMedium2",
    showFirstColumn=False,
    showLastColumn=False,
    showRowStripes=True,
    showColumnStripes=False,
)


def create_workbook():
    """Create the BRIMIS IMS workbook with all sheets, tables, and formatting."""
    wb = Workbook()

    # =========================================================================
    # 1. Create sheets in tab order: Dashboard, Incident Log, Settings
    # =========================================================================
    # Default sheet is renamed to Dashboard
    ws_dashboard = wb.active
    ws_dashboard.title = "Dashboard"
    ws_dashboard.sheet_properties.tabColor = BRIMIS_RED

    ws_incident_log = wb.create_sheet("Incident Log")
    ws_incident_log.sheet_properties.tabColor = BRIMIS_GRAY

    ws_settings = wb.create_sheet("Settings")
    ws_settings.sheet_properties.tabColor = BRIMIS_GRAY

    # =========================================================================
    # 2. Dashboard Sheet
    # =========================================================================
    _setup_dashboard(ws_dashboard)

    # =========================================================================
    # 3. Settings Sheet
    # =========================================================================
    _setup_settings(ws_settings)

    # =========================================================================
    # 4. Incident Log Sheet
    # =========================================================================
    _setup_incident_log(ws_incident_log)

    # =========================================================================
    # 5. Save as .xlsx (VBA will be injected in Plan 03 to produce .xlsm)
    # =========================================================================
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    output_path = os.path.join(project_root, "BRIMIS_IMS.xlsx")
    wb.save(output_path)
    print(f"Workbook saved to: {output_path}")
    print(f"File size: {os.path.getsize(output_path):,} bytes")
    return output_path


# =============================================================================
# Dashboard Setup
# =============================================================================
def _setup_dashboard(ws):
    """Configure the Dashboard sheet with header bar and placeholder content."""
    # Row 1: Merged header bar A1:Z1
    ws.merge_cells("A1:Z1")
    cell = ws["A1"]
    cell.value = "BRIMIS Incident Management System"
    cell.fill = FILL_DARK
    cell.font = FONT_HEADER_LARGE
    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 36

    # Apply dark fill to all cells in the merged range
    for col in range(1, 27):  # A to Z
        c = ws.cell(row=1, column=col)
        c.fill = FILL_DARK

    # Row 3: Placeholder text
    ws["A3"].value = "Dashboard will be populated in Phase 5"
    ws["A3"].font = FONT_INSTRUCTION

    # Freeze panes at row 2
    ws.freeze_panes = "A2"

    # Column A width
    ws.column_dimensions["A"].width = 30


# =============================================================================
# Settings Sheet Setup
# =============================================================================
def _setup_settings(ws):
    """Configure the Settings sheet with all 5 configuration tables."""
    # ----- Row 1: Header bar A1:H1 -----
    ws.merge_cells("A1:H1")
    cell = ws["A1"]
    cell.value = "BRIMIS IMS - System Settings"
    cell.fill = FILL_DARK
    cell.font = FONT_HEADER_MEDIUM
    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 32

    for col in range(1, 9):  # A through H
        ws.cell(row=1, column=col).fill = FILL_DARK

    # ----- Row 2: Instructions -----
    ws["A2"].value = (
        "Administrators: Edit values in the tables below. "
        "Do NOT modify column headers."
    )
    ws["A2"].font = FONT_INSTRUCTION

    # =====================================================================
    # tblTeams at A3:C6 (header + 3 sample rows)
    # =====================================================================
    teams_headers = ["TeamName", "TeamLead", "Active"]
    teams_data = [
        ["Mechanical Team", "(Team Lead Name)", "Yes"],
        ["Electrical Team", "(Team Lead Name)", "Yes"],
        ["Safety/HSE Team", "(Team Lead Name)", "Yes"],
    ]
    _write_table_data(ws, start_row=3, start_col=1, headers=teams_headers, data=teams_data)
    _create_table(ws, "tblTeams", "A3", "C6", teams_headers)

    # Column widths for tblTeams
    ws.column_dimensions["A"].width = 20  # TeamName
    ws.column_dimensions["B"].width = 20  # TeamLead
    ws.column_dimensions["C"].width = 8   # Active

    # =====================================================================
    # tblPersonnel at E3:H6 (header + 3 sample rows)
    # =====================================================================
    personnel_headers = ["Name", "Team", "Role", "Active"]
    personnel_data = [
        ["(Personnel Name)", "Mechanical Team", "Technician", "Yes"],
        ["(Personnel Name)", "Electrical Team", "Technician", "Yes"],
        ["(Personnel Name)", "Safety/HSE Team", "Supervisor", "Yes"],
    ]
    _write_table_data(ws, start_row=3, start_col=5, headers=personnel_headers, data=personnel_data)
    _create_table(ws, "tblPersonnel", "E3", "H6", personnel_headers)

    # Column widths for tblPersonnel
    ws.column_dimensions["E"].width = 20  # Name
    ws.column_dimensions["F"].width = 20  # Team
    ws.column_dimensions["G"].width = 15  # Role
    ws.column_dimensions["H"].width = 8   # Active

    # =====================================================================
    # tblCategories at A8:C20 (header + 12 sample rows)
    # =====================================================================
    categories_headers = ["Category", "Subcategory", "Active"]
    categories_data = [
        ["Mechanical", "Pump Breakdowns", "Yes"],
        ["Mechanical", "Valve Failures", "Yes"],
        ["Mechanical", "Bearing Failures", "Yes"],
        ["Mechanical", "Pipe/Fitting Leaks", "Yes"],
        ["Electrical", "Motor Failures", "Yes"],
        ["Electrical", "Control System Faults", "Yes"],
        ["Electrical", "Power Supply Issues", "Yes"],
        ["Electrical", "Wiring/Connection Faults", "Yes"],
        ["Safety/HSE", "Near Miss", "Yes"],
        ["Safety/HSE", "First Aid Incident", "Yes"],
        ["Safety/HSE", "Environmental Spill", "Yes"],
        ["Safety/HSE", "Fire/Explosion Risk", "Yes"],
    ]
    _write_table_data(ws, start_row=8, start_col=1, headers=categories_headers, data=categories_data)
    _create_table(ws, "tblCategories", "A8", "C20", categories_headers)

    # Reuse column A width (already 20 for TeamName, but Category needs 15)
    # Keep A at 20 since TeamName is there too; Category shares column but is in A8+
    # Actually tblCategories is also in column A, so the wider width (20) is fine.

    # =====================================================================
    # tblSLAThresholds at E8:G12 (header + 4 rows)
    # =====================================================================
    sla_headers = ["Priority", "ResponseTime_Hrs", "ResolutionTime_Hrs"]
    sla_data = [
        ["P1", 1, 4],
        ["P2", 4, 24],
        ["P3", 8, 72],
        ["P4", 24, 168],
    ]
    _write_table_data(ws, start_row=8, start_col=5, headers=sla_headers, data=sla_data)
    _create_table(ws, "tblSLAThresholds", "E8", "G12", sla_headers)

    # Column widths for SLA table (E already set to 20 for Name)
    # Override for this section -- E is shared, keep wider
    ws.column_dimensions["F"].width = 18  # ResponseTime_Hrs
    ws.column_dimensions["G"].width = 18  # ResolutionTime_Hrs

    # =====================================================================
    # Priority Matrix label at A22
    # =====================================================================
    ws["A22"].value = "Priority Matrix (Impact x Urgency)"
    ws["A22"].font = FONT_LABEL_BOLD

    # =====================================================================
    # tblPriorityMatrix at A23:E27 (header + 4 rows)
    # =====================================================================
    matrix_headers = ["Impact\\Urgency", "4-Critical", "3-High", "2-Medium", "1-Low"]
    matrix_data = [
        ["4-Critical", "P1", "P1", "P2", "P3"],
        ["3-High", "P1", "P2", "P2", "P3"],
        ["2-Medium", "P2", "P2", "P3", "P4"],
        ["1-Low", "P3", "P3", "P4", "P4"],
    ]
    _write_table_data(ws, start_row=23, start_col=1, headers=matrix_headers, data=matrix_data)
    _create_table(ws, "tblPriorityMatrix", "A23", "E27", matrix_headers)

    # Column widths for priority matrix
    ws.column_dimensions["D"].width = 12  # 2-Medium
    # B, C, E already set or shared; set matrix column widths explicitly
    # B=20 (TeamLead) is wider than needed for matrix but acceptable


# =============================================================================
# Incident Log Sheet Setup
# =============================================================================
def _setup_incident_log(ws):
    """Configure the Incident Log sheet with header bar and tblIncidents."""
    # ----- Row 1: Header bar A1:AA1 -----
    ws.merge_cells("A1:AA1")
    cell = ws["A1"]
    cell.value = "BRIMIS IMS - Incident Log"
    cell.fill = FILL_DARK
    cell.font = FONT_HEADER_MEDIUM
    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 32

    for col in range(1, 28):  # A through AA (27 columns)
        ws.cell(row=1, column=col).fill = FILL_DARK

    # ----- Incident table headers (Row 2) and one empty data row (Row 3) -----
    incident_headers = [
        "IncidentID", "Title", "Description", "Category", "Subcategory",
        "Priority", "Impact", "Urgency", "Status", "ReportedBy",
        "ReportedDate", "AssignedTeam", "AssignedTo", "AssignedDate",
        "AssignedBy", "ResponseDate", "ResolutionDate", "ClosedDate",
        "ResolutionNotes", "RootCause", "CorrectiveAction",
        "PreventiveAction", "AttachmentRef", "SLAResponseStatus",
        "SLAResolutionStatus", "LastModified", "LastModifiedBy",
    ]

    # Write headers to row 2
    for col_idx, header in enumerate(incident_headers, start=1):
        ws.cell(row=2, column=col_idx, value=header)

    # Row 3: empty placeholder data row (openpyxl tables require at least 1 data row)
    # Leave cells empty -- this row serves as a structural placeholder
    for col_idx in range(1, len(incident_headers) + 1):
        ws.cell(row=3, column=col_idx, value=None)

    # Create the table tblIncidents from A2:AA3
    end_col_letter = get_column_letter(len(incident_headers))  # AA for 27 columns
    table_ref = f"A2:{end_col_letter}3"
    tbl = Table(displayName="tblIncidents", ref=table_ref)
    tbl.tableStyleInfo = TABLE_STYLE
    ws.add_table(tbl)

    # Apply BRIMIS Red header formatting to row 2
    for col_idx in range(1, len(incident_headers) + 1):
        cell = ws.cell(row=2, column=col_idx)
        cell.fill = FILL_RED
        cell.font = FONT_TABLE_HEADER
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # ----- Column widths -----
    col_widths = {
        "IncidentID": 16, "Title": 30, "Description": 40,
        "Category": 14, "Subcategory": 20, "Priority": 8,
        "Impact": 8, "Urgency": 8, "Status": 14,
        "ReportedBy": 16, "ReportedDate": 18, "AssignedTeam": 16,
        "AssignedTo": 16, "AssignedDate": 18, "AssignedBy": 16,
        "ResponseDate": 18, "ResolutionDate": 18, "ClosedDate": 18,
        "ResolutionNotes": 30, "RootCause": 30, "CorrectiveAction": 30,
        "PreventiveAction": 30, "AttachmentRef": 20,
        "SLAResponseStatus": 18, "SLAResolutionStatus": 18,
        "LastModified": 18, "LastModifiedBy": 16,
    }
    for col_idx, header in enumerate(incident_headers, start=1):
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = col_widths.get(header, 14)

    # Freeze panes at row 3 (header bar + table header always visible)
    ws.freeze_panes = "A3"


# =============================================================================
# Helper Functions
# =============================================================================
def _write_table_data(ws, start_row, start_col, headers, data):
    """Write table headers and data rows to a worksheet region."""
    # Write headers
    for col_offset, header in enumerate(headers):
        ws.cell(row=start_row, column=start_col + col_offset, value=header)

    # Write data rows
    for row_offset, row_data in enumerate(data, start=1):
        for col_offset, value in enumerate(row_data):
            ws.cell(row=start_row + row_offset, column=start_col + col_offset, value=value)


def _create_table(ws, table_name, start_ref, end_ref, headers):
    """Create a named Excel Table with BRIMIS Red header formatting."""
    table_ref = f"{start_ref}:{end_ref}"
    tbl = Table(displayName=table_name, ref=table_ref)
    tbl.tableStyleInfo = TABLE_STYLE
    ws.add_table(tbl)

    # Apply BRIMIS Red header formatting
    # Parse start_ref to get row and column
    from openpyxl.utils import coordinate_to_tuple
    start_row, start_col = coordinate_to_tuple(start_ref)

    for col_offset in range(len(headers)):
        cell = ws.cell(row=start_row, column=start_col + col_offset)
        cell.fill = FILL_RED
        cell.font = FONT_TABLE_HEADER
        cell.alignment = Alignment(horizontal="center", vertical="center")


# =============================================================================
# Entry Point
# =============================================================================
if __name__ == "__main__":
    output = create_workbook()
    print(f"\nBRIMIS IMS workbook generated successfully.")
    print(f"  Sheets: Dashboard, Incident Log, Settings")
    print(f"  Tables: tblTeams, tblPersonnel, tblCategories, tblSLAThresholds, tblPriorityMatrix, tblIncidents")
    print(f"  Output: {output}")
    print(f"  Note: Saved as .xlsx -- VBA injection (Plan 03) will convert to .xlsm")
