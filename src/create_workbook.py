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
import openpyxl.worksheet.page

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
FILL_ORANGE = PatternFill(start_color="FF8C00", end_color="FF8C00", fill_type="solid")
FILL_GOLD = PatternFill(start_color="FFD700", end_color="FFD700", fill_type="solid")
FILL_GRAY = PatternFill(start_color=BRIMIS_GRAY, end_color=BRIMIS_GRAY, fill_type="solid")

FONT_HEADER_LARGE = Font(name="Calibri", size=16, bold=True, color=BRIMIS_WHITE)
FONT_HEADER_MEDIUM = Font(name="Calibri", size=14, bold=True, color=BRIMIS_WHITE)
FONT_TABLE_HEADER = Font(name="Calibri", size=11, bold=True, color=BRIMIS_WHITE)
FONT_INSTRUCTION = Font(name="Calibri", size=11, italic=True, color="808080")
FONT_LABEL_BOLD = Font(name="Calibri", size=11, bold=True, color=BRIMIS_DARK)
FONT_KPI_VALUE = Font(name="Calibri", size=28, bold=True, color=BRIMIS_DARK)
FONT_KPI_LABEL = Font(name="Calibri", size=10, bold=True, color=BRIMIS_GRAY)
FONT_SECTION_HEADER = Font(name="Calibri", size=12, bold=True, color=BRIMIS_DARK)
FONT_SUBSECTION = Font(name="Calibri", size=11, bold=True, color=BRIMIS_DARK)
FONT_SLA_HEADER = Font(name="Calibri", size=12, bold=True, color=BRIMIS_WHITE)
FONT_SLA_COL_HEADER = Font(name="Calibri", size=10, bold=True, color=BRIMIS_WHITE)
FONT_REFRESH_LABEL = Font(name="Calibri", size=9, italic=True, color="808080")

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

    ws_assignment_tracker = wb.create_sheet("Assignment Tracker")
    ws_assignment_tracker.sheet_properties.tabColor = BRIMIS_GRAY

    ws_rca_log = wb.create_sheet("RCA Log")
    ws_rca_log.sheet_properties.tabColor = BRIMIS_GRAY

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
    # 5. Assignment Tracker Sheet
    # =========================================================================
    _setup_assignment_tracker(ws_assignment_tracker)

    # =========================================================================
    # 6. RCA Log Sheet
    # =========================================================================
    _setup_rca_log(ws_rca_log)

    # =========================================================================
    # 7. IncidentReport Sheet (hidden report template)
    # =========================================================================
    ws_report = wb.create_sheet("IncidentReport")
    _setup_incident_report(ws_report)
    ws_report.sheet_state = 'veryHidden'  # xlSheetVeryHidden equivalent in openpyxl

    # =========================================================================
    # 8. Save as .xlsx (VBA will be injected separately to produce .xlsm)
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
    """Configure the Dashboard sheet with KPI cards, breakdowns, and SLA Monitor layout."""
    CENTER = Alignment(horizontal="center", vertical="center")

    # ---- Row 1: Merged header bar A1:Z1 ----
    ws.merge_cells("A1:Z1")
    cell = ws["A1"]
    cell.value = "BRIMIS Incident Management System"
    cell.fill = FILL_DARK
    cell.font = FONT_HEADER_LARGE
    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 36

    for col in range(1, 27):  # A to Z
        ws.cell(row=1, column=col).fill = FILL_DARK

    # ---- Row 5: KPI section header ----
    ws.merge_cells("B5:P5")
    cell = ws["B5"]
    cell.value = "Key Performance Indicators"
    cell.font = FONT_SECTION_HEADER

    # ---- Row 6: KPI labels at B, F, J, N ----
    kpi_labels = [
        ("B6", "Total Open"),
        ("F6", "Overdue"),
        ("J6", "Avg Resolution (days)"),
        ("N6", "Total Closed"),
    ]
    for ref, label in kpi_labels:
        c = ws[ref]
        c.value = label
        c.font = FONT_KPI_LABEL
        c.alignment = CENTER

    # ---- Row 7: KPI value placeholders (0) at B, F, J, N ----
    for ref in ["B7", "F7", "J7", "N7"]:
        c = ws[ref]
        c.value = 0
        c.font = FONT_KPI_VALUE
        c.alignment = CENTER

    # ---- Row 9: Subsection headers ----
    ws["B9"].value = "By Priority"
    ws["B9"].font = FONT_SUBSECTION
    ws["H9"].value = "By Category"
    ws["H9"].font = FONT_SUBSECTION

    # ---- Rows 10-13: Priority breakdown (B=label, C=value) ----
    priority_rows = [
        (10, "P1", FILL_RED, BRIMIS_WHITE),
        (11, "P2", FILL_ORANGE, BRIMIS_WHITE),
        (12, "P3", FILL_GOLD, BRIMIS_DARK),
        (13, "P4", FILL_GRAY, BRIMIS_WHITE),
    ]
    for row, label, fill, font_color in priority_rows:
        lbl_cell = ws.cell(row=row, column=2)  # B
        lbl_cell.value = label
        lbl_cell.fill = fill
        lbl_cell.font = Font(name="Calibri", size=11, bold=True, color=font_color)
        lbl_cell.alignment = CENTER

        val_cell = ws.cell(row=row, column=3)  # C
        val_cell.value = 0
        val_cell.font = Font(name="Calibri", size=11, bold=True, color=BRIMIS_DARK)
        val_cell.alignment = CENTER

    # ---- Rows 10-12: Category breakdown (H=label, I=value) ----
    category_rows = [
        (10, "Mechanical"),
        (11, "Electrical"),
        (12, "Safety/HSE"),
    ]
    for row, label in category_rows:
        lbl_cell = ws.cell(row=row, column=8)  # H
        lbl_cell.value = label
        lbl_cell.font = Font(name="Calibri", size=11, color=BRIMIS_DARK)

        val_cell = ws.cell(row=row, column=9)  # I
        val_cell.value = 0
        val_cell.font = Font(name="Calibri", size=11, color=BRIMIS_DARK)
        val_cell.alignment = CENTER

    # ---- Row 16: SLA Monitor header ----
    ws.merge_cells("B16:H16")
    cell = ws["B16"]
    cell.value = "SLA Monitor - Open Incidents"
    cell.fill = FILL_RED
    cell.font = FONT_SLA_HEADER
    cell.alignment = CENTER
    # Fill all cells in merged range
    for col in range(2, 9):  # B through H
        ws.cell(row=16, column=col).fill = FILL_RED

    # ---- Row 17: SLA Monitor column headers ----
    sla_col_headers = ["ID", "Title", "Priority", "Status",
                       "Response SLA", "Resolution SLA", "Time Remaining"]
    for idx, header in enumerate(sla_col_headers, start=2):  # B=2 through H=8
        c = ws.cell(row=17, column=idx)
        c.value = header
        c.fill = FILL_GRAY
        c.font = FONT_SLA_COL_HEADER
        c.alignment = CENTER

    # Rows 18-37: empty (VBA populates SLA Monitor rows, up to 20 rows)

    # ---- Row 39: Last Refreshed label ----
    ws["B39"].value = "Last Refreshed:"
    ws["B39"].font = FONT_REFRESH_LABEL

    # ---- Column widths ----
    ws.column_dimensions["A"].width = 2    # narrow gutter
    ws.column_dimensions["B"].width = 18   # ID / labels
    ws.column_dimensions["C"].width = 35   # Title / values
    ws.column_dimensions["D"].width = 12   # Priority
    ws.column_dimensions["E"].width = 14   # Status
    ws.column_dimensions["F"].width = 18   # Response SLA / KPI
    ws.column_dimensions["G"].width = 18   # Resolution SLA
    ws.column_dimensions["H"].width = 18   # Time Remaining / category labels
    ws.column_dimensions["I"].width = 12   # category values
    for letter in ["J", "K", "L", "M", "N"]:
        ws.column_dimensions[letter].width = 14  # KPI columns

    # Freeze panes at row 2
    ws.freeze_panes = "A2"


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
# Assignment Tracker Sheet Setup
# =============================================================================
def _setup_assignment_tracker(ws):
    """Configure the Assignment Tracker sheet with header bar and tblAssignmentTracker."""
    # ----- Row 1: Header bar A1:H1 -----
    ws.merge_cells("A1:H1")
    cell = ws["A1"]
    cell.value = "BRIMIS IMS - Assignment Tracker"
    cell.fill = FILL_DARK
    cell.font = FONT_HEADER_MEDIUM
    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 32

    for col in range(1, 9):  # A through H
        ws.cell(row=1, column=col).fill = FILL_DARK

    # ----- Assignment Tracker table (Row 2 headers + Row 3 empty data row) -----
    tracker_headers = [
        "AssignedTeam", "AssignedTo", "IncidentID", "Title",
        "Priority", "Status", "AssignedDate", "DaysOpen",
    ]

    # Write headers to row 2
    for col_idx, header in enumerate(tracker_headers, start=1):
        ws.cell(row=2, column=col_idx, value=header)

    # Row 3: empty placeholder data row (openpyxl tables require at least 1 data row)
    for col_idx in range(1, len(tracker_headers) + 1):
        ws.cell(row=3, column=col_idx, value=None)

    # Create the table tblAssignmentTracker from A2:H3
    end_col_letter = get_column_letter(len(tracker_headers))  # H for 8 columns
    table_ref = f"A2:{end_col_letter}3"
    tbl = Table(displayName="tblAssignmentTracker", ref=table_ref)
    tbl.tableStyleInfo = TABLE_STYLE
    ws.add_table(tbl)

    # Apply BRIMIS Red header formatting to row 2
    for col_idx in range(1, len(tracker_headers) + 1):
        cell = ws.cell(row=2, column=col_idx)
        cell.fill = FILL_RED
        cell.font = FONT_TABLE_HEADER
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # ----- Column widths -----
    tracker_col_widths = {
        "AssignedTeam": 18, "AssignedTo": 18, "IncidentID": 14,
        "Title": 30, "Priority": 10, "Status": 14,
        "AssignedDate": 18, "DaysOpen": 10,
    }
    for col_idx, header in enumerate(tracker_headers, start=1):
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = tracker_col_widths.get(header, 14)

    # Freeze panes at row 3 (header bar + table header always visible)
    ws.freeze_panes = "A3"


# =============================================================================
# RCA Log Sheet Setup
# =============================================================================
def _setup_rca_log(ws):
    """Configure the RCA Log sheet with header bar and tblRCALog."""
    # ----- Row 1: Header bar A1:J1 -----
    ws.merge_cells("A1:J1")
    cell = ws["A1"]
    cell.value = "BRIMIS IMS - RCA Log"
    cell.fill = FILL_DARK
    cell.font = FONT_HEADER_MEDIUM
    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 32

    for col in range(1, 11):  # A through J (10 columns)
        ws.cell(row=1, column=col).fill = FILL_DARK

    # ----- RCA Log table (Row 2 headers + Row 3 empty data row) -----
    # CRITICAL: Header strings MUST match the COL_RCA_* constant values in modConstants.bas EXACTLY
    rca_headers = [
        "RCAID",            # Auto-generated RCA-NNNNN
        "IncidentID",       # Foreign key to tblIncidents
        "IncidentTitle",    # Denormalized for readability
        "RootCause",        # What caused the incident
        "CorrectiveAction", # What was done to fix it
        "PreventiveAction", # What will prevent recurrence
        "ResolutionNotes",  # Additional resolution details
        "ResolvedBy",       # Who resolved it (Application.UserName)
        "ResolvedDate",     # When it was resolved (Now)
        "LastModified",     # Audit timestamp (Now)
    ]

    # Write headers to row 2
    for col_idx, header in enumerate(rca_headers, start=1):
        ws.cell(row=2, column=col_idx, value=header)

    # Row 3: empty placeholder data row (openpyxl tables require at least 1 data row)
    for col_idx in range(1, len(rca_headers) + 1):
        ws.cell(row=3, column=col_idx, value=None)

    # Create the table tblRCALog from A2:J3
    end_col_letter = get_column_letter(len(rca_headers))  # J for 10 columns
    table_ref = f"A2:{end_col_letter}3"
    tbl = Table(displayName="tblRCALog", ref=table_ref)
    tbl.tableStyleInfo = TABLE_STYLE
    ws.add_table(tbl)

    # Apply BRIMIS Red header formatting to row 2
    for col_idx in range(1, len(rca_headers) + 1):
        cell = ws.cell(row=2, column=col_idx)
        cell.fill = FILL_RED
        cell.font = FONT_TABLE_HEADER
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # ----- Column widths -----
    rca_col_widths = {
        "RCAID": 14, "IncidentID": 14, "IncidentTitle": 30,
        "RootCause": 40, "CorrectiveAction": 40, "PreventiveAction": 40,
        "ResolutionNotes": 30, "ResolvedBy": 16, "ResolvedDate": 18,
        "LastModified": 18,
    }
    for col_idx, header in enumerate(rca_headers, start=1):
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = rca_col_widths.get(header, 14)

    # Freeze panes at row 3 (header bar + table header always visible)
    ws.freeze_panes = "A3"


# =============================================================================
# IncidentReport Template Sheet Setup
# =============================================================================
def _write_report_section_header(ws, row, text):
    """Write a section header bar across A:H on the given row."""
    ws.merge_cells(f'A{row}:H{row}')
    cell = ws[f'A{row}']
    cell.value = text
    cell.font = Font(name="Calibri", size=11, bold=True, color=BRIMIS_WHITE)
    cell.fill = FILL_DARK
    cell.alignment = Alignment(horizontal='left', vertical='center')


def _setup_incident_report(ws):
    """Set up the IncidentReport template sheet (VeryHidden).

    This sheet is populated by modReports.GenerateIncidentReport at runtime.
    The layout provides section headers, label cells, and value cells that
    VBA writes into. The logo is embedded once here and persists.
    """
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    # Column widths
    ws.column_dimensions['A'].width = 18   # Labels
    ws.column_dimensions['B'].width = 22   # Values
    ws.column_dimensions['C'].width = 5    # Spacer
    ws.column_dimensions['D'].width = 18   # Labels
    ws.column_dimensions['E'].width = 22   # Values
    ws.column_dimensions['F'].width = 5    # Spacer
    ws.column_dimensions['G'].width = 15   # Extra
    ws.column_dimensions['H'].width = 15   # Extra

    # --- Logo ---
    # Try to embed BRIMIS logo if it exists
    logo_path = os.path.join(project_root, "src", "assets", "brimis_logo.png")
    if os.path.exists(logo_path):
        from openpyxl.drawing.image import Image
        img = Image(logo_path)
        img.width = 160
        img.height = 50
        ws.add_image(img, "A1")
        print(f"  Embedded logo from {logo_path}")
    else:
        # Fallback: text-only branding
        ws['A1'] = "BRIMIS"
        ws['A1'].font = Font(name="Calibri", size=16, bold=True, color=BRIMIS_RED)
        print(f"  Logo not found at {logo_path}, using text fallback")

    # --- Report Title ---
    ws['D1'] = "INCIDENT REPORT"
    ws['D1'].font = Font(name="Calibri", size=18, bold=True, color=BRIMIS_RED)
    ws['G1'] = "Report Date:"
    ws['G1'].font = Font(name="Calibri", size=9, italic=True, color="808080")

    ws['D2'] = "BRIMIS Engineering"
    ws['D2'].font = Font(name="Calibri", size=11, color=BRIMIS_GRAY)

    # --- Section: INCIDENT OVERVIEW (Row 4) ---
    _write_report_section_header(ws, 4, "INCIDENT OVERVIEW")

    # Labels for overview section
    for row, col, label in [
        (5, 'A', 'Incident ID:'), (5, 'D', 'Status:'),
        (6, 'A', 'Title:'),
        (7, 'A', 'Priority:'), (7, 'D', 'Category:'),
        (8, 'A', 'Subcategory:'),
    ]:
        ws[f'{col}{row}'] = label
        ws[f'{col}{row}'].font = Font(name="Calibri", size=10, bold=True, color=BRIMIS_GRAY)

    # Merge title value cells B6:H6
    ws.merge_cells('B6:H6')
    ws['B6'].alignment = Alignment(wrap_text=True, vertical='top')

    # --- Section: DESCRIPTION (Row 10) ---
    _write_report_section_header(ws, 10, "DESCRIPTION")

    # Merge description area A11:H14
    ws.merge_cells('A11:H14')
    ws['A11'].alignment = Alignment(wrap_text=True, vertical='top')

    # --- Section: TIMELINE (Row 16) ---
    _write_report_section_header(ws, 16, "TIMELINE")

    for row, col, label in [
        (17, 'A', 'Reported:'), (17, 'D', 'By:'),
        (18, 'A', 'Assigned:'), (18, 'D', 'To:'),
        (19, 'A', 'Response:'),
        (20, 'A', 'Resolution:'),
        (21, 'A', 'Closed:'),
    ]:
        ws[f'{col}{row}'] = label
        ws[f'{col}{row}'].font = Font(name="Calibri", size=10, bold=True, color=BRIMIS_GRAY)

    # --- Section: ASSIGNMENT (Row 23) ---
    _write_report_section_header(ws, 23, "ASSIGNMENT")

    for row, col, label in [
        (24, 'A', 'Team:'), (24, 'D', 'Assigned To:'),
        (25, 'A', 'Assigned Date:'), (25, 'D', 'Assigned By:'),
    ]:
        ws[f'{col}{row}'] = label
        ws[f'{col}{row}'].font = Font(name="Calibri", size=10, bold=True, color=BRIMIS_GRAY)

    # --- Section: RESOLUTION & RCA (Row 27) ---
    _write_report_section_header(ws, 27, "RESOLUTION & ROOT CAUSE ANALYSIS")

    for row, col, label in [
        (28, 'A', 'Root Cause:'),
        (29, 'A', 'Corrective Action:'),
        (30, 'A', 'Preventive Action:'),
        (31, 'A', 'Resolution Notes:'),
    ]:
        ws[f'{col}{row}'] = label
        ws[f'{col}{row}'].font = Font(name="Calibri", size=10, bold=True, color=BRIMIS_GRAY)

    # Merge value cells for RCA fields
    for row in [28, 29, 30, 31]:
        ws.merge_cells(f'B{row}:H{row}')
        ws[f'B{row}'].alignment = Alignment(wrap_text=True, vertical='top')

    # --- Section: SLA PERFORMANCE (Row 33) ---
    _write_report_section_header(ws, 33, "SLA PERFORMANCE")

    for row, col, label in [
        (34, 'A', 'Response SLA:'), (34, 'D', 'Resolution SLA:'),
    ]:
        ws[f'{col}{row}'] = label
        ws[f'{col}{row}'].font = Font(name="Calibri", size=10, bold=True, color=BRIMIS_GRAY)

    # --- Print setup ---
    ws.page_setup.orientation = 'portrait'
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.print_area = 'A1:H35'
    ws.page_margins = openpyxl.worksheet.page.PageMargins(
        left=0.5, right=0.5, top=0.75, bottom=0.5,
        header=0.3, footer=0.3
    )

    print("  IncidentReport template sheet created (VeryHidden)")


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
    print(f"  Sheets: Dashboard, Incident Log, Settings, Assignment Tracker, RCA Log, IncidentReport (VeryHidden)")
    print(f"  Tables: tblTeams, tblPersonnel, tblCategories, tblSLAThresholds, tblPriorityMatrix, tblIncidents, tblAssignmentTracker, tblRCALog")
    print(f"  Output: {output}")
    print(f"  Note: Saved as .xlsx -- VBA injection will convert to .xlsm")
