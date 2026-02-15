#!/usr/bin/env python3
"""
create_user_manual.py
Generates BRIMIS_IMS_User_Manual.docx using python-docx.

Produces a comprehensive Word document with:
- Cover page with BRIMIS branding
- Quick-start guide
- Step-by-step instructions for every feature
- Use case stories for 4 user roles
- Screenshot placeholders
- Troubleshooting guide
- Appendices (status lifecycle, SLA defaults, column reference)

Usage:
    pip install python-docx
    python src/create_user_manual.py
"""

import os
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUT_FILE = os.path.join(PROJECT_ROOT, "BRIMIS_IMS_User_Manual.docx")

# BRIMIS brand colours
BRIMIS_RED = RGBColor(0xEE, 0x31, 0x24)
BRIMIS_DARK = RGBColor(0x10, 0x10, 0x10)
BRIMIS_GRAY = RGBColor(0x32, 0x37, 0x3C)
BRIMIS_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BRIMIS_LIGHT_GRAY = RGBColor(0xF0, 0xF0, 0xF0)

FONT_BODY = "Calibri"


# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------

def set_cell_shading(cell, hex_color):
    """Apply background shading to a table cell."""
    shading_elm = parse_xml(
        f'<w:shd {nsdecls("w")} w:fill="{hex_color}" w:val="clear"/>'
    )
    cell._tc.get_or_add_tcPr().append(shading_elm)


def add_screenshot_placeholder(doc, description):
    """Add a bordered placeholder for a screenshot."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(f"[Screenshot: {description}]")
    run.font.italic = True
    run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
    run.font.size = Pt(10)
    run.font.name = FONT_BODY


def add_note(doc, text):
    """Add a tip/note paragraph in italic gray."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(f"Note: {text}")
    run.font.italic = True
    run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    run.font.size = Pt(10)
    run.font.name = FONT_BODY


def add_tip(doc, text):
    """Add a tip paragraph in italic gray."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(f"Tip: {text}")
    run.font.italic = True
    run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    run.font.size = Pt(10)
    run.font.name = FONT_BODY


def add_body(doc, text):
    """Add a normal body paragraph."""
    p = doc.add_paragraph(text)
    p.style = doc.styles["Normal"]
    return p


def add_numbered_step(doc, text):
    """Add a numbered list item."""
    p = doc.add_paragraph(text, style="List Number")
    return p


def add_bullet(doc, text):
    """Add a bulleted list item."""
    p = doc.add_paragraph(text, style="List Bullet")
    return p


def add_branded_table(doc, headers, rows, col_widths=None):
    """Add a BRIMIS-branded table with dark header row."""
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.style = "Table Grid"

    # Header row
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = header
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for run in paragraph.runs:
                run.font.bold = True
                run.font.color.rgb = BRIMIS_WHITE
                run.font.size = Pt(10)
                run.font.name = FONT_BODY
        set_cell_shading(cell, "101010")

    # Data rows
    for row_idx, row_data in enumerate(rows):
        for col_idx, value in enumerate(row_data):
            cell = table.rows[row_idx + 1].cells[col_idx]
            cell.text = str(value)
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(10)
                    run.font.name = FONT_BODY
            if row_idx % 2 == 1:
                set_cell_shading(cell, "F5F5F5")

    # Apply column widths if provided
    if col_widths:
        for row in table.rows:
            for i, width in enumerate(col_widths):
                row.cells[i].width = Inches(width)

    doc.add_paragraph()  # spacer
    return table


def add_page_break(doc):
    """Insert a page break."""
    doc.add_page_break()


# ---------------------------------------------------------------------------
# Style configuration
# ---------------------------------------------------------------------------

def configure_styles(doc):
    """Set up BRIMIS-branded document styles."""
    # Normal style
    style = doc.styles["Normal"]
    font = style.font
    font.name = FONT_BODY
    font.size = Pt(11)
    font.color.rgb = BRIMIS_DARK
    pf = style.paragraph_format
    pf.space_after = Pt(6)

    # Heading 1
    h1 = doc.styles["Heading 1"]
    h1.font.name = FONT_BODY
    h1.font.size = Pt(16)
    h1.font.bold = True
    h1.font.color.rgb = BRIMIS_RED
    h1.paragraph_format.space_before = Pt(18)
    h1.paragraph_format.space_after = Pt(8)

    # Heading 2
    h2 = doc.styles["Heading 2"]
    h2.font.name = FONT_BODY
    h2.font.size = Pt(13)
    h2.font.bold = True
    h2.font.color.rgb = BRIMIS_DARK
    h2.paragraph_format.space_before = Pt(14)
    h2.paragraph_format.space_after = Pt(6)

    # Heading 3
    h3 = doc.styles["Heading 3"]
    h3.font.name = FONT_BODY
    h3.font.size = Pt(11)
    h3.font.bold = True
    h3.font.color.rgb = BRIMIS_GRAY
    h3.paragraph_format.space_before = Pt(10)
    h3.paragraph_format.space_after = Pt(4)

    # List Number
    ln = doc.styles["List Number"]
    ln.font.name = FONT_BODY
    ln.font.size = Pt(11)
    ln.font.color.rgb = BRIMIS_DARK

    # List Bullet
    lb = doc.styles["List Bullet"]
    lb.font.name = FONT_BODY
    lb.font.size = Pt(11)
    lb.font.color.rgb = BRIMIS_DARK


# ---------------------------------------------------------------------------
# Header / Footer
# ---------------------------------------------------------------------------

def configure_header_footer(doc):
    """Set header and footer for all sections."""
    for section in doc.sections:
        # Header
        header = section.header
        header.is_linked_to_previous = False
        hp = header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = hp.add_run("BRIMIS Incident Management System  --  User Manual")
        run.font.name = FONT_BODY
        run.font.size = Pt(8)
        run.font.color.rgb = BRIMIS_GRAY

        # Footer
        footer = section.footer
        footer.is_linked_to_previous = False
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = fp.add_run("Confidential  |  BRIMIS Engineering  |  Version 1.0")
        run.font.name = FONT_BODY
        run.font.size = Pt(8)
        run.font.color.rgb = BRIMIS_GRAY


# ---------------------------------------------------------------------------
# Cover page
# ---------------------------------------------------------------------------

def write_cover_page(doc):
    """Create the branded cover page."""
    # Several blank lines to push content down
    for _ in range(6):
        doc.add_paragraph()

    # Title
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("BRIMIS")
    run.font.name = FONT_BODY
    run.font.size = Pt(36)
    run.font.bold = True
    run.font.color.rgb = BRIMIS_RED

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Incident Management System")
    run.font.name = FONT_BODY
    run.font.size = Pt(28)
    run.font.bold = True
    run.font.color.rgb = BRIMIS_DARK

    # Spacer
    doc.add_paragraph()

    # Subtitle
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("User Manual")
    run.font.name = FONT_BODY
    run.font.size = Pt(22)
    run.font.color.rgb = BRIMIS_GRAY

    # Spacer
    doc.add_paragraph()
    doc.add_paragraph()

    # Version and date
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Version 1.0  |  February 2026")
    run.font.name = FONT_BODY
    run.font.size = Pt(12)
    run.font.color.rgb = BRIMIS_GRAY

    # Spacer
    for _ in range(4):
        doc.add_paragraph()

    # Company name
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("BRIMIS Engineering")
    run.font.name = FONT_BODY
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.color.rgb = BRIMIS_DARK

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Precision. Quality. Innovation. Excellence.")
    run.font.name = FONT_BODY
    run.font.size = Pt(10)
    run.font.italic = True
    run.font.color.rgb = BRIMIS_GRAY

    add_page_break(doc)


# ---------------------------------------------------------------------------
# Table of Contents
# ---------------------------------------------------------------------------

def write_table_of_contents(doc):
    """Write the table of contents page."""
    doc.add_heading("Table of Contents", level=1)

    toc_entries = [
        ("1.", "Quick-Start Guide", ""),
        ("", "1.1 Opening the Workbook", ""),
        ("", "1.2 Enabling Macros", ""),
        ("", "1.3 Dashboard Overview", ""),
        ("", "1.4 Your First Incident (5-Minute Walkthrough)", ""),
        ("2.", "System Overview", ""),
        ("", "2.1 What is BRIMIS IMS?", ""),
        ("", "2.2 Key Concepts", ""),
        ("", "2.3 Sheets and Navigation", ""),
        ("", "2.4 User Roles", ""),
        ("3.", "Dashboard", ""),
        ("", "3.1 KPI Summary Cards", ""),
        ("", "3.2 Priority and Category Breakdowns", ""),
        ("", "3.3 SLA Monitor", ""),
        ("", "3.4 Action Buttons", ""),
        ("", "3.5 Refreshing Dashboard Data", ""),
        ("4.", "Logging Incidents", ""),
        ("", "4.1 Opening the Incident Entry Form", ""),
        ("", "4.2 Step 1: Category Selection", ""),
        ("", "4.3 Step 2: Incident Details", ""),
        ("", "4.4 Step 3: Priority Assessment", ""),
        ("", "4.5 Step 4: Review and Submit", ""),
        ("", "4.6 What Happens After Submission", ""),
        ("5.", "Assigning Incidents", ""),
        ("", "5.1 Opening the Assignment Form", ""),
        ("", "5.2 Selecting an Incident", ""),
        ("", "5.3 Choosing Team and Assignee", ""),
        ("", "5.4 Completing the Assignment", ""),
        ("6.", "Updating Incident Status", ""),
        ("", "6.1 Opening the Status Update Form", ""),
        ("", "6.2 Understanding the Status Lifecycle", ""),
        ("", "6.3 Valid Status Transitions", ""),
        ("", "6.4 Resolving with Root Cause Analysis", ""),
        ("", "6.5 Cancelling or Marking as Duplicate", ""),
        ("7.", "Searching Incidents", ""),
        ("", "7.1 Opening the Search Form", ""),
        ("", "7.2 Search Criteria", ""),
        ("", "7.3 Using Multiple Criteria (AND Logic)", ""),
        ("", "7.4 Viewing Search Results", ""),
        ("", "7.5 Viewing Full Incident Details", ""),
        ("", "7.6 Generating Reports from Search", ""),
        ("8.", "Reports", ""),
        ("", "8.1 Generating an Incident Report", ""),
        ("", "8.2 Print Preview", ""),
        ("", "8.3 Exporting to PDF", ""),
        ("", "8.4 Report Contents and Layout", ""),
        ("9.", "Settings Administration", ""),
        ("", "9.1 Accessing the Settings Sheet", ""),
        ("", "9.2 Managing Teams (tblTeams)", ""),
        ("", "9.3 Managing Personnel (tblPersonnel)", ""),
        ("", "9.4 Managing Categories and Subcategories", ""),
        ("", "9.5 SLA Thresholds Configuration", ""),
        ("", "9.6 Priority Matrix", ""),
        ("10.", "Use Case Stories", ""),
        ("", "10.1 Artisan: Reporting a Pump Failure", ""),
        ("", "10.2 Supervisor: Assigning and Monitoring Work", ""),
        ("", "10.3 Technician: Resolving and Documenting Root Cause", ""),
        ("", "10.4 Management: Dashboard Review and Reporting", ""),
        ("11.", "Troubleshooting", ""),
        ("", "11.1 Macros Are Disabled", ""),
        ("", "11.2 Buttons Do Not Respond", ""),
        ("", "11.3 Forms Show Error Messages", ""),
        ("", "11.4 SLA Colors Not Showing", ""),
        ("", "11.5 Backup Location Issues", ""),
        ("", "Appendix A: Status Lifecycle Diagram", ""),
        ("", "Appendix B: SLA Threshold Defaults", ""),
        ("", "Appendix C: Column Reference (tblIncidents)", ""),
    ]

    for number, title, _ in toc_entries:
        p = doc.add_paragraph()
        if number:
            # Chapter-level entry (bold)
            run = p.add_run(f"{number} {title}")
            run.font.bold = True
            run.font.size = Pt(11)
            run.font.name = FONT_BODY
            run.font.color.rgb = BRIMIS_DARK
        else:
            # Section-level entry (indented)
            run = p.add_run(f"     {title}")
            run.font.size = Pt(10)
            run.font.name = FONT_BODY
            run.font.color.rgb = BRIMIS_GRAY
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.space_before = Pt(1)

    add_page_break(doc)


# ---------------------------------------------------------------------------
# Chapter 1: Quick-Start Guide
# ---------------------------------------------------------------------------

def write_chapter_1(doc):
    """Chapter 1: Quick-Start Guide."""
    doc.add_heading("1. Quick-Start Guide", level=1)

    add_body(doc,
        "This chapter gets you up and running with the BRIMIS Incident Management System "
        "in under five minutes. Follow these steps to open the workbook, enable macros, "
        "and log your first incident."
    )

    # 1.1 Opening the Workbook
    doc.add_heading("1.1 Opening the Workbook", level=2)
    add_numbered_step(doc, 'Locate the file "BRIMIS_IMS.xlsm" on your computer or shared drive.')
    add_numbered_step(doc, "Double-click the file to open it in Microsoft Excel.")
    add_numbered_step(doc, "The workbook opens to the Dashboard sheet automatically.")
    add_note(doc,
        'The file extension ".xlsm" indicates a macro-enabled workbook. '
        "If your organisation uses a different file-sharing system, ask your IT administrator "
        "for the correct file location."
    )

    # 1.2 Enabling Macros
    doc.add_heading("1.2 Enabling Macros", level=2)
    add_body(doc,
        "The BRIMIS IMS relies on macros (VBA code) to power its forms and automation. "
        "When you first open the workbook, Excel may display a security warning."
    )
    add_numbered_step(doc,
        'Look for the yellow Security Warning bar at the top of the Excel window that says '
        '"Macros have been disabled."'
    )
    add_numbered_step(doc, 'Click the "Enable Content" button on the warning bar.')
    add_numbered_step(doc,
        "The warning bar disappears and the system is fully active. "
        "The Dashboard will refresh automatically."
    )
    add_screenshot_placeholder(doc, "Excel Security Warning bar with Enable Content button highlighted")
    add_tip(doc,
        "If you do not see the Security Warning bar, your macro security settings may need adjustment. "
        "See Chapter 11 (Troubleshooting) for instructions."
    )

    # 1.3 Dashboard Overview
    doc.add_heading("1.3 Dashboard Overview", level=2)
    add_body(doc,
        "The Dashboard is your home base. It provides an at-a-glance summary of all incident activity."
    )
    add_bullet(doc, "KPI Cards: Total Open incidents, Overdue count, Average Resolution Days, Total Closed")
    add_bullet(doc, "Priority Breakdown: Count of open incidents by priority level (P1 through P4)")
    add_bullet(doc, "Category Breakdown: Count of open incidents by category (Mechanical, Electrical, Safety/HSE)")
    add_bullet(doc, "SLA Monitor: Table showing open incidents with their SLA response and resolution status")
    add_bullet(doc,
        "Action Buttons: Five buttons at the top -- Log New Incident, Assign Incident, "
        "Update Status, Refresh Dashboard, Search Incidents"
    )
    add_screenshot_placeholder(doc, "Dashboard sheet showing KPI cards, breakdowns, SLA monitor, and 5 action buttons")

    # 1.4 Your First Incident
    doc.add_heading("1.4 Your First Incident (5-Minute Walkthrough)", level=2)
    add_body(doc,
        "Follow these steps to log your very first incident into the system:"
    )
    add_numbered_step(doc, 'Click the "Log New Incident" button on the Dashboard.')
    add_numbered_step(doc,
        "The Incident Entry Form opens as a guided wizard with four steps."
    )
    add_numbered_step(doc,
        "Step 1 -- Category: Select a Category (e.g., Mechanical), a Subcategory "
        "(e.g., Pump Breakdowns), and your name from the Reporter dropdown. Click Next."
    )
    add_numbered_step(doc,
        "Step 2 -- Details: Enter a short Title (e.g., \"Pump leak at Station 3\") and a "
        "Description of what happened. Optionally add an Attachment Reference. Click Next."
    )
    add_numbered_step(doc,
        "Step 3 -- Priority: Select the Impact level and Urgency level. The system "
        "automatically calculates the Priority (P1-P4) based on the priority matrix. Click Next."
    )
    add_numbered_step(doc,
        "Step 4 -- Review: Check all the information on the review page. If everything "
        "looks correct, click Submit. To make changes, click Back."
    )
    add_numbered_step(doc,
        'A confirmation message appears showing the new Incident ID (e.g., "INC-00001"). '
        "Click OK."
    )
    add_numbered_step(doc,
        "The incident now appears in the Incident Log sheet and the Dashboard reflects "
        "the updated totals."
    )
    add_screenshot_placeholder(doc, "Incident Entry Form -- Step 1 (Category Selection)")
    add_tip(doc,
        "The Incident ID is automatically generated. You do not need to create or remember it -- "
        "use it as a reference when discussing the incident with colleagues."
    )

    add_page_break(doc)


# ---------------------------------------------------------------------------
# Chapter 2: System Overview
# ---------------------------------------------------------------------------

def write_chapter_2(doc):
    """Chapter 2: System Overview."""
    doc.add_heading("2. System Overview", level=1)

    # 2.1 What is BRIMIS IMS?
    doc.add_heading("2.1 What is BRIMIS IMS?", level=2)
    add_body(doc,
        "The BRIMIS Incident Management System (IMS) is an Excel-based tool that helps "
        "BRIMIS Engineering track industrial incidents from initial report through resolution "
        "and closure. It provides a single source of truth for all incident data, ensuring "
        "that every pump failure, valve malfunction, electrical fault, or safety concern is "
        "documented, assigned, resolved, and reviewed."
    )
    add_body(doc,
        "The system runs entirely within Microsoft Excel using VBA-powered forms, requiring "
        "no internet connection, no additional software, and no specialised IT training. "
        "Anyone who can use Excel can use this system."
    )

    # 2.2 Key Concepts
    doc.add_heading("2.2 Key Concepts", level=2)

    doc.add_heading("Incident", level=3)
    add_body(doc,
        "An incident is any event that disrupts or threatens normal operations -- a mechanical "
        "failure, electrical fault, safety concern, or environmental issue. Each incident is "
        "assigned a unique ID (e.g., INC-00001) and tracked through its full lifecycle."
    )

    doc.add_heading("Status Lifecycle", level=3)
    add_body(doc,
        "Every incident moves through a defined series of statuses:"
    )
    add_bullet(doc, "Open -- Newly reported, awaiting assignment")
    add_bullet(doc, "Assigned -- Allocated to a team and individual")
    add_bullet(doc, "In Progress -- Work has begun on investigating/fixing the issue")
    add_bullet(doc, "Resolved -- Fix is complete and root cause has been documented")
    add_bullet(doc, "Closed -- Resolution verified and incident formally closed")
    add_bullet(doc, "Cancelled -- Incident was reported in error or is no longer relevant")
    add_bullet(doc, "Duplicate -- Incident is a duplicate of an existing record")

    doc.add_heading("Priority (P1-P4)", level=3)
    add_body(doc,
        "Priority is automatically calculated from Impact and Urgency scores using a "
        "priority matrix configured in the Settings sheet:"
    )
    add_bullet(doc, "P1 (Critical) -- Highest priority. Immediate response required.")
    add_bullet(doc, "P2 (High) -- Significant impact. Prompt response needed.")
    add_bullet(doc, "P3 (Medium) -- Moderate impact. Scheduled response acceptable.")
    add_bullet(doc, "P4 (Low) -- Minor impact. Routine handling.")

    doc.add_heading("SLA (Service Level Agreement)", level=3)
    add_body(doc,
        "SLA defines how quickly incidents must be responded to and resolved, based on "
        "their priority. The system tracks two SLA metrics for each incident:"
    )
    add_bullet(doc,
        "Response SLA -- Time from when the incident is reported to when work begins "
        "(status changes to In Progress)"
    )
    add_bullet(doc,
        "Resolution SLA -- Time from when the incident is reported to when it is resolved"
    )
    add_body(doc,
        "SLA status is colour-coded: green (On Track / Met), amber (At Risk), "
        "and red (Overdue / Breached)."
    )

    # 2.3 Sheets and Navigation
    doc.add_heading("2.3 Sheets and Navigation", level=2)
    add_body(doc,
        "The workbook contains the following sheets. Click the sheet tabs at the bottom "
        "of the Excel window to navigate between them."
    )
    add_branded_table(doc,
        ["Sheet", "Purpose", "Who Uses It"],
        [
            ["Dashboard", "At-a-glance KPIs, priority/category breakdown, SLA monitor, action buttons", "Everyone"],
            ["Incident Log", "Master table of all incidents with full data (27 columns)", "Supervisors, Management"],
            ["Settings", "Configuration tables: teams, personnel, categories, SLA thresholds, priority matrix", "Administrators"],
            ["Assignment Tracker", "Filtered view of assigned incidents with team, person, days open", "Supervisors"],
            ["RCA Log", "Root Cause Analysis records linked to resolved incidents", "Technicians, Management"],
        ],
        col_widths=[1.5, 3.0, 1.5]
    )
    add_note(doc,
        "The workbook also contains a hidden IncidentReport sheet used internally by the "
        "report generation feature. You do not need to access it directly."
    )

    # 2.4 User Roles
    doc.add_heading("2.4 User Roles", level=2)
    add_body(doc,
        "The system supports four user roles. There is no login or access control -- all "
        "users open the same workbook. The roles describe how each person typically uses "
        "the system:"
    )
    add_branded_table(doc,
        ["Role", "Primary Tasks", "Key Features Used"],
        [
            ["Artisan / Field Worker",
             "Report incidents encountered on the job",
             "Log New Incident button, Incident Entry Form"],
            ["Supervisor / Dispatcher",
             "Review, prioritise, assign, and monitor incidents",
             "Dashboard, Assign Incident, Assignment Tracker, Search"],
            ["Technician / Engineer",
             "Investigate, resolve, and document root cause",
             "Update Status, RCA fields, Status Update Form"],
            ["Management",
             "Review metrics, SLA compliance, generate reports",
             "Dashboard, Search, Reports, RCA Log"],
        ],
        col_widths=[1.5, 2.5, 2.0]
    )

    add_page_break(doc)


# ---------------------------------------------------------------------------
# Chapter 3: Dashboard
# ---------------------------------------------------------------------------

def write_chapter_3(doc):
    """Chapter 3: Dashboard."""
    doc.add_heading("3. Dashboard", level=1)
    add_body(doc,
        "The Dashboard is the first sheet you see when opening the workbook. It provides "
        "a real-time summary of incident activity and serves as the launching point for "
        "all actions."
    )

    # 3.1 KPI Summary Cards
    doc.add_heading("3.1 KPI Summary Cards", level=2)
    add_body(doc,
        "Four large KPI cards are displayed at the top of the Dashboard:"
    )
    add_branded_table(doc,
        ["Card", "What It Shows", "How It Is Calculated"],
        [
            ["Total Open", "Number of incidents in Open, Assigned, or In Progress status",
             "Count of non-closed, non-cancelled, non-duplicate, non-resolved incidents"],
            ["Overdue", "Number of open incidents that have exceeded their SLA resolution time",
             "Count where SLA Resolution Status = Overdue or Breached"],
            ["Avg Resolution Days", "Average number of days to resolve incidents",
             "Mean of (ResolutionDate - ReportedDate) for resolved/closed incidents"],
            ["Total Closed", "Number of incidents fully closed",
             "Count where Status = Closed"],
        ],
        col_widths=[1.3, 2.2, 2.5]
    )
    add_screenshot_placeholder(doc, "Dashboard KPI cards showing Total Open, Overdue, Avg Resolution Days, Total Closed")

    # 3.2 Priority and Category Breakdowns
    doc.add_heading("3.2 Priority and Category Breakdowns", level=2)
    add_body(doc,
        "Below the KPI cards, two breakdown sections show the distribution of open incidents:"
    )
    add_bullet(doc,
        "Priority Breakdown: Shows counts for P1, P2, P3, and P4 priorities with colour coding "
        "(P1 red, P2 orange, P3 green, P4 gray)"
    )
    add_bullet(doc,
        "Category Breakdown: Shows counts for each incident category "
        "(Mechanical, Electrical, Safety/HSE)"
    )

    # 3.3 SLA Monitor
    doc.add_heading("3.3 SLA Monitor", level=2)
    add_body(doc,
        "The SLA Monitor table lists all open incidents with their current SLA performance. "
        "Each row shows:"
    )
    add_bullet(doc, "Incident ID and Title")
    add_bullet(doc, "Priority level")
    add_bullet(doc, "Current status")
    add_bullet(doc, "Response SLA status (On Track, At Risk, Overdue, Met, Breached)")
    add_bullet(doc, "Resolution SLA status")
    add_bullet(doc, "Time remaining until SLA breach")
    add_body(doc,
        "Cells are colour-coded: green for on-track or met, amber for at-risk, and red for "
        "overdue or breached."
    )
    add_screenshot_placeholder(doc, "SLA Monitor table with colour-coded SLA status cells")

    # 3.4 Action Buttons
    doc.add_heading("3.4 Action Buttons", level=2)
    add_body(doc,
        "Five buttons are positioned at the top of the Dashboard for quick access to key functions:"
    )
    add_branded_table(doc,
        ["Button", "Action", "Opens"],
        [
            ["Log New Incident", "Opens the incident entry wizard to report a new incident", "Incident Entry Form (4-step wizard)"],
            ["Assign Incident", "Opens the assignment form to allocate an incident to a team and person", "Assignment Form"],
            ["Update Status", "Opens the status update form to change an incident's status", "Status Update Form"],
            ["Refresh Dashboard", "Manually recalculates all KPIs, SLA statuses, and dashboard data", "N/A (refreshes in place)"],
            ["Search Incidents", "Opens the search form to find and review incidents", "Search Form"],
        ],
        col_widths=[1.3, 2.7, 2.0]
    )
    add_screenshot_placeholder(doc, "Dashboard action buttons row: Log New Incident, Assign Incident, Update Status, Refresh Dashboard, Search Incidents")

    # 3.5 Refreshing Dashboard Data
    doc.add_heading("3.5 Refreshing Dashboard Data", level=2)
    add_body(doc,
        "The Dashboard refreshes automatically in two situations:"
    )
    add_bullet(doc, "When the workbook is first opened")
    add_bullet(doc, "When you click the Dashboard tab to switch to it from another sheet")
    add_body(doc,
        'You can also manually refresh at any time by clicking the "Refresh Dashboard" button. '
        "This recalculates all SLA statuses, updates KPI counts, and refreshes the SLA Monitor table."
    )
    add_tip(doc,
        "If you have just updated an incident status on another sheet, switch to the Dashboard "
        "tab or click Refresh to see the updated numbers."
    )

    add_page_break(doc)


# ---------------------------------------------------------------------------
# Chapter 4: Logging Incidents
# ---------------------------------------------------------------------------

def write_chapter_4(doc):
    """Chapter 4: Logging Incidents."""
    doc.add_heading("4. Logging Incidents", level=1)
    add_body(doc,
        "Logging a new incident is the first step in the incident management process. "
        "The system uses a guided four-step wizard to collect all required information."
    )

    # 4.1 Opening the Incident Entry Form
    doc.add_heading("4.1 Opening the Incident Entry Form", level=2)
    add_numbered_step(doc, "Navigate to the Dashboard sheet.")
    add_numbered_step(doc, 'Click the "Log New Incident" button.')
    add_numbered_step(doc,
        "The Incident Entry Form opens as a modal dialog with four wizard steps."
    )
    add_screenshot_placeholder(doc, "Incident Entry Form -- initial view with Step 1 active")

    # 4.2 Step 1: Category Selection
    doc.add_heading("4.2 Step 1: Category Selection", level=2)
    add_body(doc, "On the first page of the wizard, select the incident classification:")
    add_numbered_step(doc,
        'Category: Select from the dropdown (Mechanical, Electrical, or Safety/HSE). '
        "These categories are configured in the Settings sheet."
    )
    add_numbered_step(doc,
        "Subcategory: After selecting a category, the subcategory dropdown is automatically "
        "filtered to show only relevant options. For example, selecting Mechanical shows: "
        "Valve Failures, Pump Breakdowns, Equipment Malfunction, Wear & Tear."
    )
    add_numbered_step(doc,
        "Reporter: Select your name from the dropdown. This list comes from the Personnel "
        "table in Settings."
    )
    add_numbered_step(doc, 'Click "Next" to proceed to Step 2.')
    add_screenshot_placeholder(doc, "Step 1 -- Category, Subcategory, and Reporter dropdowns")
    add_note(doc,
        "All three fields are required. The Next button will not advance until all fields are filled."
    )

    # 4.3 Step 2: Incident Details
    doc.add_heading("4.3 Step 2: Incident Details", level=2)
    add_body(doc, "On the second page, provide the details of the incident:")
    add_numbered_step(doc,
        "Title: Enter a brief, descriptive title (e.g., \"Pump leak at Station 3\"). "
        "This title appears in search results and the Dashboard."
    )
    add_numbered_step(doc,
        "Description: Provide a detailed description of what happened, where, when, and any "
        "immediate actions taken. Be as specific as possible."
    )
    add_numbered_step(doc,
        "Attachment Reference (optional): If you have photos, documents, or other files related "
        "to the incident, enter a file name or path reference here. The system stores the "
        "reference text but does not attach files directly."
    )
    add_numbered_step(doc, 'Click "Next" to proceed to Step 3.')
    add_screenshot_placeholder(doc, "Step 2 -- Title, Description, and Attachment Reference fields")
    add_note(doc,
        "Title and Description are required fields. Attachment Reference is optional."
    )

    # 4.4 Step 3: Priority Assessment
    doc.add_heading("4.4 Step 3: Priority Assessment", level=2)
    add_body(doc,
        "On the third page, assess the severity of the incident. The system automatically "
        "calculates the priority."
    )
    add_numbered_step(doc,
        "Impact: Select the level of impact from the dropdown. Options typically include "
        "High, Medium, and Low, representing how severely the incident affects operations."
    )
    add_numbered_step(doc,
        "Urgency: Select how urgently the incident needs to be addressed. Options typically "
        "include High, Medium, and Low."
    )
    add_numbered_step(doc,
        "Priority: The system automatically calculates the priority (P1 through P4) based on "
        "the combination of Impact and Urgency using the Priority Matrix in Settings. The "
        "calculated priority is displayed on screen with an explanation."
    )
    add_numbered_step(doc, 'Click "Next" to proceed to the Review step.')
    add_screenshot_placeholder(doc, "Step 3 -- Impact and Urgency dropdowns with calculated Priority display")
    add_tip(doc,
        "You cannot manually override the calculated priority. If you believe the priority "
        "is incorrect, check the Priority Matrix configuration in the Settings sheet."
    )

    # 4.5 Step 4: Review and Submit
    doc.add_heading("4.5 Step 4: Review and Submit", level=2)
    add_body(doc,
        "The final page shows a summary of all information you entered. Review each field carefully."
    )
    add_numbered_step(doc, "Review the Category, Subcategory, and Reporter.")
    add_numbered_step(doc, "Review the Title and Description.")
    add_numbered_step(doc, "Review the Priority, Impact, and Urgency.")
    add_numbered_step(doc,
        'If everything is correct, click "Submit" to log the incident.'
    )
    add_numbered_step(doc,
        'If you need to change something, click "Back" to return to the relevant step.'
    )
    add_numbered_step(doc,
        "After submission, a confirmation message appears showing the new Incident ID "
        '(e.g., "Incident INC-00001 has been logged successfully"). Click OK.'
    )
    add_screenshot_placeholder(doc, "Step 4 -- Review page showing all entered information and Submit button")

    # 4.6 What Happens After Submission
    doc.add_heading("4.6 What Happens After Submission", level=2)
    add_body(doc, "When you submit a new incident, the system automatically:")
    add_bullet(doc, "Generates a unique Incident ID (INC-00001, INC-00002, etc.)")
    add_bullet(doc, 'Sets the Status to "Open"')
    add_bullet(doc, "Records the current date and time as the Reported Date")
    add_bullet(doc, "Adds a new row to the Incident Log sheet with all the data")
    add_bullet(doc, "Records a Last Modified timestamp")
    add_body(doc,
        "The incident is now visible on the Dashboard and ready for a supervisor to assign."
    )

    add_page_break(doc)


# ---------------------------------------------------------------------------
# Chapter 5: Assigning Incidents
# ---------------------------------------------------------------------------

def write_chapter_5(doc):
    """Chapter 5: Assigning Incidents."""
    doc.add_heading("5. Assigning Incidents", level=1)
    add_body(doc,
        "After an incident is logged, a supervisor assigns it to a team and a specific "
        "individual for investigation and resolution."
    )

    # 5.1
    doc.add_heading("5.1 Opening the Assignment Form", level=2)
    add_numbered_step(doc, "Navigate to the Dashboard sheet.")
    add_numbered_step(doc, 'Click the "Assign Incident" button.')
    add_numbered_step(doc,
        "The Assignment Form opens, showing a list of all incidents with Open status."
    )
    add_screenshot_placeholder(doc, "Assignment Form showing list of Open incidents and team/person dropdowns")

    # 5.2
    doc.add_heading("5.2 Selecting an Incident", level=2)
    add_numbered_step(doc,
        "The incident list shows three columns: Incident ID, Title, and Priority."
    )
    add_numbered_step(doc,
        "Click on an incident to select it. The form displays additional details below "
        "the list: Category, Priority, Reported By, and Reported Date."
    )
    add_note(doc,
        "Only incidents with Open status appear in this list. If an incident has already "
        "been assigned, it will not appear here."
    )

    # 5.3
    doc.add_heading("5.3 Choosing Team and Assignee", level=2)
    add_numbered_step(doc,
        "Team: Select the responsible team from the dropdown (e.g., Mechanical Team, "
        "Electrical Team). Teams are configured in the Settings sheet."
    )
    add_numbered_step(doc,
        "Person: After selecting a team, choose the specific individual from the Person "
        "dropdown. This list automatically filters to show only members of the selected team."
    )
    add_screenshot_placeholder(doc, "Assignment Form with team selected and person dropdown showing team members")

    # 5.4
    doc.add_heading("5.4 Completing the Assignment", level=2)
    add_numbered_step(doc,
        'Click the "Assign" button to complete the assignment.'
    )
    add_numbered_step(doc,
        "The system automatically:"
    )
    add_bullet(doc, 'Changes the incident status from "Open" to "Assigned"')
    add_bullet(doc, "Records the Assigned Date (current date and time)")
    add_bullet(doc, "Records the Assigned Team and Assigned To fields")
    add_bullet(doc, "Records the Assigned By field (the current user)")
    add_bullet(doc, "Updates the Assignment Tracker sheet")
    add_numbered_step(doc,
        'A confirmation message appears. Click OK. The assigned incident is removed from '
        "the list since it is no longer in Open status."
    )
    add_tip(doc,
        "After assigning, check the Assignment Tracker sheet to see the incident listed "
        "with its team, assignee, and Days Open counter."
    )

    add_page_break(doc)


# ---------------------------------------------------------------------------
# Chapter 6: Updating Incident Status
# ---------------------------------------------------------------------------

def write_chapter_6(doc):
    """Chapter 6: Updating Incident Status."""
    doc.add_heading("6. Updating Incident Status", level=1)
    add_body(doc,
        "As work progresses on an incident, its status is updated to reflect the current stage. "
        "The system enforces a defined lifecycle to ensure incidents follow the correct path."
    )

    # 6.1
    doc.add_heading("6.1 Opening the Status Update Form", level=2)
    add_numbered_step(doc, "Navigate to the Dashboard sheet.")
    add_numbered_step(doc, 'Click the "Update Status" button.')
    add_numbered_step(doc,
        "The Status Update Form opens, showing all non-terminal incidents "
        "(those not yet Closed, Cancelled, or Duplicate)."
    )
    add_screenshot_placeholder(doc, "Status Update Form showing list of active incidents")

    # 6.2
    doc.add_heading("6.2 Understanding the Status Lifecycle", level=2)
    add_body(doc,
        "Each incident follows a defined path through these statuses:"
    )
    add_body(doc, "Open --> Assigned --> In Progress --> Resolved --> Closed")
    add_body(doc,
        "At certain points, incidents can also be moved to terminal states:"
    )
    add_bullet(doc, "Cancelled -- Available from Open, Assigned, or In Progress")
    add_bullet(doc, "Duplicate -- Available from Open, Assigned, or In Progress")
    add_body(doc,
        "The system only shows valid next statuses in the dropdown. You cannot skip steps "
        "or move backwards. Once an incident is Closed, Cancelled, or Duplicate, no further "
        "status changes are allowed."
    )

    # 6.3
    doc.add_heading("6.3 Valid Status Transitions", level=2)
    add_branded_table(doc,
        ["Current Status", "Valid Next Statuses"],
        [
            ["Open", "Assigned (via Assignment Form only), Cancelled, Duplicate"],
            ["Assigned", "In Progress, Cancelled, Duplicate"],
            ["In Progress", "Resolved, Cancelled, Duplicate"],
            ["Resolved", "Closed"],
            ["Closed", "(No further transitions -- terminal state)"],
            ["Cancelled", "(No further transitions -- terminal state)"],
            ["Duplicate", "(No further transitions -- terminal state)"],
        ],
        col_widths=[1.5, 4.5]
    )
    add_note(doc,
        'The transition from Open to Assigned is handled exclusively through the Assignment '
        "Form, not the Status Update Form. This ensures that team and assignee information "
        "is always recorded."
    )

    # 6.4
    doc.add_heading("6.4 Resolving with Root Cause Analysis", level=2)
    add_body(doc,
        'When you change an incident\'s status to "Resolved," additional fields appear '
        "for Root Cause Analysis (RCA):"
    )
    add_numbered_step(doc, "Select an incident from the list.")
    add_numbered_step(doc, 'Select "Resolved" from the New Status dropdown.')
    add_numbered_step(doc,
        "The RCA fields appear below the status dropdown:"
    )
    add_bullet(doc, "Root Cause (required): Describe the underlying cause of the incident")
    add_bullet(doc, "Corrective Action (optional): Describe what was done to fix the immediate problem")
    add_bullet(doc, "Preventive Action (optional): Describe steps to prevent recurrence")
    add_bullet(doc, "Resolution Notes (optional): Any additional notes about the resolution")
    add_numbered_step(doc, 'Click "Update Status" to save.')
    add_body(doc,
        "The system records the RCA data in two places: inline on the Incident Log (for quick "
        "reference) and in the dedicated RCA Log sheet (for detailed review). The Resolution "
        "Date is automatically recorded."
    )
    add_screenshot_placeholder(doc, "Status Update Form showing RCA fields when Resolved is selected")

    # 6.5
    doc.add_heading("6.5 Cancelling or Marking as Duplicate", level=2)
    add_body(doc,
        "If an incident was reported in error or is a duplicate of an existing incident:"
    )
    add_numbered_step(doc, "Select the incident from the list.")
    add_numbered_step(doc, 'Select "Cancelled" or "Duplicate" from the New Status dropdown.')
    add_numbered_step(doc,
        "A Reason field appears. Enter a mandatory explanation for why the incident is "
        "being cancelled or marked as duplicate."
    )
    add_numbered_step(doc, 'Click "Update Status" to save.')
    add_note(doc,
        "A reason is required for cancellations and duplicates. The system will not allow "
        "the status change without an explanation."
    )

    add_page_break(doc)


# ---------------------------------------------------------------------------
# Chapter 7: Searching Incidents
# ---------------------------------------------------------------------------

def write_chapter_7(doc):
    """Chapter 7: Searching Incidents."""
    doc.add_heading("7. Searching Incidents", level=1)
    add_body(doc,
        "The Search feature allows you to find specific incidents using a combination "
        "of criteria. This is useful for locating historical incidents, checking on "
        "specific categories, or preparing reports."
    )

    # 7.1
    doc.add_heading("7.1 Opening the Search Form", level=2)
    add_numbered_step(doc, "Navigate to the Dashboard sheet.")
    add_numbered_step(doc, 'Click the "Search Incidents" button.')
    add_numbered_step(doc,
        "The Search Form opens with eight criteria fields, a results list, and detail display."
    )
    add_screenshot_placeholder(doc, "Search Form -- initial view with empty criteria fields")

    # 7.2
    doc.add_heading("7.2 Search Criteria", level=2)
    add_body(doc,
        "You can search by any combination of the following eight criteria:"
    )
    add_branded_table(doc,
        ["Field", "Type", "How It Matches"],
        [
            ["ID", "Text box", "Partial match on Incident ID (e.g., typing \"001\" finds INC-00001)"],
            ["Title", "Text box", "Partial match on Title, case-insensitive"],
            ["Status", "Dropdown", "Exact match on status (Open, Assigned, In Progress, Resolved, Closed, Cancelled, Duplicate)"],
            ["Priority", "Dropdown", "Exact match on priority (P1, P2, P3, P4)"],
            ["Category", "Dropdown", "Exact match on category (populated from Settings)"],
            ["Assignee", "Dropdown", "Exact match on the assigned person"],
            ["Date From", "Text box", "Only shows incidents reported on or after this date"],
            ["Date To", "Text box", "Only shows incidents reported on or before this date"],
        ],
        col_widths=[1.0, 1.0, 4.0]
    )
    add_note(doc,
        "Leave a field empty to match all values for that criterion. For example, leaving "
        "Status empty and setting Priority to P1 finds all P1 incidents regardless of status."
    )

    # 7.3
    doc.add_heading("7.3 Using Multiple Criteria (AND Logic)", level=2)
    add_body(doc,
        "When you fill in multiple criteria, all conditions must match (AND logic). For example:"
    )
    add_bullet(doc, 'Status = "Open" AND Priority = "P1" finds only open P1 incidents')
    add_bullet(doc, 'Category = "Mechanical" AND Assignee = "Jane Doe" finds mechanical incidents assigned to Jane')
    add_bullet(doc, "Date From = 2026-01-01 AND Date To = 2026-01-31 finds incidents from January 2026")
    add_numbered_step(doc, "Fill in one or more criteria fields.")
    add_numbered_step(doc, 'Click the "Search" button.')
    add_numbered_step(doc,
        "The results list below shows matching incidents with their ID, Title, Status, and Priority."
    )
    add_numbered_step(doc,
        "The results count is displayed above the list (e.g., \"Results: 5 matches found\")."
    )
    add_tip(doc,
        'Click "Clear All" to reset all criteria and start a new search.'
    )

    # 7.4
    doc.add_heading("7.4 Viewing Search Results", level=2)
    add_body(doc,
        "Search results appear in a multi-column list showing:"
    )
    add_bullet(doc, "Incident ID (e.g., INC-00001)")
    add_bullet(doc, "Title (first 50 characters)")
    add_bullet(doc, "Status")
    add_bullet(doc, "Priority")
    add_body(doc,
        "Click on any row to see more details about that incident in the detail area "
        "below the results list."
    )
    add_screenshot_placeholder(doc, "Search results list with an incident selected and details displayed below")

    # 7.5
    doc.add_heading("7.5 Viewing Full Incident Details", level=2)
    add_body(doc,
        "When you select an incident in the results list, the detail area shows:"
    )
    add_bullet(doc, "Incident ID, Status, and Priority")
    add_bullet(doc, "Category and Subcategory")
    add_bullet(doc, "Reported date and reporter name")
    add_bullet(doc, "Assignment information (person and team)")
    add_bullet(doc, "Description (first 200 characters)")
    add_body(doc,
        'For a complete view of all fields, click the "View Full Details" button. '
        "This opens a detailed pop-up showing the full incident record including timeline, "
        "RCA data, and all metadata."
    )

    # 7.6
    doc.add_heading("7.6 Generating Reports from Search", level=2)
    add_body(doc,
        "From the search results, you can generate a formatted incident report:"
    )
    add_numbered_step(doc, "Select an incident in the results list.")
    add_numbered_step(doc, 'Click the "Generate Report" button.')
    add_numbered_step(doc,
        "The search form temporarily hides while the report is generated."
    )
    add_numbered_step(doc,
        "A print preview appears showing the formatted incident report."
    )
    add_numbered_step(doc,
        "From Print Preview, you can print the report or close the preview."
    )
    add_numbered_step(doc, "The search form reappears after you close the preview.")
    add_tip(doc,
        "This is the quickest way to generate a report: search for the incident, "
        "select it, and click Generate Report."
    )

    add_page_break(doc)


# ---------------------------------------------------------------------------
# Chapter 8: Reports
# ---------------------------------------------------------------------------

def write_chapter_8(doc):
    """Chapter 8: Reports."""
    doc.add_heading("8. Reports", level=1)
    add_body(doc,
        "The system can generate professional, print-ready incident reports with full "
        "BRIMIS branding. Reports are generated from the hidden IncidentReport template "
        "sheet and include all incident data."
    )

    # 8.1
    doc.add_heading("8.1 Generating an Incident Report", level=2)
    add_body(doc,
        "There are two ways to generate a report:"
    )
    add_bullet(doc,
        'From the Search Form: Select an incident and click "Generate Report" '
        "(see Section 7.6)"
    )
    add_bullet(doc,
        "Programmatically: Advanced users can call the VBA macro directly "
        "(modReports.GenerateIncidentReport \"INC-00001\")"
    )

    # 8.2
    doc.add_heading("8.2 Print Preview", level=2)
    add_body(doc,
        "The default report generation opens a Print Preview window:"
    )
    add_numbered_step(doc, "The formatted report appears in full-page preview.")
    add_numbered_step(doc,
        'To print, click the "Print" button in the Print Preview toolbar.'
    )
    add_numbered_step(doc, "To close without printing, press Escape or click Close Print Preview.")
    add_screenshot_placeholder(doc, "Print Preview showing a formatted incident report with BRIMIS branding")

    # 8.3
    doc.add_heading("8.3 Exporting to PDF", level=2)
    add_body(doc,
        "Reports can also be exported as PDF files for sharing or archiving. The PDF export "
        "feature saves the report to the same folder as the workbook."
    )
    add_body(doc,
        "The exported PDF is named using the Incident ID (e.g., INC-00001_Report.pdf) "
        "and opens automatically after export."
    )
    add_note(doc,
        "PDF export requires the Microsoft PDF add-in, which is included by default in "
        "Excel 2010 and later."
    )

    # 8.4
    doc.add_heading("8.4 Report Contents and Layout", level=2)
    add_body(doc,
        "Each report is a single-page A4 document with the following sections:"
    )
    add_branded_table(doc,
        ["Section", "Contents"],
        [
            ["Header", "BRIMIS logo (or text branding), \"INCIDENT REPORT\" title, report generation date"],
            ["Incident Overview", "Incident ID, Status, Title, Priority, Category, Subcategory"],
            ["Description", "Full incident description text"],
            ["Timeline", "Reported, Assigned, Response, Resolution, and Closed dates with associated personnel"],
            ["Assignment", "Team name, Assigned To, Assigned Date, Assigned By"],
            ["Resolution & RCA", "Root Cause, Corrective Action, Preventive Action, Resolution Notes"],
            ["SLA Performance", "Response SLA status and Resolution SLA status"],
            ["Footer", "\"Generated by BRIMIS IMS\" with date, page number, and \"Confidential\" marking"],
        ],
        col_widths=[1.5, 4.5]
    )
    add_body(doc,
        "Section headers use BRIMIS dark background with white text. The report is designed "
        "to fit on a single A4 page, centered horizontally."
    )
    add_screenshot_placeholder(doc, "Full incident report showing all six sections with BRIMIS branding")

    add_page_break(doc)


# ---------------------------------------------------------------------------
# Chapter 9: Settings Administration
# ---------------------------------------------------------------------------

def write_chapter_9(doc):
    """Chapter 9: Settings Administration."""
    doc.add_heading("9. Settings Administration", level=1)
    add_body(doc,
        "The Settings sheet contains configuration tables that control the behaviour of the "
        "system. Only administrators should modify these tables. Changes take effect "
        "immediately -- the next time a form is opened, it will reflect the updated data."
    )

    # 9.1
    doc.add_heading("9.1 Accessing the Settings Sheet", level=2)
    add_numbered_step(doc, 'Click the "Settings" tab at the bottom of the Excel window.')
    add_numbered_step(doc,
        "The sheet contains multiple tables arranged vertically. Scroll down to see all tables."
    )
    add_note(doc,
        "The Settings sheet is protected to prevent accidental changes. The protection allows "
        "editing within the tables but prevents modifying the sheet structure."
    )
    add_screenshot_placeholder(doc, "Settings sheet overview showing configuration tables")

    # 9.2
    doc.add_heading("9.2 Managing Teams (tblTeams)", level=2)
    add_body(doc,
        "The Teams table defines the teams available for incident assignment."
    )
    add_branded_table(doc,
        ["Column", "Description"],
        [
            ["TeamName", "The name of the team (e.g., Mechanical Team, Electrical Team)"],
            ["Active", 'Set to "Yes" for active teams or "No" to hide from dropdowns'],
        ],
        col_widths=[1.5, 4.5]
    )
    add_body(doc, "To add a new team:")
    add_numbered_step(doc, "Click in the last row of the tblTeams table.")
    add_numbered_step(doc, "Type the new team name.")
    add_numbered_step(doc, 'Set Active to "Yes".')
    add_numbered_step(doc, "Press Tab or Enter. The table automatically expands.")
    add_body(doc,
        'To deactivate a team, change its Active column to "No". The team will no longer '
        "appear in the Assignment Form dropdown but historical records are preserved."
    )

    # 9.3
    doc.add_heading("9.3 Managing Personnel (tblPersonnel)", level=2)
    add_body(doc,
        "The Personnel table lists all individuals who can report or be assigned incidents."
    )
    add_branded_table(doc,
        ["Column", "Description"],
        [
            ["PersonnelName", "Full name of the employee"],
            ["Team", "The team this person belongs to (must match a TeamName in tblTeams)"],
            ["Active", 'Set to "Yes" for active personnel or "No" to hide from dropdowns'],
        ],
        col_widths=[1.5, 4.5]
    )
    add_body(doc,
        "Personnel appear in two places: the Reporter dropdown on the Incident Entry Form, "
        "and the Person dropdown on the Assignment Form (filtered by selected team)."
    )

    # 9.4
    doc.add_heading("9.4 Managing Categories and Subcategories", level=2)
    add_body(doc,
        "The Categories table controls the incident classification options."
    )
    add_branded_table(doc,
        ["Column", "Description"],
        [
            ["Category", "The main category (e.g., Mechanical, Electrical, Safety/HSE)"],
            ["Subcategory", "The specific subcategory within the category"],
            ["Active", 'Set to "Yes" to make available or "No" to hide'],
        ],
        col_widths=[1.5, 4.5]
    )
    add_body(doc,
        "Each row represents one subcategory. Multiple rows share the same Category value. "
        "For example, the Mechanical category has four rows: Valve Failures, Pump Breakdowns, "
        "Equipment Malfunction, and Wear & Tear."
    )
    add_body(doc, "To add a new subcategory:")
    add_numbered_step(doc, "Add a new row to the tblCategories table.")
    add_numbered_step(doc,
        "Enter the Category name (must match an existing category name exactly, or use a "
        "new category name to create a new top-level category)."
    )
    add_numbered_step(doc, "Enter the Subcategory name.")
    add_numbered_step(doc, 'Set Active to "Yes".')

    # 9.5
    doc.add_heading("9.5 SLA Thresholds Configuration", level=2)
    add_body(doc,
        "The SLA Thresholds table defines how quickly incidents must be responded to and "
        "resolved, based on priority level."
    )
    add_branded_table(doc,
        ["Priority", "Response Time (hours)", "Resolution Time (hours)"],
        [
            ["P1 (Critical)", "1", "4"],
            ["P2 (High)", "4", "24"],
            ["P3 (Medium)", "8", "72"],
            ["P4 (Low)", "24", "168"],
        ],
        col_widths=[2.0, 2.0, 2.0]
    )
    add_body(doc,
        "These values are used by the SLA calculation engine to determine whether an incident "
        "is On Track, At Risk, or Overdue. Modify these values to match your organisation's "
        "service level agreements."
    )
    add_note(doc,
        "All times are in hours. The At Risk threshold is 75% of the allowed time "
        "(e.g., a P1 response is At Risk after 45 minutes of the 1-hour window)."
    )

    # 9.6
    doc.add_heading("9.6 Priority Matrix", level=2)
    add_body(doc,
        "The Priority Matrix maps combinations of Impact and Urgency to a Priority level (P1-P4). "
        "This table is used by the Incident Entry Form to automatically calculate priority."
    )
    add_body(doc,
        "The matrix is structured as a lookup table with Impact as rows and Urgency as columns. "
        "Each cell contains the resulting priority. For example, High Impact + High Urgency = P1."
    )
    add_screenshot_placeholder(doc, "Priority Matrix table in Settings sheet")
    add_tip(doc,
        "If you change the Priority Matrix, the change only affects new incidents. "
        "Existing incidents retain their originally assigned priority."
    )

    add_page_break(doc)


# ---------------------------------------------------------------------------
# Chapter 10: Use Case Stories
# ---------------------------------------------------------------------------

def write_chapter_10(doc):
    """Chapter 10: Use Case Stories."""
    doc.add_heading("10. Use Case Stories", level=1)
    add_body(doc,
        "The following stories illustrate how different roles within BRIMIS Engineering "
        "use the Incident Management System in their daily work. Each story walks through "
        "a realistic scenario with step-by-step actions."
    )

    # 10.1 Artisan: Reporting a Pump Failure
    doc.add_heading("10.1 Artisan: Reporting a Pump Failure", level=2)

    doc.add_heading("The Scenario", level=3)
    add_body(doc,
        "Maria is an artisan at BRIMIS Engineering's Pump Refurbishment Workshop. During her "
        "morning inspection of Station 3, she notices a significant coolant leak from the "
        "main circulation pump. The leak is pooling on the floor and creates both a slip "
        "hazard and a risk of pump damage if the coolant level drops further."
    )

    doc.add_heading("What Maria Does", level=3)
    add_numbered_step(doc,
        "Maria walks to the nearest workstation and opens BRIMIS_IMS.xlsm. "
        'She clicks "Enable Content" when prompted.'
    )
    add_numbered_step(doc,
        'On the Dashboard, she clicks "Log New Incident."'
    )
    add_numbered_step(doc,
        "Step 1 -- Category: Maria selects Category: Mechanical, Subcategory: Pump Breakdowns, "
        "and Reporter: Maria Ndlovu. She clicks Next."
    )
    add_numbered_step(doc,
        'Step 2 -- Details: She enters the Title: "Coolant leak on Station 3 circulation pump" '
        'and Description: "During morning inspection found significant coolant leak from the '
        "main seal area of the Station 3 circulation pump. Coolant pooling on floor approx. "
        '50cm diameter. Pump still running but coolant level dropping. Immediate containment '
        'tray placed." She clicks Next.'
    )
    add_numbered_step(doc,
        "Step 3 -- Priority: She selects Impact: High (production pump at risk) and "
        "Urgency: High (active leak getting worse). The system calculates Priority: P1 (Critical). "
        "She clicks Next."
    )
    add_numbered_step(doc,
        "Step 4 -- Review: Maria reviews all the information and clicks Submit."
    )
    add_numbered_step(doc,
        'The system confirms: "Incident INC-00003 has been logged successfully." Maria notes '
        "the ID and informs her supervisor verbally."
    )

    doc.add_heading("The Outcome", level=3)
    add_body(doc,
        "The incident is now in the system with Open status. It appears on the Dashboard with "
        "P1 priority, triggering the shortest SLA window (1-hour response, 4-hour resolution). "
        "Maria's supervisor will see it immediately when they check the Dashboard."
    )

    # 10.2 Supervisor: Assigning and Monitoring Work
    doc.add_heading("10.2 Supervisor: Assigning and Monitoring Work", level=2)

    doc.add_heading("The Scenario", level=3)
    add_body(doc,
        "James is a workshop supervisor at BRIMIS Engineering. He arrives at his desk and opens "
        "the IMS to check the morning's activity. The Dashboard shows 2 new Open incidents "
        "and 1 overdue SLA warning."
    )

    doc.add_heading("What James Does", level=3)
    add_numbered_step(doc,
        "James opens BRIMIS_IMS.xlsm and sees the Dashboard. The Total Open card shows 5, "
        "and the Overdue card shows 1 (highlighted in red). He notices Maria's P1 pump leak "
        "in the SLA Monitor with a Response SLA showing \"At Risk\" in amber."
    )
    add_numbered_step(doc,
        'He clicks "Assign Incident." The Assignment Form shows two Open incidents.'
    )
    add_numbered_step(doc,
        "He selects INC-00003 (Maria's pump leak). The form shows details: Category: Mechanical, "
        "Priority: P1, Reported By: Maria Ndlovu."
    )
    add_numbered_step(doc,
        "He selects Team: Mechanical Team and Person: Sarah Mokoena (the most experienced "
        "pump technician available). He clicks Assign."
    )
    add_numbered_step(doc,
        "The system confirms the assignment. James then handles the second open incident."
    )
    add_numbered_step(doc,
        "Returning to the Dashboard, James clicks the Refresh button. The SLA Monitor now "
        "shows INC-00003 as Assigned with the Response SLA still tracking."
    )
    add_numbered_step(doc,
        "He checks the Assignment Tracker sheet to see all currently assigned incidents, "
        "including Sarah's new assignment showing Days Open: 0."
    )
    add_numbered_step(doc,
        'Later that afternoon, James uses the Search form to check the status of the '
        "overdue incident. He searches by Status: Open, Priority: P1. He selects the "
        'overdue incident and clicks "View Full Details" to review the timeline.'
    )

    doc.add_heading("The Outcome", level=3)
    add_body(doc,
        "James has assigned the P1 incident to the right team within the SLA response window. "
        "He can monitor progress throughout the day via the Dashboard and Assignment Tracker."
    )

    # 10.3 Technician: Resolving and Documenting Root Cause
    doc.add_heading("10.3 Technician: Resolving and Documenting Root Cause", level=2)

    doc.add_heading("The Scenario", level=3)
    add_body(doc,
        "Sarah Mokoena is a pump technician who has been assigned INC-00003 -- the Station 3 "
        "coolant leak. She has investigated the cause, replaced the failed seal, and needs "
        "to document her work in the system."
    )

    doc.add_heading("What Sarah Does", level=3)
    add_numbered_step(doc,
        "Sarah opens BRIMIS_IMS.xlsm. First, she needs to mark that she has started work. "
        'She clicks "Update Status" on the Dashboard.'
    )
    add_numbered_step(doc,
        'She selects INC-00003 from the list. The current status shows "Assigned." '
        'She selects New Status: "In Progress" and clicks Update Status.'
    )
    add_numbered_step(doc,
        "The system records the Response Date (the SLA response clock stops). "
        "Sarah proceeds to the workshop to complete the repair."
    )
    add_numbered_step(doc,
        "After completing the repair, Sarah returns to the workstation. She opens the "
        'Status Update Form again and selects INC-00003 (now showing "In Progress").'
    )
    add_numbered_step(doc,
        'She selects New Status: "Resolved." The RCA fields appear on the form.'
    )
    add_numbered_step(doc,
        "She fills in the Root Cause Analysis:"
    )
    add_bullet(doc,
        "Root Cause: \"Main circulation pump mechanical seal failure due to excessive wear. "
        "Seal age approximately 18 months, exceeding recommended 12-month replacement interval.\""
    )
    add_bullet(doc,
        "Corrective Action: \"Replaced mechanical seal assembly (part #MS-3421). "
        "Flushed coolant system and refilled with fresh coolant. Tested for 30 minutes at "
        "full pressure -- no leaks detected.\""
    )
    add_bullet(doc,
        "Preventive Action: \"Added Station 3 pump to quarterly inspection schedule. "
        "Set 12-month seal replacement reminder in maintenance calendar.\""
    )
    add_bullet(doc,
        "Resolution Notes: \"Pump restored to full operation. Floor cleaned and slip hazard removed.\""
    )
    add_numbered_step(doc,
        'Sarah clicks "Update Status." The system records the resolution, writes the RCA '
        "data to both the Incident Log and the RCA Log sheet."
    )

    doc.add_heading("The Outcome", level=3)
    add_body(doc,
        "The incident is now Resolved with a complete root cause analysis. The Resolution SLA "
        "status shows \"Met\" (resolved within the 4-hour P1 window). The RCA record appears in "
        "the RCA Log sheet with a unique RCA ID (e.g., RCA-00001) linked to INC-00003."
    )

    # 10.4 Management: Dashboard Review and Reporting
    doc.add_heading("10.4 Management: Dashboard Review and Reporting", level=2)

    doc.add_heading("The Scenario", level=3)
    add_body(doc,
        "Director Okafor conducts a weekly review of operational incidents. He needs to "
        "understand the current state of incidents, check SLA compliance, and generate "
        "a report for a specific incident discussed in a management meeting."
    )

    doc.add_heading("What Director Okafor Does", level=3)
    add_numbered_step(doc,
        "Director Okafor opens BRIMIS_IMS.xlsm. The Dashboard loads automatically with "
        "current data."
    )
    add_numbered_step(doc,
        "He reviews the KPI cards: Total Open: 3, Overdue: 0, Avg Resolution Days: 1.2, "
        "Total Closed: 12. He is pleased to see no overdue incidents this week."
    )
    add_numbered_step(doc,
        "He checks the Priority Breakdown: P1: 0, P2: 1, P3: 2, P4: 0. Only one high-priority "
        "incident remains open."
    )
    add_numbered_step(doc,
        "He scrolls down to the SLA Monitor to see the details of the three open incidents. "
        "All show green status -- on track for their SLA targets."
    )
    add_numbered_step(doc,
        "For the management meeting, he needs a report on the resolved pump incident (INC-00003). "
        'He clicks "Search Incidents."'
    )
    add_numbered_step(doc,
        'In the Search Form, he types "INC-00003" in the ID field and clicks Search.'
    )
    add_numbered_step(doc,
        "The incident appears in the results. He clicks on it to see the details, then "
        'clicks "Generate Report."'
    )
    add_numbered_step(doc,
        "A professional, branded incident report opens in Print Preview, showing the full "
        "incident history including the RCA. He prints a copy for the meeting."
    )
    add_numbered_step(doc,
        "He also checks the RCA Log sheet to review all root cause analyses from the past "
        "month, looking for recurring patterns."
    )

    doc.add_heading("The Outcome", level=3)
    add_body(doc,
        "Director Okafor has a comprehensive view of operational health. The printed incident "
        "report provides documented evidence for the management meeting, and the RCA Log "
        "review helps identify systematic issues that may require process changes."
    )

    add_page_break(doc)


# ---------------------------------------------------------------------------
# Chapter 11: Troubleshooting
# ---------------------------------------------------------------------------

def write_chapter_11(doc):
    """Chapter 11: Troubleshooting."""
    doc.add_heading("11. Troubleshooting", level=1)
    add_body(doc,
        "This chapter addresses common issues you may encounter when using the BRIMIS IMS "
        "and provides step-by-step solutions."
    )

    # 11.1
    doc.add_heading("11.1 Macros Are Disabled", level=2)
    doc.add_heading("Symptom", level=3)
    add_body(doc,
        "The Dashboard buttons do not work, forms do not open, and you may see a Security "
        "Warning bar at the top of the window."
    )
    doc.add_heading("Solution", level=3)
    add_numbered_step(doc,
        'If you see the yellow Security Warning bar, click "Enable Content."'
    )
    add_numbered_step(doc,
        "If the warning bar does not appear, you need to adjust your macro security settings:"
    )
    add_bullet(doc, "Go to File > Options > Trust Center > Trust Center Settings")
    add_bullet(doc, 'Under Macro Settings, select "Disable all macros with notification"')
    add_bullet(doc, "Click OK, then close and re-open the workbook")
    add_numbered_step(doc,
        "After changing the setting, the Security Warning bar should appear next time "
        'you open the file. Click "Enable Content."'
    )
    add_note(doc,
        "Your IT department may manage macro security centrally via Group Policy. If you "
        "cannot change the setting, contact your IT administrator."
    )

    # 11.2
    doc.add_heading("11.2 Buttons Do Not Respond", level=2)
    doc.add_heading("Symptom", level=3)
    add_body(doc,
        "Clicking a Dashboard button does nothing -- no form appears, no error message."
    )
    doc.add_heading("Solution", level=3)
    add_numbered_step(doc, "First, ensure macros are enabled (see Section 11.1).")
    add_numbered_step(doc,
        "Check if you are in Edit mode (cell editing). Press Escape to exit edit mode, "
        "then try clicking the button again."
    )
    add_numbered_step(doc,
        'Check if you are in "Design Mode" -- look for a highlighted Design Mode button '
        "in the Developer tab. If so, click it to toggle Design Mode off."
    )
    add_numbered_step(doc,
        "If buttons still do not work, close and re-open the workbook. This reinitialises "
        "the VBA code."
    )

    # 11.3
    doc.add_heading("11.3 Forms Show Error Messages", level=2)
    doc.add_heading("Symptom", level=3)
    add_body(doc,
        "A message box appears with an error description when opening a form or clicking "
        "a button."
    )
    doc.add_heading("Solution", level=3)
    add_numbered_step(doc,
        "Read the error message carefully. It typically indicates the module and procedure "
        "where the error occurred."
    )
    add_numbered_step(doc,
        'Common causes: corrupted workbook file, missing settings data (empty tables), '
        "or Excel version compatibility issues."
    )
    add_numbered_step(doc,
        "Try closing and re-opening the workbook. If the error persists, check the Settings "
        "sheet to ensure all tables have data."
    )
    add_numbered_step(doc,
        "If the problem continues, contact your IT administrator or the system developer "
        "with the exact error message text."
    )

    # 11.4
    doc.add_heading("11.4 SLA Colors Not Showing", level=2)
    doc.add_heading("Symptom", level=3)
    add_body(doc,
        "The SLA Response and Resolution columns on the Incident Log sheet do not show "
        "green, amber, or red colour coding."
    )
    doc.add_heading("Solution", level=3)
    add_numbered_step(doc,
        "Switch to the Dashboard sheet or click Refresh Dashboard. The SLA conditional "
        "formatting is applied during the dashboard refresh cycle."
    )
    add_numbered_step(doc,
        "Check that the SLA Thresholds table in Settings has valid data for all four "
        "priority levels."
    )
    add_numbered_step(doc,
        "Ensure the incidents have a valid Priority value (P1, P2, P3, or P4). Incidents "
        "without a priority cannot be evaluated against SLA thresholds."
    )

    # 11.5
    doc.add_heading("11.5 Backup Location Issues", level=2)
    doc.add_heading("Symptom", level=3)
    add_body(doc,
        "A message appears about backup failure, or you cannot find backup files."
    )
    doc.add_heading("Solution", level=3)
    add_body(doc,
        "The system creates automatic backups each time the workbook is opened. Backups are "
        "stored in a \"Backups\" subfolder next to the workbook file."
    )
    add_numbered_step(doc,
        "If the workbook is stored on OneDrive or SharePoint, backups are saved to a local "
        "fallback folder at: Documents\\BRIMIS_Backups"
    )
    add_numbered_step(doc,
        "The system keeps a maximum of 30 backup files. Older backups are automatically deleted."
    )
    add_numbered_step(doc,
        "If backup creation fails, the system continues normally -- it does not block "
        "your work. A debug message is logged but no error dialog appears."
    )

    add_page_break(doc)


# ---------------------------------------------------------------------------
# Appendices
# ---------------------------------------------------------------------------

def write_appendices(doc):
    """Write Appendix A, B, and C."""

    # Appendix A: Status Lifecycle Diagram
    doc.add_heading("Appendix A: Status Lifecycle Diagram", level=1)
    add_body(doc,
        "The following diagram shows the complete incident status lifecycle with all "
        "valid transitions:"
    )
    add_body(doc, "")

    # Text-based lifecycle diagram
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    lines = [
        "                    +--------+",
        "                    |  Open  |",
        "                    +---+----+",
        "                        |",
        "           +------------+------------+",
        "           |            |            |",
        "           v            v            v",
        "     +-----------+  +----------+  +-----------+",
        "     | Cancelled |  | Assigned |  | Duplicate |",
        "     +-----------+  +----+-----+  +-----------+",
        "                        |",
        "           +------------+------------+",
        "           |            |            |",
        "           v            v            v",
        "     +-----------+  +-------------+  +-----------+",
        "     | Cancelled |  | In Progress |  | Duplicate |",
        "     +-----------+  +------+------+  +-----------+",
        "                          |",
        "             +------------+------------+",
        "             |            |            |",
        "             v            v            v",
        "       +-----------+  +----------+  +-----------+",
        "       | Cancelled |  | Resolved |  | Duplicate |",
        "       +-----------+  +----+-----+  +-----------+",
        "                          |",
        "                          v",
        "                     +--------+",
        "                     | Closed |",
        "                     +--------+",
    ]
    for line in lines:
        run = p.add_run(line + "\n")
        run.font.name = "Consolas"
        run.font.size = Pt(8)
        run.font.color.rgb = BRIMIS_DARK

    add_body(doc, "")
    add_body(doc, "Key rules:")
    add_bullet(doc, "Open to Assigned transition is only available through the Assignment Form (not Status Update)")
    add_bullet(doc, "Cancelled and Duplicate are terminal states -- no further transitions allowed")
    add_bullet(doc, "Cancelled and Duplicate require a mandatory reason")
    add_bullet(doc, "Resolved requires Root Cause Analysis (at minimum, the Root Cause field)")
    add_bullet(doc, "Closed is the final positive-outcome terminal state")

    add_page_break(doc)

    # Appendix B: SLA Threshold Defaults
    doc.add_heading("Appendix B: SLA Threshold Defaults", level=1)
    add_body(doc,
        "The following table shows the default SLA thresholds configured in the system. "
        "These can be modified by an administrator in the Settings sheet (tblSLAThresholds)."
    )
    add_branded_table(doc,
        ["Priority", "Response Time", "Resolution Time", "At Risk After (Response)", "At Risk After (Resolution)"],
        [
            ["P1 (Critical)", "1 hour", "4 hours", "45 minutes", "3 hours"],
            ["P2 (High)", "4 hours", "24 hours", "3 hours", "18 hours"],
            ["P3 (Medium)", "8 hours", "72 hours (3 days)", "6 hours", "54 hours"],
            ["P4 (Low)", "24 hours", "168 hours (7 days)", "18 hours", "126 hours"],
        ],
        col_widths=[1.0, 1.2, 1.2, 1.3, 1.3]
    )
    add_body(doc,
        "SLA status definitions:"
    )
    add_bullet(doc, "On Track (green): Time elapsed is within the allowed window")
    add_bullet(doc, "At Risk (amber): More than 75% of the allowed time has elapsed")
    add_bullet(doc, "Overdue (red): The allowed time has been exceeded (incident still open)")
    add_bullet(doc, "Met (green): The incident met its SLA target (final status for resolved/closed)")
    add_bullet(doc, "Breached (red): The incident did not meet its SLA target (final status)")

    add_page_break(doc)

    # Appendix C: Column Reference (tblIncidents)
    doc.add_heading("Appendix C: Column Reference (tblIncidents)", level=1)
    add_body(doc,
        "The following table lists all 27 columns in the tblIncidents table on the Incident "
        "Log sheet. This is the master data table for all incidents."
    )
    add_branded_table(doc,
        ["#", "Column Name", "Description", "Set By"],
        [
            ["1", "IncidentID", "Unique identifier (INC-00001 format)", "System (auto-generated)"],
            ["2", "Title", "Brief description of the incident", "Reporter (Entry Form)"],
            ["3", "Description", "Detailed account of the incident", "Reporter (Entry Form)"],
            ["4", "Category", "Top-level classification (Mechanical, Electrical, Safety/HSE)", "Reporter (Entry Form)"],
            ["5", "Subcategory", "Specific classification within category", "Reporter (Entry Form)"],
            ["6", "Priority", "Calculated priority level (P1-P4)", "System (from matrix)"],
            ["7", "Impact", "Business impact level (High, Medium, Low)", "Reporter (Entry Form)"],
            ["8", "Urgency", "Time-sensitivity level (High, Medium, Low)", "Reporter (Entry Form)"],
            ["9", "Status", "Current lifecycle status", "System / Supervisor"],
            ["10", "ReportedBy", "Name of the person who reported the incident", "Reporter (Entry Form)"],
            ["11", "ReportedDate", "Date and time the incident was logged", "System (auto-timestamp)"],
            ["12", "AssignedTeam", "Team responsible for the incident", "Supervisor (Assignment Form)"],
            ["13", "AssignedTo", "Individual assigned to the incident", "Supervisor (Assignment Form)"],
            ["14", "AssignedDate", "Date and time of assignment", "System (auto-timestamp)"],
            ["15", "AssignedBy", "Supervisor who made the assignment", "Supervisor (Assignment Form)"],
            ["16", "ResponseDate", "Date work began (status changed to In Progress)", "System (auto-timestamp)"],
            ["17", "ResolutionDate", "Date incident was resolved", "System (auto-timestamp)"],
            ["18", "ClosedDate", "Date incident was formally closed", "System (auto-timestamp)"],
            ["19", "ResolutionNotes", "Notes about resolution or cancellation/duplicate reason", "Technician / Supervisor"],
            ["20", "RootCause", "Root cause identified during RCA", "Technician (Status Update)"],
            ["21", "CorrectiveAction", "Actions taken to fix the immediate issue", "Technician (Status Update)"],
            ["22", "PreventiveAction", "Actions to prevent recurrence", "Technician (Status Update)"],
            ["23", "AttachmentRef", "Reference to related files or documents", "Reporter (Entry Form)"],
            ["24", "SLAResponseStatus", "SLA status for response time (On Track/At Risk/Overdue/Met/Breached)", "System (calculated)"],
            ["25", "SLAResolutionStatus", "SLA status for resolution time", "System (calculated)"],
            ["26", "LastModified", "Timestamp of last change to any field", "System (auto-timestamp)"],
            ["27", "LastModifiedBy", "User/process that made the last change", "System"],
        ],
        col_widths=[0.3, 1.3, 2.5, 1.9]
    )


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    """Generate the BRIMIS IMS User Manual."""
    print("Creating BRIMIS IMS User Manual...")
    print(f"  Output: {OUTPUT_FILE}")

    doc = Document()

    # Page setup: A4 with reasonable margins
    for section in doc.sections:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)

    # Configure styles
    configure_styles(doc)

    # Build document
    write_cover_page(doc)
    write_table_of_contents(doc)
    write_chapter_1(doc)
    write_chapter_2(doc)
    write_chapter_3(doc)
    write_chapter_4(doc)
    write_chapter_5(doc)
    write_chapter_6(doc)
    write_chapter_7(doc)
    write_chapter_8(doc)
    write_chapter_9(doc)
    write_chapter_10(doc)
    write_chapter_11(doc)
    write_appendices(doc)

    # Configure header/footer
    configure_header_footer(doc)

    # Save
    doc.save(OUTPUT_FILE)
    print(f"  DONE: {OUTPUT_FILE}")
    print(f"  File size: {os.path.getsize(OUTPUT_FILE):,} bytes")


if __name__ == "__main__":
    main()
