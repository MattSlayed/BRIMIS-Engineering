# BRIMIS Incident Management System

An Excel-based incident management system built for **BRIMIS Engineering**, an industrial valve and pump engineering firm serving mining and energy clients including Anglo American, Glencore, Eskom, and Exxaro.

The system tracks industrial incidents from initial logging through categorization, assignment, resolution, and post-incident review — all within a branded Excel workbook with VBA-powered forms.

---

## Overview

Artisans and technicians log, track, and resolve industrial incidents through one intuitive Excel workbook — creating a single source of truth for BRIMIS operations that requires no training beyond basic computer literacy.

**Key Capabilities:**
- Incident logging via VBA UserForms (title, description, category, priority, impact, urgency, reporter)
- Priority matrix system (P1–P4) based on impact and urgency scoring
- SLA tracking with visual indicators for response and resolution times
- Dashboard with KPI metrics, category breakdowns, and resolution trends
- Post-incident review with root cause, corrective action, and preventive action fields

---

## Incident Categories

| Category | Subcategories |
|----------|---------------|
| **Mechanical** | Valve failures, pump breakdowns, equipment malfunction, wear & tear |
| **Electrical** | Motor failures, control panel issues, wiring faults |
| **Safety/HSE** | Near-misses, injuries, environmental spills, PPE issues |

## Status Lifecycle

```
Open → Assigned → In Progress → Resolved → Closed
```

Additional statuses: Cancelled, Duplicate

---

## Workbook Structure

| Sheet | Purpose |
|-------|---------|
| **Incident Log** | Master record of all logged incidents |
| **Dashboard** | KPI metrics — open incidents, overdue SLAs, category distribution, trends |
| **Assignment Tracker** | Team and individual assignment management |
| **SLA Monitor** | Response and resolution time tracking with conditional formatting |
| **RCA Log** | Post-incident root cause analysis and corrective actions |
| **Settings** | Configuration for teams, personnel, categories, SLA thresholds |

---

## Project Structure

```
├── BRIMIS_IMS.xlsm              # Production workbook with VBA
├── BRIMIS_IMS.xlsx               # Data-only version (no macros)
├── BRIMIS_IMS_User_Manual.docx   # Step-by-step user guide
└── src/
    ├── create_workbook.py        # Workbook generation script
    ├── create_user_manual.py     # User manual generation script
    ├── inject_vba.py             # VBA module injection into workbook
    └── vba_modules/              # VBA source code
        ├── frmIncidentEntry.bas  # Incident logging form
        ├── frmAssignment.bas     # Assignment form
        ├── frmStatusUpdate.bas   # Status update form
        ├── frmSearch.bas         # Search & filter form
        ├── modDashboard.bas      # Dashboard calculations
        ├── modSLA.bas            # SLA tracking logic
        ├── modReports.bas        # Report generation
        └── ...                   # Supporting modules
```

---

## Users

- **Artisans / Field Workers** — Report incidents via simple forms
- **Supervisors / Dispatchers** — Triage, prioritize, and assign incidents
- **Technicians / Engineers** — Document investigation and resolution
- **Management** — Review dashboards, SLA compliance, and post-incident summaries

---

## Design Decisions

| Decision | Rationale |
|----------|-----------|
| Excel + VBA (no web app) | Artisans need familiar tools; no internet dependency on-site |
| VBA forms for data entry | Guided input reduces errors for low-tech-comfort users |
| Visual indicators only (no email) | Self-contained system with no Outlook dependency |
| Team + individual assignment | Reflects real workflow: dispatcher → team lead → technician |
| Simple RCA (not 5-Why template) | Proportional to incident volume (<50/month) |

---

## Branding

Built to BRIMIS corporate identity:
- **Primary Red:** `#ee3124`
- **Dark:** `#101010`
- **White:** `#ffffff`
- **Gray:** `#32373c`

---

## Built By

**Matthew Koeberg — [NOVATEK LLC](https://github.com/MattSlayed)**
