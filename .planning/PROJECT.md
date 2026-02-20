# BRIMIS Incident Management System

## What This Is

An Excel-based incident management system for BRIMIS Engineering, an industrial valve and pump engineering firm serving mining and energy clients. The system tracks industrial incidents from initial logging through categorization, assignment, resolution, and post-incident review — all within a branded Excel workbook with VBA-powered forms for data entry and structured sheets for tracking and reporting.

## Core Value

Artisans and technicians can log, track, and resolve industrial incidents through one intuitive Excel workbook — creating a single source of truth for BRIMIS operations that requires no training beyond basic computer literacy.

## Requirements

### Validated

(None yet — ship to validate)

### Active

- [ ] Incident logging via VBA UserForm (title, description, category, priority, impact, urgency, reporter, attachments reference)
- [ ] Three incident categories with subcategories: Mechanical (valve failures, pump breakdowns, equipment malfunction, wear & tear), Electrical (motor failures, control panel issues, wiring faults), Safety/HSE (near-misses, injuries, environmental spills, PPE issues)
- [ ] Priority matrix system (P1-P4) based on impact and urgency scoring
- [ ] Dual assignment: incidents assigned to both a team (e.g., Pump Team) and a specific individual within that team
- [ ] Status lifecycle: Open → Assigned → In Progress → Resolved → Closed (plus Cancelled, Duplicate)
- [ ] SLA tracking with visual indicators (conditional formatting) for response time and resolution time per priority level
- [ ] Dashboard sheet with key metrics: open incidents, overdue SLAs, incidents by category, priority distribution, resolution trends
- [ ] Post-incident review with simple fields: root cause, corrective action, preventive action
- [ ] Settings/configuration sheet for managing teams, personnel, categories, SLA thresholds
- [ ] BRIMIS branding throughout: Red #ee3124, Dark #101010, White #ffffff, Gray #32373c
- [ ] Separate Word/PDF user manual with use case stories and step-by-step walkthroughs

### Out of Scope

- Email/Outlook notifications — visual indicators only, keeps system self-contained and simple
- Automated workflow routing — manual dispatch by supervisors, artisans don't need automation complexity
- Mobile or web access — Excel desktop workbook only
- Cloud/multi-site deployment — single workbook for single operation
- Integration with other BRIMIS systems — standalone tool

## Context

**Company:** BRIMIS Engineering — industrial valve and pump solutions (fitment, refurbishment, installation, repair, maintenance). Clients include Anglo American, Glencore, Eskom, CSIR, Exxaro, Columbus Stainless.

**Brand identity:** "Precision, quality, innovation and excellence." Modern industrial aesthetic — red accents on dark/neutral backgrounds. Colors: Primary Red #ee3124, Dark #101010, White #ffffff, Gray #32373c.

**Users:**
- **Artisans/field workers** — report incidents they encounter on the job. Low tech comfort, need the simplest possible interface.
- **Supervisors/dispatchers** — triage, categorize, prioritize, and assign incidents. Need clear overview of all open incidents.
- **Technicians/engineers** — receive assignments, document investigation and resolution. Need easy status updates.
- **Management** — review dashboards, SLA compliance, post-incident summaries. Need at-a-glance metrics.

**Volume:** Under 50 incidents per month — manageable in Excel without performance concerns.

**Workbook structure:** Hybrid approach — VBA UserForms for data entry (artisan-friendly), with separate functional sheets storing and displaying data (Incident Log, Dashboard, Assignment Tracker, SLA Monitor, RCA Log, Settings).

**Deliverables:** Excel workbook (.xlsm) + separate user manual document (Word/PDF) with use case stories.

## Constraints

- **Platform**: Microsoft Excel (.xlsm) with VBA — must work on standard Office installations
- **Users**: Artisans with basic computer literacy — every interaction must be obvious and forgiving
- **Branding**: Must reflect BRIMIS corporate identity (colors, logo placement)
- **Complexity**: Under 50 incidents/month — no need for database-level architecture
- **Self-contained**: No external dependencies (no add-ins, no internet required, no email integration)

## Key Decisions

| Decision | Rationale | Outcome |
|----------|-----------|---------|
| Excel hybrid (sheets + VBA forms) | Artisans need simple forms; supervisors need structured data views | — Complete |
| Visual indicators only, no email | Keeps system self-contained, no Outlook dependency | — Complete |
| 3 categories with subcategories | Covers Mechanical, Electrical, Safety/HSE — subcategories handle specifics | — Complete |
| Simple RCA (not 5-Why template) | Proportional to incident volume; detailed RCA would be overkill for <50/month | — Complete |
| Team + individual assignment | Reflects real workflow: dispatcher assigns to team lead who assigns to technician | — Complete |
| Separate user manual (Word/PDF) | Artisans need printed/offline reference; Excel Help sheet too limited | — Complete |

---
*Last updated: 2026-02-12 after initialization*
