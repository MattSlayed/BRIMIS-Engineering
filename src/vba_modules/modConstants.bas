Attribute VB_Name = "modConstants"
Option Explicit

' ============================================================================
' Module:  modConstants
' Purpose: All project-wide Public Const declarations.
'          No other module should contain magic strings for structural
'          references. Change a value here and it propagates everywhere.
' ============================================================================

' === Sheet Names ===
Public Const SHT_DASHBOARD As String = "Dashboard"
Public Const SHT_INCIDENT_LOG As String = "Incident Log"
Public Const SHT_SETTINGS As String = "Settings"

' === Table Names ===
Public Const TBL_INCIDENTS As String = "tblIncidents"
Public Const TBL_TEAMS As String = "tblTeams"
Public Const TBL_PERSONNEL As String = "tblPersonnel"
Public Const TBL_CATEGORIES As String = "tblCategories"
Public Const TBL_SLA_THRESHOLDS As String = "tblSLAThresholds"
Public Const TBL_PRIORITY_MATRIX As String = "tblPriorityMatrix"

' === Assignment Tracker ===
Public Const SHT_ASSIGNMENT_TRACKER As String = "Assignment Tracker"
Public Const TBL_ASSIGNMENT_TRACKER As String = "tblAssignmentTracker"

' === RCA Log ===
Public Const SHT_RCA_LOG As String = "RCA Log"
Public Const TBL_RCA_LOG As String = "tblRCALog"

' === RCA Log Column Names ===
Public Const COL_RCA_ID As String = "RCAID"
Public Const COL_RCA_INCIDENT_ID As String = "IncidentID"
Public Const COL_RCA_INCIDENT_TITLE As String = "IncidentTitle"
Public Const COL_RCA_ROOT_CAUSE As String = "RootCause"
Public Const COL_RCA_CORRECTIVE_ACTION As String = "CorrectiveAction"
Public Const COL_RCA_PREVENTIVE_ACTION As String = "PreventiveAction"
Public Const COL_RCA_RESOLUTION_NOTES As String = "ResolutionNotes"
Public Const COL_RCA_RESOLVED_BY As String = "ResolvedBy"
Public Const COL_RCA_RESOLVED_DATE As String = "ResolvedDate"
Public Const COL_RCA_LAST_MODIFIED As String = "LastModified"

Public Const RCA_ID_PREFIX As String = "RCA-"

' === Incident Log Column Names (27 columns) ===
Public Const COL_INCIDENT_ID As String = "IncidentID"
Public Const COL_TITLE As String = "Title"
Public Const COL_DESCRIPTION As String = "Description"
Public Const COL_CATEGORY As String = "Category"
Public Const COL_SUBCATEGORY As String = "Subcategory"
Public Const COL_PRIORITY As String = "Priority"
Public Const COL_IMPACT As String = "Impact"
Public Const COL_URGENCY As String = "Urgency"
Public Const COL_STATUS As String = "Status"
Public Const COL_REPORTED_BY As String = "ReportedBy"
Public Const COL_REPORTED_DATE As String = "ReportedDate"
Public Const COL_ASSIGNED_TEAM As String = "AssignedTeam"
Public Const COL_ASSIGNED_TO As String = "AssignedTo"
Public Const COL_ASSIGNED_DATE As String = "AssignedDate"
Public Const COL_ASSIGNED_BY As String = "AssignedBy"
Public Const COL_RESPONSE_DATE As String = "ResponseDate"
Public Const COL_RESOLUTION_DATE As String = "ResolutionDate"
Public Const COL_CLOSED_DATE As String = "ClosedDate"
Public Const COL_RESOLUTION_NOTES As String = "ResolutionNotes"
Public Const COL_ROOT_CAUSE As String = "RootCause"
Public Const COL_CORRECTIVE_ACTION As String = "CorrectiveAction"
Public Const COL_PREVENTIVE_ACTION As String = "PreventiveAction"
Public Const COL_ATTACHMENT_REF As String = "AttachmentRef"
Public Const COL_SLA_RESPONSE As String = "SLAResponseStatus"
Public Const COL_SLA_RESOLUTION As String = "SLAResolutionStatus"
Public Const COL_LAST_MODIFIED As String = "LastModified"
Public Const COL_LAST_MODIFIED_BY As String = "LastModifiedBy"

' === BRIMIS Branding Colors (as Long for Excel color properties) ===
' Formula: Long = Red + (Green * 256) + (Blue * 65536)
Public Const CLR_BRIMIS_RED As Long = 2372078       ' RGB(238, 49, 36)  -> 238 + 12544 + 2359296
Public Const CLR_BRIMIS_DARK As Long = 1052688      ' RGB(16, 16, 16)   -> 16 + 4096 + 1048576
Public Const CLR_BRIMIS_WHITE As Long = 16777215    ' RGB(255, 255, 255) -> 255 + 65280 + 16711680
Public Const CLR_BRIMIS_GRAY As Long = 3946290      ' RGB(50, 55, 60)   -> 50 + 14080 + 3932160

' === Status Values ===
Public Const STATUS_OPEN As String = "Open"
Public Const STATUS_ASSIGNED As String = "Assigned"
Public Const STATUS_IN_PROGRESS As String = "In Progress"
Public Const STATUS_RESOLVED As String = "Resolved"
Public Const STATUS_CLOSED As String = "Closed"
Public Const STATUS_CANCELLED As String = "Cancelled"
Public Const STATUS_DUPLICATE As String = "Duplicate"

' === Priority Values ===
Public Const PRIORITY_P1 As String = "P1"
Public Const PRIORITY_P2 As String = "P2"
Public Const PRIORITY_P3 As String = "P3"
Public Const PRIORITY_P4 As String = "P4"

' === Backup Configuration ===
Public Const BACKUP_SUBFOLDER As String = "Backups"
Public Const BACKUP_PREFIX As String = "BRIMIS_IMS_Backup_"
Public Const MAX_BACKUPS As Long = 30

' === System ===
Public Const SHEET_PWD As String = "BRIMIS2026"
Public Const APP_TITLE As String = "BRIMIS Incident Management System"
Public Const ID_PREFIX As String = "INC-"

' === SLA Status Values ===
Public Const SLA_STATUS_ON_TRACK As String = "On Track"
Public Const SLA_STATUS_AT_RISK As String = "At Risk"
Public Const SLA_STATUS_OVERDUE As String = "Overdue"
Public Const SLA_STATUS_MET As String = "Met"
Public Const SLA_STATUS_BREACHED As String = "Breached"

' === SLA Threshold ===
Public Const SLA_AMBER_THRESHOLD_PCT As Double = 0.75  ' Amber when 75% of time elapsed

' === SLA Conditional Formatting Colors (Long = R + G*256 + B*65536) ===
Public Const CLR_SLA_GREEN As Long = 13561798   ' RGB(198, 239, 206) = 198 + (239*256) + (206*65536)
Public Const CLR_SLA_GREEN_TEXT As Long = 24832  ' RGB(0, 97, 0) dark green text
Public Const CLR_SLA_AMBER As Long = 10284031   ' RGB(255, 235, 156) = 255 + (235*256) + (156*65536)
Public Const CLR_SLA_AMBER_TEXT As Long = 32896  ' RGB(128, 128, 0) dark amber text

' === Dashboard Layout Constants ===
Public Const DASH_KPI_SECTION_ROW As Long = 5
Public Const DASH_KPI_LABEL_ROW As Long = 6
Public Const DASH_KPI_VALUE_ROW As Long = 7
Public Const DASH_KPI_COL_OPEN As Long = 2       ' Column B
Public Const DASH_KPI_COL_OVERDUE As Long = 6    ' Column F
Public Const DASH_KPI_COL_AVGRES As Long = 10    ' Column J
Public Const DASH_KPI_COL_CLOSED As Long = 14    ' Column N

Public Const DASH_BREAKDOWN_HEADER_ROW As Long = 9
Public Const DASH_PRIORITY_START_ROW As Long = 10
Public Const DASH_PRIORITY_COL As Long = 2       ' Column B (label)
Public Const DASH_PRIORITY_VAL_COL As Long = 3   ' Column C (value)
Public Const DASH_CATEGORY_START_ROW As Long = 10
Public Const DASH_CATEGORY_COL As Long = 8       ' Column H (label)
Public Const DASH_CATEGORY_VAL_COL As Long = 9   ' Column I (value)

Public Const DASH_SLA_HEADER_ROW As Long = 16
Public Const DASH_SLA_COL_HEADERS_ROW As Long = 17
Public Const DASH_SLA_DATA_START_ROW As Long = 18
Public Const DASH_SLA_MAX_ROWS As Long = 20      ' Max incidents shown in SLA monitor
Public Const DASH_SLA_COL_ID As Long = 2         ' Column B
Public Const DASH_SLA_COL_TITLE As Long = 3
Public Const DASH_SLA_COL_PRIORITY As Long = 4
Public Const DASH_SLA_COL_STATUS As Long = 5
Public Const DASH_SLA_COL_RESPONSE As Long = 6
Public Const DASH_SLA_COL_RESOLUTION As Long = 7
Public Const DASH_SLA_COL_TIMEREM As Long = 8

Public Const DASH_REFRESH_ROW As Long = 39
