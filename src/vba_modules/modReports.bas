Attribute VB_Name = "modReports"
Option Explicit

' ============================================================================
' Module:  modReports
' Purpose: Generates print-ready incident reports on the hidden IncidentReport
'          sheet. Supports both PrintPreview and PDF export. Uses BRIMIS
'          branding with professional layout including all incident fields:
'          overview, description, timeline, assignment, RCA, and SLA.
' Dependencies: modConstants (SHT_REPORT, COL_*, CLR_*, APP_TITLE)
'               modDataAccess (ReadIncident, GetIncidentColIdx)
'               modUtilities (UnprotectSheet, ProtectSheet)
'               modErrorHandler (HandleError)
' ============================================================================

' ----------------------------------------------------------------------------
' GenerateIncidentReport
' Populates the IncidentReport sheet with full incident data and shows
' PrintPreview for printing.
'
' Parameters:
'   sIncidentID - The incident ID to generate a report for
' ----------------------------------------------------------------------------
Public Sub GenerateIncidentReport(ByVal sIncidentID As String)
    On Error GoTo ErrHandler

    Dim vData As Variant
    vData = modDataAccess.ReadIncident(sIncidentID)
    If IsEmpty(vData) Then
        MsgBox "Incident " & sIncidentID & " not found.", vbExclamation, APP_TITLE
        Exit Sub
    End If

    Dim wsReport As Worksheet
    Set wsReport = ThisWorkbook.Sheets(SHT_REPORT)

    Application.ScreenUpdating = False
    modUtilities.UnprotectSheet wsReport

    PopulateReportSheet wsReport, vData
    ConfigureReportPageSetup wsReport

    modUtilities.ProtectSheet wsReport
    Application.ScreenUpdating = True

    wsReport.PrintPreview

    Exit Sub
ErrHandler:
    On Error Resume Next
    modUtilities.ProtectSheet ThisWorkbook.Sheets(SHT_REPORT)
    Application.ScreenUpdating = True
    On Error GoTo 0
    modErrorHandler.HandleError "modReports", "GenerateIncidentReport", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' ExportReportToPDF
' Populates the IncidentReport sheet with full incident data and exports
' to PDF in the workbook directory.
'
' Parameters:
'   sIncidentID - The incident ID to export
' ----------------------------------------------------------------------------
Public Sub ExportReportToPDF(ByVal sIncidentID As String)
    On Error GoTo ErrHandler

    Dim vData As Variant
    vData = modDataAccess.ReadIncident(sIncidentID)
    If IsEmpty(vData) Then
        MsgBox "Incident " & sIncidentID & " not found.", vbExclamation, APP_TITLE
        Exit Sub
    End If

    Dim wsReport As Worksheet
    Set wsReport = ThisWorkbook.Sheets(SHT_REPORT)

    Application.ScreenUpdating = False
    modUtilities.UnprotectSheet wsReport

    ' Populate report (call shared helper)
    PopulateReportSheet wsReport, vData

    ' Configure PageSetup
    ConfigureReportPageSetup wsReport

    modUtilities.ProtectSheet wsReport

    ' Build PDF path
    Dim sPath As String
    sPath = ThisWorkbook.Path & "\" & sIncidentID & "_Report.pdf"

    ' Export
    wsReport.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=sPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True

    Application.ScreenUpdating = True
    MsgBox "Report exported to:" & vbCrLf & sPath, vbInformation, APP_TITLE

    Exit Sub
ErrHandler:
    On Error Resume Next
    modUtilities.ProtectSheet ThisWorkbook.Sheets(SHT_REPORT)
    Application.ScreenUpdating = True
    On Error GoTo 0
    modErrorHandler.HandleError "modReports", "ExportReportToPDF", _
                                 Err.Number, Err.Description
End Sub

' ============================================================================
' Private Helpers
' ============================================================================

' ----------------------------------------------------------------------------
' PopulateReportSheet
' Fills the IncidentReport sheet with all incident data using fixed cell
' positions and BRIMIS branding.
' ----------------------------------------------------------------------------
Private Sub PopulateReportSheet(wsReport As Worksheet, vData As Variant)
    ' Clear previous report data (preserve any logo in A1)
    wsReport.Range("B1:H50").ClearContents
    wsReport.Range("A2:A50").ClearContents
    wsReport.Range("A4:H50").ClearFormats

    ' --- Row 1: Report header ---
    ' A1: Logo area (placed by create_workbook.py, do not overwrite)
    With wsReport.Range("D1")
        .Value = "INCIDENT REPORT"
        .Font.Name = "Calibri"
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = CLR_BRIMIS_RED
    End With
    With wsReport.Range("G1")
        .Value = "Report Date: " & Format(Now, "dd-mmm-yyyy")
        .Font.Name = "Calibri"
        .Font.Size = 9
        .Font.Color = CLR_BRIMIS_GRAY
        .HorizontalAlignment = xlRight
    End With

    ' --- Row 2: Company name ---
    With wsReport.Range("D2")
        .Value = "BRIMIS Engineering"
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Color = CLR_BRIMIS_GRAY
    End With

    ' --- Row 3: Spacer ---

    ' --- Row 4: INCIDENT OVERVIEW header bar ---
    WriteReportHeader wsReport, 4, "INCIDENT OVERVIEW"

    ' --- Row 5: ID and Status ---
    WriteReportLabel wsReport, 5, 1, "Incident ID:"
    WriteReportValue wsReport, 5, 2, GetVal(vData, COL_INCIDENT_ID)
    WriteReportLabel wsReport, 5, 4, "Status:"
    WriteReportValue wsReport, 5, 5, GetVal(vData, COL_STATUS)

    ' --- Row 6: Title (merged B6:H6) ---
    WriteReportLabel wsReport, 6, 1, "Title:"
    wsReport.Range("B6:H6").Merge
    With wsReport.Range("B6")
        .Value = GetVal(vData, COL_TITLE)
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Color = CLR_BRIMIS_DARK
        .WrapText = True
    End With

    ' --- Row 7: Priority and Category ---
    WriteReportLabel wsReport, 7, 1, "Priority:"
    WriteReportValue wsReport, 7, 2, GetVal(vData, COL_PRIORITY)
    WriteReportLabel wsReport, 7, 4, "Category:"
    WriteReportValue wsReport, 7, 5, GetVal(vData, COL_CATEGORY)

    ' --- Row 8: Subcategory ---
    WriteReportLabel wsReport, 8, 1, "Subcategory:"
    WriteReportValue wsReport, 8, 2, GetVal(vData, COL_SUBCATEGORY)

    ' --- Row 9: Spacer ---

    ' --- Row 10: DESCRIPTION header bar ---
    WriteReportHeader wsReport, 10, "DESCRIPTION"

    ' --- Rows 11-14: Description text (merged A11:H14) ---
    wsReport.Range("A11:H14").Merge
    With wsReport.Range("A11")
        .Value = GetVal(vData, COL_DESCRIPTION)
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Color = CLR_BRIMIS_DARK
        .WrapText = True
        .VerticalAlignment = xlTop
    End With

    ' --- Row 15: Spacer ---

    ' --- Row 16: TIMELINE header bar ---
    WriteReportHeader wsReport, 16, "TIMELINE"

    ' --- Row 17: Reported ---
    WriteReportLabel wsReport, 17, 1, "Reported:"
    WriteReportValue wsReport, 17, 2, GetDateVal(vData, COL_REPORTED_DATE)
    WriteReportLabel wsReport, 17, 4, "By:"
    WriteReportValue wsReport, 17, 5, GetVal(vData, COL_REPORTED_BY)

    ' --- Row 18: Assigned ---
    WriteReportLabel wsReport, 18, 1, "Assigned:"
    WriteReportValue wsReport, 18, 2, GetDateVal(vData, COL_ASSIGNED_DATE)
    WriteReportLabel wsReport, 18, 4, "To:"
    Dim sAssignee As String
    sAssignee = GetVal(vData, COL_ASSIGNED_TO)
    Dim sTeam As String
    sTeam = GetVal(vData, COL_ASSIGNED_TEAM)
    If Len(sAssignee) > 0 And Len(sTeam) > 0 Then
        WriteReportValue wsReport, 18, 5, sAssignee & " (" & sTeam & ")"
    ElseIf Len(sAssignee) > 0 Then
        WriteReportValue wsReport, 18, 5, sAssignee
    End If

    ' --- Row 19: Response ---
    WriteReportLabel wsReport, 19, 1, "Response:"
    WriteReportValue wsReport, 19, 2, GetDateVal(vData, COL_RESPONSE_DATE)

    ' --- Row 20: Resolution ---
    WriteReportLabel wsReport, 20, 1, "Resolution:"
    WriteReportValue wsReport, 20, 2, GetDateVal(vData, COL_RESOLUTION_DATE)

    ' --- Row 21: Closed ---
    WriteReportLabel wsReport, 21, 1, "Closed:"
    WriteReportValue wsReport, 21, 2, GetDateVal(vData, COL_CLOSED_DATE)

    ' --- Row 22: Spacer ---

    ' --- Row 23: ASSIGNMENT header bar ---
    WriteReportHeader wsReport, 23, "ASSIGNMENT"

    ' --- Row 24: Team and Assigned To ---
    WriteReportLabel wsReport, 24, 1, "Team:"
    WriteReportValue wsReport, 24, 2, GetVal(vData, COL_ASSIGNED_TEAM)
    WriteReportLabel wsReport, 24, 4, "Assigned To:"
    WriteReportValue wsReport, 24, 5, GetVal(vData, COL_ASSIGNED_TO)

    ' --- Row 25: Assigned Date and Assigned By ---
    WriteReportLabel wsReport, 25, 1, "Assigned Date:"
    WriteReportValue wsReport, 25, 2, GetDateVal(vData, COL_ASSIGNED_DATE)
    WriteReportLabel wsReport, 25, 4, "Assigned By:"
    WriteReportValue wsReport, 25, 5, GetVal(vData, COL_ASSIGNED_BY)

    ' --- Row 26: Spacer ---

    ' --- Row 27: RESOLUTION & ROOT CAUSE ANALYSIS header bar ---
    WriteReportHeader wsReport, 27, "RESOLUTION & ROOT CAUSE ANALYSIS"

    ' --- Row 28: Root Cause (merged B28:H28) ---
    WriteReportLabel wsReport, 28, 1, "Root Cause:"
    wsReport.Range("B28:H28").Merge
    With wsReport.Range("B28")
        .Value = GetVal(vData, COL_ROOT_CAUSE)
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Color = CLR_BRIMIS_DARK
        .WrapText = True
    End With

    ' --- Row 29: Corrective Action (merged B29:H29) ---
    WriteReportLabel wsReport, 29, 1, "Corrective Action:"
    wsReport.Range("B29:H29").Merge
    With wsReport.Range("B29")
        .Value = GetVal(vData, COL_CORRECTIVE_ACTION)
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Color = CLR_BRIMIS_DARK
        .WrapText = True
    End With

    ' --- Row 30: Preventive Action (merged B30:H30) ---
    WriteReportLabel wsReport, 30, 1, "Preventive Action:"
    wsReport.Range("B30:H30").Merge
    With wsReport.Range("B30")
        .Value = GetVal(vData, COL_PREVENTIVE_ACTION)
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Color = CLR_BRIMIS_DARK
        .WrapText = True
    End With

    ' --- Row 31: Resolution Notes (merged B31:H31) ---
    WriteReportLabel wsReport, 31, 1, "Resolution Notes:"
    wsReport.Range("B31:H31").Merge
    With wsReport.Range("B31")
        .Value = GetVal(vData, COL_RESOLUTION_NOTES)
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Color = CLR_BRIMIS_DARK
        .WrapText = True
    End With

    ' --- Row 32: Spacer ---

    ' --- Row 33: SLA PERFORMANCE header bar ---
    WriteReportHeader wsReport, 33, "SLA PERFORMANCE"

    ' --- Row 34: SLA statuses ---
    WriteReportLabel wsReport, 34, 1, "Response SLA:"
    WriteReportValue wsReport, 34, 2, GetVal(vData, COL_SLA_RESPONSE)
    WriteReportLabel wsReport, 34, 4, "Resolution SLA:"
    WriteReportValue wsReport, 34, 5, GetVal(vData, COL_SLA_RESOLUTION)

    ' --- Row 35: Spacer (end of report area) ---
End Sub

' ----------------------------------------------------------------------------
' ConfigureReportPageSetup
' Sets print layout: A4 portrait, fit to one page, BRIMIS footer.
' ----------------------------------------------------------------------------
Private Sub ConfigureReportPageSetup(wsReport As Worksheet)
    With wsReport.PageSetup
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .Zoom = False
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.5)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .CenterHorizontally = True
        .PrintArea = "A1:H35"
        .LeftFooter = "Generated by BRIMIS IMS on " & Format(Now, "dd-mmm-yyyy")
        .CenterFooter = "Page &P of &N"
        .RightFooter = "Confidential"
    End With
End Sub

' ----------------------------------------------------------------------------
' GetVal
' Safely extracts a string value from the incident data array.
' Returns "" for empty cells (which appear as 0 in Variant arrays).
' ----------------------------------------------------------------------------
Private Function GetVal(vData As Variant, ByVal sColName As String) As String
    On Error Resume Next
    Dim cIdx As Long
    cIdx = modDataAccess.GetIncidentColIdx(sColName)
    If cIdx > 0 Then
        GetVal = CStr(vData(1, cIdx))
        If GetVal = "0" Then GetVal = ""  ' Empty cells return 0
    Else
        GetVal = ""
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' GetDateVal
' Safely extracts and formats a date value from the incident data array.
' Returns "" if the value is not a valid date.
' ----------------------------------------------------------------------------
Private Function GetDateVal(vData As Variant, ByVal sColName As String) As String
    On Error Resume Next
    Dim cIdx As Long
    cIdx = modDataAccess.GetIncidentColIdx(sColName)
    If cIdx > 0 Then
        Dim vVal As Variant
        vVal = vData(1, cIdx)
        If IsDate(vVal) And Not IsEmpty(vVal) Then
            GetDateVal = Format(CDate(vVal), "yyyy-mm-dd hh:nn")
        Else
            GetDateVal = ""
        End If
    Else
        GetDateVal = ""
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' WriteReportHeader
' Writes a section header bar: merged A:H, BRIMIS_DARK fill, white bold text.
' Used for all 6 section headers (Overview, Description, Timeline, Assignment,
' Resolution/RCA, SLA Performance).
' ----------------------------------------------------------------------------
Private Sub WriteReportHeader(wsReport As Worksheet, lRow As Long, sText As String)
    Dim rng As Range
    Set rng = wsReport.Range(wsReport.Cells(lRow, 1), wsReport.Cells(lRow, 8))
    rng.Merge
    With rng
        .Value = sText
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = CLR_BRIMIS_WHITE
        .Interior.Color = CLR_BRIMIS_DARK
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .RowHeight = 22
    End With
End Sub

' ----------------------------------------------------------------------------
' WriteReportLabel
' Writes a field label with Calibri 10pt bold BRIMIS_GRAY font.
' ----------------------------------------------------------------------------
Private Sub WriteReportLabel(wsReport As Worksheet, lRow As Long, lCol As Long, sText As String)
    With wsReport.Cells(lRow, lCol)
        .Value = sText
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = CLR_BRIMIS_GRAY
    End With
End Sub

' ----------------------------------------------------------------------------
' WriteReportValue
' Writes a field value with Calibri 10pt BRIMIS_DARK font.
' ----------------------------------------------------------------------------
Private Sub WriteReportValue(wsReport As Worksheet, lRow As Long, lCol As Long, sText As String)
    With wsReport.Cells(lRow, lCol)
        .Value = sText
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = CLR_BRIMIS_DARK
    End With
End Sub
