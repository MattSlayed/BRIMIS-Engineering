Attribute VB_Name = "modDashboard"
Option Explicit

' ============================================================================
' Module:  modDashboard
' Purpose: Dashboard KPI rendering and SLA Monitor table. RefreshDashboard is
'          the master refresh routine that recalculates SLA statuses, applies
'          conditional formatting, computes KPIs from tblIncidents, and writes
'          the SLA Monitor table to the Dashboard sheet.
'
'          Called from: Worksheet_Activate on Dashboard, Refresh button,
'          and modInitialize.InitializeWorkbook.
'
' Dependencies: modConstants (SHT_*, TBL_*, COL_*, DASH_*, CLR_*, STATUS_*,
'                SLA_STATUS_*, PRIORITY_*)
'               modSLA (RecalculateAllSLAStatuses, LoadSLAThresholds,
'                FormatTimeRemaining)
'               modFormatting (ApplySLAConditionalFormatting)
'               modUtilities (UnprotectSheet, ProtectSheet)
'               modErrorHandler (HandleError)
' ============================================================================

' Module-level re-entrancy guard to prevent triple-refresh on workbook open.
' During InitializeWorkbook, multiple refresh triggers fire in sequence:
' (1) the direct RefreshDashboard call in InitializeWorkbook, and
' (2) the Worksheet_Activate event when navigating to the Dashboard tab.
' The guard ensures only the first call executes; subsequent re-entrant calls
' exit immediately.
Private m_bRefreshing As Boolean

' ----------------------------------------------------------------------------
' RefreshDashboard
' Master refresh routine. Recalculates SLA, applies conditional formatting,
' computes all KPIs in a single loop, and writes the Dashboard content.
' ----------------------------------------------------------------------------
Public Sub RefreshDashboard()
    On Error GoTo ErrHandler

    ' Re-entrancy guard
    If m_bRefreshing Then Exit Sub
    m_bRefreshing = True

    Application.ScreenUpdating = False

    ' Save and set calculation mode
    Dim lCalcMode As Long
    lCalcMode = Application.Calculation
    Application.Calculation = xlCalculationManual

    ' Step 1: Recalculate all SLA statuses in tblIncidents
    modSLA.RecalculateAllSLAStatuses

    ' Step 2: Apply SLA conditional formatting to Incident Log columns
    modFormatting.ApplySLAConditionalFormatting

    ' Step 3: Get Dashboard worksheet
    Dim wsDash As Worksheet
    On Error Resume Next
    Set wsDash = ThisWorkbook.Sheets(SHT_DASHBOARD)
    On Error GoTo ErrHandler
    If wsDash Is Nothing Then GoTo CleanUp

    modUtilities.UnprotectSheet wsDash

    ' Step 4: Get tblIncidents for KPI computation
    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = ThisWorkbook.Sheets(SHT_INCIDENT_LOG).ListObjects(TBL_INCIDENTS)
    On Error GoTo ErrHandler

    ' Load SLA thresholds for time remaining calculation
    Dim dictThresholds As Object
    Set dictThresholds = modSLA.LoadSLAThresholds()

    ' Initialize KPI accumulators
    Dim lTotalOpen As Long, lTotalOverdue As Long, lTotalClosed As Long
    Dim dTotalResDays As Double, lResCount As Long
    lTotalOpen = 0
    lTotalOverdue = 0
    lTotalClosed = 0
    dTotalResDays = 0
    lResCount = 0

    ' Breakdown dictionaries
    Dim dictPriority As Object, dictCategory As Object
    Set dictPriority = CreateObject("Scripting.Dictionary")
    Set dictCategory = CreateObject("Scripting.Dictionary")

    ' SLA Monitor data collection (capped at DASH_SLA_MAX_ROWS)
    ' Each entry: Array(ID, Title, Priority, Status, SLAResponse, SLAResolution, ReportedDate)
    Dim arrSLAMonitor() As Variant
    Dim lSLACount As Long, lTotalOpenForMore As Long
    lSLACount = 0
    lTotalOpenForMore = 0
    ReDim arrSLAMonitor(1 To DASH_SLA_MAX_ROWS, 1 To 7)

    ' Step 5: Single loop through tblIncidents to accumulate all KPI values
    If Not tbl Is Nothing Then
        If Not tbl.DataBodyRange Is Nothing Then
            If tbl.ListRows.Count > 0 Then
                ' Column indices
                Dim cID As Long, cTitle As Long, cPriority As Long
                Dim cCategory As Long, cStatus As Long
                Dim cReportedDate As Long, cResponseDate As Long
                Dim cResolutionDate As Long
                Dim cSLAResponse As Long, cSLAResolution As Long

                cID = tbl.ListColumns(COL_INCIDENT_ID).Index
                cTitle = tbl.ListColumns(COL_TITLE).Index
                cPriority = tbl.ListColumns(COL_PRIORITY).Index
                cCategory = tbl.ListColumns(COL_CATEGORY).Index
                cStatus = tbl.ListColumns(COL_STATUS).Index
                cReportedDate = tbl.ListColumns(COL_REPORTED_DATE).Index
                cResponseDate = tbl.ListColumns(COL_RESPONSE_DATE).Index
                cResolutionDate = tbl.ListColumns(COL_RESOLUTION_DATE).Index
                cSLAResponse = tbl.ListColumns(COL_SLA_RESPONSE).Index
                cSLAResolution = tbl.ListColumns(COL_SLA_RESOLUTION).Index

                Dim i As Long
                For i = 1 To tbl.ListRows.Count
                    Dim sStatus As String
                    sStatus = CStr(tbl.ListRows(i).Range(1, cStatus).Value)

                    Dim sPriority As String
                    sPriority = CStr(tbl.ListRows(i).Range(1, cPriority).Value)

                    Dim sCategory As String
                    sCategory = CStr(tbl.ListRows(i).Range(1, cCategory).Value)

                    ' Count open incidents (Open, Assigned, In Progress)
                    If sStatus = STATUS_OPEN Or sStatus = STATUS_ASSIGNED Or sStatus = STATUS_IN_PROGRESS Then
                        lTotalOpen = lTotalOpen + 1

                        ' Priority breakdown (open only)
                        If dictPriority.Exists(sPriority) Then
                            dictPriority(sPriority) = dictPriority(sPriority) + 1
                        Else
                            dictPriority.Add sPriority, 1
                        End If

                        ' Category breakdown (open only)
                        If Len(sCategory) > 0 Then
                            If dictCategory.Exists(sCategory) Then
                                dictCategory(sCategory) = dictCategory(sCategory) + 1
                            Else
                                dictCategory.Add sCategory, 1
                            End If
                        End If

                        ' Overdue count (among non-terminal)
                        Dim sSLAResp As String, sSLARes As String
                        sSLAResp = CStr(tbl.ListRows(i).Range(1, cSLAResponse).Value)
                        sSLARes = CStr(tbl.ListRows(i).Range(1, cSLAResolution).Value)
                        If sSLAResp = SLA_STATUS_OVERDUE Or sSLARes = SLA_STATUS_OVERDUE Then
                            lTotalOverdue = lTotalOverdue + 1
                        End If

                        ' Collect for SLA Monitor (capped)
                        lTotalOpenForMore = lTotalOpenForMore + 1
                        If lSLACount < DASH_SLA_MAX_ROWS Then
                            lSLACount = lSLACount + 1
                            arrSLAMonitor(lSLACount, 1) = CStr(tbl.ListRows(i).Range(1, cID).Value)
                            arrSLAMonitor(lSLACount, 2) = CStr(tbl.ListRows(i).Range(1, cTitle).Value)
                            arrSLAMonitor(lSLACount, 3) = sPriority
                            arrSLAMonitor(lSLACount, 4) = sStatus
                            arrSLAMonitor(lSLACount, 5) = sSLAResp
                            arrSLAMonitor(lSLACount, 6) = sSLARes
                            arrSLAMonitor(lSLACount, 7) = tbl.ListRows(i).Range(1, cReportedDate).Value
                        End If
                    End If

                    ' Count closed
                    If sStatus = STATUS_CLOSED Then
                        lTotalClosed = lTotalClosed + 1
                    End If

                    ' Average resolution time (Resolved or Closed with both dates)
                    If sStatus = STATUS_RESOLVED Or sStatus = STATUS_CLOSED Then
                        Dim vResDate As Variant, vRepDate As Variant
                        vResDate = tbl.ListRows(i).Range(1, cResolutionDate).Value
                        vRepDate = tbl.ListRows(i).Range(1, cReportedDate).Value
                        If IsDate(vResDate) And Not IsEmpty(vResDate) And _
                           IsDate(vRepDate) And Not IsEmpty(vRepDate) Then
                            dTotalResDays = dTotalResDays + (CDate(vResDate) - CDate(vRepDate))
                            lResCount = lResCount + 1
                        End If
                    End If
                Next i
            End If
        End If
    End If

    ' Calculate average resolution days
    Dim dAvgResDays As Double
    If lResCount > 0 Then
        dAvgResDays = dTotalResDays / lResCount
    Else
        dAvgResDays = -1  ' Sentinel for "N/A"
    End If

    ' Step 6: Write KPIs
    WriteDashboardKPIs wsDash, lTotalOpen, lTotalOverdue, dAvgResDays, lTotalClosed

    ' Step 7: Write breakdowns
    WritePriorityBreakdown wsDash, dictPriority
    WriteCategoryBreakdown wsDash, dictCategory

    ' Step 8: Write SLA Monitor table
    WriteSLAMonitorTable wsDash, arrSLAMonitor, lSLACount, lTotalOpenForMore, dictThresholds

    ' Step 9: Write last refreshed timestamp
    With wsDash.Cells(DASH_REFRESH_ROW, 2)
        .Value = "Last Refreshed: " & Format(Now, "dd-mmm-yyyy hh:nn:ss")
        .Font.Name = "Calibri"
        .Font.Size = 9
        .Font.Italic = True
        .Font.Color = CLR_BRIMIS_GRAY
    End With

    ' Re-protect Dashboard
    modUtilities.ProtectSheet wsDash

CleanUp:
    ' Restore calculation mode and screen updating
    Application.Calculation = lCalcMode
    Application.ScreenUpdating = True
    m_bRefreshing = False
    Exit Sub

ErrHandler:
    ' Re-protect and restore state before surfacing error
    On Error Resume Next
    Dim wsClean As Worksheet
    Set wsClean = ThisWorkbook.Sheets(SHT_DASHBOARD)
    If Not wsClean Is Nothing Then modUtilities.ProtectSheet wsClean
    Application.Calculation = lCalcMode
    Application.ScreenUpdating = True
    m_bRefreshing = False
    On Error GoTo 0

    modErrorHandler.HandleError "modDashboard", "RefreshDashboard", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' WriteDashboardKPIs
' Writes KPI labels and large-font values to the Dashboard sheet.
'
' Parameters:
'   ws          - Dashboard worksheet
'   lOpen       - Total open incidents
'   lOverdue    - Total overdue incidents
'   dAvgResDays - Average resolution days (-1 for N/A)
'   lClosed     - Total closed incidents
' ----------------------------------------------------------------------------
Private Sub WriteDashboardKPIs(ws As Worksheet, ByVal lOpen As Long, _
                                ByVal lOverdue As Long, _
                                ByVal dAvgResDays As Double, _
                                ByVal lClosed As Long)

    ' Section header
    With ws.Cells(DASH_KPI_SECTION_ROW, DASH_KPI_COL_OPEN)
        .Value = "Key Performance Indicators"
        .Font.Name = "Calibri"
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = CLR_BRIMIS_DARK
    End With

    ' --- Total Open ---
    WriteKPILabel ws, DASH_KPI_LABEL_ROW, DASH_KPI_COL_OPEN, "Total Open"
    WriteKPIValue ws, DASH_KPI_VALUE_ROW, DASH_KPI_COL_OPEN, lOpen, CLR_BRIMIS_DARK

    ' --- Overdue ---
    WriteKPILabel ws, DASH_KPI_LABEL_ROW, DASH_KPI_COL_OVERDUE, "Overdue"
    Dim lOverdueColor As Long
    If lOverdue > 0 Then
        lOverdueColor = CLR_BRIMIS_RED
    Else
        lOverdueColor = CLR_BRIMIS_DARK
    End If
    WriteKPIValue ws, DASH_KPI_VALUE_ROW, DASH_KPI_COL_OVERDUE, lOverdue, lOverdueColor

    ' --- Avg Resolution ---
    WriteKPILabel ws, DASH_KPI_LABEL_ROW, DASH_KPI_COL_AVGRES, "Avg Resolution (days)"
    If dAvgResDays < 0 Then
        WriteKPIValue ws, DASH_KPI_VALUE_ROW, DASH_KPI_COL_AVGRES, "N/A", CLR_BRIMIS_DARK
    Else
        WriteKPIValue ws, DASH_KPI_VALUE_ROW, DASH_KPI_COL_AVGRES, Format(dAvgResDays, "0.0"), CLR_BRIMIS_DARK
    End If

    ' --- Total Closed ---
    WriteKPILabel ws, DASH_KPI_LABEL_ROW, DASH_KPI_COL_CLOSED, "Total Closed"
    WriteKPIValue ws, DASH_KPI_VALUE_ROW, DASH_KPI_COL_CLOSED, lClosed, CLR_BRIMIS_DARK

End Sub

' ----------------------------------------------------------------------------
' WriteKPILabel
' Writes a KPI label in smaller bold font.
' ----------------------------------------------------------------------------
Private Sub WriteKPILabel(ws As Worksheet, ByVal lRow As Long, _
                           ByVal lCol As Long, ByVal sLabel As String)
    With ws.Cells(lRow, lCol)
        .Value = sLabel
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = CLR_BRIMIS_GRAY
        .HorizontalAlignment = xlCenter
    End With
End Sub

' ----------------------------------------------------------------------------
' WriteKPIValue
' Writes a KPI value in large bold font with color coding.
' ----------------------------------------------------------------------------
Private Sub WriteKPIValue(ws As Worksheet, ByVal lRow As Long, _
                           ByVal lCol As Long, ByVal vValue As Variant, _
                           ByVal lColor As Long)
    With ws.Cells(lRow, lCol)
        .Value = vValue
        .Font.Name = "Calibri"
        .Font.Size = 28
        .Font.Bold = True
        .Font.Color = lColor
        .HorizontalAlignment = xlCenter
    End With
End Sub

' ----------------------------------------------------------------------------
' WritePriorityBreakdown
' Writes P1-P4 open incident counts with color-coded labels.
'
' Parameters:
'   ws           - Dashboard worksheet
'   dictPriority - Dictionary keyed by priority string, value = count
' ----------------------------------------------------------------------------
Private Sub WritePriorityBreakdown(ws As Worksheet, dictPriority As Object)

    ' Header
    With ws.Cells(DASH_BREAKDOWN_HEADER_ROW, DASH_PRIORITY_COL)
        .Value = "By Priority"
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = CLR_BRIMIS_DARK
    End With

    ' Priority labels and counts
    Dim arrPriorities As Variant
    arrPriorities = Array(PRIORITY_P1, PRIORITY_P2, PRIORITY_P3, PRIORITY_P4)

    ' Priority colors: background and text for each level
    ' P1: CLR_BRIMIS_RED bg, white text
    ' P2: Orange (36095) bg, white text
    ' P3: Gold (55295) bg, dark text
    ' P4: CLR_BRIMIS_GRAY bg, white text
    Dim arrBgColors As Variant, arrTxtColors As Variant
    arrBgColors = Array(CLR_BRIMIS_RED, 36095, 55295, CLR_BRIMIS_GRAY)
    arrTxtColors = Array(CLR_BRIMIS_WHITE, CLR_BRIMIS_WHITE, CLR_BRIMIS_DARK, CLR_BRIMIS_WHITE)

    Dim j As Long
    For j = 0 To 3
        Dim lRow As Long
        lRow = DASH_PRIORITY_START_ROW + j

        ' Label cell with color coding
        With ws.Cells(lRow, DASH_PRIORITY_COL)
            .Value = arrPriorities(j)
            .Font.Name = "Calibri"
            .Font.Size = 10
            .Font.Bold = True
            .Font.Color = arrTxtColors(j)
            .Interior.Color = arrBgColors(j)
            .HorizontalAlignment = xlCenter
        End With

        ' Count cell
        Dim lCount As Long
        If dictPriority.Exists(CStr(arrPriorities(j))) Then
            lCount = dictPriority(CStr(arrPriorities(j)))
        Else
            lCount = 0
        End If

        With ws.Cells(lRow, DASH_PRIORITY_VAL_COL)
            .Value = lCount
            .Font.Name = "Calibri"
            .Font.Size = 12
            .Font.Bold = True
            .Font.Color = CLR_BRIMIS_DARK
            .HorizontalAlignment = xlCenter
        End With
    Next j

End Sub

' ----------------------------------------------------------------------------
' WriteCategoryBreakdown
' Writes category names and open incident counts.
'
' Parameters:
'   ws           - Dashboard worksheet
'   dictCategory - Dictionary keyed by category string, value = count
' ----------------------------------------------------------------------------
Private Sub WriteCategoryBreakdown(ws As Worksheet, dictCategory As Object)

    ' Header
    With ws.Cells(DASH_BREAKDOWN_HEADER_ROW, DASH_CATEGORY_COL)
        .Value = "By Category"
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = CLR_BRIMIS_DARK
    End With

    ' Clear previous category data (up to 6 rows)
    Dim k As Long
    For k = 0 To 5
        ws.Cells(DASH_CATEGORY_START_ROW + k, DASH_CATEGORY_COL).ClearContents
        ws.Cells(DASH_CATEGORY_START_ROW + k, DASH_CATEGORY_VAL_COL).ClearContents
    Next k

    ' Write category entries
    If dictCategory.Count = 0 Then Exit Sub

    Dim vKeys As Variant
    vKeys = dictCategory.Keys

    Dim m As Long
    For m = 0 To dictCategory.Count - 1
        If m > 5 Then Exit For  ' Max 6 categories displayed

        With ws.Cells(DASH_CATEGORY_START_ROW + m, DASH_CATEGORY_COL)
            .Value = vKeys(m)
            .Font.Name = "Calibri"
            .Font.Size = 10
            .Font.Bold = True
            .Font.Color = CLR_BRIMIS_DARK
        End With

        With ws.Cells(DASH_CATEGORY_START_ROW + m, DASH_CATEGORY_VAL_COL)
            .Value = dictCategory(vKeys(m))
            .Font.Name = "Calibri"
            .Font.Size = 12
            .Font.Bold = True
            .Font.Color = CLR_BRIMIS_DARK
            .HorizontalAlignment = xlCenter
        End With
    Next m

End Sub

' ----------------------------------------------------------------------------
' WriteSLAMonitorTable
' Writes the SLA Monitor table on the Dashboard showing open incidents with
' their SLA status and time remaining. Receives pre-collected data from
' RefreshDashboard and the SLA thresholds dictionary.
'
' Parameters:
'   ws              - Dashboard worksheet
'   arrIncidents    - 2D array of open incident data (max DASH_SLA_MAX_ROWS x 7)
'   lCount          - Actual number of incidents in array
'   lTotalOpen      - Total open incidents (for "...and X more" message)
'   dictThresholds  - SLA thresholds dictionary from modSLA.LoadSLAThresholds
' ----------------------------------------------------------------------------
Private Sub WriteSLAMonitorTable(ws As Worksheet, ByRef arrIncidents() As Variant, _
                                  ByVal lCount As Long, ByVal lTotalOpen As Long, _
                                  dictThresholds As Object)

    ' SLA Monitor section header
    With ws.Cells(DASH_SLA_HEADER_ROW, DASH_SLA_COL_ID)
        .Value = "SLA Monitor - Open Incidents"
        .Font.Name = "Calibri"
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = CLR_BRIMIS_WHITE
    End With
    ' Apply header bar background across columns
    Dim c As Long
    For c = DASH_SLA_COL_ID To DASH_SLA_COL_TIMEREM
        ws.Cells(DASH_SLA_HEADER_ROW, c).Interior.Color = CLR_BRIMIS_RED
        ws.Cells(DASH_SLA_HEADER_ROW, c).Font.Color = CLR_BRIMIS_WHITE
    Next c

    ' Column headers row
    Dim arrHeaders As Variant
    arrHeaders = Array("ID", "Title", "Priority", "Status", "Response SLA", "Resolution SLA", "Time Remaining")
    Dim h As Long
    For h = 0 To 6
        With ws.Cells(DASH_SLA_COL_HEADERS_ROW, DASH_SLA_COL_ID + h)
            .Value = arrHeaders(h)
            .Font.Name = "Calibri"
            .Font.Size = 10
            .Font.Bold = True
            .Font.Color = CLR_BRIMIS_WHITE
            .Interior.Color = CLR_BRIMIS_GRAY
        End With
    Next h

    ' Clear previous SLA monitor data rows
    Dim r As Long
    For r = DASH_SLA_DATA_START_ROW To DASH_SLA_DATA_START_ROW + DASH_SLA_MAX_ROWS
        For c = DASH_SLA_COL_ID To DASH_SLA_COL_TIMEREM
            ws.Cells(r, c).ClearContents
            ws.Cells(r, c).Interior.ColorIndex = xlNone
            ws.Cells(r, c).Font.Color = CLR_BRIMIS_DARK
            ws.Cells(r, c).Font.Bold = False
        Next c
    Next r

    ' Write incident data rows
    Dim n As Long
    For n = 1 To lCount
        Dim lDataRow As Long
        lDataRow = DASH_SLA_DATA_START_ROW + n - 1

        ' ID
        ws.Cells(lDataRow, DASH_SLA_COL_ID).Value = arrIncidents(n, 1)
        ws.Cells(lDataRow, DASH_SLA_COL_ID).Font.Name = "Calibri"
        ws.Cells(lDataRow, DASH_SLA_COL_ID).Font.Size = 9

        ' Title (truncated to 40 chars)
        ws.Cells(lDataRow, DASH_SLA_COL_TITLE).Value = Left(CStr(arrIncidents(n, 2)), 40)
        ws.Cells(lDataRow, DASH_SLA_COL_TITLE).Font.Name = "Calibri"
        ws.Cells(lDataRow, DASH_SLA_COL_TITLE).Font.Size = 9

        ' Priority
        ws.Cells(lDataRow, DASH_SLA_COL_PRIORITY).Value = arrIncidents(n, 3)
        ws.Cells(lDataRow, DASH_SLA_COL_PRIORITY).Font.Name = "Calibri"
        ws.Cells(lDataRow, DASH_SLA_COL_PRIORITY).Font.Size = 9

        ' Status
        ws.Cells(lDataRow, DASH_SLA_COL_STATUS).Value = arrIncidents(n, 4)
        ws.Cells(lDataRow, DASH_SLA_COL_STATUS).Font.Name = "Calibri"
        ws.Cells(lDataRow, DASH_SLA_COL_STATUS).Font.Size = 9

        ' Response SLA status
        Dim sRespSLA As String
        sRespSLA = CStr(arrIncidents(n, 5))
        ws.Cells(lDataRow, DASH_SLA_COL_RESPONSE).Value = sRespSLA
        ws.Cells(lDataRow, DASH_SLA_COL_RESPONSE).Font.Name = "Calibri"
        ws.Cells(lDataRow, DASH_SLA_COL_RESPONSE).Font.Size = 9
        ColorSLACell ws.Cells(lDataRow, DASH_SLA_COL_RESPONSE), sRespSLA

        ' Resolution SLA status
        Dim sResSLA As String
        sResSLA = CStr(arrIncidents(n, 6))
        ws.Cells(lDataRow, DASH_SLA_COL_RESOLUTION).Value = sResSLA
        ws.Cells(lDataRow, DASH_SLA_COL_RESOLUTION).Font.Name = "Calibri"
        ws.Cells(lDataRow, DASH_SLA_COL_RESOLUTION).Font.Size = 9
        ColorSLACell ws.Cells(lDataRow, DASH_SLA_COL_RESOLUTION), sResSLA

        ' Time remaining: calculate resolution due date from ReportedDate + threshold
        Dim sTimeRem As String
        sTimeRem = "N/A"
        Dim sPri As String
        sPri = CStr(arrIncidents(n, 3))
        If dictThresholds.Exists(sPri) Then
            Dim vThresh As Variant
            vThresh = dictThresholds(sPri)
            Dim dblResHrs As Double
            dblResHrs = CDbl(vThresh(1))  ' Resolution threshold hours
            Dim vRepDate As Variant
            vRepDate = arrIncidents(n, 7)
            If IsDate(vRepDate) And Not IsEmpty(vRepDate) Then
                Dim dResDue As Date
                dResDue = CDate(vRepDate) + (dblResHrs / 24)
                sTimeRem = modSLA.FormatTimeRemaining(dResDue)
            End If
        End If
        ws.Cells(lDataRow, DASH_SLA_COL_TIMEREM).Value = sTimeRem
        ws.Cells(lDataRow, DASH_SLA_COL_TIMEREM).Font.Name = "Calibri"
        ws.Cells(lDataRow, DASH_SLA_COL_TIMEREM).Font.Size = 9
        ws.Cells(lDataRow, DASH_SLA_COL_TIMEREM).Font.Bold = True
    Next n

    ' "...and X more" message if there are more open incidents than displayed
    If lTotalOpen > lCount And lCount > 0 Then
        Dim lMoreRow As Long
        lMoreRow = DASH_SLA_DATA_START_ROW + lCount
        With ws.Cells(lMoreRow, DASH_SLA_COL_ID)
            .Value = "...and " & (lTotalOpen - lCount) & " more"
            .Font.Name = "Calibri"
            .Font.Size = 9
            .Font.Italic = True
            .Font.Color = CLR_BRIMIS_GRAY
        End With
    End If

End Sub

' ----------------------------------------------------------------------------
' ColorSLACell
' Applies background and text color to a cell based on SLA status value.
'
' Parameters:
'   rng     - The cell to color
'   sSLA    - SLA status string
' ----------------------------------------------------------------------------
Private Sub ColorSLACell(rng As Range, ByVal sSLA As String)
    Select Case sSLA
        Case SLA_STATUS_ON_TRACK, SLA_STATUS_MET
            rng.Interior.Color = CLR_SLA_GREEN
            rng.Font.Color = CLR_SLA_GREEN_TEXT
        Case SLA_STATUS_AT_RISK
            rng.Interior.Color = CLR_SLA_AMBER
            rng.Font.Color = CLR_SLA_AMBER_TEXT
        Case SLA_STATUS_OVERDUE, SLA_STATUS_BREACHED
            rng.Interior.Color = CLR_BRIMIS_RED
            rng.Font.Color = CLR_BRIMIS_WHITE
            rng.Font.Bold = True
        Case Else
            ' No formatting for empty/unknown
            rng.Interior.ColorIndex = xlNone
            rng.Font.Color = CLR_BRIMIS_DARK
    End Select
End Sub
