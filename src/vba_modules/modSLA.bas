Attribute VB_Name = "modSLA"
Option Explicit

' ============================================================================
' Module:  modSLA
' Purpose: SLA calculation engine. Reads thresholds from tblSLAThresholds,
'          computes due dates based on ReportedDate + threshold hours, and
'          writes SLAResponseStatus / SLAResolutionStatus to each incident row.
'          Provides time-remaining formatting for the Dashboard SLA Monitor.
' Dependencies: modConstants (TBL_SLA_THRESHOLDS, SHT_SETTINGS, SHT_INCIDENT_LOG,
'                TBL_INCIDENTS, COL_*, SLA_STATUS_*, SLA_AMBER_THRESHOLD_PCT,
'                STATUS_CLOSED, STATUS_CANCELLED, STATUS_DUPLICATE)
'               modUtilities (UnprotectSheet, ProtectSheet)
'               modErrorHandler (HandleError)
' ============================================================================

' ----------------------------------------------------------------------------
' LoadSLAThresholds
' Reads tblSLAThresholds from the Settings sheet and returns a Dictionary
' keyed by priority string (e.g., "P1"), where each value is an Array of
' (ResponseTime_Hrs, ResolutionTime_Hrs).
'
' PUBLIC so that modDashboard.WriteSLAMonitorTable can call it to look up
' threshold hours when computing time remaining for open incidents.
'
' Returns:
'   Scripting.Dictionary keyed by priority string
' ----------------------------------------------------------------------------
Public Function LoadSLAThresholds() As Object
    On Error GoTo ErrHandler

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets(SHT_SETTINGS).ListObjects(TBL_SLA_THRESHOLDS)

    ' Guard against empty table
    If tbl.ListRows.Count = 0 Then
        Set LoadSLAThresholds = dict
        Exit Function
    End If

    ' Column indices
    Dim cPriority As Long, cResponse As Long, cResolution As Long
    cPriority = tbl.ListColumns("Priority").Index
    cResponse = tbl.ListColumns("ResponseTime_Hrs").Index
    cResolution = tbl.ListColumns("ResolutionTime_Hrs").Index

    ' Load each row into dictionary
    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        Dim sPri As String
        sPri = CStr(tbl.ListRows(i).Range(1, cPriority).Value)
        If Len(sPri) > 0 And Not dict.Exists(sPri) Then
            dict.Add sPri, Array( _
                CDbl(tbl.ListRows(i).Range(1, cResponse).Value), _
                CDbl(tbl.ListRows(i).Range(1, cResolution).Value))
        End If
    Next i

    Set LoadSLAThresholds = dict
    Exit Function

ErrHandler:
    modErrorHandler.HandleError "modSLA", "LoadSLAThresholds", _
                                 Err.Number, Err.Description
    Set LoadSLAThresholds = CreateObject("Scripting.Dictionary")
End Function

' ----------------------------------------------------------------------------
' GetSLAStatus
' Determines the SLA status for a single metric (response or resolution).
'
' Parameters:
'   dReportedDate   - When the incident was reported
'   dActualDate     - When response/resolution actually occurred (Empty if not yet)
'   dblThresholdHrs - SLA threshold in hours from tblSLAThresholds
'
' Returns:
'   String - SLA_STATUS_MET, SLA_STATUS_BREACHED, SLA_STATUS_ON_TRACK,
'            SLA_STATUS_AT_RISK, or SLA_STATUS_OVERDUE
' ----------------------------------------------------------------------------
Private Function GetSLAStatus(ByVal dReportedDate As Date, _
                               ByVal dActualDate As Variant, _
                               ByVal dblThresholdHrs As Double) As String

    Dim dDue As Date
    dDue = dReportedDate + (dblThresholdHrs / 24)

    ' If the actual date is recorded, it is a final determination
    If IsDate(dActualDate) And Not IsEmpty(dActualDate) Then
        If CDate(dActualDate) <= dDue Then
            GetSLAStatus = SLA_STATUS_MET
        Else
            GetSLAStatus = SLA_STATUS_BREACHED
        End If
        Exit Function
    End If

    ' Still open -- check against Now
    If Now > dDue Then
        GetSLAStatus = SLA_STATUS_OVERDUE
    ElseIf Now > dReportedDate + (dblThresholdHrs * SLA_AMBER_THRESHOLD_PCT / 24) Then
        GetSLAStatus = SLA_STATUS_AT_RISK
    Else
        GetSLAStatus = SLA_STATUS_ON_TRACK
    End If
End Function

' ----------------------------------------------------------------------------
' FormatTimeRemaining
' Formats the time difference between Now and a due date as a human-readable
' string for the Dashboard SLA Monitor.
'
' Parameters:
'   dDue - The SLA due date
'
' Returns:
'   String - e.g., "2d 5h left", "OVERDUE 3d 1h", "5h 30m left", "OVERDUE 2h 15m"
' ----------------------------------------------------------------------------
Public Function FormatTimeRemaining(ByVal dDue As Date) As String
    On Error GoTo ErrHandler

    Dim dblMinutesLeft As Double
    dblMinutesLeft = DateDiff("n", Now, dDue)

    Dim dblHoursLeft As Double
    dblHoursLeft = dblMinutesLeft / 60

    If dblHoursLeft >= 0 Then
        ' Time remaining
        If dblHoursLeft >= 24 Then
            Dim lDays As Long, lHours As Long
            lDays = Int(dblHoursLeft / 24)
            lHours = Int(dblHoursLeft) Mod 24
            FormatTimeRemaining = lDays & "d " & lHours & "h left"
        Else
            Dim lHrs As Long, lMins As Long
            lHrs = Int(dblHoursLeft)
            lMins = Int(Abs(dblMinutesLeft)) Mod 60
            FormatTimeRemaining = lHrs & "h " & lMins & "m left"
        End If
    Else
        ' Overdue
        Dim dblOverdueHrs As Double
        dblOverdueHrs = Abs(dblHoursLeft)
        If dblOverdueHrs >= 24 Then
            Dim lODays As Long, lOHours As Long
            lODays = Int(dblOverdueHrs / 24)
            lOHours = Int(dblOverdueHrs) Mod 24
            FormatTimeRemaining = "OVERDUE " & lODays & "d " & lOHours & "h"
        Else
            Dim lOHrs As Long, lOMins As Long
            lOHrs = Int(dblOverdueHrs)
            lOMins = Int(Abs(dblMinutesLeft)) Mod 60
            FormatTimeRemaining = "OVERDUE " & lOHrs & "h " & lOMins & "m"
        End If
    End If

    Exit Function

ErrHandler:
    modErrorHandler.HandleError "modSLA", "FormatTimeRemaining", _
                                 Err.Number, Err.Description
    FormatTimeRemaining = "N/A"
End Function

' ----------------------------------------------------------------------------
' RecalculateAllSLAStatuses
' Loops all rows in tblIncidents and writes SLAResponseStatus and
' SLAResolutionStatus based on SLA thresholds from tblSLAThresholds.
'
' Terminal statuses (Closed, Cancelled, Duplicate) are skipped -- they keep
' whatever SLA status they already had when they transitioned out of active.
'
' Uses a single UnprotectSheet/ProtectSheet cycle for the entire operation.
' ScreenUpdating is disabled during processing for performance.
' ----------------------------------------------------------------------------
Public Sub RecalculateAllSLAStatuses()
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False

    ' Load SLA thresholds
    Dim dictThresholds As Object
    Set dictThresholds = LoadSLAThresholds()

    ' Guard: if no thresholds defined, nothing to calculate
    If dictThresholds.Count = 0 Then
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Get tblIncidents
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHT_INCIDENT_LOG)

    Dim tbl As ListObject
    Set tbl = ws.ListObjects(TBL_INCIDENTS)

    ' Guard: if table has no data rows, exit
    If tbl.DataBodyRange Is Nothing Then
        Application.ScreenUpdating = True
        Exit Sub
    End If

    If tbl.ListRows.Count = 0 Then
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Column indices
    Dim cStatus As Long, cPriority As Long, cReportedDate As Long
    Dim cResponseDate As Long, cResolutionDate As Long
    Dim cSLAResponse As Long, cSLAResolution As Long

    cStatus = tbl.ListColumns(COL_STATUS).Index
    cPriority = tbl.ListColumns(COL_PRIORITY).Index
    cReportedDate = tbl.ListColumns(COL_REPORTED_DATE).Index
    cResponseDate = tbl.ListColumns(COL_RESPONSE_DATE).Index
    cResolutionDate = tbl.ListColumns(COL_RESOLUTION_DATE).Index
    cSLAResponse = tbl.ListColumns(COL_SLA_RESPONSE).Index
    cSLAResolution = tbl.ListColumns(COL_SLA_RESOLUTION).Index

    ' Single unprotect cycle
    modUtilities.UnprotectSheet ws

    ' Loop all rows
    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        Dim sStatus As String
        sStatus = CStr(tbl.ListRows(i).Range(1, cStatus).Value)

        ' Skip terminal statuses -- they keep their existing SLA status
        If sStatus = STATUS_CLOSED Or sStatus = STATUS_CANCELLED Or sStatus = STATUS_DUPLICATE Then
            GoTo NextRow
        End If

        ' Read priority and look up thresholds
        Dim sPriority As String
        sPriority = CStr(tbl.ListRows(i).Range(1, cPriority).Value)

        If Not dictThresholds.Exists(sPriority) Then
            GoTo NextRow
        End If

        ' Read reported date -- guard against invalid/missing values
        Dim vReportedDate As Variant
        vReportedDate = tbl.ListRows(i).Range(1, cReportedDate).Value
        If Not IsDate(vReportedDate) Or IsEmpty(vReportedDate) Then
            GoTo NextRow
        End If

        Dim dReportedDate As Date
        dReportedDate = CDate(vReportedDate)

        ' Get threshold values: Array(ResponseHrs, ResolutionHrs)
        Dim vThreshold As Variant
        vThreshold = dictThresholds(sPriority)
        Dim dblResponseHrs As Double, dblResolutionHrs As Double
        dblResponseHrs = CDbl(vThreshold(0))
        dblResolutionHrs = CDbl(vThreshold(1))

        ' Response SLA
        Dim vResponseDate As Variant
        vResponseDate = tbl.ListRows(i).Range(1, cResponseDate).Value
        tbl.ListRows(i).Range(1, cSLAResponse).Value = _
            GetSLAStatus(dReportedDate, vResponseDate, dblResponseHrs)

        ' Resolution SLA
        Dim vResolutionDate As Variant
        vResolutionDate = tbl.ListRows(i).Range(1, cResolutionDate).Value
        tbl.ListRows(i).Range(1, cSLAResolution).Value = _
            GetSLAStatus(dReportedDate, vResolutionDate, dblResolutionHrs)

NextRow:
    Next i

    ' Re-protect
    modUtilities.ProtectSheet ws

    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    ' Re-protect sheet before surfacing the error
    On Error Resume Next
    Dim wsErr As Worksheet
    Set wsErr = ThisWorkbook.Sheets(SHT_INCIDENT_LOG)
    modUtilities.ProtectSheet wsErr
    Application.ScreenUpdating = True
    On Error GoTo 0

    modErrorHandler.HandleError "modSLA", "RecalculateAllSLAStatuses", _
                                 Err.Number, Err.Description
End Sub
