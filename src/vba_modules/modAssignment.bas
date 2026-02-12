Attribute VB_Name = "modAssignment"
Option Explicit

' ============================================================================
' Module:  modAssignment
' Purpose: Entry points for the Assignment and Status Lifecycle forms.
'          Provides public macros assigned to Dashboard buttons via OnAction.
'          Contains the state machine (valid transitions), timestamp mapping,
'          and Assignment Tracker refresh logic.
' Dependencies: frmAssignment (UserForm)
'               frmStatusUpdate (UserForm)
'               modConstants (SHT_*, TBL_*, COL_*, STATUS_*)
'               modUtilities (UnprotectSheet, ProtectSheet)
'               modFormatting (ApplyBrandingToSheet, ApplyTableHeaderBranding)
'               modErrorHandler (HandleError)
' ============================================================================

' ----------------------------------------------------------------------------
' ShowAssignmentForm
' Launches the incident assignment form as a modal dialog.
' Called by the Dashboard "Assign Incident" button shape's OnAction property.
' ----------------------------------------------------------------------------
Public Sub ShowAssignmentForm()
    On Error GoTo ErrHandler
    frmAssignment.Show vbModal
    Exit Sub
ErrHandler:
    modErrorHandler.HandleError "modAssignment", "ShowAssignmentForm", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' GetValidNextStatuses
' Returns a comma-delimited string of valid next statuses for the given
' current status, based on the incident lifecycle state machine.
'
' State machine:
'   Open        -> Cancelled, Duplicate  (Assigned is via assignment form only)
'   Assigned    -> In Progress, Cancelled
'   In Progress -> Resolved, Cancelled
'   Resolved    -> Closed, In Progress (reopen)
'   Closed, Cancelled, Duplicate -> (terminal, no transitions)
'
' Parameters:
'   sCurrentStatus - Current status value (use STATUS_* constants)
'
' Returns:
'   Comma-delimited string of valid next statuses, or "" if terminal
' ----------------------------------------------------------------------------
Public Function GetValidNextStatuses(ByVal sCurrentStatus As String) As String
    Select Case sCurrentStatus
        Case STATUS_OPEN
            GetValidNextStatuses = STATUS_CANCELLED & "," & STATUS_DUPLICATE
        Case STATUS_ASSIGNED
            GetValidNextStatuses = STATUS_IN_PROGRESS & "," & STATUS_CANCELLED
        Case STATUS_IN_PROGRESS
            GetValidNextStatuses = STATUS_RESOLVED & "," & STATUS_CANCELLED
        Case STATUS_RESOLVED
            GetValidNextStatuses = STATUS_CLOSED & "," & STATUS_IN_PROGRESS
        Case Else
            GetValidNextStatuses = ""
    End Select
End Function

' ----------------------------------------------------------------------------
' GetTimestampColumn
' Returns the COL_* constant name for the timestamp column that corresponds
' to the given new status.
'
' Mapping:
'   Assigned    -> COL_ASSIGNED_DATE
'   In Progress -> COL_RESPONSE_DATE (SLA response = work started)
'   Resolved    -> COL_RESOLUTION_DATE
'   Closed      -> COL_CLOSED_DATE
'   Cancelled   -> COL_CLOSED_DATE (terminal states use ClosedDate)
'   Duplicate   -> COL_CLOSED_DATE
'
' Parameters:
'   sNewStatus - The status being transitioned TO
'
' Returns:
'   Column name string (COL_* value), or "" if no timestamp applies
' ----------------------------------------------------------------------------
Public Function GetTimestampColumn(ByVal sNewStatus As String) As String
    Select Case sNewStatus
        Case STATUS_ASSIGNED:    GetTimestampColumn = COL_ASSIGNED_DATE
        Case STATUS_IN_PROGRESS: GetTimestampColumn = COL_RESPONSE_DATE
        Case STATUS_RESOLVED:    GetTimestampColumn = COL_RESOLUTION_DATE
        Case STATUS_CLOSED:      GetTimestampColumn = COL_CLOSED_DATE
        Case STATUS_CANCELLED:   GetTimestampColumn = COL_CLOSED_DATE
        Case STATUS_DUPLICATE:   GetTimestampColumn = COL_CLOSED_DATE
        Case Else:               GetTimestampColumn = ""
    End Select
End Function

' ----------------------------------------------------------------------------
' RefreshAssignmentTracker
' Clears and repopulates tblAssignmentTracker with all active assigned
' incidents (Assigned, In Progress, Resolved) from tblIncidents, sorted
' by team then assignee.
'
' Called after: assignment, status update, and from Refresh button on tracker sheet.
' ----------------------------------------------------------------------------
Public Sub RefreshAssignmentTracker()
    On Error GoTo ErrHandler

    Dim wsDest As Worksheet
    Set wsDest = ThisWorkbook.Sheets(SHT_ASSIGNMENT_TRACKER)

    Dim tblTracker As ListObject
    Set tblTracker = wsDest.ListObjects(TBL_ASSIGNMENT_TRACKER)

    Dim tblIncidents As ListObject
    Set tblIncidents = ThisWorkbook.Sheets(SHT_INCIDENT_LOG).ListObjects(TBL_INCIDENTS)

    ' Unprotect for writing
    modUtilities.UnprotectSheet wsDest

    ' Clear existing tracker data (keep header)
    If Not tblTracker.DataBodyRange Is Nothing Then
        tblTracker.DataBodyRange.Delete
    End If

    If tblIncidents.ListRows.Count = 0 Then
        modUtilities.ProtectSheet wsDest
        Exit Sub
    End If

    ' Column indices from tblIncidents
    Dim cTeam As Long, cPerson As Long, cID As Long, cTitle As Long
    Dim cPriority As Long, cStatus As Long, cAssignDate As Long
    cTeam = tblIncidents.ListColumns(COL_ASSIGNED_TEAM).Index
    cPerson = tblIncidents.ListColumns(COL_ASSIGNED_TO).Index
    cID = tblIncidents.ListColumns(COL_INCIDENT_ID).Index
    cTitle = tblIncidents.ListColumns(COL_TITLE).Index
    cPriority = tblIncidents.ListColumns(COL_PRIORITY).Index
    cStatus = tblIncidents.ListColumns(COL_STATUS).Index
    cAssignDate = tblIncidents.ListColumns(COL_ASSIGNED_DATE).Index

    ' Collect active assigned incidents (not Open, not terminal)
    Dim i As Long
    Dim sStatus As String
    For i = 1 To tblIncidents.ListRows.Count
        sStatus = CStr(tblIncidents.ListRows(i).Range(1, cStatus).Value)
        ' Include: Assigned, In Progress, Resolved
        ' Exclude: Open, Closed, Cancelled, Duplicate
        If sStatus = STATUS_ASSIGNED Or sStatus = STATUS_IN_PROGRESS Or sStatus = STATUS_RESOLVED Then
            Dim newRow As ListRow
            Set newRow = tblTracker.ListRows.Add
            newRow.Range(1, 1).Value = tblIncidents.ListRows(i).Range(1, cTeam).Value    ' AssignedTeam
            newRow.Range(1, 2).Value = tblIncidents.ListRows(i).Range(1, cPerson).Value  ' AssignedTo
            newRow.Range(1, 3).Value = tblIncidents.ListRows(i).Range(1, cID).Value      ' IncidentID
            newRow.Range(1, 4).Value = tblIncidents.ListRows(i).Range(1, cTitle).Value   ' Title
            newRow.Range(1, 5).Value = tblIncidents.ListRows(i).Range(1, cPriority).Value ' Priority
            newRow.Range(1, 6).Value = sStatus                                            ' Status
            newRow.Range(1, 7).Value = tblIncidents.ListRows(i).Range(1, cAssignDate).Value ' AssignedDate
            ' DaysOpen = Now - AssignedDate (calculated during refresh)
            Dim dAssigned As Variant
            dAssigned = tblIncidents.ListRows(i).Range(1, cAssignDate).Value
            If IsDate(dAssigned) Then
                newRow.Range(1, 8).Value = Int(Now - CDate(dAssigned))
            Else
                newRow.Range(1, 8).Value = ""
            End If
        End If
    Next i

    ' Re-protect
    modUtilities.ProtectSheet wsDest

    Exit Sub
ErrHandler:
    On Error Resume Next
    modUtilities.ProtectSheet wsDest
    On Error GoTo 0
    modErrorHandler.HandleError "modAssignment", "RefreshAssignmentTracker", _
                                 Err.Number, Err.Description
End Sub
