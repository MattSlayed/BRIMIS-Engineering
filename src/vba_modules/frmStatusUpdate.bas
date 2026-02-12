Option Explicit

' ============================================================================
' Form:    frmStatusUpdate
' Purpose: Single-page form for updating incident status through the enforced
'          lifecycle state machine. Shows non-terminal incidents, presents only
'          valid next statuses, conditionally shows reason field for Cancelled
'          and Duplicate transitions.
'
' Control Names: lstIncidents, lblCurrentStatus, lblCurrentAssignee,
'   lblCurrentDate, cboNewStatus, lblReason, txtReason,
'   btnUpdateStatus, btnCancel
'
' Dependencies: modConstants (SHT_*, TBL_*, COL_*, STATUS_*, APP_TITLE)
'               modDataAccess (UpdateIncidentFields)
'               modAssignment (GetValidNextStatuses, GetTimestampColumn,
'                              RefreshAssignmentTracker)
'               modErrorHandler (HandleError)
' ============================================================================

' ----------------------------------------------------------------------------
' UserForm_Initialize
' Populates lstIncidents with all non-terminal incidents. Hides reason field.
' ----------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    On Error GoTo ErrHandler

    ' --- Configure ListBox columns ---
    Me.lstIncidents.ColumnCount = 3
    Me.lstIncidents.ColumnWidths = "80;220;80"
    Me.lstIncidents.BoundColumn = 1

    ' --- Populate incident list with non-terminal incidents ---
    PopulateIncidentList

    ' --- Hide reason field by default ---
    Me.lblReason.Visible = False
    Me.txtReason.Visible = False
    Me.txtReason.Value = ""

    ' --- Clear info labels ---
    Me.lblCurrentStatus.Caption = ""
    Me.lblCurrentAssignee.Caption = ""
    Me.lblCurrentDate.Caption = ""

    ' --- Clear status combo ---
    Me.cboNewStatus.Clear

    Exit Sub
ErrHandler:
    modErrorHandler.HandleError "frmStatusUpdate", "UserForm_Initialize", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' PopulateIncidentList
' Fills lstIncidents with non-terminal incidents (Open, Assigned, In Progress,
' Resolved). Terminal statuses (Closed, Cancelled, Duplicate) are excluded.
' ----------------------------------------------------------------------------
Private Sub PopulateIncidentList()
    On Error GoTo ErrHandler

    Me.lstIncidents.Clear

    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets(SHT_INCIDENT_LOG).ListObjects(TBL_INCIDENTS)

    If tbl.ListRows.Count = 0 Then Exit Sub

    Dim idCol As Long, titleCol As Long, statusCol As Long
    idCol = tbl.ListColumns(COL_INCIDENT_ID).Index
    titleCol = tbl.ListColumns(COL_TITLE).Index
    statusCol = tbl.ListColumns(COL_STATUS).Index

    Dim i As Long
    Dim sStatus As String
    For i = 1 To tbl.ListRows.Count
        sStatus = CStr(tbl.ListRows(i).Range(1, statusCol).Value)
        ' Include non-terminal statuses only
        If sStatus = STATUS_OPEN Or sStatus = STATUS_ASSIGNED Or _
           sStatus = STATUS_IN_PROGRESS Or sStatus = STATUS_RESOLVED Then
            Me.lstIncidents.AddItem CStr(tbl.ListRows(i).Range(1, idCol).Value)
            Me.lstIncidents.List(Me.lstIncidents.ListCount - 1, 1) = _
                CStr(tbl.ListRows(i).Range(1, titleCol).Value)
            Me.lstIncidents.List(Me.lstIncidents.ListCount - 1, 2) = sStatus
        End If
    Next i

    Exit Sub
ErrHandler:
    modErrorHandler.HandleError "frmStatusUpdate", "PopulateIncidentList", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' lstIncidents_Click
' When an incident is selected, display its current status info and populate
' cboNewStatus with valid next statuses from the state machine.
' ----------------------------------------------------------------------------
Private Sub lstIncidents_Click()
    On Error GoTo ErrHandler

    If Me.lstIncidents.ListIndex = -1 Then Exit Sub

    Dim sID As String
    sID = Me.lstIncidents.Value

    ' Read incident data for display
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets(SHT_INCIDENT_LOG).ListObjects(TBL_INCIDENTS)

    Dim idCol As Long, statusCol As Long, teamCol As Long, personCol As Long, assignDateCol As Long
    idCol = tbl.ListColumns(COL_INCIDENT_ID).Index
    statusCol = tbl.ListColumns(COL_STATUS).Index
    teamCol = tbl.ListColumns(COL_ASSIGNED_TEAM).Index
    personCol = tbl.ListColumns(COL_ASSIGNED_TO).Index
    assignDateCol = tbl.ListColumns(COL_ASSIGNED_DATE).Index

    Dim i As Long
    Dim sCurrentStatus As String
    For i = 1 To tbl.ListRows.Count
        If CStr(tbl.ListRows(i).Range(1, idCol).Value) = sID Then
            sCurrentStatus = CStr(tbl.ListRows(i).Range(1, statusCol).Value)
            Me.lblCurrentStatus.Caption = "Current Status: " & sCurrentStatus

            Dim sTeam As String, sPerson As String
            sTeam = CStr(tbl.ListRows(i).Range(1, teamCol).Value)
            sPerson = CStr(tbl.ListRows(i).Range(1, personCol).Value)
            If Len(sPerson) > 0 Then
                Me.lblCurrentAssignee.Caption = "Assigned To: " & sPerson & " (" & sTeam & ")"
            Else
                Me.lblCurrentAssignee.Caption = "Assigned To: (not assigned)"
            End If

            Dim dAssign As Variant
            dAssign = tbl.ListRows(i).Range(1, assignDateCol).Value
            If IsDate(dAssign) Then
                Me.lblCurrentDate.Caption = "Assigned Date: " & Format(CDate(dAssign), "yyyy-mm-dd hh:nn")
            Else
                Me.lblCurrentDate.Caption = "Assigned Date: (n/a)"
            End If

            Exit For
        End If
    Next i

    ' Populate valid next statuses
    Me.cboNewStatus.Clear
    Me.cboNewStatus.Value = ""
    Dim sValidStatuses As String
    sValidStatuses = modAssignment.GetValidNextStatuses(sCurrentStatus)

    If Len(sValidStatuses) > 0 Then
        Dim vStatuses As Variant
        vStatuses = Split(sValidStatuses, ",")
        Dim j As Long
        For j = LBound(vStatuses) To UBound(vStatuses)
            Me.cboNewStatus.AddItem Trim(CStr(vStatuses(j)))
        Next j
    End If

    ' Reset reason field
    Me.lblReason.Visible = False
    Me.txtReason.Visible = False
    Me.txtReason.Value = ""

    Exit Sub
ErrHandler:
    modErrorHandler.HandleError "frmStatusUpdate", "lstIncidents_Click", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' cboNewStatus_Change
' Shows or hides the reason TextBox based on whether the selected new status
' is Cancelled or Duplicate (which require a mandatory reason).
' ----------------------------------------------------------------------------
Private Sub cboNewStatus_Change()
    On Error GoTo ErrHandler

    Dim bShowReason As Boolean
    bShowReason = (Me.cboNewStatus.Value = STATUS_CANCELLED Or _
                   Me.cboNewStatus.Value = STATUS_DUPLICATE)
    Me.txtReason.Visible = bShowReason
    Me.lblReason.Visible = bShowReason

    If Not bShowReason Then
        Me.txtReason.Value = ""
    End If

    Exit Sub
ErrHandler:
    modErrorHandler.HandleError "frmStatusUpdate", "cboNewStatus_Change", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' btnUpdateStatus_Click
' Validates selections, writes status + timestamp (+ reason if applicable)
' atomically via UpdateIncidentFields, refreshes tracker, and closes form.
' ----------------------------------------------------------------------------
Private Sub btnUpdateStatus_Click()
    On Error GoTo ErrHandler

    ' Validate selections
    If Me.lstIncidents.ListIndex = -1 Then
        MsgBox "Please select an incident.", vbExclamation, APP_TITLE
        Exit Sub
    End If
    If Len(Me.cboNewStatus.Value) = 0 Then
        MsgBox "Please select a new status.", vbExclamation, APP_TITLE
        Me.cboNewStatus.SetFocus
        Exit Sub
    End If

    ' Check if reason is required
    Dim sNewStatus As String
    sNewStatus = Me.cboNewStatus.Value
    If (sNewStatus = STATUS_CANCELLED Or sNewStatus = STATUS_DUPLICATE) Then
        If Len(Trim(Me.txtReason.Value)) = 0 Then
            MsgBox "Please provide a reason for " & sNewStatus & ".", _
                   vbExclamation, APP_TITLE
            Me.txtReason.SetFocus
            Exit Sub
        End If
    End If

    Dim sIncidentID As String
    sIncidentID = Me.lstIncidents.Value

    ' Determine timestamp column
    Dim sTimestampCol As String
    sTimestampCol = modAssignment.GetTimestampColumn(sNewStatus)

    ' Build field arrays
    Dim vNames As Variant
    Dim vValues As Variant

    If (sNewStatus = STATUS_CANCELLED Or sNewStatus = STATUS_DUPLICATE) Then
        ' Include cancellation/duplicate reason in ResolutionNotes column
        vNames = Array(COL_STATUS, sTimestampCol, COL_RESOLUTION_NOTES)
        vValues = Array(sNewStatus, Now, Trim(Me.txtReason.Value))
    Else
        vNames = Array(COL_STATUS, sTimestampCol)
        vValues = Array(sNewStatus, Now)
    End If

    Dim bSuccess As Boolean
    bSuccess = modDataAccess.UpdateIncidentFields(sIncidentID, vNames, vValues)

    If bSuccess Then
        MsgBox "Incident " & sIncidentID & " status updated to " & sNewStatus & ".", _
               vbInformation, APP_TITLE
        modAssignment.RefreshAssignmentTracker
        Unload Me
    Else
        MsgBox "Failed to update status. Please try again.", _
               vbExclamation, APP_TITLE
    End If

    Exit Sub
ErrHandler:
    modErrorHandler.HandleError "frmStatusUpdate", "btnUpdateStatus_Click", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' btnCancel_Click
' Closes the form without making changes.
' ----------------------------------------------------------------------------
Private Sub btnCancel_Click()
    On Error GoTo ErrHandler
    Unload Me
    Exit Sub
ErrHandler:
    modErrorHandler.HandleError "frmStatusUpdate", "btnCancel_Click", _
                                 Err.Number, Err.Description
End Sub
