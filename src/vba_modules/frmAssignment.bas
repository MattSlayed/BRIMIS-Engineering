Option Explicit

' ============================================================================
' Form:    frmAssignment
' Purpose: Single-page form for supervisors to assign an open incident to a
'          team and specific individual. Shows only Open-status incidents.
'          On assignment: writes AssignedTeam, AssignedTo, AssignedDate,
'          AssignedBy, Status=Assigned atomically via UpdateIncidentFields.
'
' Control Names: lstIncidents, lblInfoCategory, lblInfoPriority,
'   lblInfoReportedBy, lblInfoReportedDate, cboTeam, cboPerson,
'   btnAssign, btnCancel
'
' Dependencies: modConstants (SHT_*, TBL_*, COL_*, STATUS_*, APP_TITLE)
'               modDataAccess (UpdateIncidentFields)
'               modAssignment (RefreshAssignmentTracker)
'               modErrorHandler (HandleError)
' ============================================================================

' ----------------------------------------------------------------------------
' UserForm_Initialize
' Populates lstIncidents with Open-status incidents and cboTeam with active teams.
' ----------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    On Error GoTo ErrHandler

    ' --- Configure ListBox columns ---
    Me.lstIncidents.ColumnCount = 3
    Me.lstIncidents.ColumnWidths = "80;220;50"
    Me.lstIncidents.BoundColumn = 1

    ' --- Populate incident list with Open incidents only ---
    PopulateIncidentList

    ' --- Populate team dropdown from tblTeams ---
    Dim tblTeams As ListObject
    Set tblTeams = ThisWorkbook.Sheets(SHT_SETTINGS).ListObjects(TBL_TEAMS)

    Dim nameCol As Long, activeCol As Long
    nameCol = tblTeams.ListColumns("TeamName").Index
    activeCol = tblTeams.ListColumns("Active").Index

    Dim i As Long
    For i = 1 To tblTeams.ListRows.Count
        If CStr(tblTeams.ListRows(i).Range(1, activeCol).Value) = "Yes" Then
            Me.cboTeam.AddItem CStr(tblTeams.ListRows(i).Range(1, nameCol).Value)
        End If
    Next i

    ' --- Clear info labels ---
    Me.lblInfoCategory.Caption = ""
    Me.lblInfoPriority.Caption = ""
    Me.lblInfoReportedBy.Caption = ""
    Me.lblInfoReportedDate.Caption = ""

    Exit Sub
ErrHandler:
    modErrorHandler.HandleError "frmAssignment", "UserForm_Initialize", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' PopulateIncidentList
' Fills lstIncidents with Open-status incidents (ID, Title, Priority).
' ----------------------------------------------------------------------------
Private Sub PopulateIncidentList()
    On Error GoTo ErrHandler

    Me.lstIncidents.Clear

    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets(SHT_INCIDENT_LOG).ListObjects(TBL_INCIDENTS)

    If tbl.ListRows.Count = 0 Then Exit Sub

    Dim idCol As Long, titleCol As Long, statusCol As Long, priorityCol As Long
    idCol = tbl.ListColumns(COL_INCIDENT_ID).Index
    titleCol = tbl.ListColumns(COL_TITLE).Index
    statusCol = tbl.ListColumns(COL_STATUS).Index
    priorityCol = tbl.ListColumns(COL_PRIORITY).Index

    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        If CStr(tbl.ListRows(i).Range(1, statusCol).Value) = STATUS_OPEN Then
            Me.lstIncidents.AddItem CStr(tbl.ListRows(i).Range(1, idCol).Value)
            Me.lstIncidents.List(Me.lstIncidents.ListCount - 1, 1) = _
                CStr(tbl.ListRows(i).Range(1, titleCol).Value)
            Me.lstIncidents.List(Me.lstIncidents.ListCount - 1, 2) = _
                CStr(tbl.ListRows(i).Range(1, priorityCol).Value)
        End If
    Next i

    Exit Sub
ErrHandler:
    modErrorHandler.HandleError "frmAssignment", "PopulateIncidentList", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' lstIncidents_Click
' When an incident is selected, display its details in the info labels.
' ----------------------------------------------------------------------------
Private Sub lstIncidents_Click()
    On Error GoTo ErrHandler

    If Me.lstIncidents.ListIndex = -1 Then Exit Sub

    Dim sID As String
    sID = Me.lstIncidents.Value

    ' Read incident data
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets(SHT_INCIDENT_LOG).ListObjects(TBL_INCIDENTS)

    Dim idCol As Long, catCol As Long, priCol As Long, repByCol As Long, repDateCol As Long
    idCol = tbl.ListColumns(COL_INCIDENT_ID).Index
    catCol = tbl.ListColumns(COL_CATEGORY).Index
    priCol = tbl.ListColumns(COL_PRIORITY).Index
    repByCol = tbl.ListColumns(COL_REPORTED_BY).Index
    repDateCol = tbl.ListColumns(COL_REPORTED_DATE).Index

    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        If CStr(tbl.ListRows(i).Range(1, idCol).Value) = sID Then
            Me.lblInfoCategory.Caption = "Category: " & CStr(tbl.ListRows(i).Range(1, catCol).Value)
            Me.lblInfoPriority.Caption = "Priority: " & CStr(tbl.ListRows(i).Range(1, priCol).Value)
            Me.lblInfoReportedBy.Caption = "Reported By: " & CStr(tbl.ListRows(i).Range(1, repByCol).Value)
            Dim dReported As Variant
            dReported = tbl.ListRows(i).Range(1, repDateCol).Value
            If IsDate(dReported) Then
                Me.lblInfoReportedDate.Caption = "Reported: " & Format(CDate(dReported), "yyyy-mm-dd hh:nn")
            Else
                Me.lblInfoReportedDate.Caption = "Reported: (unknown)"
            End If
            Exit For
        End If
    Next i

    Exit Sub
ErrHandler:
    modErrorHandler.HandleError "frmAssignment", "lstIncidents_Click", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' cboTeam_Change
' Cascading dropdown: repopulates cboPerson with active personnel matching
' the selected team from tblPersonnel.
' ----------------------------------------------------------------------------
Private Sub cboTeam_Change()
    On Error GoTo ErrHandler

    Me.cboPerson.Clear
    Me.cboPerson.Value = ""

    If Len(Me.cboTeam.Value) = 0 Then Exit Sub

    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets(SHT_SETTINGS).ListObjects(TBL_PERSONNEL)

    Dim nameCol As Long, teamCol As Long, activeCol As Long
    nameCol = tbl.ListColumns("Name").Index
    teamCol = tbl.ListColumns("Team").Index
    activeCol = tbl.ListColumns("Active").Index

    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        If CStr(tbl.ListRows(i).Range(1, teamCol).Value) = Me.cboTeam.Value Then
            If CStr(tbl.ListRows(i).Range(1, activeCol).Value) = "Yes" Then
                Me.cboPerson.AddItem CStr(tbl.ListRows(i).Range(1, nameCol).Value)
            End If
        End If
    Next i

    Exit Sub
ErrHandler:
    modErrorHandler.HandleError "frmAssignment", "cboTeam_Change", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' btnAssign_Click
' Validates selections, atomically updates 5 fields (team, person, date,
' assigned-by, status), refreshes Assignment Tracker, and closes form.
' ----------------------------------------------------------------------------
Private Sub btnAssign_Click()
    On Error GoTo ErrHandler

    ' Validate selections
    If Me.lstIncidents.ListIndex = -1 Then
        MsgBox "Please select an incident to assign.", vbExclamation, APP_TITLE
        Exit Sub
    End If
    If Len(Me.cboTeam.Value) = 0 Then
        MsgBox "Please select a team.", vbExclamation, APP_TITLE
        Me.cboTeam.SetFocus
        Exit Sub
    End If
    If Len(Me.cboPerson.Value) = 0 Then
        MsgBox "Please select an assignee.", vbExclamation, APP_TITLE
        Me.cboPerson.SetFocus
        Exit Sub
    End If

    Dim sIncidentID As String
    sIncidentID = Me.lstIncidents.Value

    ' Update 5 fields atomically
    Dim vNames As Variant
    Dim vValues As Variant
    vNames = Array(COL_ASSIGNED_TEAM, COL_ASSIGNED_TO, COL_ASSIGNED_DATE, _
                   COL_ASSIGNED_BY, COL_STATUS)
    vValues = Array(Me.cboTeam.Value, Me.cboPerson.Value, Now, _
                    Application.UserName, STATUS_ASSIGNED)

    Dim bSuccess As Boolean
    bSuccess = modDataAccess.UpdateIncidentFields(sIncidentID, vNames, vValues)

    If bSuccess Then
        MsgBox "Incident " & sIncidentID & " has been assigned to " & _
               Me.cboPerson.Value & " (" & Me.cboTeam.Value & ").", _
               vbInformation, APP_TITLE
        ' Refresh Assignment Tracker
        modAssignment.RefreshAssignmentTracker
        Unload Me
    Else
        MsgBox "Failed to assign incident. Please try again.", _
               vbExclamation, APP_TITLE
    End If

    Exit Sub
ErrHandler:
    modErrorHandler.HandleError "frmAssignment", "btnAssign_Click", _
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
    modErrorHandler.HandleError "frmAssignment", "btnCancel_Click", _
                                 Err.Number, Err.Description
End Sub
