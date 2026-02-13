Option Explicit

' ============================================================================
' Form:    frmSearch
' Purpose: Incident search form with multi-criteria filtering, result display,
'          detail preview, and report generation.
' Dependencies: modConstants, modDataAccess, modReports, modErrorHandler
' ============================================================================

' ----------------------------------------------------------------------------
' UserForm_Initialize
' Populates filter dropdowns (status, priority, category, assignee),
' configures the multi-column ListBox, and clears detail labels.
' ----------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    On Error GoTo ErrHandler

    ' --- Populate cboStatus with empty first, then all statuses ---
    Me.cboStatus.AddItem ""
    Me.cboStatus.AddItem STATUS_OPEN
    Me.cboStatus.AddItem STATUS_ASSIGNED
    Me.cboStatus.AddItem STATUS_IN_PROGRESS
    Me.cboStatus.AddItem STATUS_RESOLVED
    Me.cboStatus.AddItem STATUS_CLOSED
    Me.cboStatus.AddItem STATUS_CANCELLED
    Me.cboStatus.AddItem STATUS_DUPLICATE

    ' --- Populate cboPriority with empty first, then P1-P4 ---
    Me.cboPriority.AddItem ""
    Me.cboPriority.AddItem PRIORITY_P1
    Me.cboPriority.AddItem PRIORITY_P2
    Me.cboPriority.AddItem PRIORITY_P3
    Me.cboPriority.AddItem PRIORITY_P4

    ' --- Populate cboCategory from tblCategories ---
    Me.cboCategory.AddItem ""
    Dim wsSettings As Worksheet
    Set wsSettings = ThisWorkbook.Sheets(SHT_SETTINGS)

    Dim tblCat As ListObject
    Set tblCat = wsSettings.ListObjects(TBL_CATEGORIES)
    If Not tblCat Is Nothing Then
        If tblCat.ListRows.Count > 0 Then
            Dim cCatName As Long
            cCatName = tblCat.ListColumns("Category").Index

            Dim dictCat As Object
            Set dictCat = CreateObject("Scripting.Dictionary")

            Dim iCat As Long
            For iCat = 1 To tblCat.ListRows.Count
                Dim sCatVal As String
                sCatVal = CStr(tblCat.ListRows(iCat).Range(1, cCatName).Value)
                If Len(sCatVal) > 0 And Not dictCat.Exists(sCatVal) Then
                    dictCat.Add sCatVal, True
                    Me.cboCategory.AddItem sCatVal
                End If
            Next iCat
        End If
    End If

    ' --- Populate cboAssignee from tblPersonnel ---
    Me.cboAssignee.AddItem ""
    Dim tblPers As ListObject
    Set tblPers = wsSettings.ListObjects(TBL_PERSONNEL)
    If Not tblPers Is Nothing Then
        If tblPers.ListRows.Count > 0 Then
            Dim cPersName As Long
            cPersName = tblPers.ListColumns("PersonnelName").Index

            Dim iPers As Long
            For iPers = 1 To tblPers.ListRows.Count
                Dim sPersVal As String
                sPersVal = CStr(tblPers.ListRows(iPers).Range(1, cPersName).Value)
                If Len(sPersVal) > 0 Then
                    Me.cboAssignee.AddItem sPersVal
                End If
            Next iPers
        End If
    End If

    ' --- Configure ListBox ---
    Me.lstResults.ColumnCount = 4
    Me.lstResults.ColumnWidths = "70;220;80;60"
    Me.lstResults.ColumnHeads = False

    ' --- Clear detail labels ---
    ClearDetailLabels

    Me.lblResultCount.Caption = "Results:"

    Exit Sub
ErrHandler:
    modErrorHandler.HandleError "frmSearch", "UserForm_Initialize", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' btnSearch_Click
' Single-pass AND filter through tblIncidents. Each criterion is AND-combined;
' empty criteria match all.
' ----------------------------------------------------------------------------
Private Sub btnSearch_Click()
    On Error GoTo ErrHandler

    Me.lstResults.Clear
    ClearDetailLabels

    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets(SHT_INCIDENT_LOG).ListObjects(TBL_INCIDENTS)
    If tbl.ListRows.Count = 0 Then
        Me.lblResultCount.Caption = "Results: 0 matches found"
        Exit Sub
    End If

    ' Cache column indices once for performance
    Dim cID As Long, cTitle As Long, cStatus As Long, cPriority As Long
    Dim cCategory As Long, cAssignee As Long, cReportedDate As Long
    cID = tbl.ListColumns(COL_INCIDENT_ID).Index
    cTitle = tbl.ListColumns(COL_TITLE).Index
    cStatus = tbl.ListColumns(COL_STATUS).Index
    cPriority = tbl.ListColumns(COL_PRIORITY).Index
    cCategory = tbl.ListColumns(COL_CATEGORY).Index
    cAssignee = tbl.ListColumns(COL_ASSIGNED_TO).Index
    cReportedDate = tbl.ListColumns(COL_REPORTED_DATE).Index

    Dim lMatchCount As Long
    lMatchCount = 0

    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        Dim bMatch As Boolean
        bMatch = True

        ' ID filter (partial, case-insensitive)
        If bMatch And Len(Me.txtSearchID.Value) > 0 Then
            If InStr(1, CStr(tbl.ListRows(i).Range(1, cID).Value), _
                     Me.txtSearchID.Value, vbTextCompare) = 0 Then
                bMatch = False
            End If
        End If

        ' Title filter (partial, case-insensitive)
        If bMatch And Len(Me.txtSearchTitle.Value) > 0 Then
            If InStr(1, CStr(tbl.ListRows(i).Range(1, cTitle).Value), _
                     Me.txtSearchTitle.Value, vbTextCompare) = 0 Then
                bMatch = False
            End If
        End If

        ' Status filter (exact)
        If bMatch And Len(Me.cboStatus.Value) > 0 Then
            If CStr(tbl.ListRows(i).Range(1, cStatus).Value) <> Me.cboStatus.Value Then
                bMatch = False
            End If
        End If

        ' Priority filter (exact)
        If bMatch And Len(Me.cboPriority.Value) > 0 Then
            If CStr(tbl.ListRows(i).Range(1, cPriority).Value) <> Me.cboPriority.Value Then
                bMatch = False
            End If
        End If

        ' Category filter (exact)
        If bMatch And Len(Me.cboCategory.Value) > 0 Then
            If CStr(tbl.ListRows(i).Range(1, cCategory).Value) <> Me.cboCategory.Value Then
                bMatch = False
            End If
        End If

        ' Assignee filter (exact)
        If bMatch And Len(Me.cboAssignee.Value) > 0 Then
            If CStr(tbl.ListRows(i).Range(1, cAssignee).Value) <> Me.cboAssignee.Value Then
                bMatch = False
            End If
        End If

        ' Date From filter (ReportedDate >= value)
        If bMatch And Len(Me.txtDateFrom.Value) > 0 Then
            If IsDate(Me.txtDateFrom.Value) Then
                Dim vDateFrom As Variant
                vDateFrom = tbl.ListRows(i).Range(1, cReportedDate).Value
                If IsDate(vDateFrom) Then
                    If CDate(vDateFrom) < CDate(Me.txtDateFrom.Value) Then bMatch = False
                End If
            End If
        End If

        ' Date To filter (ReportedDate <= value)
        If bMatch And Len(Me.txtDateTo.Value) > 0 Then
            If IsDate(Me.txtDateTo.Value) Then
                Dim vDateTo As Variant
                vDateTo = tbl.ListRows(i).Range(1, cReportedDate).Value
                If IsDate(vDateTo) Then
                    If CDate(vDateTo) > CDate(Me.txtDateTo.Value) Then bMatch = False
                End If
            End If
        End If

        ' Add match to ListBox
        If bMatch Then
            lMatchCount = lMatchCount + 1
            Me.lstResults.AddItem CStr(tbl.ListRows(i).Range(1, cID).Value)
            Me.lstResults.List(Me.lstResults.ListCount - 1, 1) = _
                Left(CStr(tbl.ListRows(i).Range(1, cTitle).Value), 50)
            Me.lstResults.List(Me.lstResults.ListCount - 1, 2) = _
                CStr(tbl.ListRows(i).Range(1, cStatus).Value)
            Me.lstResults.List(Me.lstResults.ListCount - 1, 3) = _
                CStr(tbl.ListRows(i).Range(1, cPriority).Value)
        End If
    Next i

    Me.lblResultCount.Caption = "Results: " & lMatchCount & " matches found"
    Exit Sub
ErrHandler:
    modErrorHandler.HandleError "frmSearch", "btnSearch_Click", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' btnClear_Click
' Clears all search criteria, results, and detail labels.
' ----------------------------------------------------------------------------
Private Sub btnClear_Click()
    Me.txtSearchID.Value = ""
    Me.txtSearchTitle.Value = ""
    Me.cboStatus.Value = ""
    Me.cboPriority.Value = ""
    Me.cboCategory.Value = ""
    Me.cboAssignee.Value = ""
    Me.txtDateFrom.Value = ""
    Me.txtDateTo.Value = ""
    Me.lstResults.Clear
    ClearDetailLabels
    Me.lblResultCount.Caption = "Results:"
End Sub

' ----------------------------------------------------------------------------
' lstResults_Click
' When user clicks a result row, populate detail labels below using
' ReadIncident + GetIncidentColIdx for 2D array access.
' ----------------------------------------------------------------------------
Private Sub lstResults_Click()
    On Error GoTo ErrHandler

    If Me.lstResults.ListIndex = -1 Then Exit Sub

    Dim sID As String
    sID = Me.lstResults.Value

    Dim vData As Variant
    vData = modDataAccess.ReadIncident(sID)
    If IsEmpty(vData) Then
        ClearDetailLabels
        Exit Sub
    End If

    ' Use GetIncidentColIdx to decode the 2D array vData(1, colIndex)
    Me.lblDetailID.Caption = "ID: " & CStr(vData(1, modDataAccess.GetIncidentColIdx(COL_INCIDENT_ID))) & _
        "     Status: " & CStr(vData(1, modDataAccess.GetIncidentColIdx(COL_STATUS))) & _
        "     Priority: " & CStr(vData(1, modDataAccess.GetIncidentColIdx(COL_PRIORITY)))

    Me.lblDetailCategory.Caption = "Category: " & _
        CStr(vData(1, modDataAccess.GetIncidentColIdx(COL_CATEGORY))) & " / " & _
        CStr(vData(1, modDataAccess.GetIncidentColIdx(COL_SUBCATEGORY)))

    ' Reported info
    Dim sReported As String
    sReported = "Reported: "
    Dim vRepDate As Variant
    vRepDate = vData(1, modDataAccess.GetIncidentColIdx(COL_REPORTED_DATE))
    If IsDate(vRepDate) Then sReported = sReported & Format(CDate(vRepDate), "yyyy-mm-dd hh:nn")
    sReported = sReported & " by " & CStr(vData(1, modDataAccess.GetIncidentColIdx(COL_REPORTED_BY)))
    Me.lblDetailReported.Caption = sReported

    ' Assignment info
    Dim sAssigned As String
    Dim vAssignTo As Variant
    vAssignTo = vData(1, modDataAccess.GetIncidentColIdx(COL_ASSIGNED_TO))
    If Len(Trim(CStr(vAssignTo))) > 0 And CStr(vAssignTo) <> "0" Then
        sAssigned = "Assigned: " & CStr(vAssignTo) & " (" & _
            CStr(vData(1, modDataAccess.GetIncidentColIdx(COL_ASSIGNED_TEAM))) & ")"
    Else
        sAssigned = "Assigned: (not yet assigned)"
    End If
    Me.lblDetailAssigned.Caption = sAssigned

    ' Description (first 200 chars)
    Dim sDesc As String
    sDesc = CStr(vData(1, modDataAccess.GetIncidentColIdx(COL_DESCRIPTION)))
    If Len(sDesc) > 200 Then sDesc = Left(sDesc, 200) & "..."
    Me.lblDetailDescription.Caption = "Description: " & sDesc

    ' Enable action buttons
    Me.btnViewDetails.Enabled = True
    Me.btnGenerateReport.Enabled = True

    Exit Sub
ErrHandler:
    modErrorHandler.HandleError "frmSearch", "lstResults_Click", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' btnViewDetails_Click
' Shows full incident details in a formatted MsgBox.
' ----------------------------------------------------------------------------
Private Sub btnViewDetails_Click()
    On Error GoTo ErrHandler

    If Me.lstResults.ListIndex = -1 Then
        MsgBox "Please select an incident first.", vbInformation, APP_TITLE
        Exit Sub
    End If

    Dim sID As String
    sID = Me.lstResults.Value

    Dim vData As Variant
    vData = modDataAccess.ReadIncident(sID)
    If IsEmpty(vData) Then
        MsgBox "Incident not found.", vbExclamation, APP_TITLE
        Exit Sub
    End If

    ' Build detail string
    Dim sMsg As String
    sMsg = "INCIDENT DETAILS" & vbCrLf & String(40, "-") & vbCrLf & vbCrLf

    sMsg = sMsg & "ID: " & CStr(vData(1, modDataAccess.GetIncidentColIdx(COL_INCIDENT_ID))) & vbCrLf
    sMsg = sMsg & "Title: " & CStr(vData(1, modDataAccess.GetIncidentColIdx(COL_TITLE))) & vbCrLf
    sMsg = sMsg & "Status: " & CStr(vData(1, modDataAccess.GetIncidentColIdx(COL_STATUS))) & vbCrLf
    sMsg = sMsg & "Priority: " & CStr(vData(1, modDataAccess.GetIncidentColIdx(COL_PRIORITY))) & vbCrLf
    sMsg = sMsg & "Category: " & CStr(vData(1, modDataAccess.GetIncidentColIdx(COL_CATEGORY))) & " / " & _
        CStr(vData(1, modDataAccess.GetIncidentColIdx(COL_SUBCATEGORY))) & vbCrLf & vbCrLf

    sMsg = sMsg & "Description:" & vbCrLf
    sMsg = sMsg & CStr(vData(1, modDataAccess.GetIncidentColIdx(COL_DESCRIPTION))) & vbCrLf & vbCrLf

    sMsg = sMsg & "TIMELINE" & vbCrLf & String(40, "-") & vbCrLf
    ' Reported
    Dim vVal As Variant
    vVal = vData(1, modDataAccess.GetIncidentColIdx(COL_REPORTED_DATE))
    If IsDate(vVal) Then sMsg = sMsg & "Reported: " & Format(CDate(vVal), "yyyy-mm-dd hh:nn") & _
        " by " & CStr(vData(1, modDataAccess.GetIncidentColIdx(COL_REPORTED_BY))) & vbCrLf
    ' Assigned
    vVal = vData(1, modDataAccess.GetIncidentColIdx(COL_ASSIGNED_DATE))
    If IsDate(vVal) Then sMsg = sMsg & "Assigned: " & Format(CDate(vVal), "yyyy-mm-dd hh:nn") & _
        " to " & CStr(vData(1, modDataAccess.GetIncidentColIdx(COL_ASSIGNED_TO))) & vbCrLf
    ' Response
    vVal = vData(1, modDataAccess.GetIncidentColIdx(COL_RESPONSE_DATE))
    If IsDate(vVal) Then sMsg = sMsg & "Response: " & Format(CDate(vVal), "yyyy-mm-dd hh:nn") & vbCrLf
    ' Resolution
    vVal = vData(1, modDataAccess.GetIncidentColIdx(COL_RESOLUTION_DATE))
    If IsDate(vVal) Then sMsg = sMsg & "Resolution: " & Format(CDate(vVal), "yyyy-mm-dd hh:nn") & vbCrLf
    ' Closed
    vVal = vData(1, modDataAccess.GetIncidentColIdx(COL_CLOSED_DATE))
    If IsDate(vVal) Then sMsg = sMsg & "Closed: " & Format(CDate(vVal), "yyyy-mm-dd hh:nn") & vbCrLf

    sMsg = sMsg & vbCrLf & "RCA" & vbCrLf & String(40, "-") & vbCrLf
    Dim sRC As String
    sRC = CStr(vData(1, modDataAccess.GetIncidentColIdx(COL_ROOT_CAUSE)))
    If Len(Trim(sRC)) > 0 And sRC <> "0" Then
        sMsg = sMsg & "Root Cause: " & sRC & vbCrLf
        sMsg = sMsg & "Corrective: " & CStr(vData(1, modDataAccess.GetIncidentColIdx(COL_CORRECTIVE_ACTION))) & vbCrLf
        sMsg = sMsg & "Preventive: " & CStr(vData(1, modDataAccess.GetIncidentColIdx(COL_PREVENTIVE_ACTION))) & vbCrLf
    Else
        sMsg = sMsg & "(No RCA recorded)" & vbCrLf
    End If

    MsgBox sMsg, vbInformation, APP_TITLE & " - " & sID

    Exit Sub
ErrHandler:
    modErrorHandler.HandleError "frmSearch", "btnViewDetails_Click", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' btnGenerateReport_Click
' Generates a print-ready incident report for the selected incident.
' Hides the form temporarily while the report generates.
' ----------------------------------------------------------------------------
Private Sub btnGenerateReport_Click()
    On Error GoTo ErrHandler

    If Me.lstResults.ListIndex = -1 Then
        MsgBox "Please select an incident first.", vbInformation, APP_TITLE
        Exit Sub
    End If

    Dim sID As String
    sID = Me.lstResults.Value

    ' Hide form temporarily while report generates
    Me.Hide
    modReports.GenerateIncidentReport sID
    Me.Show

    Exit Sub
ErrHandler:
    On Error Resume Next
    Me.Show
    On Error GoTo 0
    modErrorHandler.HandleError "frmSearch", "btnGenerateReport_Click", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' btnClose_Click
' Closes and unloads the search form.
' ----------------------------------------------------------------------------
Private Sub btnClose_Click()
    Unload Me
End Sub

' ----------------------------------------------------------------------------
' txtDateFrom_Exit
' Validates the Date From field on exit.
' ----------------------------------------------------------------------------
Private Sub txtDateFrom_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Len(Me.txtDateFrom.Value) > 0 Then
        If Not IsDate(Me.txtDateFrom.Value) Then
            MsgBox "Please enter a valid date (e.g., 2026-01-15).", vbExclamation, APP_TITLE
            Cancel = True
        End If
    End If
End Sub

' ----------------------------------------------------------------------------
' txtDateTo_Exit
' Validates the Date To field on exit.
' ----------------------------------------------------------------------------
Private Sub txtDateTo_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Len(Me.txtDateTo.Value) > 0 Then
        If Not IsDate(Me.txtDateTo.Value) Then
            MsgBox "Please enter a valid date (e.g., 2026-01-15).", vbExclamation, APP_TITLE
            Cancel = True
        End If
    End If
End Sub

' ----------------------------------------------------------------------------
' ClearDetailLabels
' Resets all detail labels and disables action buttons.
' ----------------------------------------------------------------------------
Private Sub ClearDetailLabels()
    Me.lblDetailID.Caption = ""
    Me.lblDetailCategory.Caption = ""
    Me.lblDetailReported.Caption = ""
    Me.lblDetailAssigned.Caption = ""
    Me.lblDetailDescription.Caption = ""
    Me.btnViewDetails.Enabled = False
    Me.btnGenerateReport.Enabled = False
End Sub
