Attribute VB_Name = "modInitialize"
Option Explicit

' ============================================================================
' Module:  modInitialize
' Purpose: Workbook setup, structural verification, and initialization.
'          Called from ThisWorkbook.Workbook_Open to ensure the workbook is
'          in a consistent state every time it opens.
' Dependencies: modConstants (SHT_*, TBL_*, SHEET_PWD)
'               modFormatting (ApplyAllBranding)
'               modErrorHandler (HandleError)
' ============================================================================

' ----------------------------------------------------------------------------
' VerifyWorkbookStructure
' Checks that all required sheets and tables exist. Logs missing items via
' Debug.Print and shows a MsgBox warning if critical components are absent.
' ----------------------------------------------------------------------------
Public Sub VerifyWorkbookStructure()
    On Error GoTo ErrHandler

    Dim bCriticalMissing As Boolean
    bCriticalMissing = False

    Dim sMissingItems As String
    sMissingItems = ""

    ' --- Check required sheets ---
    If Not SheetExists(SHT_DASHBOARD) Then
        sMissingItems = sMissingItems & vbCrLf & "  - Sheet: " & SHT_DASHBOARD
        Debug.Print "VERIFY: Missing sheet - " & SHT_DASHBOARD
    End If

    If Not SheetExists(SHT_INCIDENT_LOG) Then
        sMissingItems = sMissingItems & vbCrLf & "  - Sheet: " & SHT_INCIDENT_LOG
        bCriticalMissing = True
        Debug.Print "VERIFY: Missing CRITICAL sheet - " & SHT_INCIDENT_LOG
    End If

    If Not SheetExists(SHT_SETTINGS) Then
        sMissingItems = sMissingItems & vbCrLf & "  - Sheet: " & SHT_SETTINGS
        Debug.Print "VERIFY: Missing sheet - " & SHT_SETTINGS
    End If

    ' --- Check required tables ---
    If SheetExists(SHT_INCIDENT_LOG) Then
        If Not TableExists(SHT_INCIDENT_LOG, TBL_INCIDENTS) Then
            sMissingItems = sMissingItems & vbCrLf & "  - Table: " & TBL_INCIDENTS
            bCriticalMissing = True
            Debug.Print "VERIFY: Missing CRITICAL table - " & TBL_INCIDENTS
        End If
    End If

    If SheetExists(SHT_SETTINGS) Then
        If Not TableExists(SHT_SETTINGS, TBL_TEAMS) Then
            sMissingItems = sMissingItems & vbCrLf & "  - Table: " & TBL_TEAMS
            Debug.Print "VERIFY: Missing table - " & TBL_TEAMS
        End If
        If Not TableExists(SHT_SETTINGS, TBL_PERSONNEL) Then
            sMissingItems = sMissingItems & vbCrLf & "  - Table: " & TBL_PERSONNEL
            Debug.Print "VERIFY: Missing table - " & TBL_PERSONNEL
        End If
        If Not TableExists(SHT_SETTINGS, TBL_CATEGORIES) Then
            sMissingItems = sMissingItems & vbCrLf & "  - Table: " & TBL_CATEGORIES
            Debug.Print "VERIFY: Missing table - " & TBL_CATEGORIES
        End If
        If Not TableExists(SHT_SETTINGS, TBL_SLA_THRESHOLDS) Then
            sMissingItems = sMissingItems & vbCrLf & "  - Table: " & TBL_SLA_THRESHOLDS
            Debug.Print "VERIFY: Missing table - " & TBL_SLA_THRESHOLDS
        End If
        If Not TableExists(SHT_SETTINGS, TBL_PRIORITY_MATRIX) Then
            sMissingItems = sMissingItems & vbCrLf & "  - Table: " & TBL_PRIORITY_MATRIX
            Debug.Print "VERIFY: Missing table - " & TBL_PRIORITY_MATRIX
        End If
    End If

    ' --- Warn user if critical items are missing ---
    If bCriticalMissing Then
        MsgBox "WARNING: Critical workbook components are missing:" & vbCrLf & _
               sMissingItems & vbCrLf & vbCrLf & _
               "The Incident Management System may not function correctly. " & _
               "Please contact your system administrator.", _
               vbExclamation, APP_TITLE
    ElseIf Len(sMissingItems) > 0 Then
        Debug.Print "VERIFY: Non-critical items missing:" & sMissingItems
    Else
        Debug.Print "VERIFY: All workbook components present."
    End If

    Exit Sub

ErrHandler:
    modErrorHandler.HandleError "modInitialize", "VerifyWorkbookStructure", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' ApplyAllProtection
' Applies appropriate sheet protection to all key sheets. Called from
' Workbook_Open to re-apply UserInterfaceOnly (does not persist across
' save/close cycles).
'
' Protection strategy:
'   - Settings: Protected with AllowFiltering (admins edit table data directly,
'     so UserInterfaceOnly is False -- they unlock via password if needed)
'   - Incident Log: Protected with UserInterfaceOnly=True (VBA writes freely)
'   - Dashboard: Protected with UserInterfaceOnly=True (display only)
' ----------------------------------------------------------------------------
Public Sub ApplyAllProtection()
    On Error Resume Next

    Dim ws As Worksheet

    ' --- Settings sheet: admin-editable ---
    If SheetExists(SHT_SETTINGS) Then
        Set ws = ThisWorkbook.Sheets(SHT_SETTINGS)
        ws.Unprotect Password:=SHEET_PWD
        ws.Protect Password:=SHEET_PWD, _
            UserInterfaceOnly:=False, _
            AllowFiltering:=True
    End If

    ' --- Incident Log: VBA-writable, UI-locked ---
    If SheetExists(SHT_INCIDENT_LOG) Then
        Set ws = ThisWorkbook.Sheets(SHT_INCIDENT_LOG)
        ws.Unprotect Password:=SHEET_PWD
        ws.Protect Password:=SHEET_PWD, _
            UserInterfaceOnly:=True, _
            AllowFiltering:=True, _
            AllowSorting:=True
    End If

    ' --- Dashboard: display only ---
    If SheetExists(SHT_DASHBOARD) Then
        Set ws = ThisWorkbook.Sheets(SHT_DASHBOARD)
        ws.Unprotect Password:=SHEET_PWD
        ws.Protect Password:=SHEET_PWD, _
            UserInterfaceOnly:=True
    End If

    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' InitializeWorkbook
' Master initialization routine called from Workbook_Open. Verifies structure,
' applies protection, optionally refreshes branding, and navigates to Dashboard.
' ----------------------------------------------------------------------------
Public Sub InitializeWorkbook()
    On Error GoTo ErrHandler

    ' Step 1: Verify all required sheets and tables exist
    VerifyWorkbookStructure

    ' Step 2: Refresh branding BEFORE protection (formatting fails on protected sheets)
    modFormatting.ApplyAllBranding

    ' Step 3: Re-apply protection (UserInterfaceOnly does not persist)
    ApplyAllProtection

    ' Step 4: Navigate to Dashboard
    If SheetExists(SHT_DASHBOARD) Then
        ThisWorkbook.Sheets(SHT_DASHBOARD).Activate
    End If

    Exit Sub

ErrHandler:
    modErrorHandler.HandleError "modInitialize", "InitializeWorkbook", _
                                 Err.Number, Err.Description
End Sub

' ============================================================================
' Private Helpers
' ============================================================================

' ----------------------------------------------------------------------------
' SheetExists
' Checks whether a sheet with the given name exists in ThisWorkbook.
'
' Parameters:
'   sName - The sheet name to look for
'
' Returns:
'   True if the sheet exists, False otherwise
' ----------------------------------------------------------------------------
Private Function SheetExists(ByVal sName As String) As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sName)
    SheetExists = (Not ws Is Nothing)
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' TableExists
' Checks whether a ListObject with the given name exists on the specified sheet.
'
' Parameters:
'   sSheetName - The sheet to check
'   sTableName - The table name to look for
'
' Returns:
'   True if the table exists, False otherwise
' ----------------------------------------------------------------------------
Private Function TableExists(ByVal sSheetName As String, _
                              ByVal sTableName As String) As Boolean
    On Error Resume Next
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets(sSheetName).ListObjects(sTableName)
    TableExists = (Not tbl Is Nothing)
    On Error GoTo 0
End Function
