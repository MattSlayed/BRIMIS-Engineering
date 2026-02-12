Attribute VB_Name = "modDataAccess"
Option Explicit

' ============================================================================
' Module:  modDataAccess
' Purpose: Central data access layer for the tblIncidents table.
'          All incident CRUD operations go through this module.
'          Sheet/table names are resolved once via GetIncidentTable();
'          individual procedures use only the returned ListObject reference
'          and column-name constants from modConstants.
' Dependencies: modConstants (TBL_INCIDENTS, SHT_INCIDENT_LOG, COL_*, ID_PREFIX, STATUS_OPEN)
'               modUtilities (UnprotectSheet, ProtectSheet)
'               modErrorHandler (HandleError)
' ============================================================================

' ----------------------------------------------------------------------------
' GetIncidentTable
' Returns the tblIncidents ListObject from the Incident Log sheet.
' This is the ONLY place in the module that references sheet/table names.
'
' Returns:
'   ListObject reference to tblIncidents
' ----------------------------------------------------------------------------
Private Function GetIncidentTable() As ListObject
    On Error GoTo ErrHandler

    Set GetIncidentTable = ThisWorkbook.Sheets(SHT_INCIDENT_LOG).ListObjects(TBL_INCIDENTS)
    Exit Function

ErrHandler:
    modErrorHandler.HandleError "modDataAccess", "GetIncidentTable", _
                                 Err.Number, Err.Description
    Set GetIncidentTable = Nothing
End Function

' ----------------------------------------------------------------------------
' ColIdx
' Returns the 1-based column index for a named column in the given table.
' Used by all read/write operations to avoid hard-coded column positions.
'
' Parameters:
'   tbl      - The ListObject to look up
'   colName  - The column header name (use COL_* constants)
'
' Returns:
'   1-based column index within the table
' ----------------------------------------------------------------------------
Private Function ColIdx(tbl As ListObject, ByVal colName As String) As Long
    On Error GoTo ErrHandler

    ColIdx = tbl.ListColumns(colName).Index
    Exit Function

ErrHandler:
    modErrorHandler.HandleError "modDataAccess", "ColIdx", _
                                 Err.Number, Err.Description
    ColIdx = 0
End Function

' ----------------------------------------------------------------------------
' WriteIncident
' Adds a new incident row to tblIncidents with all required fields populated.
' Automatically sets Status=Open, ReportedDate=Now, LastModified=Now,
' LastModifiedBy=Application.UserName.
'
' Parameters:
'   sID           - Incident ID (e.g., "INC-00001")
'   sTitle        - Short title
'   sDesc         - Full description
'   sCategory     - Category value (from tblCategories)
'   sSubcategory  - Subcategory value
'   sPriority     - Priority level (P1-P4)
'   lImpact       - Impact score (1-5)
'   lUrgency      - Urgency score (1-5)
'   sReportedBy   - Name of person reporting
'   sAttachmentRef - (Optional) Attachment reference/path
'
' Returns:
'   True on success, False on failure
' ----------------------------------------------------------------------------
Public Function WriteIncident(ByVal sID As String, _
                              ByVal sTitle As String, _
                              ByVal sDesc As String, _
                              ByVal sCategory As String, _
                              ByVal sSubcategory As String, _
                              ByVal sPriority As String, _
                              ByVal lImpact As Long, _
                              ByVal lUrgency As Long, _
                              ByVal sReportedBy As String, _
                              Optional ByVal sAttachmentRef As String = "") As Boolean
    On Error GoTo ErrHandler

    Dim tbl As ListObject
    Set tbl = GetIncidentTable()
    If tbl Is Nothing Then
        WriteIncident = False
        Exit Function
    End If

    ' Unprotect sheet for writing
    modUtilities.UnprotectSheet tbl.Parent

    ' Add a new row
    Dim newRow As ListRow
    Set newRow = tbl.ListRows.Add

    ' Populate fields
    newRow.Range(1, ColIdx(tbl, COL_INCIDENT_ID)).Value = sID
    newRow.Range(1, ColIdx(tbl, COL_TITLE)).Value = sTitle
    newRow.Range(1, ColIdx(tbl, COL_DESCRIPTION)).Value = sDesc
    newRow.Range(1, ColIdx(tbl, COL_CATEGORY)).Value = sCategory
    newRow.Range(1, ColIdx(tbl, COL_SUBCATEGORY)).Value = sSubcategory
    newRow.Range(1, ColIdx(tbl, COL_PRIORITY)).Value = sPriority
    newRow.Range(1, ColIdx(tbl, COL_IMPACT)).Value = lImpact
    newRow.Range(1, ColIdx(tbl, COL_URGENCY)).Value = lUrgency
    newRow.Range(1, ColIdx(tbl, COL_STATUS)).Value = STATUS_OPEN
    newRow.Range(1, ColIdx(tbl, COL_REPORTED_BY)).Value = sReportedBy
    newRow.Range(1, ColIdx(tbl, COL_REPORTED_DATE)).Value = Now
    newRow.Range(1, ColIdx(tbl, COL_LAST_MODIFIED)).Value = Now
    newRow.Range(1, ColIdx(tbl, COL_LAST_MODIFIED_BY)).Value = Application.UserName

    ' Set attachment reference if provided
    If Len(sAttachmentRef) > 0 Then
        newRow.Range(1, ColIdx(tbl, COL_ATTACHMENT_REF)).Value = sAttachmentRef
    End If

    ' Re-protect sheet
    modUtilities.ProtectSheet tbl.Parent

    WriteIncident = True
    Exit Function

ErrHandler:
    ' Always re-protect before surfacing the error
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHT_INCIDENT_LOG)
    modUtilities.ProtectSheet ws
    On Error GoTo 0

    modErrorHandler.HandleError "modDataAccess", "WriteIncident", _
                                 Err.Number, Err.Description
    WriteIncident = False
End Function

' ----------------------------------------------------------------------------
' ReadIncident
' Retrieves a complete incident row by IncidentID.
'
' Parameters:
'   sIncidentID - The incident ID to look up (e.g., "INC-00001")
'
' Returns:
'   Variant array of the entire row if found, Empty if not found
' ----------------------------------------------------------------------------
Public Function ReadIncident(ByVal sIncidentID As String) As Variant
    On Error GoTo ErrHandler

    Dim tbl As ListObject
    Set tbl = GetIncidentTable()
    If tbl Is Nothing Then
        ReadIncident = Empty
        Exit Function
    End If

    Dim idCol As Long
    idCol = ColIdx(tbl, COL_INCIDENT_ID)
    If idCol = 0 Then
        ReadIncident = Empty
        Exit Function
    End If

    ' Loop through rows to find matching ID
    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        If CStr(tbl.ListRows(i).Range(1, idCol).Value) = sIncidentID Then
            ReadIncident = tbl.ListRows(i).Range.Value
            Exit Function
        End If
    Next i

    ' Not found
    ReadIncident = Empty
    Exit Function

ErrHandler:
    modErrorHandler.HandleError "modDataAccess", "ReadIncident", _
                                 Err.Number, Err.Description
    ReadIncident = Empty
End Function

' ----------------------------------------------------------------------------
' UpdateIncidentField
' Updates a single field for a given incident, identified by IncidentID.
' Automatically updates LastModified and LastModifiedBy.
'
' Parameters:
'   sIncidentID - The incident ID to update
'   sColumnName - The column name to update (use COL_* constants)
'   vNewValue   - The new value to set
'
' Returns:
'   True on success, False if not found or on error
' ----------------------------------------------------------------------------
Public Function UpdateIncidentField(ByVal sIncidentID As String, _
                                     ByVal sColumnName As String, _
                                     ByVal vNewValue As Variant) As Boolean
    On Error GoTo ErrHandler

    Dim tbl As ListObject
    Set tbl = GetIncidentTable()
    If tbl Is Nothing Then
        UpdateIncidentField = False
        Exit Function
    End If

    Dim idCol As Long
    idCol = ColIdx(tbl, COL_INCIDENT_ID)
    If idCol = 0 Then
        UpdateIncidentField = False
        Exit Function
    End If

    Dim targetCol As Long
    targetCol = ColIdx(tbl, sColumnName)
    If targetCol = 0 Then
        UpdateIncidentField = False
        Exit Function
    End If

    ' Find the row
    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        If CStr(tbl.ListRows(i).Range(1, idCol).Value) = sIncidentID Then
            ' Found -- unprotect, update, re-protect
            modUtilities.UnprotectSheet tbl.Parent

            tbl.ListRows(i).Range(1, targetCol).Value = vNewValue
            tbl.ListRows(i).Range(1, ColIdx(tbl, COL_LAST_MODIFIED)).Value = Now
            tbl.ListRows(i).Range(1, ColIdx(tbl, COL_LAST_MODIFIED_BY)).Value = Application.UserName

            modUtilities.ProtectSheet tbl.Parent

            UpdateIncidentField = True
            Exit Function
        End If
    Next i

    ' Not found
    UpdateIncidentField = False
    Exit Function

ErrHandler:
    ' Always re-protect before surfacing the error
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHT_INCIDENT_LOG)
    modUtilities.ProtectSheet ws
    On Error GoTo 0

    modErrorHandler.HandleError "modDataAccess", "UpdateIncidentField", _
                                 Err.Number, Err.Description
    UpdateIncidentField = False
End Function

' ----------------------------------------------------------------------------
' GetNextIncidentID
' Generates the next sequential incident ID in the format INC-NNNNN.
' Scans all existing IDs, finds the maximum numeric portion, and returns
' the next value (zero-padded to 5 digits).
'
' Returns:
'   Next incident ID string (e.g., "INC-00001" for an empty table)
' ----------------------------------------------------------------------------
Public Function GetNextIncidentID() As String
    On Error GoTo ErrHandler

    Dim tbl As ListObject
    Set tbl = GetIncidentTable()
    If tbl Is Nothing Then
        GetNextIncidentID = ID_PREFIX & Format(1, "00000")
        Exit Function
    End If

    ' Empty table -- first incident
    If tbl.ListRows.Count = 0 Then
        GetNextIncidentID = ID_PREFIX & Format(1, "00000")
        Exit Function
    End If

    Dim idCol As Long
    idCol = ColIdx(tbl, COL_INCIDENT_ID)
    If idCol = 0 Then
        GetNextIncidentID = ID_PREFIX & Format(1, "00000")
        Exit Function
    End If

    ' Scan all rows for the maximum numeric portion
    Dim maxNum As Long
    maxNum = 0

    Dim i As Long
    Dim sVal As String
    Dim sNumPart As String
    Dim lNum As Long

    For i = 1 To tbl.ListRows.Count
        sVal = CStr(tbl.ListRows(i).Range(1, idCol).Value)

        ' Extract numeric portion after the prefix
        If Left(sVal, Len(ID_PREFIX)) = ID_PREFIX Then
            sNumPart = Mid(sVal, Len(ID_PREFIX) + 1)

            ' Validate that the remainder is numeric
            If IsNumeric(sNumPart) Then
                lNum = CLng(sNumPart)
                If lNum > maxNum Then
                    maxNum = lNum
                End If
            End If
        End If
    Next i

    GetNextIncidentID = ID_PREFIX & Format(maxNum + 1, "00000")
    Exit Function

ErrHandler:
    modErrorHandler.HandleError "modDataAccess", "GetNextIncidentID", _
                                 Err.Number, Err.Description
    ' Fallback -- return first ID
    GetNextIncidentID = ID_PREFIX & Format(1, "00000")
End Function

' ----------------------------------------------------------------------------
' GetIncidentRowIndex
' Returns the 1-based ListRow index for a given IncidentID, or 0 if not found.
' Helper for future phases that need to locate rows.
'
' Parameters:
'   sIncidentID - The incident ID to find
'
' Returns:
'   1-based ListRow index, or 0 if not found
' ----------------------------------------------------------------------------
Public Function GetIncidentRowIndex(ByVal sIncidentID As String) As Long
    On Error GoTo ErrHandler

    Dim tbl As ListObject
    Set tbl = GetIncidentTable()
    If tbl Is Nothing Then
        GetIncidentRowIndex = 0
        Exit Function
    End If

    Dim idCol As Long
    idCol = ColIdx(tbl, COL_INCIDENT_ID)
    If idCol = 0 Then
        GetIncidentRowIndex = 0
        Exit Function
    End If

    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        If CStr(tbl.ListRows(i).Range(1, idCol).Value) = sIncidentID Then
            GetIncidentRowIndex = i
            Exit Function
        End If
    Next i

    ' Not found
    GetIncidentRowIndex = 0
    Exit Function

ErrHandler:
    modErrorHandler.HandleError "modDataAccess", "GetIncidentRowIndex", _
                                 Err.Number, Err.Description
    GetIncidentRowIndex = 0
End Function
