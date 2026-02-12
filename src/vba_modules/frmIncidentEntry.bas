Option Explicit

' ============================================================================
' Form:    frmIncidentEntry
' Purpose: Multi-step wizard for logging new incidents. Contains all event
'          handler code for the 4-step incident entry form.
'          Step 1: Category Selection (category, subcategory, reporter)
'          Step 2: Incident Details (title, description, attachment)
'          Step 3: Priority Assessment (impact, urgency, auto-priority)
'          Step 4: Review and Submit
'
' Control Names: lblStepIndicator, mpWizard, btnBack, btnNext, btnSubmit,
'   btnCancel, cboCategory, cboSubcategory, cboReporter, txtTitle,
'   txtDescription, txtAttachment, btnBrowse, cboImpact, cboUrgency,
'   lblPriorityValue, lblPriorityExplain, lblReviewCategory,
'   lblReviewSubcategory, lblReviewReporter, lblReviewTitle,
'   lblReviewDescription, lblReviewAttachment, lblReviewPriority,
'   lblReviewImpact, lblReviewUrgency
'
' Dependencies: modConstants (SHT_SETTINGS, TBL_CATEGORIES, TBL_PERSONNEL,
'               TBL_PRIORITY_MATRIX, CLR_BRIMIS_RED, CLR_BRIMIS_GRAY,
'               APP_TITLE)
'               modDataAccess (WriteIncident, GetNextIncidentID)
'               modErrorHandler (HandleError)
' ============================================================================

' ----------------------------------------------------------------------------
' UserForm_Initialize
' Populates all dropdown controls from Settings sheet tables and sets the
' wizard to Step 1.
' ----------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    On Error GoTo ErrHandler

    ' --- Populate Category dropdown from tblCategories (unique active categories) ---
    Dim tblCat As ListObject
    Set tblCat = ThisWorkbook.Sheets(SHT_SETTINGS).ListObjects(TBL_CATEGORIES)

    Dim dictCat As Object
    Set dictCat = CreateObject("Scripting.Dictionary")

    Dim catCol As Long, activeCol As Long
    catCol = tblCat.ListColumns("Category").Index
    activeCol = tblCat.ListColumns("Active").Index

    Dim i As Long
    For i = 1 To tblCat.ListRows.Count
        Dim sCat As String
        sCat = CStr(tblCat.ListRows(i).Range(1, catCol).Value)
        If CStr(tblCat.ListRows(i).Range(1, activeCol).Value) = "Yes" Then
            If Not dictCat.Exists(sCat) Then
                dictCat.Add sCat, True
                Me.cboCategory.AddItem sCat
            End If
        End If
    Next i

    ' --- Populate Reporter dropdown from tblPersonnel (active personnel) ---
    Dim tblPers As ListObject
    Set tblPers = ThisWorkbook.Sheets(SHT_SETTINGS).ListObjects(TBL_PERSONNEL)

    Dim nameCol As Long, persActiveCol As Long
    nameCol = tblPers.ListColumns("Name").Index
    persActiveCol = tblPers.ListColumns("Active").Index

    For i = 1 To tblPers.ListRows.Count
        If CStr(tblPers.ListRows(i).Range(1, persActiveCol).Value) = "Yes" Then
            Me.cboReporter.AddItem CStr(tblPers.ListRows(i).Range(1, nameCol).Value)
        End If
    Next i

    ' --- Populate Impact and Urgency dropdowns with fixed items ---
    Me.cboImpact.AddItem "1 - Low"
    Me.cboImpact.AddItem "2 - Medium"
    Me.cboImpact.AddItem "3 - High"
    Me.cboImpact.AddItem "4 - Critical"

    Me.cboUrgency.AddItem "1 - Low"
    Me.cboUrgency.AddItem "2 - Medium"
    Me.cboUrgency.AddItem "3 - High"
    Me.cboUrgency.AddItem "4 - Critical"

    ' --- Set initial wizard state ---
    Me.mpWizard.Value = 0
    UpdateNavButtons

    Exit Sub
ErrHandler:
    modErrorHandler.HandleError "frmIncidentEntry", "UserForm_Initialize", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' cboCategory_Change
' Cascading dropdown: repopulates cboSubcategory with matching subcategories
' from tblCategories when the category selection changes.
' ----------------------------------------------------------------------------
Private Sub cboCategory_Change()
    On Error GoTo ErrHandler

    ' Clear existing subcategory items
    Me.cboSubcategory.Clear
    Me.cboSubcategory.Value = ""

    If Len(Me.cboCategory.Value) = 0 Then Exit Sub

    ' Get tblCategories from Settings sheet
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets(SHT_SETTINGS).ListObjects(TBL_CATEGORIES)

    If tbl.ListRows.Count = 0 Then Exit Sub

    ' Find column indices
    Dim catCol As Long, subCol As Long, activeCol As Long
    catCol = tbl.ListColumns("Category").Index
    subCol = tbl.ListColumns("Subcategory").Index
    activeCol = tbl.ListColumns("Active").Index

    ' Loop through rows and add matching subcategories
    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        If CStr(tbl.ListRows(i).Range(1, catCol).Value) = Me.cboCategory.Value Then
            If CStr(tbl.ListRows(i).Range(1, activeCol).Value) = "Yes" Then
                Me.cboSubcategory.AddItem CStr(tbl.ListRows(i).Range(1, subCol).Value)
            End If
        End If
    Next i

    Exit Sub
ErrHandler:
    modErrorHandler.HandleError "frmIncidentEntry", "cboCategory_Change", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' cboImpact_Change
' Recalculates priority when the impact selection changes.
' ----------------------------------------------------------------------------
Private Sub cboImpact_Change()
    On Error GoTo ErrHandler
    CalculatePriority
    Exit Sub
ErrHandler:
    modErrorHandler.HandleError "frmIncidentEntry", "cboImpact_Change", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' cboUrgency_Change
' Recalculates priority when the urgency selection changes.
' ----------------------------------------------------------------------------
Private Sub cboUrgency_Change()
    On Error GoTo ErrHandler
    CalculatePriority
    Exit Sub
ErrHandler:
    modErrorHandler.HandleError "frmIncidentEntry", "cboUrgency_Change", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' CalculatePriority
' Looks up the priority value (P1-P4) from tblPriorityMatrix based on the
' selected impact and urgency levels. Updates lblPriorityValue with the
' result and applies color coding.
' ----------------------------------------------------------------------------
Private Sub CalculatePriority()
    On Error GoTo ErrHandler

    ' If either selection is empty, show placeholder text
    If Len(Me.cboImpact.Value) = 0 Or Len(Me.cboUrgency.Value) = 0 Then
        Me.lblPriorityValue.Caption = "(select both Impact and Urgency)"
        Me.lblPriorityValue.ForeColor = CLR_BRIMIS_GRAY
        Exit Sub
    End If

    ' Extract numeric portion from selections like "4 - Critical"
    Dim lImpact As Long, lUrgency As Long
    lImpact = CLng(Left(Me.cboImpact.Value, 1))
    lUrgency = CLng(Left(Me.cboUrgency.Value, 1))

    ' Build labels to match tblPriorityMatrix row/column headers
    Dim impactLabel As String
    Select Case lImpact
        Case 4: impactLabel = "4-Critical"
        Case 3: impactLabel = "3-High"
        Case 2: impactLabel = "2-Medium"
        Case 1: impactLabel = "1-Low"
    End Select

    Dim urgencyLabel As String
    Select Case lUrgency
        Case 4: urgencyLabel = "4-Critical"
        Case 3: urgencyLabel = "3-High"
        Case 2: urgencyLabel = "2-Medium"
        Case 1: urgencyLabel = "1-Low"
    End Select

    ' Lookup in tblPriorityMatrix on Settings sheet
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets(SHT_SETTINGS).ListObjects(TBL_PRIORITY_MATRIX)

    Dim labelCol As Long
    labelCol = tbl.ListColumns("Impact\Urgency").Index

    Dim urgCol As Long
    urgCol = tbl.ListColumns(urgencyLabel).Index

    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        If CStr(tbl.ListRows(i).Range(1, labelCol).Value) = impactLabel Then
            Dim sPriority As String
            sPriority = CStr(tbl.ListRows(i).Range(1, urgCol).Value)
            Me.lblPriorityValue.Caption = sPriority

            ' Color the priority label based on level
            Select Case sPriority
                Case "P1"
                    Me.lblPriorityValue.ForeColor = CLR_BRIMIS_RED   ' 2372078
                Case "P2"
                    Me.lblPriorityValue.ForeColor = 42495            ' RGB(255, 165, 0) orange
                Case "P3"
                    Me.lblPriorityValue.ForeColor = 32768            ' RGB(0, 128, 0) green
                Case "P4"
                    Me.lblPriorityValue.ForeColor = CLR_BRIMIS_GRAY  ' 3946290
            End Select

            Exit Sub
        End If
    Next i

    Me.lblPriorityValue.Caption = "(not found)"

    Exit Sub
ErrHandler:
    modErrorHandler.HandleError "frmIncidentEntry", "CalculatePriority", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' ValidateCurrentPage
' Per-step validation called before advancing to the next wizard page.
' Returns True if the current page passes all required field checks.
' ----------------------------------------------------------------------------
Private Function ValidateCurrentPage() As Boolean
    On Error GoTo ErrHandler

    ValidateCurrentPage = False

    Select Case Me.mpWizard.Value
        Case 0  ' Category Selection
            If Len(Me.cboCategory.Value) = 0 Then
                MsgBox "Please select a category.", vbExclamation, APP_TITLE
                Me.cboCategory.SetFocus
                Exit Function
            End If
            If Len(Me.cboSubcategory.Value) = 0 Then
                MsgBox "Please select a subcategory.", vbExclamation, APP_TITLE
                Me.cboSubcategory.SetFocus
                Exit Function
            End If
            If Len(Me.cboReporter.Value) = 0 Then
                MsgBox "Please select your name.", vbExclamation, APP_TITLE
                Me.cboReporter.SetFocus
                Exit Function
            End If

        Case 1  ' Incident Details
            If Len(Trim(Me.txtTitle.Value)) = 0 Then
                MsgBox "Please enter an incident title.", vbExclamation, APP_TITLE
                Me.txtTitle.SetFocus
                Exit Function
            End If
            If Len(Trim(Me.txtDescription.Value)) = 0 Then
                MsgBox "Please enter a description.", vbExclamation, APP_TITLE
                Me.txtDescription.SetFocus
                Exit Function
            End If

        Case 2  ' Priority Assessment
            If Len(Me.cboImpact.Value) = 0 Then
                MsgBox "Please select an impact level.", vbExclamation, APP_TITLE
                Me.cboImpact.SetFocus
                Exit Function
            End If
            If Len(Me.cboUrgency.Value) = 0 Then
                MsgBox "Please select an urgency level.", vbExclamation, APP_TITLE
                Me.cboUrgency.SetFocus
                Exit Function
            End If

        Case 3  ' Review -- always valid (read-only)
            ' No validation needed on review page
    End Select

    ValidateCurrentPage = True
    Exit Function
ErrHandler:
    modErrorHandler.HandleError "frmIncidentEntry", "ValidateCurrentPage", _
                                 Err.Number, Err.Description
    ValidateCurrentPage = False
End Function

' ----------------------------------------------------------------------------
' PopulateReviewPage
' Called when navigating TO page 3 (Review). Sets all lblReview* captions
' with a summary of the data entered on the previous pages.
' ----------------------------------------------------------------------------
Private Sub PopulateReviewPage()
    On Error GoTo ErrHandler

    Me.lblReviewCategory.Caption = "Category: " & Me.cboCategory.Value
    Me.lblReviewSubcategory.Caption = "Subcategory: " & Me.cboSubcategory.Value
    Me.lblReviewReporter.Caption = "Reported By: " & Me.cboReporter.Value
    Me.lblReviewTitle.Caption = "Title: " & Me.txtTitle.Value
    Me.lblReviewDescription.Caption = "Description: " & Left(Me.txtDescription.Value, 200) & _
                                       IIf(Len(Me.txtDescription.Value) > 200, "...", "")
    Me.lblReviewAttachment.Caption = "Attachment: " & _
                                      IIf(Len(Me.txtAttachment.Value) > 0, Me.txtAttachment.Value, "(none)")
    Me.lblReviewPriority.Caption = "Priority: " & Me.lblPriorityValue.Caption
    Me.lblReviewImpact.Caption = "Impact: " & Me.cboImpact.Value
    Me.lblReviewUrgency.Caption = "Urgency: " & Me.cboUrgency.Value

    Exit Sub
ErrHandler:
    modErrorHandler.HandleError "frmIncidentEntry", "PopulateReviewPage", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' UpdateNavButtons
' Updates the visibility and enabled state of navigation buttons based on
' the current wizard page, and refreshes the step indicator label.
' ----------------------------------------------------------------------------
Private Sub UpdateNavButtons()
    On Error GoTo ErrHandler

    Me.btnBack.Enabled = (Me.mpWizard.Value > 0)
    Me.btnNext.Visible = (Me.mpWizard.Value < 3)
    Me.btnSubmit.Visible = (Me.mpWizard.Value = 3)
    Me.lblStepIndicator.Caption = "Step " & (Me.mpWizard.Value + 1) & " of 4"

    Exit Sub
ErrHandler:
    modErrorHandler.HandleError "frmIncidentEntry", "UpdateNavButtons", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' btnBack_Click
' Navigates the wizard one page backward.
' ----------------------------------------------------------------------------
Private Sub btnBack_Click()
    On Error GoTo ErrHandler

    If Me.mpWizard.Value > 0 Then
        Me.mpWizard.Value = Me.mpWizard.Value - 1
    End If
    UpdateNavButtons

    Exit Sub
ErrHandler:
    modErrorHandler.HandleError "frmIncidentEntry", "btnBack_Click", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' btnNext_Click
' Validates the current page and advances the wizard one page forward.
' If navigating to page 3 (Review), populates the review summary.
' ----------------------------------------------------------------------------
Private Sub btnNext_Click()
    On Error GoTo ErrHandler

    If Not ValidateCurrentPage() Then Exit Sub

    If Me.mpWizard.Value < 3 Then
        Me.mpWizard.Value = Me.mpWizard.Value + 1
    End If

    ' If now on the review page, populate the summary
    If Me.mpWizard.Value = 3 Then
        PopulateReviewPage
    End If

    UpdateNavButtons

    Exit Sub
ErrHandler:
    modErrorHandler.HandleError "frmIncidentEntry", "btnNext_Click", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' btnSubmit_Click
' Generates a new incident ID, collects all form data, and calls
' modDataAccess.WriteIncident to save the incident to tblIncidents.
' Shows success or failure message and closes the form on success.
' ----------------------------------------------------------------------------
Private Sub btnSubmit_Click()
    On Error GoTo ErrHandler

    ' Generate the next incident ID
    Dim sID As String
    sID = modDataAccess.GetNextIncidentID()
    If Len(sID) = 0 Then
        MsgBox "Could not generate incident ID. Please try again.", _
               vbExclamation, APP_TITLE
        Exit Sub
    End If

    ' Extract numeric impact/urgency from "4 - Critical" format
    Dim lImpact As Long, lUrgency As Long
    lImpact = CLng(Left(Me.cboImpact.Value, 1))
    lUrgency = CLng(Left(Me.cboUrgency.Value, 1))

    ' Get priority from the already-calculated label
    Dim sPriority As String
    sPriority = Me.lblPriorityValue.Caption

    ' Call the data access layer to write the incident
    Dim bSuccess As Boolean
    bSuccess = modDataAccess.WriteIncident( _
        sID, _
        Trim(Me.txtTitle.Value), _
        Trim(Me.txtDescription.Value), _
        Me.cboCategory.Value, _
        Me.cboSubcategory.Value, _
        sPriority, _
        lImpact, _
        lUrgency, _
        Me.cboReporter.Value, _
        Trim(Me.txtAttachment.Value))

    If bSuccess Then
        MsgBox "Incident " & sID & " has been logged successfully.", _
               vbInformation, APP_TITLE
        Unload Me
    Else
        MsgBox "Failed to save incident. Please try again.", _
               vbExclamation, APP_TITLE
    End If

    Exit Sub
ErrHandler:
    modErrorHandler.HandleError "frmIncidentEntry", "btnSubmit_Click", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' btnCancel_Click
' Closes the form without saving. No confirmation needed.
' ----------------------------------------------------------------------------
Private Sub btnCancel_Click()
    On Error GoTo ErrHandler
    Unload Me
    Exit Sub
ErrHandler:
    modErrorHandler.HandleError "frmIncidentEntry", "btnCancel_Click", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' btnBrowse_Click
' Opens a file browser dialog for selecting an attachment file path.
' Sets the txtAttachment TextBox with the selected file path.
' ----------------------------------------------------------------------------
Private Sub btnBrowse_Click()
    On Error GoTo ErrHandler

    Dim sFile As Variant
    sFile = Application.GetOpenFilename( _
        FileFilter:="All Files (*.*), *.*", _
        Title:="Select Attachment File")

    If sFile <> False Then
        Me.txtAttachment.Value = CStr(sFile)
    End If

    Exit Sub
ErrHandler:
    modErrorHandler.HandleError "frmIncidentEntry", "btnBrowse_Click", _
                                 Err.Number, Err.Description
End Sub
