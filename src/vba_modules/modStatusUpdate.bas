Attribute VB_Name = "modStatusUpdate"
Option Explicit

' ============================================================================
' Module:  modStatusUpdate
' Purpose: Entry point for the Status Update form.
'          Provides the public macro (ShowStatusUpdateForm) that is assigned
'          to the Dashboard "Update Status" button shape via OnAction.
' Dependencies: frmStatusUpdate (UserForm)
'               modErrorHandler (HandleError)
' ============================================================================

' ----------------------------------------------------------------------------
' ShowStatusUpdateForm
' Launches the status update form as a modal dialog.
' Called by the Dashboard "Update Status" button shape's OnAction property.
' ----------------------------------------------------------------------------
Public Sub ShowStatusUpdateForm()
    On Error GoTo ErrHandler
    frmStatusUpdate.Show vbModal
    Exit Sub
ErrHandler:
    modErrorHandler.HandleError "modStatusUpdate", "ShowStatusUpdateForm", _
                                 Err.Number, Err.Description
End Sub
