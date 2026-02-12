Attribute VB_Name = "modIncidentEntry"
Option Explicit

' ============================================================================
' Module:  modIncidentEntry
' Purpose: Entry point for the Incident Entry wizard form.
'          Provides the public macro (ShowIncidentEntryForm) that is assigned
'          to the Dashboard "Log New Incident" button shape via OnAction.
' Dependencies: frmIncidentEntry (UserForm)
'               modErrorHandler (HandleError)
' ============================================================================

' ----------------------------------------------------------------------------
' ShowIncidentEntryForm
' Launches the incident entry wizard as a modal dialog. Called by the
' Dashboard button shape's OnAction property.
' ----------------------------------------------------------------------------
Public Sub ShowIncidentEntryForm()
    On Error GoTo ErrHandler
    frmIncidentEntry.Show vbModal
    Exit Sub
ErrHandler:
    modErrorHandler.HandleError "modIncidentEntry", "ShowIncidentEntryForm", _
                                 Err.Number, Err.Description
End Sub
