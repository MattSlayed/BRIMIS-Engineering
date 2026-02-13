Attribute VB_Name = "modSearch"
Option Explicit

' ============================================================================
' Module:  modSearch
' Purpose: Entry point for the Incident Search form.
'          Provides the public macro (ShowSearchForm) that is assigned
'          to the Dashboard "Search Incidents" button shape via OnAction.
' Dependencies: frmSearch (UserForm)
'               modErrorHandler (HandleError)
' ============================================================================

' ----------------------------------------------------------------------------
' ShowSearchForm
' Launches the incident search form as a modal dialog. Called by the
' Dashboard button shape's OnAction property.
' ----------------------------------------------------------------------------
Public Sub ShowSearchForm()
    On Error GoTo ErrHandler
    frmSearch.Show vbModal
    Exit Sub
ErrHandler:
    modErrorHandler.HandleError "modSearch", "ShowSearchForm", _
                                 Err.Number, Err.Description
End Sub
