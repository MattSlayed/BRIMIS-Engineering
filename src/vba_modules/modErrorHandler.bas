Attribute VB_Name = "modErrorHandler"
Option Explicit

' ============================================================================
' Module:  modErrorHandler
' Purpose: Centralized error handling for the entire application.
'          Every Public Sub/Function calls HandleError from its ErrHandler
'          label. This prevents the Debug/End dialog from ever appearing
'          and ensures Application state is always cleaned up.
' ============================================================================

' ----------------------------------------------------------------------------
' HandleError
' Called from every procedure's ErrHandler label. Logs the error, shows a
' user-friendly message, and resets Application state.
'
' Parameters:
'   sModule  - Name of the calling module (e.g., "modDataAccess")
'   sProc    - Name of the calling procedure (e.g., "WriteIncident")
'   lErrNum  - Err.Number from the runtime error
'   sErrDesc - Err.Description from the runtime error
' ----------------------------------------------------------------------------
Public Sub HandleError(ByVal sModule As String, ByVal sProc As String, _
                       ByVal lErrNum As Long, ByVal sErrDesc As String)

    ' Step 1: Log to Immediate window for development/debugging
    Debug.Print Now & " | " & sModule & "." & sProc & _
                " | Error " & lErrNum & ": " & sErrDesc

    ' Step 2: Build user-friendly message (never show error numbers or VBA jargon)
    Dim sUserMsg As String
    sUserMsg = "Something went wrong while performing this operation." & vbCrLf & vbCrLf & _
               "What happened: " & GetFriendlyMessage(lErrNum, sErrDesc) & vbCrLf & vbCrLf & _
               "What to do: Please try again. If the problem persists, " & _
               "contact your system administrator." & vbCrLf & _
               "Reference: " & sModule & "." & sProc

    ' Step 3: Show message box with warning icon
    MsgBox sUserMsg, vbExclamation, APP_TITLE & " - Error"

    ' Step 4: Clean up Application state to prevent lingering side effects
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    On Error GoTo 0

End Sub

' ----------------------------------------------------------------------------
' GetFriendlyMessage
' Maps common VBA runtime error numbers to plain-English descriptions that
' make sense to non-technical users.
'
' Parameters:
'   lErrNum  - The VBA error number
'   sErrDesc - The original VBA error description (fallback)
'
' Returns:
'   A user-friendly string describing the error
' ----------------------------------------------------------------------------
Private Function GetFriendlyMessage(ByVal lErrNum As Long, _
                                     ByVal sErrDesc As String) As String
    Select Case lErrNum
        Case 9
            GetFriendlyMessage = "A data reference was out of range. " & _
                                  "The expected data may have been moved or deleted."
        Case 13
            GetFriendlyMessage = "A value was not in the expected format " & _
                                  "(e.g., text where a number was expected)."
        Case 53
            GetFriendlyMessage = "A required file could not be found. " & _
                                  "It may have been moved or deleted."
        Case 70
            GetFriendlyMessage = "Permission was denied when trying to access a file. " & _
                                  "The file may be open in another program or read-only."
        Case 76
            GetFriendlyMessage = "The specified folder path could not be found. " & _
                                  "The folder may have been moved or deleted."
        Case 91
            GetFriendlyMessage = "A required component was not found. " & _
                                  "The workbook structure may need to be repaired."
        Case 1004
            GetFriendlyMessage = "An Excel operation failed. " & _
                                  "A sheet may be protected or a range may be invalid."
        Case Else
            GetFriendlyMessage = sErrDesc
    End Select
End Function
