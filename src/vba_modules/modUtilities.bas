Attribute VB_Name = "modUtilities"
Option Explicit

' ============================================================================
' Module:  modUtilities
' Purpose: Sheet protection toggle, folder management, backup path resolution,
'          and general-purpose helper functions used across the application.
' Dependencies: modConstants (SHEET_PWD, BACKUP_SUBFOLDER)
'               modErrorHandler (HandleError)
' ============================================================================

' ----------------------------------------------------------------------------
' UnprotectSheet
' Removes worksheet protection using the standard password. Safe to call
' even if the sheet is already unprotected.
'
' Parameters:
'   ws - The worksheet to unprotect
' ----------------------------------------------------------------------------
Public Sub UnprotectSheet(ws As Worksheet)
    On Error Resume Next
    If ws.ProtectContents Then
        ws.Unprotect Password:=SHEET_PWD
    End If
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' ProtectSheet
' Applies worksheet protection with UserInterfaceOnly:=True so VBA can still
' write while the UI is locked. Allows filtering and sorting for users.
' Safe to call even if the sheet is already protected.
'
' Parameters:
'   ws - The worksheet to protect
' ----------------------------------------------------------------------------
Public Sub ProtectSheet(ws As Worksheet)
    On Error Resume Next
    ws.Protect Password:=SHEET_PWD, _
        UserInterfaceOnly:=True, _
        AllowFiltering:=True, _
        AllowSorting:=True
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' EnsureFolderExists
' Checks whether a folder exists at the given path and creates it if not.
' Uses Dir() + MkDir (zero external dependencies).
'
' Parameters:
'   sPath - Full path to the folder to ensure exists
'
' Returns:
'   True if the folder exists or was created successfully, False on failure
' ----------------------------------------------------------------------------
Public Function EnsureFolderExists(ByVal sPath As String) As Boolean
    On Error GoTo ErrHandler

    If Len(Dir(sPath, vbDirectory)) = 0 Then
        MkDir sPath
    End If

    EnsureFolderExists = True
    Exit Function

ErrHandler:
    modErrorHandler.HandleError "modUtilities", "EnsureFolderExists", _
                                 Err.Number, Err.Description
    EnsureFolderExists = False
End Function

' ----------------------------------------------------------------------------
' GetLocalBackupPath
' Returns the local filesystem path for storing backup copies of the workbook.
' First tries a Backups subfolder next to the workbook. If the workbook is
' stored on OneDrive/SharePoint (URL-style path), falls back to the user's
' Documents folder.
'
' Returns:
'   Full path string to the backup folder
' ----------------------------------------------------------------------------
Public Function GetLocalBackupPath() As String
    On Error GoTo ErrHandler

    Dim sBasePath As String
    sBasePath = ThisWorkbook.Path

    ' Check for OneDrive/SharePoint URL paths
    If IsCloudPath(sBasePath) Then
        ' Fall back to user's Documents folder
        GetLocalBackupPath = Environ("USERPROFILE") & "\Documents\BRIMIS_Backups"
    Else
        GetLocalBackupPath = sBasePath & "\" & BACKUP_SUBFOLDER
    End If

    Exit Function

ErrHandler:
    modErrorHandler.HandleError "modUtilities", "GetLocalBackupPath", _
                                 Err.Number, Err.Description
    ' Fallback on error: use user Documents
    GetLocalBackupPath = Environ("USERPROFILE") & "\Documents\BRIMIS_Backups"
End Function

' ----------------------------------------------------------------------------
' IsCloudPath
' Checks whether a path is a cloud/URL path (OneDrive or SharePoint).
'
' Parameters:
'   sPath - The path to check
'
' Returns:
'   True if the path starts with "http://" or "https://"
' ----------------------------------------------------------------------------
Public Function IsCloudPath(ByVal sPath As String) As Boolean
    IsCloudPath = (Left(LCase(sPath), 7) = "http://" Or _
                   Left(LCase(sPath), 8) = "https://")
End Function
