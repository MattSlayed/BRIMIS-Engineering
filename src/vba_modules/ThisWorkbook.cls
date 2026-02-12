Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' ============================================================================
' Class:   ThisWorkbook
' Purpose: Workbook-level event handlers. Workbook_Open is the entry point
'          for all automatic initialization: backup, protection, navigation.
' Dependencies: modConstants (BACKUP_PREFIX, MAX_BACKUPS)
'               modUtilities (GetLocalBackupPath, EnsureFolderExists)
'               modInitialize (InitializeWorkbook)
'               modErrorHandler (HandleError)
' ============================================================================

' ----------------------------------------------------------------------------
' Workbook_Open
' Fires every time the workbook is opened. Creates an auto-backup, then
' initializes the workbook (verify structure, apply protection, navigate
' to Dashboard).
' ----------------------------------------------------------------------------
Private Sub Workbook_Open()
    On Error GoTo ErrHandler

    ' Step 1: Create timestamped backup
    CreateAutoBackup

    ' Step 2: Initialize workbook (structure check, protection, branding, navigation)
    modInitialize.InitializeWorkbook

    Exit Sub

ErrHandler:
    modErrorHandler.HandleError "ThisWorkbook", "Workbook_Open", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' CreateAutoBackup
' Creates a timestamped backup copy of the workbook in the Backups folder.
' Backup failure is non-critical -- errors are logged but do NOT show a
' MsgBox or block workbook usage.
' ----------------------------------------------------------------------------
Private Sub CreateAutoBackup()
    On Error GoTo ErrHandler

    ' Get the backup folder path (handles OneDrive/SharePoint fallback)
    Dim sBackupFolder As String
    sBackupFolder = modUtilities.GetLocalBackupPath()

    ' Ensure the folder exists
    If Not modUtilities.EnsureFolderExists(sBackupFolder) Then
        Debug.Print "BACKUP: Could not create backup folder: " & sBackupFolder
        Exit Sub
    End If

    ' Build timestamped filename
    Dim sBackupFile As String
    sBackupFile = sBackupFolder & "\" & BACKUP_PREFIX & Format(Now, "YYYYMMDD\_HHNNSS") & ".xlsm"

    ' Save a copy
    ThisWorkbook.SaveCopyAs sBackupFile
    Debug.Print "BACKUP: Created " & sBackupFile

    ' Clean up old backups
    CleanOldBackups sBackupFolder

    Exit Sub

ErrHandler:
    ' Backup failure is non-critical -- log and continue silently
    Debug.Print "BACKUP: Failed - Error " & Err.Number & ": " & Err.Description
End Sub

' ----------------------------------------------------------------------------
' CleanOldBackups
' Removes the oldest backup files when the count exceeds MAX_BACKUPS.
' Files are sorted chronologically by their timestamp in the filename.
'
' Parameters:
'   sFolderPath - Full path to the backup folder
' ----------------------------------------------------------------------------
Private Sub CleanOldBackups(ByVal sFolderPath As String)
    On Error Resume Next

    Dim sPattern As String
    sPattern = sFolderPath & "\" & BACKUP_PREFIX & "*.xlsm"

    ' Collect all backup filenames into a Collection
    Dim colFiles As New Collection
    Dim sFile As String
    sFile = Dir(sPattern)

    Do While Len(sFile) > 0
        colFiles.Add sFile
        sFile = Dir()
    Loop

    ' If within limit, nothing to do
    If colFiles.Count <= MAX_BACKUPS Then
        Exit Sub
    End If

    ' Sort filenames (they sort chronologically due to YYYYMMDD_HHNNSS format)
    ' Simple bubble sort for small collection
    Dim arrFiles() As String
    ReDim arrFiles(1 To colFiles.Count)

    Dim i As Long
    For i = 1 To colFiles.Count
        arrFiles(i) = colFiles(i)
    Next i

    Dim j As Long
    Dim sTemp As String
    For i = 1 To UBound(arrFiles) - 1
        For j = i + 1 To UBound(arrFiles)
            If arrFiles(i) > arrFiles(j) Then
                sTemp = arrFiles(i)
                arrFiles(i) = arrFiles(j)
                arrFiles(j) = sTemp
            End If
        Next j
    Next i

    ' Delete the oldest files (those at the beginning of the sorted array)
    Dim lToDelete As Long
    lToDelete = UBound(arrFiles) - MAX_BACKUPS

    For i = 1 To lToDelete
        Kill sFolderPath & "\" & arrFiles(i)
        Debug.Print "BACKUP: Deleted old backup - " & arrFiles(i)
    Next i

End Sub
