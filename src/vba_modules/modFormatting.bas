Attribute VB_Name = "modFormatting"
Option Explicit

' ============================================================================
' Module:  modFormatting
' Purpose: BRIMIS branding application across sheets, tables, and cells.
'          All color values come from modConstants (CLR_BRIMIS_*).
'          Provides conditional formatting for priority and status columns.
' Dependencies: modConstants (CLR_BRIMIS_*, SHT_*, STATUS_*, PRIORITY_*)
'               modErrorHandler (HandleError)
' ============================================================================

' ----------------------------------------------------------------------------
' ApplyBrandingToSheet
' Applies BRIMIS branding to a worksheet: header row styling (row 1) and
' sheet tab color based on the sheet name.
'
' Parameters:
'   ws - The worksheet to brand
' ----------------------------------------------------------------------------
Public Sub ApplyBrandingToSheet(ws As Worksheet)
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False

    ' Apply header row styling (row 1 = sheet title/branding bar)
    ' Use the merged area of A1 to handle different merge widths per sheet
    ' (Dashboard=A1:Z1, Incident Log=A1:AA1, Settings=A1:H1)
    Dim rngHeader As Range
    Set rngHeader = ws.Range("A1").MergeArea

    With rngHeader
        .Interior.Color = CLR_BRIMIS_DARK
        .Font.Color = CLR_BRIMIS_WHITE
        .Font.Name = "Calibri"
        .Font.Size = 14
        .Font.Bold = True
    End With

    ' Apply sheet tab color based on sheet name
    Select Case ws.Name
        Case SHT_DASHBOARD
            ws.Tab.Color = CLR_BRIMIS_RED
        Case SHT_INCIDENT_LOG
            ws.Tab.Color = CLR_BRIMIS_GRAY
        Case SHT_SETTINGS
            ws.Tab.Color = CLR_BRIMIS_GRAY
        Case Else
            ws.Tab.Color = CLR_BRIMIS_GRAY
    End Select

    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    modErrorHandler.HandleError "modFormatting", "ApplyBrandingToSheet", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' ApplyTableHeaderBranding
' Styles a ListObject's header row with BRIMIS red background, white bold text,
' and enables row stripes for readability.
'
' Parameters:
'   tbl - The ListObject (Excel Table) to brand
' ----------------------------------------------------------------------------
Public Sub ApplyTableHeaderBranding(tbl As ListObject)
    On Error GoTo ErrHandler

    ' Style the table header row
    With tbl.HeaderRowRange
        .Interior.Color = CLR_BRIMIS_RED
        .Font.Color = CLR_BRIMIS_WHITE
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
    End With

    ' Enable alternating row stripes for readability
    tbl.ShowTableStyleRowStripes = True

    Exit Sub

ErrHandler:
    modErrorHandler.HandleError "modFormatting", "ApplyTableHeaderBranding", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' ApplyPriorityConditionalFormatting
' Applies conditional formatting rules for priority values P1-P4 on a range.
' Clears any existing format conditions on the range first.
'
' Parameters:
'   rng - The range to apply priority formatting to (typically the Priority column)
' ----------------------------------------------------------------------------
Public Sub ApplyPriorityConditionalFormatting(rng As Range)
    On Error GoTo ErrHandler

    ' Clear existing rules on this range
    rng.FormatConditions.Delete

    ' P1 - Critical: BRIMIS Red background, white bold text
    With rng.FormatConditions.Add(xlCellValue, xlEqual, "=""P1""")
        .Interior.Color = CLR_BRIMIS_RED
        .Font.Color = CLR_BRIMIS_WHITE
        .Font.Bold = True
    End With

    ' P2 - High: Orange background, white bold text
    With rng.FormatConditions.Add(xlCellValue, xlEqual, "=""P2""")
        .Interior.Color = RGB(255, 140, 0)
        .Font.Color = CLR_BRIMIS_WHITE
        .Font.Bold = True
    End With

    ' P3 - Medium: Gold background, dark text
    With rng.FormatConditions.Add(xlCellValue, xlEqual, "=""P3""")
        .Interior.Color = RGB(255, 215, 0)
        .Font.Color = CLR_BRIMIS_DARK
    End With

    ' P4 - Low: BRIMIS Gray background, white text
    With rng.FormatConditions.Add(xlCellValue, xlEqual, "=""P4""")
        .Interior.Color = CLR_BRIMIS_GRAY
        .Font.Color = CLR_BRIMIS_WHITE
    End With

    Exit Sub

ErrHandler:
    modErrorHandler.HandleError "modFormatting", "ApplyPriorityConditionalFormatting", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' ApplyStatusConditionalFormatting
' Applies conditional formatting rules for all 7 status values on a range.
' Clears any existing format conditions on the range first.
'
' Parameters:
'   rng - The range to apply status formatting to (typically the Status column)
' ----------------------------------------------------------------------------
Public Sub ApplyStatusConditionalFormatting(rng As Range)
    On Error GoTo ErrHandler

    ' Clear existing rules on this range
    rng.FormatConditions.Delete

    ' Open - White background, dark bold text
    With rng.FormatConditions.Add(xlCellValue, xlEqual, "=""Open""")
        .Interior.Color = CLR_BRIMIS_WHITE
        .Font.Color = CLR_BRIMIS_DARK
        .Font.Bold = True
    End With

    ' Assigned - Light blue background, dark text
    With rng.FormatConditions.Add(xlCellValue, xlEqual, "=""Assigned""")
        .Interior.Color = RGB(173, 216, 230)
        .Font.Color = CLR_BRIMIS_DARK
    End With

    ' In Progress - Light amber background, dark text
    With rng.FormatConditions.Add(xlCellValue, xlEqual, "=""In Progress""")
        .Interior.Color = RGB(255, 235, 156)
        .Font.Color = CLR_BRIMIS_DARK
    End With

    ' Resolved - Light green background, dark text
    With rng.FormatConditions.Add(xlCellValue, xlEqual, "=""Resolved""")
        .Interior.Color = RGB(198, 239, 206)
        .Font.Color = CLR_BRIMIS_DARK
    End With

    ' Closed - Light gray background, gray text
    With rng.FormatConditions.Add(xlCellValue, xlEqual, "=""Closed""")
        .Interior.Color = RGB(217, 217, 217)
        .Font.Color = CLR_BRIMIS_GRAY
    End With

    ' Cancelled - Light red background, dark text
    With rng.FormatConditions.Add(xlCellValue, xlEqual, "=""Cancelled""")
        .Interior.Color = RGB(255, 199, 206)
        .Font.Color = CLR_BRIMIS_DARK
    End With

    ' Duplicate - Light gray background, gray italic text
    With rng.FormatConditions.Add(xlCellValue, xlEqual, "=""Duplicate""")
        .Interior.Color = RGB(217, 217, 217)
        .Font.Color = CLR_BRIMIS_GRAY
        .Font.Italic = True
    End With

    Exit Sub

ErrHandler:
    modErrorHandler.HandleError "modFormatting", "ApplyStatusConditionalFormatting", _
                                 Err.Number, Err.Description
End Sub

' ----------------------------------------------------------------------------
' ApplyAllBranding
' Convenience wrapper that applies BRIMIS branding to all visible worksheets
' and all ListObjects found on each sheet.
' ----------------------------------------------------------------------------
Public Sub ApplyAllBranding()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim tbl As ListObject

    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Worksheets
        ' Apply sheet-level branding (header row + tab color)
        ApplyBrandingToSheet ws

        ' Apply table header branding to every ListObject on this sheet
        For Each tbl In ws.ListObjects
            ApplyTableHeaderBranding tbl
        Next tbl
    Next ws

    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    modErrorHandler.HandleError "modFormatting", "ApplyAllBranding", _
                                 Err.Number, Err.Description
End Sub
