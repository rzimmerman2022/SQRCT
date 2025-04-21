Option Explicit

'=====================================================================
' Module :  modFormatting
' Purpose: Contains functions for applying and verifying sheet formatting,
'          focusing on EXACT replication of the main Dashboard layout.
' REVISED: 04/20/2025 - Created module. Added ExactlyCloneDashboardFormatting
'                      to replicate Dashboard Row 1/2 style via PasteSpecial
'                      and explicit overrides. Added verification helpers.
'                      Corrected A2 style to steel blue. Added O2+ clear.
'                      Adjusted unlock/button/count call order.
'=====================================================================

Public Sub ExactlyCloneDashboardFormatting(targetSheet As Worksheet, viewType As String)
    ' Purpose: Create an EXACT clone of the Dashboard Row 2 formatting
    ' viewType can be "Active" or "Archive" to determine the title text only

    Dim sourceDashboard As Worksheet

    On Error Resume Next
    Set sourceDashboard = ThisWorkbook.Worksheets(Module_Dashboard.DASHBOARD_SHEET_NAME)
    If Err.Number <> 0 Or sourceDashboard Is Nothing Then
        MsgBox "Cannot find source Dashboard sheet '" & Module_Dashboard.DASHBOARD_SHEET_NAME & "' to clone formatting from!", vbCritical
        Exit Sub
    End If
    On Error GoTo 0 ' Restore default handler for the rest of the sub

    ' Ensure sheets are unprotected for formatting operations
    On Error Resume Next
    sourceDashboard.Unprotect Password:=Module_Dashboard.PW_WORKBOOK
    targetSheet.Unprotect Password:=Module_Dashboard.PW_WORKBOOK
    If Err.Number <> 0 Then
        Module_Dashboard.DebugLog "ExactlyCloneDashboardFormatting", "Warning: Failed to unprotect source or target sheet. Err=" & Err.Number
        Err.Clear
    End If
    On Error GoTo CloneErrorHandler ' Use specific handler for cloning process

    Module_Dashboard.DebugLog "ExactlyCloneDashboardFormatting", "Beginning exact replication for " & targetSheet.Name

    ' --- Step 1: Copy Row Height from Dashboard ---
    targetSheet.Rows(2).RowHeight = sourceDashboard.Rows(2).RowHeight
    Module_Dashboard.DebugLog "ExactlyCloneDashboardFormatting", "Copied Row 2 height: " & targetSheet.Rows(2).RowHeight

    ' --- Step 2: Copy Column Widths from Dashboard to match exactly ---
    Dim col As Range
    For Each col In sourceDashboard.Range("A:N").Columns
        targetSheet.Columns(col.Column).ColumnWidth = col.ColumnWidth
    Next col
    Module_Dashboard.DebugLog "ExactlyCloneDashboardFormatting", "Copied column widths A:N"

    ' --- Step 3: EXACT copy of Row 2 formatting (using PasteSpecial) ---
    ' First clear any existing formatting or content on target to prevent issues
    targetSheet.Range("A2:N2").ClearContents
    targetSheet.Range("A2:N2").ClearFormats

    ' Then copy ONLY formats from the source Dashboard's Row 2
    sourceDashboard.Range("A2:N2").Copy
    targetSheet.Range("A2").PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    Module_Dashboard.DebugLog "ExactlyCloneDashboardFormatting", "Copied exact Row 2 formatting via PasteSpecial"

    ' --- Step 4: Set A2 "CONTROL PANEL" text and specific formatting ---
    ' Explicitly set A2 style AFTER PasteSpecial to ensure desired look.
    With targetSheet.Range("A2")
        .Value = "CONTROL PANEL" ' Set text
        .Font.Bold = True
        .Font.Size = 10
        .Font.Name = "Segoe UI"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(70, 130, 180) ' Steel blue - CORRECTED based on original SetupDashboard
        .Font.Color = RGB(255, 255, 255)   ' White text - CORRECTED based on original SetupDashboard
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlEdgeRight).Color = RGB(200, 200, 200)
    End With
    Module_Dashboard.DebugLog "ExactlyCloneDashboardFormatting", "Set A2 cell text and specific formatting"

    ' --- Step 5: Set custom title with the correct view suffix ---
    Dim titleText As String
    If viewType = "Active" Then
        titleText = "STRATEGIC QUOTE RECOVERY & CONVERSION TRACKER – ACTIVE VIEW"
    ElseIf viewType = "Archive" Then
        titleText = "STRATEGIC QUOTE RECOVERY & CONVERSION TRACKER – ARCHIVE VIEW"
    Else ' Default or Dashboard case
        titleText = "STRATEGIC QUOTE RECOVERY & CONVERSION TRACKER"
    End If

    With targetSheet.Range("A1:N1")
        If .MergeCells Then .UnMerge
        .ClearContents ' Clear before merging
        .Merge
        .Value = titleText
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Size = 18
        .Font.Bold = True
        ' Set the color based on view type
        If viewType = "Active" Then
            .Interior.Color = RGB(0, 110, 0) ' Dark Green for Active
        ElseIf viewType = "Archive" Then
            .Interior.Color = RGB(150, 40, 40) ' Dark Red for Archive
        Else ' Default or Dashboard case
            .Interior.Color = RGB(16, 107, 193) ' Blue for main Dashboard
        End If
        .Font.Color = RGB(255, 255, 255) ' White text
        .RowHeight = 32
    End With
    Module_Dashboard.DebugLog "ExactlyCloneDashboardFormatting", "Set title banner with text: " & titleText

    ' --- Step 6: Ensure Row 2 Light Grey Background (B2:N2) ---
    ' Explicitly set B2:N2 background AFTER PasteSpecial and A2 formatting.
    With targetSheet.Range("B2:N2")
        .Interior.Color = RGB(245, 245, 245) ' Light grey background
    End With
    Module_Dashboard.DebugLog "ExactlyCloneDashboardFormatting", "Set B2:N2 background to light grey"

    ' --- Step 6b: Explicitly clear background beyond N2 ---
    ' Prevents color bleed on Active/Archive sheets.
    On Error Resume Next ' Ignore error if columns don't exist or other issues
    targetSheet.Range("O2:" & targetSheet.Columns(targetSheet.Columns.Count).Address).Interior.ColorIndex = xlNone ' Clear background from O2 to end
    If Err.Number <> 0 Then Module_Dashboard.DebugLog "ExactlyCloneDashboardFormatting", "Warning: Could not clear background O2 onwards. Err=" & Err.Number: Err.Clear
    On Error GoTo CloneErrorHandler ' Restore handler
    Module_Dashboard.DebugLog "ExactlyCloneDashboardFormatting", "Cleared background color from O2 onwards"

    ' --- Step 7: Unlock Count Cells (BEFORE adding buttons) ---
    ' Unlock J2:L2 so UpdateAllViewCounts can write to them later.
    ' Moved BEFORE AddNavigationButtons as that sub used to re-protect prematurely.
    Module_Dashboard.DebugLog "ExactlyCloneDashboardFormatting", "Unlocking J2:L2..."
    On Error Resume Next ' Handle error unlocking
    targetSheet.Range("J2:L2").Locked = False
    If Err.Number <> 0 Then Module_Dashboard.DebugLog "ExactlyCloneDashboardFormatting", "Warning: Failed to unlock J2:L2. Err=" & Err.Number: Err.Clear
    On Error GoTo CloneErrorHandler ' Restore handler

    ' --- Step 8: Add Navigation Buttons (always identical) ---
    ' Calls the standard button creation routine. Assumes it's Public in modArchival.
    ' NOTE: AddNavigationButtons should NOT protect the sheet itself.
    modArchival.AddNavigationButtons targetSheet
    Module_Dashboard.DebugLog "ExactlyCloneDashboardFormatting", "Added navigation buttons"

    ' --- Step 9: Update counts display ---
    ' Calls the standard count display routine. Cells J2:L2 were unlocked in Step 7.
    modUtilities.UpdateAllViewCounts targetSheet
    Module_Dashboard.DebugLog "ExactlyCloneDashboardFormatting", "Updated count display"

    ' --- Step 10: Take a direct snapshot copy of all Row 2 colors (Verification Log) ---
    ' Commented out 04/20/2025 - Too verbose for standard logging.
    ' Dim cellColorLog As String
    ' cellColorLog = "Row 2 Colors: "
    ' Dim c As Range
    ' For Each c In targetSheet.Range("A2:N2").Cells
    '     cellColorLog = cellColorLog & " | " & c.Address(False, False) & ":" & c.Interior.Color
    ' Next c
    ' Module_Dashboard.DebugLog "ExactlyCloneDashboardFormatting", cellColorLog

    Module_Dashboard.DebugLog "ExactlyCloneDashboardFormatting", "Completed exact Dashboard formatting clone for " & targetSheet.Name
    Exit Sub ' Normal Exit

CloneErrorHandler:
    Module_Dashboard.DebugLog "ExactlyCloneDashboardFormatting", "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    Module_Dashboard.DebugLog "ExactlyCloneDashboardFormatting", "ERROR Handler! Sheet='" & targetSheet.Name & "'. Err=" & Err.Number & ": " & Err.Description & " near line " & Erl
    Module_Dashboard.DebugLog "ExactlyCloneDashboardFormatting", "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    Application.CutCopyMode = False ' Ensure clipboard is cleared on error
    ' Note: Protection is handled by the caller (ApplyViewFormatting)
End Sub


' --- Verification and Debugging Helpers ---

Public Sub VerifyExactFormatting()
    ' Purpose: Visually verify that formatting is 100% identical across all sheets

    Dim wsDash As Worksheet
    Dim wsActive As Worksheet
    Dim wsArchive As Worksheet
    Dim msg As String

    On Error Resume Next
    Set wsDash = ThisWorkbook.Worksheets(Module_Dashboard.DASHBOARD_SHEET_NAME)
    Set wsActive = ThisWorkbook.Worksheets(modArchival.SH_ACTIVE)
    Set wsArchive = ThisWorkbook.Worksheets(modArchival.SH_ARCHIVE)
    On Error GoTo 0

    If wsDash Is Nothing Or wsActive Is Nothing Or wsArchive Is Nothing Then
        MsgBox "One or more sheets not found (Dashboard, Active, or Archive)! Cannot verify.", vbCritical
        Exit Sub
    End If

    msg = "FORMATTING VERIFICATION REPORT" & vbCrLf & vbCrLf

    ' Compare Row 2 background colors
    msg = msg & "Row 2 Background Colors:" & vbCrLf
    Dim col As Long, addr As String
    For col = 1 To 14 ' A through N
        addr = Chr(64 + col) & "2" ' Convert column number to letter + "2"
        msg = msg & addr & ": " & vbCrLf
        msg = msg & "  Dashboard: " & wsDash.Range(addr).Interior.Color & vbCrLf
        msg = msg & "  Active:    " & wsActive.Range(addr).Interior.Color & vbCrLf
        msg = msg & "  Archive:   " & wsArchive.Range(addr).Interior.Color & vbCrLf & vbCrLf
    Next col

    ' Compare Row Heights
    msg = msg & "Row 2 Heights:" & vbCrLf
    msg = msg & "  Dashboard: " & wsDash.Rows(2).RowHeight & vbCrLf
    msg = msg & "  Active:    " & wsActive.Rows(2).RowHeight & vbCrLf
    msg = msg & "  Archive:   " & wsArchive.Rows(2).RowHeight & vbCrLf & vbCrLf

    ' Compare A2 formatting in detail
    msg = msg & "A2 Cell Detail:" & vbCrLf
    msg = msg & "  Dashboard: " & FormatCellDetails(wsDash.Range("A2")) & vbCrLf
    msg = msg & "  Active:    " & FormatCellDetails(wsActive.Range("A2")) & vbCrLf
    msg = msg & "  Archive:   " & FormatCellDetails(wsArchive.Range("A2")) & vbCrLf & vbCrLf

    MsgBox msg, vbInformation, "Format Verification Report"
End Sub

Private Function FormatCellDetails(rng As Range) As String
    ' Helper to gather all formatting details of a cell
    On Error Resume Next
    Dim s As String
    s = "BGColor=" & rng.Interior.Color & ", "
    s = s & "Font=" & rng.Font.Name & ", "
    s = s & "Size=" & rng.Font.Size & ", "
    s = s & "Bold=" & rng.Font.Bold & ", "
    s = s & "TextColor=" & rng.Font.Color
    If Err.Number <> 0 Then FormatCellDetails = "Error reading details" Else FormatCellDetails = s
    On Error GoTo 0
End Function

Public Sub LogAllSheetFormatting()
    ' Purpose: Logs detailed formatting for all sheets to help debug any differences

    Dim wsDash As Worksheet
    Dim wsActive As Worksheet
    Dim wsArchive As Worksheet

    On Error Resume Next
    Set wsDash = ThisWorkbook.Worksheets(Module_Dashboard.DASHBOARD_SHEET_NAME)
    Set wsActive = ThisWorkbook.Worksheets(modArchival.SH_ACTIVE)
    Set wsArchive = ThisWorkbook.Worksheets(modArchival.SH_ARCHIVE)
    On Error GoTo 0

    Debug.Print String(80, "=")
    Debug.Print "START FORMATTING LOG DUMP (" & Now() & ")"
    Debug.Print String(80, "=")

    If Not wsDash Is Nothing Then LogSheetFormatting wsDash, "Dashboard"
    If Not wsActive Is Nothing Then LogSheetFormatting wsActive, "Active"
    If Not wsArchive Is Nothing Then LogSheetFormatting wsArchive, "Archive"

    Debug.Print String(80, "=")
    Debug.Print "END FORMATTING LOG DUMP"
    Debug.Print String(80, "=") & vbCrLf

    MsgBox "Format logging complete. Check the Immediate Window (Ctrl+G).", vbInformation
End Sub

Private Sub LogSheetFormatting(ws As Worksheet, sheetType As String)
    ' Helper to log detailed formatting for a specific sheet
    On Error Resume Next ' Prevent errors reading properties from halting log
    
    Debug.Print "---- " & sheetType & " SHEET: " & ws.Name & " ----"
    Debug.Print "ROW 2 FORMATS:"

    Dim c As Range
    For Each c In ws.Range("A2:N2").Cells
        Debug.Print "  " & c.Address(False, False) & ": " & _
            "BGColor=" & c.Interior.Color & ", " & _
            "Content='" & c.Value & "', " & _
            "Font=" & c.Font.Name & ", " & _
            "Size=" & c.Font.Size & ", " & _
            "Bold=" & c.Font.Bold & ", " & _
            "TextColor=" & c.Font.Color
    Next c

    Debug.Print "ROW HEIGHTS:"
    Debug.Print "  Row 1: " & ws.Rows(1).RowHeight
    Debug.Print "  Row 2: " & ws.Rows(2).RowHeight
    Debug.Print "  Row 3: " & ws.Rows(3).RowHeight

    Debug.Print "COLUMN WIDTHS (A:N):"
    Dim col As Long
    For col = 1 To 14
        Debug.Print "  " & Chr(64 + col) & ": " & ws.Columns(col).ColumnWidth
    Next col

    Debug.Print "----------------------------------------" & vbCrLf
    On Error GoTo 0
End Sub
