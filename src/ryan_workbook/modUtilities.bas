Option Explicit

'==================================================================================
' Module: modUtilities
' Purpose: Contains general helper functions and subroutines used across the
'          SQRCT application, particularly for data validation, list handling,
'          and shared UI updates.
'==================================================================================

'----------------------------------------------------------------------------------
Public Function GetPhaseFromPrefix(txt As String) As String
'----------------------------------------------------------------------------------
' Purpose:      Finds the unique, full phase name from the master 'PHASE_LIST'.
'               Matches based on case-insensitive prefix or exact match.
' Arguments:    txt (String): The text typed by the user.
' Returns:      String: The full, correctly-cased phase name from PHASE_LIST if a
'                       unique match is found; returns an empty string ("") if no
'                       match is found or if the prefix is ambiguous (matches more
'                       than one entry).
' Assumptions:  - Named range "PHASE_LIST" exists and refers to a single column
'                 list of all valid phase names on the "Lists" sheet (or similar).
' Called By:    Workbook_SheetChange (in ThisWorkbook module)
' Location:     modUtilities Module
'----------------------------------------------------------------------------------
    Dim rngList As Range, cell As Range
    Dim candidate As String, hit As String
    Dim matchCount As Long
    Dim normalizedInput As String

    On Error GoTo ErrorHandler_GPFP

    normalizedInput = LCase$(Trim$(txt)) ' Normalize input: Lowercase and Trim
    If Len(normalizedInput) = 0 Then Exit Function ' Ignore empty input

    ' --- Get the master list range ---
    On Error Resume Next ' Handle error if named range doesn't exist
    Set rngList = ThisWorkbook.Names("PHASE_LIST").RefersToRange
    On Error GoTo ErrorHandler_GPFP ' Restore proper error handling
    If rngList Is Nothing Then
        MsgBox "Error: Named range 'PHASE_LIST' not found. Auto-complete cannot function.", vbCritical
        Exit Function
    End If

    ' --- Look for matches in the list ---
    hit = ""          ' Stores the first (and potentially only) match found
    matchCount = 0    ' Counts how many items in the list match the prefix

    For Each cell In rngList.Cells
         If Len(Trim$(CStr(cell.Value))) > 0 Then ' Ensure cell in list not empty
            candidate = LCase$(Trim$(CStr(cell.Value))) ' Normalize list value for comparison

            ' Check if input is an exact match (case-insensitive) OR a prefix match
            If candidate = normalizedInput Or Left$(candidate, Len(normalizedInput)) = normalizedInput Then
                If matchCount = 0 Then
                    ' First match found
                    hit = CStr(cell.Value) ' Store the correctly cased value from the list
                    matchCount = 1
                    ' If it was an exact match, we don't need to check further in the list.
                    If candidate = normalizedInput Then Exit For
                Else
                    ' This is the second (or more) prefix match found
                    matchCount = matchCount + 1
                    ' If we find a second match, AND the input wasn't an EXACT match
                    ' to the *first* hit we stored, then the prefix is ambiguous.
                    If LCase$(hit) <> normalizedInput Then
                         GetPhaseFromPrefix = "" ' Return empty string for ambiguous prefix
                         Exit Function
                    End If
                    ' If we get here, it means the input EXACTLY matched the first hit,
                    ' but is also a prefix of this second hit (e.g., input "A", list has "A", "Apple").
                    ' The exact match wins, so we keep the original 'hit' and continue checking
                    ' just in case there's *another* exact match later (highly unlikely).
                End If
            End If
         End If
    Next cell

    ' --- Return Result ---
    ' Return the single hit found (or "" if no matches or ambiguous prefix)
    GetPhaseFromPrefix = hit
    Exit Function

ErrorHandler_GPFP:
     MsgBox "Error #" & Err.Number & " in GetPhaseFromPrefix: " & Err.Description, vbCritical
     GetPhaseFromPrefix = "" ' Return empty on error
End Function

'----------------------------------------------------------------------------------
Public Sub AddPhaseValidation()
'----------------------------------------------------------------------------------
' Purpose:      Applies Data Validation rules (List type) to the Engagement Phase
'               columns on the main Dashboard and UserEdits sheets. This creates
'               the dropdown arrow and enforces selection from the master list.
'               Should be run ONCE during setup or if validation needs resetting.
' Arguments:    None.
' Returns:      None.
' Assumptions:  - Named range "PHASE_LIST" exists and is correctly defined.
'               - Sheet names defined in constants (DASH_SHEET, EDITS_SHEET) are correct.
'               - Column letters/numbers defined in constants are correct.
' Called By:    Manually run by developer/admin during setup.
' Location:     modUtilities Module
'----------------------------------------------------------------------------------
    Dim wsDash As Worksheet, wsEdits As Worksheet
    Dim validationFormula As String
    Const DASH_SHEET As String = "SQRCT Dashboard" ' Use actual name
    Const EDITS_SHEET As String = "UserEdits"      ' Use actual name
    Const DASH_PHASE_COL As String = "L" ' Use letter for Range object
    Const EDITS_PHASE_COL As String = "B" ' Use letter for Range object
    ' Note: Start rows are handled within Workbook_SheetChange for event firing,
    '       but applying validation to the whole column is generally robust.

    validationFormula = "=PHASE_LIST" ' The named range containing all valid phases

    On Error Resume Next ' Ignore errors if sheets don't exist
    Set wsDash = ThisWorkbook.Worksheets(DASH_SHEET)
    Set wsEdits = ThisWorkbook.Worksheets(EDITS_SHEET)
    On Error GoTo 0 ' Restore default error handling

    Application.EnableEvents = False ' Temporarily disable events during validation changes
    On Error Resume Next ' Handle errors applying validation individually

    ' --- Apply to Dashboard Sheet ---
    If Not wsDash Is Nothing Then
        Module_Dashboard.DebugLog "AddPhaseValidation", "Applying validation to " & DASH_SHEET & " Column " & DASH_PHASE_COL
        With wsDash.Columns(DASH_PHASE_COL) ' Apply to whole column
            .Validation.Delete ' Clear existing validation first
            ' Add List validation, using Stop style (VBA handler provides custom message/logic)
            .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                            Operator:=xlBetween, Formula1:=validationFormula
            If Err.Number = 0 Then ' Only set properties if Add succeeded
                .Validation.IgnoreBlank = True ' Allow blank cells
                .Validation.InCellDropdown = True ' Show dropdown arrow
                .Validation.ErrorTitle = "Invalid Phase" ' Title for Excel's potential error
                .Validation.ErrorMessage = "Please select a phase from the list or type a recognized prefix." ' Standard error message
            Else
                Module_Dashboard.DebugLog "AddPhaseValidation", "ERROR applying validation to Dashboard: " & Err.Description: Err.Clear
            End If
        End With
    Else
         Module_Dashboard.DebugLog "AddPhaseValidation", "WARNING: Sheet not found - " & DASH_SHEET
    End If
    Err.Clear ' Clear any error from Dashboard validation attempt

    ' --- Apply to UserEdits Sheet ---
    If Not wsEdits Is Nothing Then
         Module_Dashboard.DebugLog "AddPhaseValidation", "Applying validation to " & EDITS_SHEET & " Column " & EDITS_PHASE_COL
         With wsEdits.Columns(EDITS_PHASE_COL) ' Apply to whole column
            .Validation.Delete
            .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                            Operator:=xlBetween, Formula1:=validationFormula
             If Err.Number = 0 Then
                .Validation.IgnoreBlank = True
                .Validation.InCellDropdown = True
                .Validation.ErrorTitle = "Invalid Phase"
                .Validation.ErrorMessage = "Please select a phase from the list or type a recognized prefix."
             Else
                Module_Dashboard.DebugLog "AddPhaseValidation", "ERROR applying validation to UserEdits: " & Err.Description: Err.Clear
             End If
        End With
    Else
         Module_Dashboard.DebugLog "AddPhaseValidation", "WARNING: Sheet not found - " & EDITS_SHEET
    End If
    Err.Clear ' Clear any error from UserEdits validation attempt

    ' --- Finish Up ---
    On Error GoTo 0 ' Restore default error handling
    Application.EnableEvents = True
    Module_Dashboard.DebugLog "AddPhaseValidation", "Finished applying phase validation."
    MsgBox "Phase data validation rules applied to Dashboard (Col L) and UserEdits (Col B).", vbInformation

End Sub

'----------------------------------------------------------------------------------
Sub ApplyPhaseValidationToListColumn(ws As Worksheet, colLetter As String, startDataRow As Long)
'----------------------------------------------------------------------------------
' Purpose:      Re-applies the List Data Validation rule (using PHASE_LIST) to a
'               specific column on a worksheet after data has been refreshed.
'               Ensures the dropdown arrow reappears on overwritten cells.
' Arguments:    ws (Worksheet): The target worksheet.
'               colLetter (String): The letter of the column to apply validation to (e.g., "L").
'               startDataRow (Long): The first row containing data in that column.
' Returns:      None.
' Assumptions:  - Named range "PHASE_LIST" exists and is correctly defined.
'               - Column A is a reliable indicator of the last data row.
' Called By:    RefreshDashboard, ApplyViewFormatting
' Location:     modUtilities Module
'----------------------------------------------------------------------------------
    If ws Is Nothing Then Exit Sub
    If Len(colLetter) = 0 Then Exit Sub
    If startDataRow < 1 Then startDataRow = 1 ' Basic sanity check

    Dim lastRow As Long
    Dim validationRange As Range
    Dim validationFormula As String
    validationFormula = "=PHASE_LIST" ' Assumes named range exists

    ' --- Determine Range ---
    On Error Resume Next ' Handle errors getting last row
    lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).Row ' Base last row on Col A data
    If Err.Number <> 0 Or lastRow < startDataRow Then
        Module_Dashboard.DebugLog "ApplyPhaseValidationToListColumn", "No data rows found for validation on " & ws.Name & "!" & colLetter & startDataRow
        Exit Sub ' No data rows to apply validation to
    End If
    On Error GoTo 0 ' Restore error handling
    Set validationRange = ws.Range(colLetter & startDataRow & ":" & colLetter & lastRow)

    ' --- Apply Validation ---
    Module_Dashboard.DebugLog "ApplyPhaseValidationToListColumn", "Applying List validation to " & ws.Name & "!" & validationRange.Address(False, False)
    On Error Resume Next ' Handle errors applying validation
    validationRange.Validation.Delete ' Clear any existing rule first
    validationRange.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                                   Operator:=xlBetween, Formula1:=validationFormula
    If Err.Number = 0 Then ' Only set these if Add succeeded
         validationRange.Validation.IgnoreBlank = True
         validationRange.Validation.InCellDropdown = True
         validationRange.Validation.ErrorTitle = "Invalid Phase"
         validationRange.Validation.ErrorMessage = "Please select a phase from the list or type a recognized prefix."
    Else
         Module_Dashboard.DebugLog "ApplyPhaseValidationToListColumn", "ERROR applying validation rule to " & ws.Name & "!" & validationRange.Address(False, False) & ". Error: " & Err.Description
         Err.Clear ' Clear the error
    End If
    On Error GoTo 0 ' Restore default error handling

    Set validationRange = Nothing

End Sub

'----------------------------------------------------------------------------------
Public Function GetDataRowCount(ws As Worksheet) As Long
'----------------------------------------------------------------------------------
' Purpose:      Calculates the number of actual data rows on a given worksheet.
' Arguments:    ws (Worksheet): The worksheet object to check.
' Returns:      Long: The count of data rows found. Returns 0 if sheet is invalid,
'               no data exists, or last row is within header rows.
' Assumptions:  - Data starts on Row 4 on Dashboard/Active/Archive sheets.
'               - Headers occupy Rows 1 through 3 on these sheets.
'               - Column A is a reliable indicator of the last used data row.
' Called By:    UpdateAllViewCounts
' Location:     modUtilities Module
'----------------------------------------------------------------------------------
    Dim lastRow As Long
    Const HEADER_ROWS As Long = 3 ' Number of rows before data starts

    ' --- Input Validation ---
    If ws Is Nothing Then
        DebugLog "GetDataRowCount", "ERROR: Worksheet object provided was Nothing."
        GetDataRowCount = 0 ' Return 0 if no valid sheet provided
        Exit Function
    End If

    ' --- Find Last Row ---
    On Error Resume Next ' Handle potential errors if sheet is empty or protected
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If Err.Number <> 0 Then
        DebugLog "GetDataRowCount", "ERROR finding last row on '" & ws.Name & "'. Err: " & Err.Description
        lastRow = 0 ' Reset lastRow if error occurred
        Err.Clear
    End If
    On Error GoTo 0 ' Restore default error handling

    ' --- Calculate Data Row Count ---
    If lastRow <= HEADER_ROWS Then
        ' If last used row is within header rows, there are no data rows
        GetDataRowCount = 0
    Else
        ' Subtract the header rows from the last used row number
        GetDataRowCount = lastRow - HEADER_ROWS
    End If

    ' --- Logging ---
    DebugLog "GetDataRowCount", "Sheet '" & ws.Name & "' has " & GetDataRowCount & " data rows (LastRow in Col A = " & lastRow & ")."

End Function

'----------------------------------------------------------------------------------
Public Sub UpdateAllViewCounts()
'----------------------------------------------------------------------------------
' Purpose:      Calculates the data row counts for the main Dashboard, Active,
'               and Archive views and updates the corresponding labels in Row 2
'               (Cells J2, K2, L2) on all three sheets consistently.
' Arguments:    None.
' Returns:      None.
' Assumptions:  - Public Constants for sheet names (DASHBOARD_SHEET_NAME in Module_Dashboard,
'                 SH_ACTIVE, SH_ARCHIVE in modArchival) are defined and accessible.
'               - Public Constant PW_WORKBOOK in Module_Dashboard is defined for protection.
'               - GetDataRowCount function exists and is accessible (expected in this module).
'               - Target cells J2, K2, L2 exist on all three sheets.
' Called By:    RefreshDashboard (typically near the end)
' Location:     modUtilities Module
'----------------------------------------------------------------------------------
    Dim wsDash As Worksheet, wsAct As Worksheet, wsArc As Worksheet
    Dim cntDash As Long, cntAct As Long, cntArc As Long
    Dim ws As Worksheet ' Loop variable
    Const TOP_ROW As Long = 2 ' Row where count labels are placed
    Const DASH_COUNT_COL As String = "J" ' Column for Dashboard count
    Const ACT_COUNT_COL As String = "K"  ' Column for Active count
    Const ARC_COUNT_COL As String = "L"  ' Column for Archive count

    On Error GoTo CountErrorHandler

    DebugLog "UpdateAllViewCounts", "Starting count update process..."

    ' --- Get Sheet Objects ---
    On Error Resume Next ' Temporarily ignore errors if a sheet doesn't exist
    Set wsDash = ThisWorkbook.Worksheets(Module_Dashboard.DASHBOARD_SHEET_NAME)
    Set wsAct = ThisWorkbook.Worksheets(modArchival.SH_ACTIVE)   ' Requires SH_ACTIVE to be Public Const in modArchival
    Set wsArc = ThisWorkbook.Worksheets(modArchival.SH_ARCHIVE)  ' Requires SH_ARCHIVE to be Public Const in modArchival
    On Error GoTo CountErrorHandler ' Restore proper error handling

    ' --- Calculate Counts using Helper Function ---
    If Not wsDash Is Nothing Then cntDash = GetDataRowCount(wsDash) Else cntDash = -1
    If Not wsAct Is Nothing Then cntAct = GetDataRowCount(wsAct) Else cntAct = -1
    If Not wsArc Is Nothing Then cntArc = GetDataRowCount(wsArc) Else cntArc = -1
    DebugLog "UpdateAllViewCounts", "Counts Calculated: Dashboard=" & cntDash & ", Active=" & cntAct & ", Archive=" & cntArc

    ' --- Update Labels on All Three Sheets ---
    DebugLog "UpdateAllViewCounts", "Updating Row 2 count labels on relevant sheets..."
    Application.EnableEvents = False ' Prevent triggering sheet change events

    For Each ws In Array(wsDash, wsAct, wsArc) ' Loop through the sheet objects
        If Not ws Is Nothing Then ' Proceed only if the sheet object exists
            On Error Resume Next ' Handle errors during unprotect/write/protect for each sheet

            ws.Unprotect Password:=Module_Dashboard.PW_WORKBOOK ' Assumes PW_WORKBOOK is Public Const
            If Err.Number <> 0 Then DebugLog "UpdateAllViewCounts", "Warning: Could not unprotect '" & ws.Name & "' (Err#" & Err.Number & ")": Err.Clear

            ' Write the counts
            ws.Range(DASH_COUNT_COL & TOP_ROW).Value = "Dashboard: " & IIf(cntDash = -1, "ERR", cntDash)
            ws.Range(ACT_COUNT_COL & TOP_ROW).Value = "Active: " & IIf(cntAct = -1, "ERR", cntAct)
            ws.Range(ARC_COUNT_COL & TOP_ROW).Value = "Archive: " & IIf(cntArc = -1, "ERR", cntArc)

            ' Format the count labels
            With ws.Range(DASH_COUNT_COL & TOP_ROW & ":" & ARC_COUNT_COL & TOP_ROW) ' e.g., J2:L2
                .Font.Size = 9
                .Font.Italic = True
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlCenter
                .NumberFormat = "@" ' Treat as Text
            End With

            ' Re-protect the sheet
            ws.Protect Password:=Module_Dashboard.PW_WORKBOOK, DrawingObjects:=True, Contents:=True, Scenarios:=True
            If Err.Number <> 0 Then DebugLog "UpdateAllViewCounts", "Warning: Could not re-protect '" & ws.Name & "' (Err#" & Err.Number & ")": Err.Clear

            On Error GoTo CountErrorHandler ' Restore main error handler after handling sheet-specific errors
        Else
             DebugLog "UpdateAllViewCounts", "Skipping count update for a sheet object that was Nothing."
        End If
    Next ws

    Application.EnableEvents = True
    DebugLog "UpdateAllViewCounts", "Finished count update process."
    Exit Sub ' Normal exit

CountErrorHandler:
    DebugLog "UpdateAllViewCounts", "ERROR updating counts: #" & Err.Number & " - " & Err.Description & " (Line: " & Erl & ")"
    Application.EnableEvents = True ' Ensure events are always re-enabled on error exit

End Sub

