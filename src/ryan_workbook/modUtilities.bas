Option Explicit

'=====================================================================
' Module :  modUtilities
' Purpose: Contains general utility functions used across the project,
'          including validation helpers and UI updaters.
' REVISED: 04/20/2025 - Added ApplyStandardControlRow (later removed and replaced
'                      by modFormatting.ExactlyCloneDashboardFormatting).
'                      Added logging to UpdateAllViewCounts to debug count display.
'                      Removed IndentLevel setting from UpdateAllViewCounts.
'=====================================================================

'----------------------------------------------------------------------------------
Public Function GetPhaseFromPrefix(txt As String) As String
'----------------------------------------------------------------------------------
' Purpose:      Finds the unique, full phase name from the master 'PHASE_LIST'.
'               Matches based on case-insensitive prefix or exact match. Handles
'               ambiguity by returning an empty string.
' Arguments:    txt (String): The text typed by the user into a phase cell.
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
        GetPhaseFromPrefix = "" ' Return empty string if list not found
        Exit Function
    End If

    ' --- Look for matches in the list ---
    hit = ""          ' Stores the first (and potentially only) match found
    matchCount = 0    ' Counts how many items in the list match the prefix

    For Each cell In rngList.Cells
         If Len(Trim$(CStr(cell.value))) > 0 Then ' Ensure cell in list not empty
             candidate = LCase$(Trim$(CStr(cell.value))) ' Normalize list value for comparison

             ' Check if input is an exact match (case-insensitive) OR a prefix match
             If candidate = normalizedInput Or Left$(candidate, Len(normalizedInput)) = normalizedInput Then
                 If matchCount = 0 Then
                     ' First match found
                     hit = CStr(cell.value) ' Store the correctly cased value from the list
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
                     ' The exact match wins, so we keep the original 'hit' and continue checking.
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
' Purpose:      Applies Data Validation rules (List type based on PHASE_LIST)
'               to the Engagement Phase columns on the main Dashboard (L) and
'               UserEdits (B) sheets. Creates the dropdown arrow and sets the
'               Stop alert style (VBA SheetChange handler provides user messages).
'               Should be run ONCE during setup or if validation needs resetting.
' Arguments:    None.
' Returns:      None.
' Assumptions:  - Named range "PHASE_LIST" exists and is correctly defined.
'               - Sheet names defined in constants (DASH_SHEET, EDITS_SHEET) are correct.
'               - Column letters defined in constants (DASH_PHASE_COL, EDITS_PHASE_COL) are correct.
'               - Module_Dashboard.DebugLog exists and is accessible.
' Called By:    Manually run by developer/admin during setup.
' Location:     modUtilities Module
'----------------------------------------------------------------------------------
    Dim wsDash As Worksheet, wsEdits As Worksheet
    Dim validationFormula As String
    Const DASH_SHEET As String = "SQRCT Dashboard" ' Use actual name
    Const EDITS_SHEET As String = "UserEdits"     ' Use actual name
    Const DASH_PHASE_COL As String = "L"
    Const EDITS_PHASE_COL As String = "B"
    ' Note: Start rows are not needed here as validation is applied to whole column

    validationFormula = "=" & Module_Dashboard.PHASE_LIST_NAMED_RANGE ' Use constant from Module_Dashboard

    On Error Resume Next ' Ignore errors if sheets don't exist yet
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
Public Sub ApplyPhaseValidationToListColumn(ws As Worksheet, colLetter As String, startDataRow As Long)
'----------------------------------------------------------------------------------
' Purpose:      Re-applies the List Data Validation rule (using PHASE_LIST) to a
'               specific column on a worksheet AFTER data has been refreshed.
'               Ensures the dropdown arrow reappears on overwritten cells.
' Arguments:    ws (Worksheet): The target worksheet.
'               colLetter (String): The letter of the column to apply validation to (e.g., "L").
'               startDataRow (Long): The first row containing data in that column.
' Returns:      None.
' Assumptions:  - Named range "PHASE_LIST" exists and is correctly defined.
'               - Column A is a reliable indicator of the last data row on the sheet 'ws'.
'               - Module_Dashboard.DebugLog exists and is accessible.
' Called By:    RefreshDashboard, ApplyViewFormatting
' Location:     modUtilities Module
'----------------------------------------------------------------------------------
    If ws Is Nothing Then Exit Sub
    If Len(colLetter) = 0 Then Exit Sub
    If startDataRow < 1 Then startDataRow = 1 ' Basic sanity check

    Dim lastRow As Long
    Dim validationRange As Range
    Dim validationFormula As String
    validationFormula = "=" & Module_Dashboard.PHASE_LIST_NAMED_RANGE ' Use constant from Module_Dashboard

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
    On Error Resume Next ' Handle errors applying validation (e.g., sheet protected)
    validationRange.Validation.Delete ' Clear any existing rule first
    validationRange.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                                   Operator:=xlBetween, Formula1:=validationFormula
    If Err.Number = 0 Then ' Only set these properties if Add succeeded
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
Public Function GetDataRowCount(ws As Object) As Long ' Changed Worksheet to Object
'----------------------------------------------------------------------------------
' Purpose:      Calculates the number of actual data rows on a given worksheet.
' Arguments:    ws (Object): The worksheet object to check (passed as Object
'                           to avoid compile error with Public Function signature).
' Returns:      Long: The count of data rows found. Returns 0 if object is invalid,
'                 not a worksheet, no data exists, or last row is within header rows.
' Assumptions:  - Data starts on Row 4 on Dashboard/Active/Archive sheets.
'               - Headers occupy Rows 1 through 3 on these sheets.
'               - Column A is a reliable indicator of the last used data row.
'               - Module_Dashboard.DebugLog exists and is accessible.
' Called By:    UpdateAllViewCounts
' Location:     modUtilities Module
'----------------------------------------------------------------------------------
    Dim lastRow As Long
    Dim actualWS As Worksheet ' Variable to hold worksheet reference

    ' --- Input Validation ---
    If ws Is Nothing Then
        Module_Dashboard.DebugLog "GetDataRowCount", "ERROR: Object provided was Nothing."
        GetDataRowCount = 0
        Exit Function
    End If

    ' --- Type Check ---
    If Not TypeOf ws Is Worksheet Then
        Module_Dashboard.DebugLog "GetDataRowCount", "ERROR: Object provided is not a Worksheet. Type: " & TypeName(ws)
        GetDataRowCount = 0
        Exit Function
    End If
    ' If TypeOf check passes, it's safe to treat ws as a Worksheet
    Set actualWS = ws

    ' --- Find Last Row ---
    Const HEADER_ROWS As Long = 3 ' Number of rows before data starts
    On Error Resume Next ' Handle potential errors if sheet is empty or protected
    lastRow = actualWS.Cells(actualWS.rows.Count, "A").End(xlUp).Row
    If Err.Number <> 0 Then
        Module_Dashboard.DebugLog "GetDataRowCount", "ERROR finding last row on '" & actualWS.Name & "'. Err: " & Err.Description
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
    Module_Dashboard.DebugLog "GetDataRowCount", "Sheet '" & actualWS.Name & "' has " & GetDataRowCount & " data rows (LastRow in Col A = " & lastRow & ")."

    Set actualWS = Nothing ' Clean up object variable
End Function


'---------------------------------------------------------------------
' UpdateAllViewCounts - Reads counts from modArchival properties and
'                       writes them to the specified sheet's Row 2.
'---------------------------------------------------------------------
Public Sub UpdateAllViewCounts(ws As Worksheet)
' Purpose: Reads the stored record counts from modArchival's properties
'          and updates the display cells (J2:L2) on the provided worksheet.
' Called By: modFormatting.ExactlyCloneDashboardFormatting ' Updated 04/20/2025

    If ws Is Nothing Then
        Module_Dashboard.DebugLog "UpdateAllViewCounts", "ERROR: Worksheet object is Nothing. Cannot update counts."
        Exit Sub
    End If

    Dim totalCount As Long
    Dim activeCount As Long
    Dim archiveCount As Long
    Dim strTotal As String, strActive As String, strArchive As String

    On Error Resume Next ' Handle errors reading properties (e.g., if modArchival had compile error)
    totalCount = modArchival.TotalRecords
    activeCount = modArchival.ActiveRecords
    archiveCount = modArchival.ArchiveRecords
    If Err.Number <> 0 Then
        Module_Dashboard.DebugLog "UpdateAllViewCounts", "ERROR reading count properties from modArchival. Err: " & Err.Description
        strTotal = "Total: ERR"
        strActive = "Active: ERR"
        strArchive = "Archive: ERR"
        Err.Clear
    Else
        strTotal = "Total: " & totalCount
        strActive = "Active: " & activeCount
        strArchive = "Archive: " & archiveCount
    End If
    On Error GoTo 0 ' Restore default error handling

    ' *** ADDED (04/20/2025): Log the values READ from properties BEFORE writing ***
    Module_Dashboard.DebugLog "UpdateAllViewCounts", "Values READ for sheet '" & ws.Name & "': Total=" & totalCount & ", Active=" & activeCount & ", Archive=" & archiveCount

    ' --- Write Counts to Row 2 ---
    On Error Resume Next ' Handle errors writing to sheet (e.g., protection)
    ' Clear previous counts first
    ws.Range("J2:L2").ClearContents

    ' Populate based on the sheet type (Show all on Dashboard, relevant on others)
    Select Case ws.Name
        Case Module_Dashboard.DASHBOARD_SHEET_NAME ' Main Dashboard
            ws.Range("J2").value = strTotal
            ws.Range("K2").value = strActive
            ws.Range("L2").value = strArchive
        Case modArchival.SH_ACTIVE ' Active View
            ws.Range("J2").value = strTotal  ' Show total for context
            ws.Range("K2").value = strActive
            ' ws.Range("L2").Value = "" ' Leave Archive blank
        Case modArchival.SH_ARCHIVE ' Archive View
            ws.Range("J2").value = strTotal  ' Show total for context
            ' ws.Range("K2").Value = "" ' Leave Active blank
            ws.Range("L2").value = strArchive
        Case Else
            ' Apply to other sheets if needed, or do nothing
            Module_Dashboard.DebugLog "UpdateAllViewCounts", "Sheet '" & ws.Name & "' not recognized for specific count display. Writing all counts."
             ws.Range("J2").value = strTotal
             ws.Range("K2").value = strActive
             ws.Range("L2").value = strArchive
    End Select

    ' --- Apply Formatting to Count Cells ---
    With ws.Range("J2:L2")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = False ' Make counts normal weight
        .Font.Size = 9     ' Smaller font for counts
        .Font.Italic = True
        ' .IndentLevel = 1 ' REMOVED (04/20/2025) - Caused error 1004 on protected sheets if unlock failed.
        ' Optional: Set specific font color, e.g., .Font.Color = RGB(100, 100, 100)
    End With

    If Err.Number <> 0 Then
        Module_Dashboard.DebugLog "UpdateAllViewCounts", "ERROR writing or formatting counts on sheet '" & ws.Name & "'. Err: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0 ' Restore default error handling

    Module_Dashboard.DebugLog "UpdateAllViewCounts", "Finished updating counts display for sheet '" & ws.Name & "'."

End Sub


' --- ApplyStandardControlRow REMOVED (04/20/2025) ---
' Centralized formatting logic moved to modFormatting.ExactlyCloneDashboardFormatting.


' --- Dummy DebugLog Sub (if not already present in modUtilities) ---
' Add this simple version if your modUtilities doesn't have logging setup,
' otherwise remove this and ensure your existing DebugLog handles two string arguments.
' Assumes Module_Dashboard and its DEBUG_LOGGING constant are accessible.
' *** REMOVED DUPLICATE DEFINITION ***
Private Sub DebugLog(procedureName As String, message As String)
    If Module_Dashboard.DEBUG_LOGGING Then ' Assumes DEBUG_LOGGING constant exists in Module_Dashboard
        Debug.Print Format$(Now(), "hh:nn:ss") & " [" & procedureName & "] " & message
    End If
End Sub
