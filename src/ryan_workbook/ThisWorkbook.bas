Option Explicit

' ===============  Auto-snap Engagement Phase & Prompt for 'Other' ===============
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)

    ' --- Exit if multiple cells changed ---
    If Target.Cells.CountLarge > 1 Then Exit Sub

    ' --- Define columns and sheets to monitor ---
    ' Ensure these constants match your actual sheet names and Module_Dashboard constants
    Const DASH_SHEET As String = "SQRCT Dashboard"
    Const DASH_PHASE_COL As Long = 12 ' Column L
    Const DASH_START_ROW As Long = 4
    Const EDITS_SHEET As String = "UserEdits"
    Const EDITS_PHASE_COL As Long = 2  ' Column B
    Const EDITS_START_ROW As Long = 2

    Dim isDash As Boolean, isUE As Boolean
    isDash = (LCase$(Sh.Name) = LCase$(DASH_SHEET) And Target.Column = DASH_PHASE_COL)
    isUE = (LCase$(Sh.Name) = LCase$(EDITS_SHEET) And Target.Column = EDITS_PHASE_COL)

    ' --- Exit if change wasn't in a monitored Phase column ---
    If Not (isDash Or isUE) Then Exit Sub

    ' --- Exit if change was in header rows ---
    If isDash And Target.Row < DASH_START_ROW Then Exit Sub
    If isUE And Target.Row < EDITS_START_ROW Then Exit Sub

    ' --- Process the change ---
    Dim originalValue As Variant: originalValue = Target.Value ' Store original in case of Undo
    Application.EnableEvents = False ' Prevent this event from firing itself
    On Error GoTo SafeExit_SheetChange ' Use error handler within this sub

    Dim raw As String: raw = Trim$(CStr(Target.Value)) ' Use CStr for safety

    If Len(raw) > 0 Then ' Only process if not blank
        ' --- Find best match using helper function ---
        Dim bestMatch As String
        ' Assumes GetPhaseFromPrefix is in a standard module (e.g., modUtilities)
        bestMatch = GetPhaseFromPrefix(raw) ' Call Helper Function

        If bestMatch = "" Then
            ' No unique/exact match found in PHASE_LIST - alert user and undo
            MsgBox "“" & raw & "” isn’t a recognised or unique Engagement Phase prefix." & vbCrLf & vbCrLf & _
                   "Please choose from the dropdown list or type a more specific prefix.", _
                   vbExclamation, "Invalid Phase Entry"
            Application.Undo ' Revert the user's typing to original value
            Target.Select ' Re-select the cell

        ' Check if the successfully matched/completed phase starts with "Other ("
        ElseIf Left$(LCase$(bestMatch), 7) = "other (" Then
            ' Handling AFTER EITHER "Other (Active)" OR "Other (Archive)" is entered/completed
            ' Check if the cell value actually changed (prevents prompt if user selects same value again)
             If Target.Value <> bestMatch Then
                 Target.Value = bestMatch ' Ensure correct casing first
             End If

            MsgBox "You specified """ & bestMatch & """." & vbCrLf & vbCrLf & _
                   "Please describe the specific engagement phase or status in the " & _
                   "'User Comments' column (N)." & vbCrLf & vbCrLf & _
                   "Tip: Use clear notes for future reference and filtering.", _
                   vbInformation, "Additional Details Recommended for 'Other' Phase"

            ' Optional: Jump cursor to Comments column (N) on the same row
            On Error Resume Next ' Ignore error if selecting cell fails
            ' Ensure DB_COL_COMMENTS constant is Public in Module_Dashboard or use column number (14)
            Sh.Cells(Target.Row, Module_Dashboard.DB_COL_COMMENTS).Select
            On Error GoTo SafeExit_SheetChange ' Restore proper error handling

        ElseIf Target.Value <> bestMatch Then
             ' Unique VALID standard match found (not "Other...")
             ' AND it's different from current cell value (could be prefix or just case correction)
             ' Snap to the proper text/casing
             Target.Value = bestMatch
        End If
        ' If bestMatch is same as Target.Value, do nothing (already correct)

    End If ' Len(raw) > 0

SafeExit_SheetChange:
    If Err.Number <> 0 Then
         MsgBox "An error occurred processing the phase change: #" & Err.Number & " - " & Err.Description, vbCritical
         On Error Resume Next ' Attempt to re-enable events even if error occurred
         Application.Undo ' Try to undo if an error happened during processing
         On Error GoTo 0
    End If
    Application.EnableEvents = True ' Re-enable events

End Sub